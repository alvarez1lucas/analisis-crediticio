"""Sitio estático para GitHub Pages.

`python -m empresas.sitio` escribe site/ con la portada, los gráficos, el memo en HTML y el Excel para descargar.
Todas las cifras salen de outputs/ y del Excel calculado.
"""

from __future__ import annotations

import shutil

import markdown
import pandas as pd
import plotly.graph_objects as go

from empresas import memo, plantilla as p
from empresas.sql import OUTPUTS

SITIO = p.ROOT / "site"
COLOR_EMPRESA = p.COLOR_EMPRESA
COLOR_INSTRUMENTO = {"Dólar duro": "#2A78D6", "Dólar linked": "#E34948"}
GRIS = "#898781"
ESTILO = """
:root { --fondo: #FBFBF9; --texto: #0B0B0B; --suave: #52514E; --linea: #E1E0D9; --acento: #2A78D6; }
@media (prefers-color-scheme: dark) { :root { --fondo: #141413; --texto: #F2F1EC; --suave: #B5B3AC; --linea: #33322F; --acento: #6EA8EC; } }
* { box-sizing: border-box; }
body { margin: 0; background: var(--fondo); color: var(--texto); font: 16px/1.55 system-ui, -apple-system, "Segoe UI", sans-serif; }
main { max-width: 920px; margin: 0 auto; padding: 32px 16px 64px; }
h1 { font-size: 1.9rem; line-height: 1.2; margin: 0 0 8px; }
h2 { font-size: 1.25rem; margin: 40px 0 8px; }
p, li { color: var(--texto); }
.bajada { color: var(--suave); margin: 0 0 24px; }
nav a, a { color: var(--acento); }
nav { display: flex; flex-wrap: wrap; gap: 16px; margin: 0 0 24px; }
table { border-collapse: collapse; width: 100%; font-size: .92rem; margin: 12px 0; }
th, td { border-bottom: 1px solid var(--linea); padding: 6px 8px; text-align: right; }
th:first-child, td:first-child { text-align: left; }
.grafico { width: 100%; min-height: 360px; }
.nota { color: var(--suave); font-size: .85rem; }
.tablas { overflow-x: auto; }
"""


def numero(x: float, d: int = 1) -> str:
    return memo.numero(x, d)


def pagina(titulo: str, cuerpo: str) -> str:
    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titulo}</title><style>{ESTILO}</style></head>
<body><main>{cuerpo}</main></body></html>"""


def estilo_grafico(fig: go.Figure, titulo: str, eje_y: str) -> str:
    fig.update_layout(title=dict(text=titulo, font=dict(size=15)), template="simple_white", height=380, margin=dict(l=56, r=16, t=56, b=48),
                      legend=dict(orientation="h", y=-0.18), hovermode="closest", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="system-ui, sans-serif", color="#52514E"))
    fig.update_yaxes(title=eje_y, gridcolor="#E1E0D9", zeroline=False)
    fig.update_xaxes(showgrid=False)
    return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False, "responsive": True}, div_id=None)


def grafico_emisiones(mercado: pd.DataFrame, e: pd.DataFrame) -> str:
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=mercado.fecha, y=mercado.riesgo_pais, name="Riesgo país", mode="lines", line=dict(color=GRIS, width=1.5),
                             hovertemplate="%{x|%d-%m-%Y}: %{y:.0f} pb<extra></extra>"))
    for instrumento, color in COLOR_INSTRUMENTO.items():
        x = e[e.instrumento == instrumento]
        fig.add_trace(go.Scatter(x=x.fecha, y=x.riesgo_pais, name=f"ON en {instrumento.lower()}", mode="markers",
                                 marker=dict(color=color, size=(x.monto ** 0.5).clip(8, 30), line=dict(color="#FFFFFF", width=1.5)),
                                 customdata=x[["empresa", "clase", "monto", "tasa"]], hovertemplate="%{customdata[0]} %{customdata[1]}<br>US$ %{customdata[2]:.0f} M al %{customdata[3]:.2%}<extra></extra>"))
    return estilo_grafico(fig, "Con riesgo país alto, las empresas emitieron dólar linked", "Riesgo país (pb)")


def grafico_apalancamiento(d: dict, ind: pd.DataFrame) -> str:
    fig = go.Figure()
    real = ind[ind.indicador == "deuda_neta_ebitda"]
    for e, color in COLOR_EMPRESA.items():
        r = real[real.empresa == e]
        fig.add_trace(go.Scatter(x=r.anio, y=r.valor, name=e, mode="lines+markers", line=dict(color=color, width=2), marker=dict(size=8),
                                 hovertemplate=f"{e} %{{x}}: %{{y:.1f}}x<extra></extra>"))
        fig.add_trace(go.Scatter(x=[p.ANIOS[-1]] + p.PROYECCION, y=[r.valor.iloc[-1]] + d[e]["estres"]["apal"], name=f"{e} estrés", mode="lines+markers",
                                 line=dict(color=color, width=2, dash="dot"), marker=dict(size=7), showlegend=False,
                                 hovertemplate=f"{e} estrés %{{x}}: %{{y:.1f}}x<extra></extra>"))
    fig.add_hline(y=d["covenants"]["apal"], line=dict(color="#E34948", width=1, dash="dash"), annotation_text="Covenant supuesto", annotation_position="top left")
    fig.update_xaxes(dtick=1)
    return estilo_grafico(fig, "Deuda neta / EBITDA: real y escenario de estrés (punteado)", "Veces")


def grafico_modelo(e: pd.DataFrame, m: pd.DataFrame) -> str:
    from empresas.tendencias import ESCALA, ajustar
    modelo = ajustar(e)
    xs = list(range(500, 2701, 50))
    probs = modelo.predict(pd.DataFrame({"const": 1.0, "riesgo_pais": [x / ESCALA for x in xs]}))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=xs, y=probs, name="Probabilidad estimada", mode="lines", line=dict(color="#2A78D6", width=2),
                             hovertemplate="Riesgo país %{x}: %{y:.0%}<extra></extra>"))
    y = (e.instrumento == "Dólar linked").astype(int)
    fig.add_trace(go.Scatter(x=e.riesgo_pais, y=y, name="Cada emisión (1 = dólar linked)", mode="markers", marker=dict(color=GRIS, size=8, opacity=0.6),
                             customdata=e[["empresa", "clase"]], hovertemplate="%{customdata[0]} %{customdata[1]}<extra></extra>"))
    fig.update_yaxes(tickformat=".0%", range=[-0.05, 1.05])
    return estilo_grafico(fig, "Probabilidad de emitir dólar linked según el riesgo país", "Probabilidad")


def portada(d: dict, m: pd.DataFrame, graficos: list[str], emisiones: int) -> str:
    v = m.set_index("medida").valor
    pa = d["Pampa"]
    descubiertas = [e for e in p.EMPRESAS if min(d[e]["estres"]["dscr"]) < 1]
    negativas = [e for e in p.EMPRESAS if all(x < 0 for x in d[e]["fcf_3"])]
    sin_linked = "ninguna emitió" if v["dolar_linked_con_riesgo_pais_bajo_1500"] == 0 else f"sólo {int(v['dolar_linked_con_riesgo_pais_bajo_1500'])} emisiones fueron"
    filas = "".join(f"<tr><td>{e}</td><td>{memo.veces(d[e]['ratio']['Deuda neta / EBITDA'])}</td><td>{memo.veces(d[e]['ratio']['EBITDA / intereses'])}</td>"
                    f"<td>{memo.pct(d[e]['ratio']['ON / deuda financiera'])}</td><td>{d[e]['nota']}</td><td>{d[e]['nota_estres']}</td></tr>" for e in p.EMPRESAS)
    return f"""
<h1>¿Cuándo conviene una ON y cuándo un crédito?</h1>
<p class="bajada">Pampa Energía, YPF y Vista, 2023-2025, con sus estados contables en dólares, {emisiones} emisiones de ON y series de mercado diarias.</p>
<nav><a href="memo.html">Memo de crédito</a><a href="analisis_crediticio.xlsx">Modelo en Excel</a><a href="https://github.com/alvarez1lucas/analisis-crediticio">Código</a></nav>

<h2>Hallazgos</h2>
<ul>
<li><strong>Las tres se financian sobre todo en el mercado:</strong> entre {memo.pct(min(d[e]['ratio']['ON / deuda financiera'] for e in p.EMPRESAS))} y
{memo.pct(max(d[e]['ratio']['ON / deuda financiera'] for e in p.EMPRESAS))} de su deuda son ON. El banco aporta corto plazo y prefinanciación.</li>
<li><strong>El riesgo país decide el instrumento:</strong> {sin_linked} dólar linked con riesgo país debajo de 1.500 puntos. Con 2.000 puntos, la probabilidad estimada
es {memo.pct(v['probabilidad_con_2000'])}; con 700, {memo.pct(v['probabilidad_con_700'])}.</li>
<li><strong>El riesgo es refinanciar, no ganar plata:</strong> {', '.join(negativas[:-1]) + ' y ' + negativas[-1] if len(negativas) > 1 else ''.join(negativas)} tuvieron flujo libre negativo los tres años. Bajo estrés, {' y '.join(descubiertas)} no cubren sus vencimientos con caja propia.</li>
<li><strong>Recomendación:</strong> línea de hasta US$ {memo.numero(10 * int(memo.PARTE_DEL_BANCO * memo.capacidad(pa) // 10), 0)} M a 3 años para Pampa. Ver el <a href="memo.html">memo</a>.</li>
</ul>

<div class="tablas"><table>
<tr><th>2025</th><th>Deuda neta / EBITDA</th><th>EBITDA / intereses</th><th>ON / deuda</th><th>Rating interno</th><th>Bajo estrés</th></tr>
{filas}
</table></div>

<h2>Instrumento y momento</h2>
{graficos[0]}
<p class="nota">Cada punto es una ON, con el riesgo país del día y tamaño según el monto. Las fechas de Vista y YPF son del mes informado.</p>
{graficos[2]}
<p class="nota">Regresión logística con {emisiones} emisiones: coeficiente {memo.numero(v['coeficiente_por_1000_pb'], 2)} por cada 1.000 puntos (p = {memo.numero(v['p_valor'], 3)}).
Son tres empresas y muchas emisiones del mismo mes, así que describe el patrón y no sirve para predecir.</p>

<h2>Capacidad de pago</h2>
{graficos[1]}
<p class="nota">Estrés: ventas {memo.pct(d['supuestos']['caida'])} en 2026, margen {memo.numero(100 * d['supuestos']['margen'], 0)} pp menor y deuda nueva
{memo.numero(100 * d['supuestos']['tasa'], 0)} pp más cara. El covenant de {memo.veces(d['covenants']['apal'])} es un supuesto.</p>

<h2>Fuentes</h2>
<p class="nota">20-F 2025 de Pampa e YPF (SEC), estados contables de Vista (Bolsa Mexicana de Valores), riesgo país (ArgentinaDatos), Tesoro de EE.UU. (FRED),
tasas de crédito y BADLAR (BCRA) y plazo fijo (datos.gob.ar).</p>
"""


def main() -> None:
    SITIO.mkdir(exist_ok=True)
    d = memo.leer()
    e = pd.read_csv(OUTPUTS / "emisiones.csv", parse_dates=["fecha"])
    m = pd.read_csv(OUTPUTS / "modelo_instrumento.csv")
    ind = pd.read_csv(OUTPUTS / "indicadores.csv")
    mercado = pd.read_csv(p.MERCADO, parse_dates=["fecha"])
    graficos = [grafico_emisiones(mercado, e), grafico_apalancamiento(d, ind), grafico_modelo(e, m)]
    script = '<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js"></script>'
    (SITIO / "index.html").write_text(pagina("ON o crédito", script + portada(d, m, graficos, len(e))), encoding="utf-8")
    texto = memo.SALIDA.read_text(encoding="utf-8").replace("(../excel/analisis_crediticio.xlsx)", "(analisis_crediticio.xlsx)")
    cuerpo = '<nav><a href="index.html">Volver</a></nav>' + markdown.markdown(texto, extensions=["tables"])
    (SITIO / "memo.html").write_text(pagina("Memo de crédito", cuerpo), encoding="utf-8")
    shutil.copy(p.ARCHIVO, SITIO / p.ARCHIVO.name)
    print(f"sitio: {', '.join(x.name for x in sorted(SITIO.iterdir()))}")


if __name__ == "__main__":
    main()
