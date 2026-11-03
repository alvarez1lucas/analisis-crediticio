"""Sitio estático bilingüe para GitHub Pages: inglés en la raíz y español en /es/.

`python -m empresas.sitio` escribe site/ con la portada, el memo, la guía y el Excel para descargar.
Todas las cifras salen de outputs/, data/ y del Excel calculado.
"""

from __future__ import annotations

import re
import shutil

import markdown
import pandas as pd
import plotly.graph_objects as go

from empresas import guia, memo, plantilla as p
from empresas.sql import OUTPUTS

SITIO = p.ROOT / "site"
REPO = "https://github.com/alvarez1lucas/analisis-crediticio"
COLOR_EMPRESA = p.COLOR_EMPRESA
COLOR_INSTRUMENTO = {"Dólar duro": "#2A78D6", "Dólar linked": "#E34948"}
COLOR_TIPO = {"Banco": "#2A78D6", "ALyC": "#898781", "ALyC de un banco": "#6F6E69", "Plataforma digital": "#1BAF7A", "Banco de inversión del exterior": "#EDA100"}
GRIS = "#898781"
ESTILO = """
:root { --fondo: #FBFBF9; --superficie: #FFFFFF; --texto: #0B0B0B; --suave: #52514E; --linea: #E1E0D9; --acento: #2A78D6; }
@media (prefers-color-scheme: dark) { :root { --fondo: #141413; --superficie: #1C1C1B; --texto: #F2F1EC; --suave: #B5B3AC; --linea: #33322F; --acento: #6EA8EC; } }
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { margin: 0; background: var(--fondo); color: var(--texto); font: 16px/1.55 system-ui, -apple-system, "Segoe UI", sans-serif; }
main { max-width: 980px; margin: 0 auto; padding: 24px 16px 64px; }
header .arriba { display: flex; justify-content: space-between; gap: 12px; align-items: center; font-size: .9rem; color: var(--suave); }
h1 { font-size: 2rem; line-height: 1.15; margin: 16px 0 8px; }
h2 { font-size: 1.35rem; margin: 48px 0 6px; padding-top: 8px; border-top: 1px solid var(--linea); }
h3 { font-size: 1.05rem; margin: 24px 0 6px; }
.bajada { color: var(--suave); margin: 0 0 16px; max-width: 760px; }
a { color: var(--acento); }
nav.secciones { display: flex; flex-wrap: wrap; gap: 6px 16px; margin: 12px 0 24px; font-size: .92rem; }
.tarjetas { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 12px; margin: 16px 0 8px; }
.tarjeta { background: var(--superficie); border: 1px solid var(--linea); border-radius: 10px; padding: 14px 16px; }
.tarjeta .valor { font-size: 1.7rem; font-weight: 600; line-height: 1.1; }
.tarjeta .rotulo { color: var(--suave); font-size: .88rem; margin-top: 4px; }
.miradas { display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px; }
.mirada { background: var(--superficie); border: 1px solid var(--linea); border-radius: 10px; padding: 12px 16px; }
.mirada h3 { margin-top: 4px; }
table { border-collapse: collapse; width: 100%; font-size: .9rem; margin: 12px 0; }
th, td { border-bottom: 1px solid var(--linea); padding: 6px 8px; text-align: right; white-space: nowrap; }
th:first-child, td:first-child { text-align: left; }
.tablas { overflow-x: auto; }
.nota { color: var(--suave); font-size: .85rem; }
ol.pasos li { margin-bottom: 6px; }
"""

T = {
    "en": dict(
        otro="es/index.html", otro_rotulo="Español", titulo="When does a bond beat a bank loan?",
        bajada="Credit analysis of Pampa Energía, YPF and Vista (2023-2025): audited financial statements in US dollars, {n} corporate bonds and daily market data. "
               "Built as a corporate-banking credit file: balance sheets, debt, funding choices, stress and a credit memo.",
        nav=[("hallazgos", "Findings"), ("balances", "Balance sheets"), ("deuda", "Debt and cash"), ("financiamiento", "Funding"), ("bancos", "Banks and fintech"),
             ("estres", "Stress"), ("descargas", "Downloads"), ("metodo", "Method")],
        tarjetas=["corporate bonds analysed", "issued 2021-2025", "company-years of audited statements", "recommended facility for Pampa"],
        memo="Credit memo", guia="Decision guide", excel="Excel model", pbix="Power BI dashboard", codigo="Code", datos="Data (CSV)",
    ),
    "es": dict(
        otro="../index.html", otro_rotulo="English", titulo="¿Cuándo conviene una ON y cuándo un crédito?",
        bajada="Análisis crediticio de Pampa Energía, YPF y Vista (2023-2025): estados contables auditados en dólares, {n} obligaciones negociables y datos de mercado diarios. "
               "Armado como una carpeta de crédito de banca corporativa: balances, deuda, forma de financiarse, estrés y memo de crédito.",
        nav=[("hallazgos", "Hallazgos"), ("balances", "Balances"), ("deuda", "Deuda y caja"), ("financiamiento", "Financiamiento"), ("bancos", "Bancos y fintech"),
             ("estres", "Estrés"), ("descargas", "Descargas"), ("metodo", "Método")],
        tarjetas=["obligaciones negociables analizadas", "emitidos 2021-2025", "años-empresa de estados auditados", "línea recomendada para Pampa"],
        memo="Memo de crédito", guia="Guía de decisión", excel="Modelo en Excel", pbix="Tablero de Power BI", codigo="Código", datos="Datos (CSV)",
    ),
}
INDICADORES = {"deuda_neta_ebitda": ("Net debt / EBITDA", "Deuda neta / EBITDA", "x"), "ebitda_intereses": ("EBITDA / interest", "EBITDA / intereses", "x"),
               "margen_ebitda": ("EBITDA margin", "Margen EBITDA", "%"), "liquidez_corriente": ("Current ratio", "Liquidez corriente", "x"),
               "flujo_libre_deuda": ("Free cash flow / debt", "Flujo libre / deuda", "%"), "capex_ventas": ("Capex / sales", "Capex / ventas", "%"),
               "roe": ("Return on equity", "Rentabilidad sobre patrimonio", "%"), "on_deuda": ("Bonds / financial debt", "ON / deuda financiera", "%")}
TIPOS_EN = {"Banco": "Bank", "ALyC": "Broker-dealer", "ALyC de un banco": "Bank-owned broker", "Plataforma digital": "Digital platform",
            "Banco de inversión del exterior": "Foreign investment bank"}
INSTRUMENTO_EN = {"Dólar duro": "hard-dollar", "Dólar linked": "dollar-linked"}
CONCEPTO_EN = {"EBITDA": "EBITDA", "Impuestos y capital de trabajo": "Taxes and working capital", "Intereses pagados": "Interest paid", "Capex": "Capex"}


def fmt(L: str):
    """Formateadores de números según el idioma."""
    if L == "en":
        return memo.number, memo.percent, memo.times
    return memo.numero, memo.pct, memo.veces


def es(L: str, ingles, espanol):
    return ingles if L == "en" else espanol


def pagina(L: str, titulo: str, cuerpo: str, scripts: str = "") -> str:
    return f"""<!doctype html>
<html lang="{L}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{titulo}</title><style>{ESTILO}</style>{scripts}</head>
<body><main>{cuerpo}</main></body></html>"""


def estilo(fig: go.Figure, titulo: str, eje_y: str, alto: int = 400) -> str:
    fig.update_layout(title=dict(text=titulo, font=dict(size=15)), template="simple_white", height=alto, margin=dict(l=60, r=16, t=60, b=50),
                      legend=dict(orientation="h", y=-0.17), hovermode="closest", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(family="system-ui, sans-serif", color="#52514E"))
    fig.update_yaxes(title=eje_y, gridcolor="#E1E0D9", zeroline=False)
    fig.update_xaxes(showgrid=False)
    return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False, "responsive": True})


def grafico_indicadores(ind: pd.DataFrame, L: str) -> str:
    fig = go.Figure()
    claves = list(INDICADORES)
    for k, clave in enumerate(claves):
        for e, color in COLOR_EMPRESA.items():
            x = ind[(ind.indicador == clave) & (ind.empresa == e)]
            formato = ".1%" if INDICADORES[clave][2] == "%" else ".2f"
            fig.add_trace(go.Scatter(x=x.anio, y=x.valor, name=e, mode="lines+markers", line=dict(color=color, width=2), marker=dict(size=8),
                                     visible=k == 0, legendgroup=e, hovertemplate=f"{e} %{{x}}: %{{y:{formato}}}<extra></extra>"))
    botones = []
    for k, clave in enumerate(claves):
        en, sp, unidad = INDICADORES[clave]
        visible = [j // len(COLOR_EMPRESA) == k for j in range(len(claves) * len(COLOR_EMPRESA))]
        botones.append(dict(label=es(L, en, sp), method="update",
                            args=[{"visible": visible}, {"yaxis.tickformat": ".0%" if unidad == "%" else ".1f", "title.text": es(L, en, sp)}]))
    fig.update_layout(updatemenus=[dict(buttons=botones, x=1, xanchor="right", y=1.18, yanchor="top", showactive=True)])
    fig.update_xaxes(dtick=1)
    en, sp, _ = INDICADORES[claves[0]]
    return estilo(fig, es(L, en, sp), "")


def grafico_vencimientos(v: pd.DataFrame, L: str) -> str:
    total = v.groupby("empresa").monto.sum()
    x = v.copy()
    x["tramo"] = x.anio.where(x.anio < 2036, 2036)
    x = x.groupby(["empresa", "tramo"], as_index=False).monto.sum()
    x["parte"] = x.monto / x.empresa.map(total)
    x["rotulo"] = x.tramo.astype(str).replace("2036", es(L, "2036+", "2036 y más"))
    fig = go.Figure()
    for e, color in COLOR_EMPRESA.items():
        y = x[x.empresa == e]
        fig.add_trace(go.Bar(x=y.rotulo, y=y.parte, name=e, marker=dict(color=color), customdata=y.monto,
                             hovertemplate=f"{e} %{{x}}: %{{y:.0%}} (US$ %{{customdata:,.0f}} M)<extra></extra>"))
    fig.update_layout(barmode="group", bargap=0.25)
    fig.update_yaxes(tickformat=".0%")
    orden = x.drop_duplicates("tramo").sort_values("tramo").rotulo.tolist()
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=orden)
    return estilo(fig, es(L, "Share of end-2025 debt maturing each year", "Qué parte de la deuda de fin de 2025 vence cada año"), es(L, "Share of debt", "Parte de la deuda"))


def grafico_puente(pu: pd.DataFrame, L: str) -> str:
    anio = p.ANIOS[-1]
    fig = go.Figure()
    for k, e in enumerate(p.EMPRESAS):
        x = pu[(pu.empresa == e) & (pu.anio == anio)].sort_values("orden")
        etiquetas = [es(L, CONCEPTO_EN[c], c) for c in x.concepto] + [es(L, "Free cash flow", "Flujo libre")]
        fig.add_trace(go.Waterfall(x=etiquetas, y=list(x.monto) + [0], measure=["absolute"] + ["relative"] * (len(x) - 1) + ["total"], name=e, visible=k == 0,
                                   increasing=dict(marker=dict(color="#1BAF7A")), decreasing=dict(marker=dict(color="#E34948")), totals=dict(marker=dict(color="#2A78D6")),
                                   connector=dict(line=dict(color="#E1E0D9")), hovertemplate="%{x}: US$ %{y:,.0f} M<extra></extra>"))
    botones = [dict(label=e, method="update", args=[{"visible": [j == k for j in range(len(p.EMPRESAS))]}]) for k, e in enumerate(p.EMPRESAS)]
    fig.update_layout(updatemenus=[dict(buttons=botones, x=1, xanchor="right", y=1.18, yanchor="top")], showlegend=False)
    return estilo(fig, es(L, f"From EBITDA to free cash flow, {anio} (US$ M)", f"Del EBITDA al flujo libre, {anio} (US$ M)"), "US$ M")


def grafico_emisiones(mercado: pd.DataFrame, e: pd.DataFrame, L: str) -> str:
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=mercado.fecha, y=mercado.riesgo_pais, name=es(L, "Country risk", "Riesgo país"), mode="lines", line=dict(color=GRIS, width=1.5),
                             hovertemplate="%{x|%Y-%m-%d}: %{y:.0f} bp<extra></extra>"))
    for instrumento, color in COLOR_INSTRUMENTO.items():
        x = e[e.instrumento == instrumento]
        nombre = es(L, f"{INSTRUMENTO_EN[instrumento].capitalize()} bond", f"ON en {instrumento.lower()}")
        fig.add_trace(go.Scatter(x=x.fecha, y=x.riesgo_pais, name=nombre, mode="markers",
                                 marker=dict(color=color, size=(x.monto ** 0.5).clip(8, 30), line=dict(color="#FFFFFF", width=1.5)),
                                 customdata=x[["empresa", "clase", "monto", "tasa"]],
                                 hovertemplate="%{customdata[0]} %{customdata[1]}<br>US$ %{customdata[2]:.0f} M, %{customdata[3]:.2%}<extra></extra>"))
    return estilo(fig, es(L, "With high country risk, companies issued dollar-linked bonds", "Con riesgo país alto, las empresas emitieron dólar linked"),
                  es(L, "Country risk (bp)", "Riesgo país (pb)"))


def grafico_modelo(e: pd.DataFrame, L: str) -> str:
    from empresas.tendencias import ESCALA, ajustar
    modelo = ajustar(e)
    xs = list(range(500, 2701, 50))
    probs = modelo.predict(pd.DataFrame({"const": 1.0, "riesgo_pais": [x / ESCALA for x in xs]}))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=xs, y=probs, name=es(L, "Estimated probability", "Probabilidad estimada"), mode="lines", line=dict(color="#2A78D6", width=2),
                             hovertemplate="%{x}: %{y:.0%}<extra></extra>"))
    fig.add_trace(go.Scatter(x=e.riesgo_pais, y=(e.instrumento == "Dólar linked").astype(int), name=es(L, "Each bond (1 = dollar-linked)", "Cada ON (1 = dólar linked)"),
                             mode="markers", marker=dict(color=GRIS, size=8, opacity=0.6), customdata=e[["empresa", "clase"]],
                             hovertemplate="%{customdata[0]} %{customdata[1]}<extra></extra>"))
    fig.update_yaxes(tickformat=".0%", range=[-0.05, 1.05])
    fig.update_xaxes(title=es(L, "Country risk on the issue date (bp)", "Riesgo país el día de emisión (pb)"))
    return estilo(fig, es(L, "Probability of issuing dollar-linked, by country risk", "Probabilidad de emitir dólar linked según el riesgo país"),
                  es(L, "Probability", "Probabilidad"))


def grafico_colocadores(ent: pd.DataFrame, L: str) -> str:
    x = ent.head(15).iloc[::-1]
    fig = go.Figure()
    for tipo, color in COLOR_TIPO.items():
        y = x[x.tipo_entidad == tipo]
        if y.empty:
            continue
        fig.add_trace(go.Bar(y=y.entidad, x=y.emisiones, orientation="h", name=es(L, TIPOS_EN[tipo], tipo), marker=dict(color=color),
                             hovertemplate="%{y}: %{x}<extra></extra>"))
    fig.update_layout(barmode="stack")
    fig.update_yaxes(categoryorder="array", categoryarray=list(x.entidad))
    fig.update_xaxes(dtick=1, title=es(L, "Placements", "Emisiones"))
    return estilo(fig, es(L, "Who places the bonds: participations by entity", "Quién coloca las ON: participaciones por entidad"), "", 460)


def grafico_estres(d: dict, ind: pd.DataFrame, L: str) -> str:
    fig = go.Figure()
    real = ind[ind.indicador == "deuda_neta_ebitda"]
    for e, color in COLOR_EMPRESA.items():
        r = real[real.empresa == e]
        fig.add_trace(go.Scatter(x=r.anio, y=r.valor, name=e, mode="lines+markers", line=dict(color=color, width=2), marker=dict(size=8),
                                 hovertemplate=f"{e} %{{x}}: %{{y:.1f}}x<extra></extra>"))
        fig.add_trace(go.Scatter(x=[p.ANIOS[-1]] + p.PROYECCION, y=[r.valor.iloc[-1]] + d[e]["estres"]["apal"], name=f"{e} stress", mode="lines+markers",
                                 line=dict(color=color, width=2, dash="dot"), marker=dict(size=7), showlegend=False,
                                 hovertemplate=f"{e} " + es(L, "stress", "estrés") + " %{x}: %{y:.1f}x<extra></extra>"))
    fig.add_hline(y=d["covenants"]["apal"], line=dict(color="#E34948", width=1, dash="dash"),
                  annotation_text=es(L, "Assumed covenant", "Covenant supuesto"), annotation_position="top left")
    fig.update_xaxes(dtick=1)
    return estilo(fig, es(L, "Net debt / EBITDA: actual and stress scenario (dotted)", "Deuda neta / EBITDA: real y escenario de estrés (punteado)"), es(L, "Times", "Veces"))


def tabla(filas: list[list[str]], encabezado: list[str]) -> str:
    cabeza = "".join(f"<th>{h}</th>" for h in encabezado)
    cuerpo = "".join("<tr>" + "".join(f"<td>{c}</td>" for c in f) + "</tr>" for f in filas)
    return f'<div class="tablas"><table><tr>{cabeza}</tr>{cuerpo}</table></div>'


def portada(L: str, base: dict) -> str:
    t = T[L]
    n, pc, x_ = fmt(L)
    d, g, e, m, ind, ten = base["d"], base["g"], base["e"], base["m"], base["ind"], base["ten"]
    graf = base["graficos"][L]
    pa = d["Pampa"]
    linea = 10 * int(memo.PARTE_DEL_BANCO * memo.capacidad(pa) // 10)
    co, si, li = g["colocacion"], g["sistema"], g["linked"]
    pref = "" if L == "en" else "../"
    descubiertas = [x for x in p.EMPRESAS if min(d[x]["estres"]["dscr"]) < 1]
    ve = {x: d["vencimientos"][x] for x in p.EMPRESAS}
    y_ = es(L, " and ", " y ")
    on_min = min(d[x]["ratio"]["ON / deuda financiera"] for x in p.EMPRESAS)
    on_max = max(d[x]["ratio"]["ON / deuda financiera"] for x in p.EMPRESAS)
    hueco = ve["Pampa"]["primer_hueco"] - 1
    ypf_v, ypf_c = ve["YPF"]["primer_anio"], ve["YPF"]["caja"]

    tarjetas = [f"{len(e)}", f"US$ {n(e.monto.sum() / 1000, 1)} {es(L, 'bn', 'mil M')}", f"{len(p.EMPRESAS) * len(p.ANIOS)}", f"US$ {n(linea, 0)} M"]
    tarjetas_html = "".join(f'<div class="tarjeta"><div class="valor">{v}</div><div class="rotulo">{r}</div></div>' for v, r in zip(tarjetas, t["tarjetas"]))
    nav = "".join(f'<a href="#{a}">{r}</a>' for a, r in t["nav"])

    hallazgos = es(L,
                   [f"<strong>All three fund themselves in the market:</strong> between {pc(on_min)} and {pc(on_max)} of their debt is bonds. "
                    "Banks provide short tenors, pre-financing and placement.",
                    f"<strong>Country risk drives the instrument:</strong> no dollar-linked bond was issued below {n(li['rp_min'], 0)} points. "
                    f"At 2,000 points the estimated probability is {pc(m['probabilidad_con_2000'])}; at 700, {pc(m['probabilidad_con_700'])}.",
                    f"<strong>The risk is refinancing, not profitability:</strong> all three had negative free cash flow every year. With its cash, Pampa covers maturities "
                    f"until {hueco}; YPF faces US$ {n(ypf_v, 0)} M in {p.PROYECCION[0]}, {x_(ypf_v / ypf_c)} its cash.",
                    f"<strong>Recommendation:</strong> a facility of up to US$ {n(linea, 0)} M, 3 years, for Pampa (internal rating {pa['nota']}). "
                    "See the <a href=\"memo.html\">credit memo</a>."],
                   [f"<strong>Las tres se financian en el mercado:</strong> entre {pc(on_min)} y {pc(on_max)} de su deuda son ON. "
                    "El banco aporta plazos cortos, prefinanciación y colocación.",
                    f"<strong>El riesgo país decide el instrumento:</strong> ninguna ON dólar linked se emitió con menos de {n(li['rp_min'], 0)} puntos. "
                    f"Con 2.000 puntos la probabilidad estimada es {pc(m['probabilidad_con_2000'])}; con 700, {pc(m['probabilidad_con_700'])}.",
                    f"<strong>El riesgo es refinanciar, no ganar plata:</strong> las tres tuvieron flujo libre negativo todos los años. Con su caja, Pampa cubre lo que vence "
                    f"hasta {hueco}; a YPF le vencen US$ {n(ypf_v, 0)} M en {p.PROYECCION[0]}, {x_(ypf_v / ypf_c)} su caja.",
                    f"<strong>Recomendación:</strong> línea de hasta US$ {n(linea, 0)} M a 3 años para Pampa (rating interno {pa['nota']}). "
                    "Ver el <a href=\"memo.html\">memo de crédito</a>."])

    ratios = [("Deuda neta / EBITDA", x_, "Net debt / EBITDA"), ("EBITDA / intereses", x_, "EBITDA / interest"), ("Margen EBITDA", pc, "EBITDA margin"),
              ("Liquidez corriente", x_, "Current ratio"), ("ON / deuda financiera", pc, "Bonds / debt"), ("Altman Z'' mercados emergentes", x_, "Altman Z''")]
    filas_ratios = [[x] + [f(d[x]["ratio"][r]) for r, f, _ in ratios] + [d[x]["nota"]] for x in p.EMPRESAS]
    enc_ratios = [str(p.ANIOS[-1])] + [es(L, en, r) for r, _, en in ratios] + [es(L, "Internal rating", "Rating interno")]

    dup = ind[ind.anio == p.ANIOS[-1]].pivot(index="empresa", columns="indicador", values="valor").reindex(p.EMPRESAS)
    filas_dupont = [[x, pc(dup.loc[x, "roe"], 1), pc(dup.loc[x, "margen_neto"], 1), n(dup.loc[x, "rotacion_activos"], 2), x_(dup.loc[x, "activo_patrimonio"])] for x in p.EMPRESAS]
    enc_dupont = [str(p.ANIOS[-1]), es(L, "Return on equity", "Rentabilidad sobre patrimonio"), es(L, "Net margin", "Margen neto"),
                  es(L, "Asset turnover", "Rotación de activos"), es(L, "Assets / equity", "Activo / patrimonio")]

    filas_venc = [[x, f"US$ {n(ve[x]['caja'], 0)} M", f"US$ {n(ve[x]['primer_anio'], 0)} M", x_(ve[x]["caja"] / ve[x]["primer_anio"]),
                   str((ve[x]["primer_hueco"] or max(ve[x]["por_anio"]) + 1) - p.PROYECCION[0])] for x in p.EMPRESAS]
    enc_venc = [es(L, "End-2025", "Cierre 2025"), es(L, "Cash and investments", "Caja e inversiones"), es(L, f"Maturing in {p.PROYECCION[0]}", f"Vence en {p.PROYECCION[0]}"),
                es(L, "Cash / maturities", "Caja / vencimientos"), es(L, "Years of maturities covered by cash", "Años de vencimientos que cubre la caja")]

    filas_ten = [[r.empresa, str(r.anio), str(r.emisiones), f"US$ {n(r.monto, 0)} M", pc(0 if pd.isna(r.parte_dolar_linked) else r.parte_dolar_linked),
                  n(r.plazo_medio, 1), pc(r.tasa_media, 1), n(r.riesgo_pais_medio, 0)] for r in ten.itertuples()]
    enc_ten = [es(L, "Company", "Empresa"), es(L, "Year", "Año"), es(L, "Bonds", "ON"), es(L, "Amount", "Monto"), es(L, "Dollar-linked", "Dólar linked"),
               es(L, "Tenor (years)", "Plazo (años)"), es(L, "Rate", "Tasa"), es(L, "Country risk", "Riesgo país")]

    filas_estres = [[x, x_(max(d[x]["estres"]["apal"])), x_(min(d[x]["estres"]["cob"])), x_(min(d[x]["estres"]["dscr"]), 2), f"{d[x]['nota']} → {d[x]['nota_estres']}"]
                    for x in p.EMPRESAS]
    enc_estres = [es(L, "Worst year, stress", "Peor año, estrés"), es(L, "Net debt / EBITDA", "Deuda neta / EBITDA"), es(L, "EBITDA / interest", "EBITDA / intereses"),
                  es(L, "Debt service coverage", "Cobertura de servicio"), es(L, "Internal rating", "Rating interno")]
    s = d["supuestos"]
    margen_banco = n(si["credito"] - si["pf"], 1)
    siempre = y_.join(co["siempre"])

    miradas = es(L,
                 [("Company", f"Hard-dollar bonds when country risk falls; dollar-linked only above {n(li['rp_min'], 0)} points; banks for short tenors and export pre-financing."),
                  ("Bank", f"Places bonds without taking the credit risk: {siempre} were in every placement. Lending in dollars earned about {margen_banco} points over "
                           f"dollar time deposits in {p.ANIOS[-1]}."),
                  ("Fintech", f"Does not lend to these companies; it distributes their bonds. Digital platforms were in {co['con_plataforma']} of the {co['locales']} local bonds in the sample.")],
                 [("Empresa", f"ON en dólar duro cuando baja el riesgo país; dólar linked sólo arriba de {n(li['rp_min'], 0)} puntos; el banco para plazos cortos y prefinanciación."),
                  ("Banco", f"Coloca sin asumir el riesgo: {siempre} estuvieron en todas. Prestando en dólares ganó unos {margen_banco} puntos sobre el plazo fijo "
                            f"en dólares en {p.ANIOS[-1]}."),
                  ("Fintech", f"No les presta: distribuye sus ON. Las plataformas digitales estuvieron en {co['con_plataforma']} de las {co['locales']} ON locales de la muestra.")])
    miradas_html = "".join(f'<div class="mirada"><h3>{a}</h3><p>{b}</p></div>' for a, b in miradas)

    pasos = es(L,
               ["<strong>Data:</strong> financial statements from 20-F and Mexican exchange filings, debt notes, results notices, country risk (ArgentinaDatos), "
                "US Treasury yields (FRED), and BCRA and datos.gob.ar rates.",
                "<strong>Excel:</strong> one sheet per company with checks against reported totals, ratios, vertical analysis, sources and uses, maturities, "
                "rating, projection and covenants.",
                "<strong>SQL (DuckDB):</strong> PIVOT and UNPIVOT for the statements, ASOF JOIN to match each bond with that day's market, maturity wall, DuPont and placement agents.",
                "<strong>Python:</strong> logistic regression on instrument choice; the memo, the guide and this site are written from the model.",
                "<strong>Power BI:</strong> six-page dashboard generated as a text project (TMDL and PBIR) from the same outputs.",
                f"<strong>Controls:</strong> {base['tests']} tests; SQL, Excel and Power BI return the same figures. GitHub Actions refreshes market data every Monday."],
               ["<strong>Datos:</strong> estados contables de los 20-F y de la bolsa mexicana, notas de deuda, avisos de resultados, riesgo país (ArgentinaDatos), "
                "Tesoro de EE.UU. (FRED) y tasas del BCRA y de datos.gob.ar.",
                "<strong>Excel:</strong> una hoja por empresa con controles de cuadre contra lo informado, ratios, análisis vertical, fuentes y usos, vencimientos, "
                "rating, proyección y covenants.",
                "<strong>SQL (DuckDB):</strong> PIVOT y UNPIVOT de los estados, ASOF JOIN para cruzar cada ON con el mercado del día, muro de vencimientos, DuPont y colocadores.",
                "<strong>Python:</strong> regresión logística sobre la elección de instrumento; el memo, la guía y este sitio se escriben desde el modelo.",
                "<strong>Power BI:</strong> tablero de seis páginas generado como proyecto de texto (TMDL y PBIR) desde los mismos datos.",
                f"<strong>Controles:</strong> {base['tests']} tests; SQL, Excel y Power BI dan las mismas cifras. GitHub Actions actualiza el mercado cada lunes."])

    descargas = [(t["memo"], "memo.html"), (t["guia"], "guia.html"), (t["excel"], f"{pref}analisis_crediticio.xlsx"),
                 (t["pbix"], f"{REPO}/raw/main/powerbi/analisis_crediticio.pbix"), (t["datos"], f"{REPO}/tree/main/outputs"), (t["codigo"], REPO)]
    nota_modelo = es(L, f"Logistic regression on {len(e)} bonds: coefficient {n(m['coeficiente_por_1000_pb'], 2)} per 1,000 points (p = {n(m['p_valor'], 3)}). "
                        "Three companies and many same-month issues: it describes the pattern and is not a forecast.",
                     f"Regresión logística con {len(e)} ON: coeficiente {n(m['coeficiente_por_1000_pb'], 2)} por cada 1.000 puntos (p = {n(m['p_valor'], 3)}). "
                     "Son tres empresas y muchas emisiones del mismo mes: describe el patrón, no predice.")
    nota_bancos = es(L, f"Sample: {co['emisiones']} of {len(e)} bonds with a results notice or prospectus supplement available. See the <a href='guia.html'>decision guide</a>.",
                     f"Muestra: {co['emisiones']} de {len(e)} ON con aviso de resultados o suplemento disponible. Ver la <a href='guia.html'>guía de decisión</a>.")
    bajada_estres = es(L, f"Stress: sales {pc(s['caida'])} in {p.PROYECCION[0]}, margin {n(100 * s['margen'], 0)} pp lower, new debt {n(100 * s['tasa'], 0)} pp more expensive, "
                          "capex equal to depreciation and the actual maturity schedule.",
                       f"Estrés: ventas {pc(s['caida'])} en {p.PROYECCION[0]}, margen {n(100 * s['margen'], 0)} pp menor, deuda nueva {n(100 * s['tasa'], 0)} pp más cara, "
                       "capex igual a la depreciación y los vencimientos reales.")
    nota_estres = es(L, f"Below 1x debt service coverage under stress: {y_.join(descubiertas) or 'none'}. The covenant is an assumption: filings do not publish contract terms.",
                     f"Debajo de 1x de cobertura de servicio bajo estrés: {y_.join(descubiertas) or 'ninguna'}. El covenant es un supuesto: los estados no publican los contratos.")
    return f"""
<header><div class="arriba"><span>{es(L, 'Portfolio project: credit analysis', 'Proyecto de portfolio: análisis crediticio')}</span><a href="{t['otro']}">{t['otro_rotulo']}</a></div>
<h1>{t['titulo']}</h1><p class="bajada">{t['bajada'].format(n=len(e))}</p>
<nav class="secciones">{nav}</nav></header>
<div class="tarjetas">{tarjetas_html}</div>

<h2 id="hallazgos">{t['nav'][0][1]}</h2>
<ul>{''.join(f'<li>{h}</li>' for h in hallazgos)}</ul>
{tabla(filas_ratios, enc_ratios)}

<h2 id="balances">{t['nav'][1][1]}</h2>
<p class="bajada">{es(L, 'Same definitions for all three: EBITDA excludes associates and one-off results. Choose the metric in the menu.',
                         'Misma definición para las tres: el EBITDA excluye asociadas y resultados no recurrentes. Elegí el indicador en el menú.')}</p>
{graf['indicadores']}
<h3>DuPont</h3>
<p class="bajada">{es(L, "Vista's return comes from margin and leverage; Pampa earns less with much less debt; YPF's tax charge turned 2025 into a loss.",
                      'La rentabilidad de Vista viene de margen y deuda; Pampa rinde menos con mucho menos apalancamiento; a YPF el cargo por impuestos le dio pérdida en 2025.')}</p>
{tabla(filas_dupont, enc_dupont)}

<h2 id="deuda">{t['nav'][2][1]}</h2>
{graf['vencimientos']}
{tabla(filas_venc, enc_venc)}
{graf['puente']}
<p class="nota">{es(L, 'Maturities: debt notes in each filing. YPF reports only the maturity year of each bond and groups its loans by range of years.',
                     'Vencimientos: notas de deuda de cada estado. YPF informa sólo el año de vencimiento de cada ON y agrupa sus préstamos por rango de años.')}</p>

<h2 id="financiamiento">{t['nav'][3][1]}</h2>
{graf['emisiones']}
{graf['modelo']}
<p class="nota">{nota_modelo}</p>
{tabla(filas_ten, enc_ten)}

<h2 id="bancos">{t['nav'][4][1]}</h2>
<div class="miradas">{miradas_html}</div>
{graf['colocadores']}
<p class="nota">{nota_bancos}</p>

<h2 id="estres">{t['nav'][5][1]}</h2>
<p class="bajada">{bajada_estres}</p>
{graf['estres']}
{tabla(filas_estres, enc_estres)}
<p class="nota">{nota_estres}</p>

<h2 id="descargas">{t['nav'][6][1]}</h2>
<ul>{''.join(f'<li><a href="{u}">{r}</a></li>' for r, u in descargas)}</ul>

<h2 id="metodo">{t['nav'][7][1]}</h2>
<ol class="pasos">{''.join(f'<li>{x}</li>' for x in pasos)}</ol>
"""


def documento(L: str, md: str) -> str:
    excel = "analisis_crediticio.xlsx" if L == "en" else "../analisis_crediticio.xlsx"
    texto = md.replace("(../excel/analisis_crediticio.xlsx)", f"({excel})")
    return f'<nav class="secciones"><a href="index.html">{es(L, "Back", "Volver")}</a></nav>' + markdown.markdown(texto, extensions=["tables"])


def main() -> None:
    shutil.rmtree(SITIO, ignore_errors=True)
    (SITIO / "es").mkdir(parents=True)
    d = memo.leer()
    e = pd.read_csv(OUTPUTS / "emisiones.csv", parse_dates=["fecha"])
    base = dict(d=d, g=guia.cifras(), e=e, m=pd.read_csv(OUTPUTS / "modelo_instrumento.csv").set_index("medida").valor,
                ind=pd.read_csv(OUTPUTS / "indicadores.csv"), ten=pd.read_csv(OUTPUTS / "tendencias.csv"),
                tests=sum(len(re.findall(r"^def test_", f.read_text(encoding="utf-8"), re.M)) for f in (p.ROOT / "tests").glob("test_*.py")))
    mercado = pd.read_csv(p.MERCADO, parse_dates=["fecha"])
    v, pu, ent = pd.read_csv(OUTPUTS / "vencimientos.csv"), pd.read_csv(OUTPUTS / "puente.csv"), pd.read_csv(OUTPUTS / "colocadores_entidad.csv")
    base["graficos"] = {L: dict(indicadores=grafico_indicadores(base["ind"], L), vencimientos=grafico_vencimientos(v, L), puente=grafico_puente(pu, L),
                                emisiones=grafico_emisiones(mercado, e, L), modelo=grafico_modelo(e, L), colocadores=grafico_colocadores(ent, L),
                                estres=grafico_estres(d, base["ind"], L)) for L in ("en", "es")}
    script = '<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js"></script>'
    for L, carpeta in (("en", SITIO), ("es", SITIO / "es")):
        (carpeta / "index.html").write_text(pagina(L, T[L]["titulo"], portada(L, base), script), encoding="utf-8")
        memo_md = (memo.SALIDA_EN if L == "en" else memo.SALIDA).read_text(encoding="utf-8")
        guia_md = (guia.SALIDA_EN if L == "en" else guia.SALIDA).read_text(encoding="utf-8")
        (carpeta / "memo.html").write_text(pagina(L, T[L]["memo"], documento(L, memo_md)), encoding="utf-8")
        (carpeta / "guia.html").write_text(pagina(L, T[L]["guia"], documento(L, guia_md)), encoding="utf-8")
    shutil.copy(p.ARCHIVO, SITIO / p.ARCHIVO.name)
    print(f"sitio: {sorted(str(x.relative_to(SITIO)) for x in SITIO.rglob('*.*'))}")


if __name__ == "__main__":
    main()
