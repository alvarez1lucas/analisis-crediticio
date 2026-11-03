"""Guía de decisión por instrumento: ON en dólar duro, ON dólar linked, préstamo bancario, prefinanciación y pesos.

`python -m empresas.guia` escribe docs/guia_instrumentos.md con las cifras de outputs/ y data/. Describe lo que hicieron
Pampa, YPF y Vista; no es una regla general.
"""

from __future__ import annotations

import pandas as pd

from empresas import memo, plantilla as p
from empresas.sql import OUTPUTS

SALIDA = p.ROOT / "docs" / "guia_instrumentos.md"


def ponderado(df: pd.DataFrame, columna: str, peso: str = "monto") -> float:
    return (df[columna] * df[peso]).sum() / df[peso].sum()


def cifras() -> dict:
    e = pd.read_csv(OUTPUTS / "emisiones.csv", parse_dates=["fecha"])
    c = pd.read_csv(p.CREDITOS, parse_dates=["inicio", "vencimiento"])
    t = pd.read_csv(p.TASAS_SISTEMA, parse_dates=["mes"])
    col = pd.read_csv(OUTPUTS / "colocadores_emision.csv")
    ent = pd.read_csv(OUTPUTS / "colocadores_entidad.csv")
    m = pd.read_csv(OUTPUTS / "modelo_instrumento.csv").set_index("medida").valor
    duro, linked = e[e.instrumento == "Dólar duro"], e[e.instrumento == "Dólar linked"]
    vista = c[(c.empresa == "Vista") & (c.tipo != "Otro")].copy()
    vista["plazo"] = (vista.vencimiento - vista.inicio).dt.days / 365.25
    fija, variable = vista[vista.tasa == "Fija"], vista[(vista.tasa == "Variable") & (vista.referencia == "SOFR")]
    anio = p.ANIOS[-1]
    t_anio = t[t.mes.dt.year == anio]
    duro_anio = duro[duro.fecha.dt.year == anio]
    locales = col[col.bancos_exterior == 0]
    bal = p.leer_balances()
    b = {e: {k: bal.get((e, anio, k)) for k in ("m_pesos", "deuda_cp", "deuda_lp")} for e in p.EMPRESAS}
    return {
        "duro": dict(n=len(duro), monto=duro.monto.sum(), plazo=ponderado(duro, "plazo"), tasa=ponderado(duro, "tasa"), rp=ponderado(duro, "riesgo_pais"),
                     rp_min=duro.riesgo_pais.min()),
        "linked": dict(n=len(linked), monto=linked.monto.sum(), plazo=ponderado(linked, "plazo"), tasa=ponderado(linked, "tasa"), rp=ponderado(linked, "riesgo_pais"),
                       rp_min=linked.riesgo_pais.min()),
        "duro_anio": dict(tasa=ponderado(duro_anio, "tasa"), plazo=ponderado(duro_anio, "plazo"), prima=ponderado(duro_anio, "prima_riesgo_pais_pb")),
        "banco_fija": dict(n=len(fija), saldo=fija.saldo.sum(), tasa=ponderado(fija, "tasa_o_margen", "saldo"), plazo=ponderado(fija.dropna(subset=["plazo"]), "plazo", "saldo")),
        "banco_sofr": dict(n=len(variable), saldo=variable.saldo.sum(), margen=ponderado(variable, "tasa_o_margen", "saldo"), plazo=ponderado(variable, "plazo", "saldo")),
        "sistema": dict(credito=t_anio.doc_usd.mean(), corto=t_anio.doc_usd_corto.mean(), largo=t_anio.doc_usd_largo.mean(), pf=t_anio.pf_usd.mean(),
                        pesos=t_anio.doc_ars.mean(), pf_pesos=t_anio.pf_ars.mean()),
        "colocacion": dict(emisiones=len(col), locales=len(locales), entidades_locales=locales.entidades.mean(),
                           con_plataforma=int((locales.plataformas_digitales > 0).sum()), siempre=ent[ent.emisiones == len(col)].entidad.tolist(),
                           exterior=col[col.bancos_exterior > 0][["empresa", "clase", "monto"]].values.tolist()),
        "modelo": m,
        "pesos": {e: (b[e]["m_pesos"] / (b[e]["deuda_cp"] + b[e]["deuda_lp"])) if b[e].get("m_pesos") is not None else None for e in p.EMPRESAS},
    }


def escribir(x: dict, d: dict) -> str:
    du, li, da, bf, bs, si, co = x["duro"], x["linked"], x["duro_anio"], x["banco_fija"], x["banco_sofr"], x["sistema"], x["colocacion"]
    pct, num, ver = memo.pct, memo.numero, memo.veces
    anio = p.ANIOS[-1]
    ext = "; ".join(f"{e} {c} (US$ {num(m, 0)} M)" for e, c, m in co["exterior"])
    return f"""# Guía de decisión: ¿ON o crédito, en dólares o en pesos?

**Para:** analistas de banca corporativa y de fintech · **Base:** {du['n'] + li['n']} ON de Pampa, YPF y Vista (2021-{anio}), sus préstamos al cierre de {anio}
y las tasas del sistema financiero. Describe lo que hicieron estas tres empresas; no es una regla para cualquier emisor.

## Qué muestran los datos

| Instrumento | Casos | Monto (US$ M) | Plazo medio | Tasa media | Riesgo país al emitir |
|---|---|---|---|---|---|
| ON en dólar duro | {du['n']} | {num(du['monto'], 0)} | {num(du['plazo'])} años | {pct(du['tasa'], 1)} | {num(du['rp'], 0)} pb |
| ON dólar linked | {li['n']} | {num(li['monto'], 0)} | {num(li['plazo'])} años | {pct(li['tasa'], 1)} | {num(li['rp'], 0)} pb |
| Préstamo bancario a tasa fija (Vista) | {bf['n']} | {num(bf['saldo'], 0)} | {num(bf['plazo'])} años | {pct(bf['tasa'], 1)} | n.a. |
| Préstamo bancario a SOFR (Vista) | {bs['n']} | {num(bs['saldo'], 0)} | {num(bs['plazo'])} años | SOFR + {pct(bs['margen'], 1)} | n.a. |
| Crédito del sistema a empresas, {anio} | promedio mensual | n.a. | hasta 89 días: {pct(si['corto'] / 100, 1)} · 90 días o más: {pct(si['largo'] / 100, 1)} | {pct(si['credito'] / 100, 1)} | n.a. |

Montos y tasas ponderados por monto. La tasa de las ON es el rendimiento de emisión cuando se informa y, si no, el cupón.
Los préstamos de YPF y de Pampa no tienen detalle por crédito: YPF informa rangos (prefinanciación de 3,50% a 8,65%; préstamos en dólares de 2,40% a 11,00%).

## Cuándo conviene cada uno

### Para la empresa

- **ON en dólar duro, cuando el riesgo país baja.** Se emitieron con riesgo país desde {num(du['rp_min'], 0)} puntos. En {anio}, las ON en dólar duro
  salieron al {pct(da['tasa'], 1)} y a {num(da['plazo'])} años en promedio, en promedio {num(-da['prima'], 0)} puntos por debajo del riesgo país del día.
- **ON dólar linked, cuando el riesgo país es alto.** Ninguna se emitió con menos de {num(li['rp_min'], 0)} puntos. Pagan casi cero en dólares
  ({pct(li['tasa'], 1)} en promedio) porque el inversor local busca cobertura cambiaria, pero son cortas ({num(li['plazo'])} años) y se pagan en pesos.
  El modelo estima {pct(x['modelo']['probabilidad_con_2000'])} de probabilidad de elegirlas con 2.000 puntos y {pct(x['modelo']['probabilidad_con_700'])} con 700.
- **Préstamo bancario, para plazos cortos o montos que no justifican una emisión.** Al cierre de {anio}, Vista debía US$ {num(bf['saldo'] + bs['saldo'], 0)} M
  a bancos, con plazos de {num(bf['plazo'])} años a tasa fija y de {num(bs['plazo'])} años contra SOFR, cuando sus ON llegaban a 8 años.
- **Prefinanciación de exportaciones, para el exportador que necesita capital de trabajo en dólares.** Está atada a cobros del exterior, así que el banco
  la ve con menos riesgo que un préstamo común. Pampa la usó y la canceló dentro de {anio}.
- **Pesos, sólo para capital de trabajo corto.** En {anio}, el crédito en pesos a empresas costó {pct(si['pesos'] / 100, 0)} en promedio, contra
  {pct(si['credito'] / 100, 1)} en dólares. Estas empresas casi no se endeudan en pesos: {pct(x['pesos']['YPF'], 1)} de la deuda de YPF y nada en Pampa y Vista.

### Para el banco

- **Colocar es recurrente:** {' y '.join(co['siempre'])} participaron en las {co['emisiones']} emisiones de la muestra. Una ON local reúne en promedio
  {num(co['entidades_locales'], 0)} colocadores: el banco que coloca no asume el riesgo de crédito, lo distribuye entre inversores y mantiene la relación con la empresa.
- **Prestar queda para el corto plazo:** con tasa activa en dólares de {pct(si['credito'] / 100, 1)} y plazo fijo en dólares de {pct(si['pf'] / 100, 1)} en {anio},
  el margen bruto fue de unos {num(si['credito'] - si['pf'], 1)} puntos. A más plazo, la empresa prefiere la ON.
- **En una ON internacional, el banco local sólo distribuye:** la emisión más grande de la muestra, {ext}, tuvo bancos de inversión del exterior
  como compradores iniciales y sólo {x['colocacion'].get('locales_ext', 4)} colocadores locales.
- **Qué mirar antes de prestar:** el muro de vencimientos y cuánto cubre la caja. Con su caja, Pampa paga lo que vence hasta
  {d['vencimientos']['Pampa']['primer_hueco'] - 1}; YPF no cubre ni {p.PROYECCION[0]}.

### Para una fintech

- **No le presta a estas empresas:** su escala y su plazo no compiten con el mercado ni con un banco.
- **Distribuye ON a minoristas:** las plataformas digitales participaron en {co['con_plataforma']} de las {co['locales']} ON locales de la muestra,
  todas de {anio}. Es el canal por el que el ahorrista llega a una ON.
- **Las usa para invertir su caja o la de sus clientes:** una ON en dólar duro de {anio} rindió {num(100 * da['tasa'] - si['pf'], 1)} puntos más que un plazo fijo en
  dólares, a cambio de plazo y de riesgo de crédito. Lo que hay que medir es lo mismo que mide un banco: rating, vencimientos y covenants.

## Límites

- Son tres empresas grandes del mismo sector. Una pyme no tiene acceso a ON internacionales y depende mucho más del banco.
- La muestra de colocadores tiene {co['emisiones']} de {du['n'] + li['n']} emisiones: las que tienen aviso de resultados o suplemento disponible.
- La tasa del sistema es un promedio de empresas no pymes a tasa fija: sirve de referencia, no es el costo de estas empresas.
"""


def main() -> None:
    x = cifras()
    col = pd.read_csv(OUTPUTS / "colocadores_emision.csv")
    x["colocacion"]["locales_ext"] = int(col[col.bancos_exterior > 0].eval("bancos + alycs + plataformas_digitales").iloc[0]) if (col.bancos_exterior > 0).any() else 0
    SALIDA.write_text(escribir(x, memo.leer()), encoding="utf-8", newline="\n")
    print(f"{SALIDA.name} escrito")


if __name__ == "__main__":
    main()
