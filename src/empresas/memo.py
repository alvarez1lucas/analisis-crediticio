"""Memo de crédito de Pampa, con YPF y Vista como comparables.

`python -m empresas.memo` escribe docs/memo_credito.md con las cifras de excel/analisis_crediticio.xlsx. Ninguna cifra se
escribe a mano: si cambia un dato o un supuesto, se regenera el Excel y después el memo.
"""

from __future__ import annotations

import openpyxl

from empresas import plantilla as p

SALIDA = p.ROOT / "docs" / "memo_credito.md"
LIMITE_INTERNO = 3.0  # SUPUESTO: deuda neta / EBITDA máxima que acepta el banco, medio punto bajo el covenant de 3,5x
PARTE_DEL_BANCO = 0.3  # SUPUESTO: el banco toma como máximo 30% de la capacidad de deuda adicional
DIRIGIDO = "Comité de crédito, banca corporativa"


def numero(x: float, decimales: int = 1) -> str:
    return f"{x:,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def veces(x, decimales: int = 1) -> str:
    return f"{numero(x, decimales)}x" if isinstance(x, (int, float)) else str(x)


def zona_altman(z: float) -> str:
    """Altman Z'' para mercados emergentes, con la constante 3,25: segura sobre 5,85, gris entre 4,15 y 5,85."""
    return "segura" if z > 5.85 else "gris" if z >= 4.15 else "de riesgo"


def pct(x, decimales: int = 0) -> str:
    return f"{numero(100 * x, decimales)}%" if isinstance(x, (int, float)) else str(x)


def leer() -> dict:
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    ultimo = p.COLUMNAS[-1]
    alto = len(p.RATIOS) + 3
    nombres = [r[0] for r in p.RATIOS]
    pr, rt = wb["Proyección"], wb["Rating"]
    alto_p = 16 + 4
    fila_nota = 7 + len(p.FACTORES)
    d = {}
    for k, e in enumerate(p.EMPRESAS):
        ini = p.FILA_BLOQUES + 1 + k * alto + 2
        h = wb[e]
        s = 20 + k * alto_p
        fila = lambda i: s + 2 + i
        d[e] = {
            "ratio": {n: wb["Ratios"][f"{ultimo}{ini + i}"].value for i, n in enumerate(nombres)},
            "linea": {c: h[f"{ultimo}{p.FILAS[c]}"].value for c in ("ventas", "ebitda", "deuda", "deuda_neta", "fcf", "cfo", "capex", "int_pag", "caja", "inv_cp", "d_bancos")},
            "fcf_3": [h[f"{c}{p.FILAS['fcf']}"].value for c in p.COLUMNAS],
            "base": {"apal": [pr[f"{c}{fila(13)}"].value for c in "CDE"], "cob": [pr[f"{c}{fila(14)}"].value for c in "CDE"], "dscr": [pr[f"{c}{fila(15)}"].value for c in "CDE"]},
            "estres": {"apal": [pr[f"{c}{fila(13)}"].value for c in "GHI"], "cob": [pr[f"{c}{fila(14)}"].value for c in "GHI"], "dscr": [pr[f"{c}{fila(15)}"].value for c in "GHI"],
                       "ebitda": [pr[f"{c}{fila(1)}"].value for c in "GHI"], "deuda": [pr[f"{c}{fila(8)}"].value for c in "GHI"], "caja": pr[f"G{fila(12)}"].value},
            "nota": rt[f"{'HIJ'[k]}{fila_nota}"].value, "nota_estres": rt[f"{'KLM'[k]}{fila_nota}"].value,
            "puntaje": rt[f"{'HIJ'[k]}{fila_nota - 1}"].value, "puntaje_estres": rt[f"{'KLM'[k]}{fila_nota - 1}"].value,
        }
    em = wb["Emisiones"]
    filas_e = [f for f in range(5, 5 + p.FILAS_EMISIONES) if em[f"A{f}"].value == "Pampa" and em[f"C{f}"].value and em[f"C{f}"].value.year == p.ANIOS[-1]]
    d["emisiones_pampa"] = [(em[f"B{f}"].value, em[f"L{f}"].value, em[f"I{f}"].value or em[f"H{f}"].value) for f in filas_e]
    cv = wb["Covenants"]
    d["covenants"] = {"apal": cv["B4"].value, "cob": cv["B5"].value,
                      "margen": {e: cv[f"H{8 + k * 5}"].value for k, e in enumerate(p.EMPRESAS)}}
    sup = wb["Proyección"]
    d["supuestos"] = {"caida": sup["B6"].value, "margen": sup["B7"].value, "tasa": sup["B8"].value}
    d["vencimientos"] = vencimientos(d)
    return d


def vencimientos(d: dict) -> dict:
    """Por empresa: vencimientos por año, hasta qué año alcanza la caja de 2025 y los años con más vencimientos."""
    import pandas as pd

    v = pd.read_csv(p.VENCIMIENTOS).groupby(["empresa", "anio"]).monto.sum()
    salida = {}
    for e in p.EMPRESAS:
        por_anio = v[e].sort_index()
        caja = d[e]["linea"]["caja"] + d[e]["linea"]["inv_cp"]
        acumulado = por_anio.cumsum()
        cubre = acumulado[acumulado <= caja].index.max() if (acumulado <= caja).any() else None
        primer_hueco = acumulado[acumulado > caja].index.min() if (acumulado > caja).any() else None
        salida[e] = {"por_anio": por_anio.to_dict(), "caja": caja, "cubre_hasta": cubre, "primer_hueco": primer_hueco,
                     "mayores": por_anio.nlargest(3).sort_index().to_dict(), "primer_anio": por_anio.get(p.PROYECCION[0], 0.0)}
    return salida


def capacidad(e: dict) -> float:
    """Deuda adicional que entra sin superar el límite interno en ningún año del estrés."""
    s = e["estres"]
    return min(LIMITE_INTERNO * eb - (de - s["caja"]) for eb, de in zip(s["ebitda"], s["deuda"]))


def escribir(d: dict) -> str:
    pa, yp, vi = d["Pampa"], d["YPF"], d["Vista"]
    r = lambda e, n: e["ratio"][n]
    cap = capacidad(pa)
    linea = 10 * int(PARTE_DEL_BANCO * cap // 10)
    sup = d["supuestos"]
    comparables = "\n".join(
        f"| {n} | " + " | ".join(f(e["ratio"][n]) for e in (pa, yp, vi)) + " |"
        for n, f in (("Margen EBITDA", pct), ("Deuda neta / EBITDA", veces), ("EBITDA / intereses", veces), ("Liquidez corriente", veces),
                     ("Flujo libre / deuda", pct), ("ON / deuda financiera", pct), ("Costo medio de la deuda", lambda x: pct(x, 1)), ("Altman Z'' mercados emergentes", veces)))
    ve = d["vencimientos"]["Pampa"]
    rend = [x[2] for x in d["emisiones_pampa"]]
    corta = min(d["emisiones_pampa"], key=lambda x: x[1])
    zonas = {e: zona_altman(x["ratio"]["Altman Z'' mercados emergentes"]) for e, x in (("Pampa", pa), ("YPF", yp), ("Vista", vi))}
    seguras = [e for e, z in zonas.items() if z == "segura"]
    otras = [e for e in zonas if e not in seguras]
    altman = f"El Altman Z'' de {' y '.join(seguras)} queda en zona segura" + (f"; el de {' y '.join(otras)}, en zona {zonas[otras[0]]}." if otras else ".")
    estres = "\n".join(
        f"| {e} | {veces(max(x['estres']['apal']))} | {veces(min(x['estres']['cob']))} | {veces(min(x['estres']['dscr']), 2)} | {x['nota']} → {x['nota_estres']} |"
        for e, x in (("Pampa", pa), ("YPF", yp), ("Vista", vi)))
    return f"""# Memo de crédito: Pampa Energía S.A.

**Para:** {DIRIGIDO} · **Datos:** estados contables 2023-2025 en dólares (20-F 2025) · **Modelo:** [excel/analisis_crediticio.xlsx](../excel/analisis_crediticio.xlsx)

## Recomendación

**Aprobar una línea de hasta US$ {numero(linea, 0)} M a 3 años en dólares**, preferentemente como prefinanciación de exportaciones o préstamo amortizable,
con covenants de deuda neta / EBITDA ≤ {veces(d['covenants']['apal'])} y EBITDA / intereses ≥ {veces(d['covenants']['cob'])}.

- **Por qué sí:** deuda neta / EBITDA de {veces(r(pa, 'Deuda neta / EBITDA'))}, liquidez corriente de {veces(r(pa, 'Liquidez corriente'))} y caja e inversiones por
  US$ {numero(pa['linea']['caja'] + pa['linea']['inv_cp'], 0)} M. Rating interno **{pa['nota']}** (puntaje {numero(pa['puntaje'], 2)} sobre 5).
- **Cuánto:** en el escenario de estrés, Pampa podría sumar US$ {numero(cap, 0)} M de deuda sin pasar {veces(LIMITE_INTERNO)} de deuda neta / EBITDA en ningún año.
  La línea propuesta es {pct(linea / cap)} de esa capacidad.
- **Por qué no más:** el flujo libre fue negativo los tres años (US$ {', '.join(numero(x, 0) for x in pa['fcf_3'])} M) y la cobertura de servicio de deuda de 2025 fue
  {veces(r(pa, 'Cobertura de servicio de deuda'), 2)}: Pampa paga sus vencimientos con deuda nueva, no con caja propia.

## Perfil financiero comparado (2025)

| Indicador | Pampa | YPF | Vista |
|---|---|---|---|
{comparables}

Pampa es la menos rentable de las tres, pero la menos endeudada y la de mayor liquidez. {altman}

## Financiamiento

- **Depende del mercado de capitales:** {pct(r(pa, 'ON / deuda financiera'))} de su deuda son obligaciones negociables y sólo US$ {numero(pa['linea']['d_bancos'], 0)} M son préstamos financieros.
  Toda su deuda de 2025 está en dólares y a tasa fija.
- **Se financia mejor que el Estado:** sus {len(rend)} colocaciones de {p.ANIOS[-1]} rindieron entre {pct(min(rend), 2)} y {pct(max(rend), 3)}, por debajo del riesgo país del día (ver hoja Emisiones).
- **Qué le aporta el banco:** liquidez de corto plazo y prefinanciación atada a exportaciones, donde la ON no compite. Para plazos largos, la empresa elige el mercado.
- **Precio de referencia:** su ON {corta[0]} se colocó a {numero(corta[1], 0)} años al {pct(corta[2], 2)}. Una línea de plazo parecido tiene que competir con esa tasa.

## Escenario de estrés (2026-2028)

Supuestos: ventas {pct(sup['caida'])} en 2026 y sin recuperación, margen EBITDA {numero(100 * sup['margen'], 0)} pp menor, deuda que se renueva {numero(100 * sup['tasa'], 0)} pp más cara y
capex igual a la depreciación. Peor año de cada indicador:

| Empresa | Deuda neta / EBITDA | EBITDA / intereses | Cobertura de servicio | Rating interno |
|---|---|---|---|---|
{estres}

- **Pampa no rompe los covenants supuestos ni bajo estrés**, y en el escenario base su EBITDA podría caer {pct(d['covenants']['margen']['Pampa'])} antes de llegar a
  {veces(d['covenants']['apal'])} de deuda neta / EBITDA.
- **El riesgo real es refinanciar, pero no es inmediato:** bajo estrés, la cobertura de servicio de deuda baja hasta {veces(min(pa['estres']['dscr']), 2)}.
  Aun así, con la caja de 2025 (US$ {numero(ve['caja'], 0)} M) Pampa paga todo lo que vence hasta {ve['primer_hueco'] - 1} aunque el mercado se cierre.
  YPF está en la situación opuesta: en {p.PROYECCION[0]} le vencen US$ {numero(d['vencimientos']['YPF']['primer_anio'], 0)} M, {veces(d['vencimientos']['YPF']['primer_anio'] / d['vencimientos']['YPF']['caja'])} su caja.

## Condiciones propuestas

1. Covenants de mantenimiento: deuda neta / EBITDA ≤ {veces(d['covenants']['apal'])} y EBITDA / intereses ≥ {veces(d['covenants']['cob'])}, medidos semestralmente.
2. Incumplimiento cruzado con las ON vigentes.
3. Si es prefinanciación, cesión de cobranzas de exportación.
4. Revisión anual, o antes si el riesgo país supera el nivel de 2023.

## Riesgos y límites del análisis

- **Riesgo soberano y de transferencia:** la empresa se financia por debajo del riesgo país, pero opera en Argentina. Este rating interno no aplica techo soberano.
- **Concentración de vencimientos:** Pampa paga poco hasta {p.PROYECCION[-1]} y después vencen bloques grandes:
  {', '.join(f'US$ {numero(m, 0)} M en {a}' for a, m in ve['mayores'].items())} (nota 12.5.2 del 20-F). La línea propuesta vence antes que esos bloques.
- **Supuestos declarados:** los covenants, el límite interno de {veces(LIMITE_INTERNO)} y la parte del banco ({pct(PARTE_DEL_BANCO)}) son supuestos, no datos de contratos.
- **Cifras propias:** el EBITDA excluye resultados no recurrentes y asociadas, así que no coincide con el EBITDA ajustado que publica la empresa.
"""


SALIDA_EN = p.ROOT / "docs" / "credit_memo.md"
ZONA_EN = {"segura": "safe", "gris": "grey", "de riesgo": "distress"}


def number(x: float, decimals: int = 1) -> str:
    return f"{x:,.{decimals}f}"


def times(x, decimals: int = 1) -> str:
    return f"{number(x, decimals)}x" if isinstance(x, (int, float)) else str(x)


def percent(x, decimals: int = 0) -> str:
    return f"{number(100 * x, decimals)}%" if isinstance(x, (int, float)) else str(x)


def write_en(d: dict) -> str:
    """The same memo in English, with the same figures."""
    pa, yp, vi = d["Pampa"], d["YPF"], d["Vista"]
    r = lambda e, n: e["ratio"][n]
    cap = capacidad(pa)
    line = 10 * int(PARTE_DEL_BANCO * cap // 10)
    sup = d["supuestos"]
    names = {"Margen EBITDA": "EBITDA margin", "Deuda neta / EBITDA": "Net debt / EBITDA", "EBITDA / intereses": "EBITDA / interest",
             "Liquidez corriente": "Current ratio", "Flujo libre / deuda": "Free cash flow / debt", "ON / deuda financiera": "Bonds / financial debt",
             "Costo medio de la deuda": "Average cost of debt", "Altman Z'' mercados emergentes": "Altman Z'' emerging markets"}
    fmt = {"Margen EBITDA": percent, "Flujo libre / deuda": percent, "ON / deuda financiera": percent, "Costo medio de la deuda": lambda x: percent(x, 1)}
    peers = "\n".join(f"| {en} | " + " | ".join(fmt.get(es, times)(e["ratio"][es]) for e in (pa, yp, vi)) + " |" for es, en in names.items())
    ve = d["vencimientos"]["Pampa"]
    yields = [x[2] for x in d["emisiones_pampa"]]
    short = min(d["emisiones_pampa"], key=lambda x: x[1])
    zones = {e: ZONA_EN[zona_altman(x["ratio"]["Altman Z'' mercados emergentes"])] for e, x in (("Pampa", pa), ("YPF", yp), ("Vista", vi))}
    safe = [e for e, z in zones.items() if z == "safe"]
    rest = [e for e in zones if e not in safe]
    altman = f"{' and '.join(safe)}'s Altman Z'' is in the safe zone" + (f"; {' and '.join(rest)}'s, in the {zones[rest[0]]} zone." if rest else ".")
    stress = "\n".join(f"| {e} | {times(max(x['estres']['apal']))} | {times(min(x['estres']['cob']))} | {times(min(x['estres']['dscr']), 2)} | {x['nota']} → {x['nota_estres']} |"
                       for e, x in (("Pampa", pa), ("YPF", yp), ("Vista", vi)))
    ypf = d["vencimientos"]["YPF"]
    return f"""# Credit memo: Pampa Energía S.A.

**To:** Credit committee, corporate banking · **Data:** 2023-2025 financial statements in US dollars (2025 20-F) · **Model:** [excel/analisis_crediticio.xlsx](../excel/analisis_crediticio.xlsx)

## Recommendation

**Approve a facility of up to US$ {number(line, 0)} M, 3 years, in US dollars**, preferably as export pre-financing or an amortizing loan,
with covenants of net debt / EBITDA ≤ {times(d['covenants']['apal'])} and EBITDA / interest ≥ {times(d['covenants']['cob'])}.

- **Why yes:** net debt / EBITDA of {times(r(pa, 'Deuda neta / EBITDA'))}, a current ratio of {times(r(pa, 'Liquidez corriente'))} and cash and investments of
  US$ {number(pa['linea']['caja'] + pa['linea']['inv_cp'], 0)} M. Internal rating **{pa['nota']}** (score {number(pa['puntaje'], 2)} out of 5).
- **How much:** in the stress scenario, Pampa could add US$ {number(cap, 0)} M of debt without exceeding {times(LIMITE_INTERNO)} net debt / EBITDA in any year.
  The proposed facility is {percent(line / cap)} of that capacity.
- **Why not more:** free cash flow was negative in all three years (US$ {', '.join(number(x, 0) for x in pa['fcf_3'])} M) and 2025 debt service coverage was
  {times(r(pa, 'Cobertura de servicio de deuda'), 2)}: Pampa repays maturities with new debt, not with its own cash.

## Peer comparison (2025)

| Metric | Pampa | YPF | Vista |
|---|---|---|---|
{peers}

Pampa is the least profitable of the three, but the least leveraged and the most liquid. {altman}

## Funding

- **It relies on the capital market:** {percent(r(pa, 'ON / deuda financiera'))} of its debt is corporate bonds (obligaciones negociables, ON) and only US$ {number(pa['linea']['d_bancos'], 0)} M is loans.
  All of its 2025 debt is in dollars and at a fixed rate.
- **It borrows below the sovereign:** its {len(yields)} placements in {p.ANIOS[-1]} yielded between {percent(min(yields), 2)} and {percent(max(yields), 3)}, below that day's country risk (Emisiones sheet).
- **What a bank adds:** short-term liquidity and export pre-financing, where bonds do not compete. For long tenors, the company chooses the market.
- **Pricing reference:** its {short[0]} bond was placed at {number(short[1], 0)} years at {percent(short[2], 2)}. A facility of similar tenor has to compete with that rate.

## Stress scenario (2026-2028)

Assumptions: sales {percent(sup['caida'])} in 2026 with no recovery, EBITDA margin {number(100 * sup['margen'], 0)} pp lower, refinanced debt {number(100 * sup['tasa'], 0)} pp more expensive
and capex equal to depreciation. Worst year of each metric:

| Company | Net debt / EBITDA | EBITDA / interest | Debt service coverage | Internal rating |
|---|---|---|---|---|
{stress}

- **Pampa does not breach the assumed covenants, even under stress**, and in the base case its EBITDA could fall {percent(d['covenants']['margen']['Pampa'])} before reaching
  {times(d['covenants']['apal'])} net debt / EBITDA.
- **The real risk is refinancing, but it is not immediate:** under stress, debt service coverage falls to {times(min(pa['estres']['dscr']), 2)}.
  Even so, with its 2025 cash (US$ {number(ve['caja'], 0)} M) Pampa can repay everything maturing until {ve['primer_hueco'] - 1} even if the market closes.
  YPF is in the opposite position: US$ {number(ypf['primer_anio'], 0)} M matures in {p.PROYECCION[0]}, {times(ypf['primer_anio'] / ypf['caja'])} its cash.

## Proposed terms

1. Maintenance covenants: net debt / EBITDA ≤ {times(d['covenants']['apal'])} and EBITDA / interest ≥ {times(d['covenants']['cob'])}, tested semi-annually.
2. Cross-default with outstanding bonds.
3. If pre-financing, assignment of export receivables.
4. Annual review, or earlier if country risk exceeds its 2023 level.

## Risks and limits of the analysis

- **Sovereign and transfer risk:** the company borrows below country risk, but operates in Argentina. This internal rating applies no sovereign ceiling.
- **Maturity concentration:** Pampa repays little until {p.PROYECCION[-1]}, then large blocks mature:
  {', '.join(f'US$ {number(m, 0)} M in {a}' for a, m in ve['mayores'].items())} (20-F note 12.5.2). The proposed facility matures before those blocks.
- **Stated assumptions:** the covenants, the {times(LIMITE_INTERNO)} internal limit and the bank's share ({percent(PARTE_DEL_BANCO)}) are assumptions, not contract data.
- **Own figures:** EBITDA excludes one-off results and associates, so it differs from the adjusted EBITDA the company reports.
"""


def exportar(d: dict) -> None:
    """Rating y proyección a outputs/, para el tablero: las mismas cifras que el memo."""
    import pandas as pd

    salida = p.ROOT / "outputs"
    pd.DataFrame({"empresa": p.EMPRESAS, "orden": range(1, len(p.EMPRESAS) + 1)}).to_csv(salida / "empresas.csv", index=False, encoding="utf-8")
    filas = []
    for k, e in enumerate(p.EMPRESAS):
        x, r = d[e], d[e]["ratio"]
        filas.append(dict(empresa=e, orden=k + 1, rating_2025=x["nota"], rating_estres=x["nota_estres"], puntaje_2025=x["puntaje"], puntaje_estres=x["puntaje_estres"],
                          deuda_neta_ebitda=r["Deuda neta / EBITDA"], ebitda_intereses=r["EBITDA / intereses"], liquidez_corriente=r["Liquidez corriente"],
                          on_deuda=r["ON / deuda financiera"], peor_apalancamiento_estres=max(x["estres"]["apal"]), peor_cobertura_estres=min(x["estres"]["cob"]),
                          peor_servicio_estres=min(x["estres"]["dscr"]), ebitda_puede_caer=d["covenants"]["margen"][e],
                          caja_2025=d["vencimientos"][e]["caja"], vence_primer_anio=d["vencimientos"][e]["primer_anio"],
                          caja_sobre_vencimientos=d["vencimientos"][e]["caja"] / d["vencimientos"][e]["primer_anio"],
                          anios_cubiertos=(d["vencimientos"][e]["primer_hueco"] or max(d["vencimientos"][e]["por_anio"]) + 1) - p.PROYECCION[0]))
    pd.DataFrame(filas).to_csv(salida / "rating.csv", index=False, encoding="utf-8")
    proy = []
    reales = {"deuda_neta_ebitda": "Deuda neta / EBITDA", "cobertura_servicio": "Cobertura de servicio de deuda"}
    for e in p.EMPRESAS:
        for escenario, clave in (("Base", "base"), ("Estrés", "estres")):
            for indicador, ratio in reales.items():
                proy.append(dict(empresa=e, escenario=escenario, anio=p.ANIOS[-1], indicador=indicador, valor=d[e]["ratio"][ratio]))
                serie = d[e][clave]["apal" if indicador == "deuda_neta_ebitda" else "dscr"]
                proy += [dict(empresa=e, escenario=escenario, anio=a, indicador=indicador, valor=v) for a, v in zip(p.PROYECCION, serie)]
    pd.DataFrame(proy).to_csv(salida / "proyeccion.csv", index=False, encoding="utf-8")


def main() -> None:
    d = leer()
    SALIDA.write_text(escribir(d), encoding="utf-8", newline="\n")
    SALIDA_EN.write_text(write_en(d), encoding="utf-8", newline="\n")
    exportar(d)
    print(f"{SALIDA.name} escrito; rating y proyección exportados a outputs/")


if __name__ == "__main__":
    main()
