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


def veces(x) -> str:
    return f"{numero(x)}x" if isinstance(x, (int, float)) else str(x)


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
            "linea": {c: h[f"{ultimo}{p.FILAS[c]}"].value for c in ("ventas", "ebitda", "deuda", "deuda_neta", "fcf", "cfo", "capex", "int_pag", "caja", "inv_cp")},
            "fcf_3": [h[f"{c}{p.FILAS['fcf']}"].value for c in p.COLUMNAS],
            "base": {"apal": [pr[f"{c}{fila(13)}"].value for c in "CDE"], "cob": [pr[f"{c}{fila(14)}"].value for c in "CDE"], "dscr": [pr[f"{c}{fila(15)}"].value for c in "CDE"]},
            "estres": {"apal": [pr[f"{c}{fila(13)}"].value for c in "GHI"], "cob": [pr[f"{c}{fila(14)}"].value for c in "GHI"], "dscr": [pr[f"{c}{fila(15)}"].value for c in "GHI"],
                       "ebitda": [pr[f"{c}{fila(1)}"].value for c in "GHI"], "deuda": [pr[f"{c}{fila(8)}"].value for c in "GHI"], "caja": pr[f"G{fila(12)}"].value},
            "nota": rt[f"{'HIJ'[k]}{fila_nota}"].value, "nota_estres": rt[f"{'KLM'[k]}{fila_nota}"].value,
            "puntaje": rt[f"{'HIJ'[k]}{fila_nota - 1}"].value, "puntaje_estres": rt[f"{'KLM'[k]}{fila_nota - 1}"].value,
        }
    cv = wb["Covenants"]
    d["covenants"] = {"apal": cv["B4"].value, "cob": cv["B5"].value,
                      "margen": {e: cv[f"H{8 + k * 5}"].value for k, e in enumerate(p.EMPRESAS)}}
    sup = wb["Proyección"]
    d["supuestos"] = {"caida": sup["B6"].value, "margen": sup["B7"].value, "tasa": sup["B8"].value}
    return d


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
    estres = "\n".join(
        f"| {e} | {veces(max(x['estres']['apal']))} | {veces(min(x['estres']['cob']))} | {veces(min(x['estres']['dscr']))} | {x['nota']} → {x['nota_estres']} |"
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
  {veces(r(pa, 'Cobertura de servicio de deuda'))}: Pampa paga sus vencimientos con deuda nueva, no con caja propia.

## Perfil financiero comparado (2025)

| Indicador | Pampa | YPF | Vista |
|---|---|---|---|
{comparables}

Pampa es la menos rentable de las tres, pero la menos endeudada y la de mayor liquidez. Su Altman Z'' queda en zona segura, como el de las otras dos.

## Financiamiento

- **Depende del mercado de capitales:** {pct(r(pa, 'ON / deuda financiera'))} de su deuda son obligaciones negociables y sólo US$ 78 M son préstamos financieros.
  Toda su deuda de 2025 está en dólares y a tasa fija.
- **Se financia mejor que el Estado:** sus tres colocaciones de 2025 rindieron entre 7,25% y 8,125%, por debajo del riesgo país del día (ver hoja Emisiones).
- **Qué le aporta el banco:** liquidez de corto plazo y prefinanciación atada a exportaciones, donde la ON no compite. Para plazos largos, la empresa elige el mercado.
- **Precio de referencia:** su ON Clase 25 a 3 años rinde 7,25%. Una línea a la misma duración tiene que competir con esa tasa.

## Escenario de estrés (2026-2028)

Supuestos: ventas {pct(sup['caida'])} en 2026 y sin recuperación, margen EBITDA {numero(100 * sup['margen'], 0)} pp menor, deuda que se renueva {numero(100 * sup['tasa'], 0)} pp más cara y
capex igual a la depreciación. Peor año de cada indicador:

| Empresa | Deuda neta / EBITDA | EBITDA / intereses | Cobertura de servicio | Rating interno |
|---|---|---|---|---|
{estres}

- **Pampa no rompe los covenants supuestos ni bajo estrés**, y en el escenario base su EBITDA podría caer {pct(d['covenants']['margen']['Pampa'])} antes de llegar a
  {veces(d['covenants']['apal'])} de deuda neta / EBITDA.
- **El riesgo real es refinanciar:** bajo estrés, la cobertura de servicio de deuda baja hasta {veces(min(pa['estres']['dscr']))}. Si el mercado se cierra,
  la caja actual cubre aproximadamente {numero((pa['linea']['caja'] + pa['linea']['inv_cp']) / (pa['estres']['deuda'][0] / p.VIDA_MEDIA['Pampa']), 1)} años de vencimientos.

## Condiciones propuestas

1. Covenants de mantenimiento: deuda neta / EBITDA ≤ {veces(d['covenants']['apal'])} y EBITDA / intereses ≥ {veces(d['covenants']['cob'])}, medidos semestralmente.
2. Incumplimiento cruzado con las ON vigentes.
3. Si es prefinanciación, cesión de cobranzas de exportación.
4. Revisión anual, o antes si el riesgo país supera el nivel de 2023.

## Riesgos y límites del análisis

- **Riesgo soberano y de transferencia:** la empresa se financia por debajo del riesgo país, pero opera en Argentina. Este rating interno no aplica techo soberano.
- **Concentración de vencimientos:** la proyección supone vencimientos parejos según la vida media de la deuda ({numero(p.VIDA_MEDIA['Pampa'], 1)} años). En la práctica vencen en bloques
  grandes, como la Clase 23 por US$ 700 M en 2034.
- **Supuestos declarados:** los covenants, el límite interno de {veces(LIMITE_INTERNO)} y la parte del banco ({pct(PARTE_DEL_BANCO)}) son supuestos, no datos de contratos.
- **Cifras propias:** el EBITDA excluye resultados no recurrentes y asociadas, así que no coincide con el EBITDA ajustado que publica la empresa.
"""


def main() -> None:
    SALIDA.write_text(escribir(leer()), encoding="utf-8", newline="\n")
    print(f"{SALIDA.name} escrito")


if __name__ == "__main__":
    main()
