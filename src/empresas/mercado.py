"""Series de mercado para ubicar cada emisión en su momento.

`python -m empresas.mercado` escribe data/mercado.csv con riesgo país (ArgentinaDatos), tasas del Tesoro de EE.UU. (FRED)
y tasas locales del BCRA, en días hábiles desde 2020, y data/tasas.csv con las tasas mensuales de crédito a empresas y de
plazo fijo, en dólares y en pesos.
"""

from __future__ import annotations

import io

import pandas as pd
import requests
import truststore
import xlrd

from empresas.plantilla import MERCADO, PLAZOS_TESORO, TASAS_SISTEMA

DESDE = "2020-01-01"
RIESGO_PAIS = "https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais"
FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS{}"
BCRA = "https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/{}"
SERIES_BCRA = {"badlar": 7, "adelantos": 145}  # adelantos: 10 millones o más, de 1 a 7 días
TASAS_BCRA = "https://www.bcra.gob.ar/archivos/Pdfs/PublicacionesEstadisticas/preser_tas.xls"
HOJA_EMPRESAS = "Tasas_pers_jur_otras"  # personas jurídicas que no son pymes; préstamos a tasa fija y repactable
COLUMNAS_TASAS = {"doc_usd": 74, "doc_usd_corto": 76, "doc_usd_largo": 78, "doc_ars": 11}  # documentos a sola firma
PERIODO_ARS, PERIODO_USD = 2, 65
PLAZO_FIJO = "https://apis.datos.gob.ar/series/api/series/?ids=174.1_T_INTERUS_0_0_43,174.1_T_INTERD_0_0_41&start_date=2020-01-01&limit=1000&format=csv"


truststore.inject_into_ssl()  # verifica con los certificados del sistema operativo


def riesgo_pais() -> pd.Series:
    r = requests.get(RIESGO_PAIS, timeout=60)
    r.raise_for_status()
    d = pd.DataFrame(r.json())
    return d.set_index(pd.to_datetime(d.fecha)).valor.rename("riesgo_pais").groupby(level=0).last()


def tesoro(plazo: int) -> pd.Series:
    r = requests.get(FRED.format(plazo), timeout=60)
    r.raise_for_status()
    d = pd.read_csv(io.StringIO(r.text), na_values=".")
    return d.set_index(pd.to_datetime(d.observation_date))[f"DGS{plazo}"].rename(f"ust_{plazo}")


def bcra(nombre: str, variable: int) -> pd.Series:
    filas, offset = [], 0
    while True:
        r = requests.get(BCRA.format(variable), params={"desde": DESDE, "limit": 1000, "offset": offset}, timeout=60)
        r.raise_for_status()
        detalle = r.json()["results"][0]["detalle"]
        filas += detalle
        offset += len(detalle)
        if len(detalle) < 1000:
            break
    d = pd.DataFrame(filas)
    return d.set_index(pd.to_datetime(d.fecha)).valor.rename(nombre).sort_index()


def tasas_credito(contenido: bytes) -> pd.DataFrame:
    """Lee el cuadro del BCRA. Falla si cambiaron las columnas, para no mezclar pesos con dólares."""
    s = xlrd.open_workbook(file_contents=contenido).sheet_by_name(HOJA_EMPRESAS)
    controles = [(5, 66, "foreign"), (6, 74, "Bills"), (7, 76, "Up to 89"), (7, 78, "From 90"), (5, 3, "domestic"), (6, 11, "Bills")]
    for fila, columna, texto in controles:
        assert texto in str(s.cell_value(fila, columna)), f"cambió el cuadro del BCRA: {fila},{columna}"
    filas = []
    for i in range(10, s.nrows):
        periodo = s.cell_value(i, PERIODO_USD)
        if not isinstance(periodo, float) or periodo != s.cell_value(i, PERIODO_ARS) or periodo < 202001 or not 1 <= int(periodo) % 100 <= 12:
            continue
        valores = {k: s.cell_value(i, c) for k, c in COLUMNAS_TASAS.items()}
        filas.append({"mes": pd.Timestamp(year=int(periodo) // 100, month=int(periodo) % 100, day=1), **valores})
    d = pd.DataFrame(filas).set_index("mes")
    return d.apply(pd.to_numeric, errors="coerce").where(lambda x: x > 0)  # cero: sin operaciones ese mes


def plazo_fijo() -> pd.DataFrame:
    r = requests.get(PLAZO_FIJO, timeout=60)
    r.raise_for_status()
    d = pd.read_csv(io.StringIO(r.text), parse_dates=["indice_tiempo"]).set_index("indice_tiempo")
    d.columns = ["pf_usd", "pf_ars"]
    return d.where(d > 0)


def main() -> None:
    series = [riesgo_pais(), *(tesoro(p) for p in PLAZOS_TESORO), *(bcra(n, v) for n, v in SERIES_BCRA.items())]
    d = pd.concat(series, axis=1, sort=True)
    d = d.ffill().loc[DESDE:]
    d = d[d.index.day_of_week < 5].dropna()
    d.index.name = "fecha"
    MERCADO.parent.mkdir(exist_ok=True)
    d.to_csv(MERCADO, encoding="utf-8")
    print(f"{MERCADO.name}: {len(d)} días, {d.index[0].date()} a {d.index[-1].date()}")
    r = requests.get(TASAS_BCRA, timeout=180)
    r.raise_for_status()
    t = tasas_credito(r.content).join(plazo_fijo(), how="left")
    t.index.name = "mes"
    t.to_csv(TASAS_SISTEMA, encoding="utf-8")
    print(f"{TASAS_SISTEMA.name}: {len(t)} meses, {t.index[0].date()} a {t.index[-1].date()}")


if __name__ == "__main__":
    main()
