"""Series de mercado para ubicar cada emisión en su momento.

`python -m empresas.mercado` escribe data/mercado.csv con riesgo país (ArgentinaDatos), tasas del Tesoro de EE.UU. (FRED)
y tasas locales del BCRA, en días hábiles desde 2020.
"""

from __future__ import annotations

import io

import pandas as pd
import requests
import truststore

from empresas.plantilla import MERCADO, PLAZOS_TESORO

DESDE = "2020-01-01"
RIESGO_PAIS = "https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais"
FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS{}"
BCRA = "https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/{}"
SERIES_BCRA = {"badlar": 7, "adelantos": 145}  # adelantos: 10 millones o más, de 1 a 7 días


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


def main() -> None:
    series = [riesgo_pais(), *(tesoro(p) for p in PLAZOS_TESORO), *(bcra(n, v) for n, v in SERIES_BCRA.items())]
    d = pd.concat(series, axis=1).sort_index()
    d = d.ffill().loc[DESDE:]
    d = d[d.index.dayofweek < 5].dropna()
    d.index.name = "fecha"
    MERCADO.parent.mkdir(exist_ok=True)
    d.to_csv(MERCADO, encoding="utf-8")
    print(f"{MERCADO.name}: {len(d)} días, {d.index[0].date()} a {d.index[-1].date()}")


if __name__ == "__main__":
    main()
