"""Tendencias de financiamiento: qué condiciones acompañan elegir dólar linked en vez de dólar duro.

`python -m empresas.tendencias` ajusta una regresión logística sobre outputs/emisiones.csv y escribe
outputs/modelo_instrumento.csv. Es descriptiva: con 42 emisiones de tres empresas, muchas del mismo mes,
las observaciones no son independientes y el modelo no sirve para predecir.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import statsmodels.api as sm

from empresas.sql import OUTPUTS

ESCALA = 1000  # el coeficiente se lee por cada 1.000 puntos básicos de riesgo país
PUNTOS = [700, 1000, 1500, 2000, 2500]


def ajustar(e: pd.DataFrame):
    y = (e.instrumento == "Dólar linked").astype(int)
    x = sm.add_constant(e.riesgo_pais / ESCALA)
    return sm.Logit(y, x).fit(disp=0)


def main() -> None:
    e = pd.read_csv(OUTPUTS / "emisiones.csv")
    m = ajustar(e)
    bajo, alto = m.conf_int().loc["riesgo_pais"]
    filas = [dict(medida="coeficiente_por_1000_pb", valor=m.params["riesgo_pais"], inferior=bajo, superior=alto),
             dict(medida="p_valor", valor=m.pvalues["riesgo_pais"]),
             dict(medida="pseudo_r2", valor=m.prsquared),
             dict(medida="emisiones", valor=len(e)),
             dict(medida="dolar_linked_con_riesgo_pais_bajo_1500", valor=int(((e.instrumento == "Dólar linked") & (e.riesgo_pais < 1500)).sum()))]
    filas += [dict(medida=f"probabilidad_con_{rp}", valor=float(m.predict(np.array([[1, rp / ESCALA]]))[0])) for rp in PUNTOS]
    pd.DataFrame(filas).to_csv(OUTPUTS / "modelo_instrumento.csv", index=False, encoding="utf-8")
    print(pd.DataFrame(filas).round(3).to_string(index=False))


if __name__ == "__main__":
    main()
