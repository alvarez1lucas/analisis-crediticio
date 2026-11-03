import re
import openpyxl
import pandas as pd
import pytest

from empresas import plantilla as p
from empresas.sql import OUTPUTS

# indicador de SQL -> fila de Ratios en el Excel
EQUIVALENCIAS = {"margen_ebitda": "Margen EBITDA", "deuda_neta_ebitda": "Deuda neta / EBITDA", "ebitda_intereses": "EBITDA / intereses",
                 "liquidez_corriente": "Liquidez corriente", "flujo_libre_deuda": "Flujo libre / deuda", "capex_ventas": "Capex / ventas",
                 "roe": "Rentabilidad sobre patrimonio", "margen_neto": "Margen neto", "rotacion_activos": "Rotación de activos",
                 "activo_patrimonio": "Activo / patrimonio"}


@pytest.fixture(scope="module")
def excel():
    return openpyxl.load_workbook(p.ARCHIVO, data_only=True)


def test_sql_y_excel_dan_los_mismos_indicadores(excel):
    ind = pd.read_csv(OUTPUTS / "indicadores.csv")
    nombres = [r[0] for r in p.RATIOS]
    alto = len(p.RATIOS) + 3
    comparados = 0
    for k, e in enumerate(p.EMPRESAS):
        ini = p.FILA_BLOQUES + 1 + k * alto + 2
        for indicador, ratio in EQUIVALENCIAS.items():
            for col, anio in zip(p.COLUMNAS, p.ANIOS):
                en_excel = excel["Ratios"][f"{col}{ini + nombres.index(ratio)}"].value
                fila = ind[(ind.empresa == e) & (ind.anio == anio) & (ind.indicador == indicador)]
                if isinstance(en_excel, str):  # n.d. o n.s.: SQL no devuelve fila
                    assert fila.empty, (e, anio, indicador)
                    continue
                assert fila.valor.iloc[0] == pytest.approx(en_excel, rel=1e-9), (e, anio, indicador)
                comparados += 1
    assert comparados == len(EQUIVALENCIAS) * len(p.EMPRESAS) * len(p.ANIOS)  # todos tienen dato


def test_sql_y_excel_dan_el_mismo_spread_de_cada_on(excel):
    e = pd.read_csv(OUTPUTS / "emisiones.csv")
    h = excel["Emisiones"]
    filas = {(h[f"A{f}"].value, h[f"B{f}"].value): f for f in range(5, 5 + p.FILAS_EMISIONES) if h[f"A{f}"].value}
    assert len(filas) == len(e) == 42
    for r in e.itertuples():
        f = filas[(r.empresa, r.clase)]
        assert r.riesgo_pais == h[f"M{f}"].value
        if r.instrumento == "Dólar duro":
            assert r.spread_pb == pytest.approx(h[f"O{f}"].value, abs=1e-6)
        else:
            assert h[f"O{f}"].value == "n.a." and pd.isna(r.spread_pb)


def test_dupont_multiplica_a_la_rentabilidad_sobre_patrimonio():
    ind = pd.read_csv(OUTPUTS / "indicadores.csv").pivot_table(index=["empresa", "anio"], columns="indicador", values="valor")
    producto = ind.margen_neto * ind.rotacion_activos * ind.activo_patrimonio
    assert (producto - ind.roe).abs().max() < 1e-9


def test_vencimientos_cuadran_con_la_deuda_del_balance_y_el_puente_con_el_flujo_libre():
    v = pd.read_csv(OUTPUTS / "vencimientos.csv").groupby("empresa").monto.sum()
    lineas = pd.read_csv(OUTPUTS / "lineas.csv")
    ultimo = lineas[lineas.anio == p.ANIOS[-1]].set_index("empresa")
    assert ((ultimo.deuda_cp + ultimo.deuda_lp) - v).abs().max() < 1
    puente = pd.read_csv(OUTPUTS / "puente.csv").groupby(["empresa", "anio"]).monto.sum()
    assert (puente - lineas.set_index(["empresa", "anio"]).fcf).abs().max() < 1e-9


def test_colocadores_salen_de_emisiones_existentes_y_separan_roles():
    c = pd.read_csv(OUTPUTS / "colocadores.csv")
    e = pd.read_csv(OUTPUTS / "emisiones.csv")
    assert set(zip(c.empresa, c.clase)) <= set(zip(e.empresa, e.clase))
    assert set(c.rol) <= {"Organizador y colocador", "Colocador", "Colocador local", "Comprador inicial"}
    assert (c[c.rol == "Comprador inicial"].tipo_entidad == "Banco de inversión del exterior").all()


def test_la_guia_no_deja_huecos():
    from empresas import guia, memo
    texto = guia.escribir(guia.cifras(), memo.leer())
    assert not re.search(r"\bnan\b|\bNone\b", texto)
    assert re.search(r"\bnan\b", "valor nan") and not re.search(r"\bnan\b", "financiero")  # el control detecta lo que tiene que detectar


def test_modelo_de_instrumento():
    m = pd.read_csv(OUTPUTS / "modelo_instrumento.csv").set_index("medida").valor
    assert m["coeficiente_por_1000_pb"] > 0 and m["p_valor"] < 0.05
    assert m["probabilidad_con_700"] < m["probabilidad_con_1500"] < m["probabilidad_con_2500"]
    assert m["emisiones"] == 42
