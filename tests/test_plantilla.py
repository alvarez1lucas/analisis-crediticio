from datetime import datetime

import openpyxl
import pandas as pd
import pytest

from empresas import plantilla as p

# Empresa inventada que cuadra: sólo prueba las fórmulas, no es un dato real.
EJEMPLO = {"ventas": 1000, "costo": 600, "comercial": 50, "admin": 40, "explor": 10, "otros_op": 5, "deterioro": 5, "amort": 100,
           "asociadas": 10, "ing_fin": 20, "gasto_int": 50, "otros_fin": -10, "impuesto": 60, "neto_inf": 210,
           "caja": 100, "inv_cp": 50, "creditos": 150, "inventarios": 80, "otros_ac": 20, "ppe": 1200, "intang": 100, "inv_asoc": 150, "otros_anc": 50,
           "activo_inf": 1900, "deuda_cp": 100, "arr_cp": 10, "proveedores": 120, "otros_pc": 70, "deuda_lp": 500, "arr_lp": 40, "otros_pnc": 160,
           "patrimonio": 900, "res_acum": 300, "cfo": 350, "capex": 200, "int_pag": 45, "dividendos": 30,
           "d_on": 500, "d_bancos": 60, "d_prefin": 30, "d_otra": 10, "m_dolar": 550, "m_linked": 20, "m_pesos": 30, "m_otra": 0,
           "v_1": 100, "v_2": 150, "v_3": 200, "v_mas": 150, "tasa_fija": 480}
MERCADO = pd.DataFrame({"fecha": ["2024-01-02", "2024-01-03"], "riesgo_pais": [1900, 800], "ust_1": [4.0, 4.1], "ust_2": [4.2, 4.3], "ust_3": [4.3, 4.4],
                        "ust_5": [4.5, 4.6], "ust_7": [4.7, 4.8], "ust_10": [4.9, 5.0], "badlar": [100, 101], "adelantos": [110, 111]})


@pytest.fixture(scope="module")
def calculado(tmp_path_factory):
    base = tmp_path_factory.mktemp("xl")
    vacio = base / "vacio.xlsx"
    p.construir(vacio, {}, MERCADO)
    wb = openpyxl.load_workbook(vacio)
    e = wb["Emisiones"]
    fila = dict(A="Pampa", B="Clase X", C=datetime(2024, 1, 3), D=datetime(2030, 1, 3), E="Dólar duro", F="Nueva York", G=100, H=0.08, I=0.085)
    for col, valor in fila.items():
        e[f"{col}5"] = valor
    fila.update(E="Pesos tasa variable", C=datetime(2024, 1, 2))
    for col, valor in fila.items():
        e[f"{col}6"] = valor
    c = wb["Créditos"]
    for col, valor in dict(A="Pampa", B=2025, C="Prefinanciación", E="Dólar", F="Variable", G="SOFR", H=0.03, I=datetime(2024, 1, 1), J=datetime(2026, 1, 1), K=65).items():
        c[f"{col}5"] = valor
    ws = wb["Pampa"]
    assert set(EJEMPLO) == set(p.CARGA)
    for col in p.COLUMNAS:
        for clave, valor in EJEMPLO.items():
            ws[f"{col}{p.FILAS[clave]}"] = valor
    lleno = base / "lleno.xlsx"
    wb.save(lleno)
    return p.calcular(lleno), p.calcular(vacio)


def bloque(fila_ratio, col, empresa=0):
    alto = len(p.RATIOS) + 3
    return f"{col}{p.FILA_BLOQUES + 1 + empresa * alto + 2 + fila_ratio}"


def indice(texto):
    return [r[0] for r in p.RATIOS].index(texto)


def test_cuadra_y_marca_ok(calculado):
    v, _ = calculado
    for col in p.COLUMNAS:
        assert v[("PAMPA", f"{col}{p.FILAS['ebit']}")] == 300
        assert v[("PAMPA", f"{col}{p.FILAS['ebitda']}")] == 405
        assert v[("PAMPA", f"{col}{p.FILAS['estado']}")] == "OK"


def test_ratios_a_mano(calculado):
    v, _ = calculado
    deuda, neta = 650, 500
    assert v[("RATIOS", bloque(indice("Deuda / EBITDA"), "F"))] == pytest.approx(deuda / 405)
    assert v[("RATIOS", bloque(indice("Deuda neta / EBITDA"), "F"))] == pytest.approx(neta / 405)
    assert v[("RATIOS", bloque(indice("EBITDA / intereses"), "F"))] == pytest.approx(405 / 50)
    assert v[("RATIOS", bloque(indice("Cobertura de servicio de deuda"), "F"))] == pytest.approx(350 / (45 + 110))
    assert v[("RATIOS", bloque(indice("Flujo libre / deuda"), "F"))] == pytest.approx((350 - 45 - 200) / 650)
    assert v[("RATIOS", bloque(indice("Liquidez corriente"), "F"))] == pytest.approx(400 / 300)


def test_el_primer_anio_no_inventa_datos_del_anterior(calculado):
    v, _ = calculado
    wb = openpyxl.load_workbook(p.ARCHIVO)
    for ratio in ("Crecimiento de ventas", "Cobertura de servicio de deuda"):
        assert wb["Ratios"][bloque(indice(ratio), "B")].value == "n.d."


def test_emision_en_dolares_toma_tesoro_y_riesgo_pais_del_dia(calculado):
    v, _ = calculado
    assert v[("EMISIONES", "L5")] == pytest.approx(6.0, abs=0.01)
    assert v[("EMISIONES", "M5")] == 800
    assert v[("EMISIONES", "N5")] == 4.6  # plazo inmediato inferior: 5 años
    assert v[("EMISIONES", "O5")] == pytest.approx(390)
    assert v[("EMISIONES", "P5")] == pytest.approx(-410)
    assert v[("EMISIONES", "Q5")] == 2024


def test_emision_en_pesos_no_calcula_spread_contra_el_tesoro(calculado):
    v, _ = calculado
    assert v[("EMISIONES", "O6")] == "n.a."
    assert v[("EMISIONES", "P6")] == "n.a."


def test_composicion_entra_en_el_control_y_en_los_ratios(calculado):
    v, _ = calculado
    assert v[("RATIOS", bloque(indice("ON / deuda financiera"), "F"))] == pytest.approx(500 / 600)
    assert v[("RATIOS", bloque(indice("Caja / vencimientos a 1 año"), "F"))] == pytest.approx(150 / 100)
    assert v[("RATIOS", bloque(indice("Costo medio de la deuda"), "F"))] == pytest.approx(50 / 650)


def test_balance_a_fondo_a_mano(calculado):
    v, _ = calculado
    assert v[("RATIOS", bloque(indice("Ciclo de caja (días)"), "F"))] == pytest.approx((150 / 1000 + 80 / 600 - 120 / 600) * 365)
    assert v[("RATIOS", bloque(indice("Rotación de activos"), "F"))] == pytest.approx(1000 / 1900)
    altman = 3.25 + 6.56 * 100 / 1900 + 3.26 * 300 / 1900 + 6.72 * 300 / 1900 + 1.05 * 900 / 1000
    assert v[("RATIOS", bloque(indice("Altman Z'' mercados emergentes"), "F"))] == pytest.approx(altman)
    assert v[("ESTRUCTURA", "F7")] == pytest.approx(405 / 1000)


def test_fuentes_y_usos_cierra_con_la_variacion_de_caja(calculado):
    v, _ = calculado
    h = "FUENTES Y USOS"
    assert v[(h, "E6")] == 350
    assert v[(h, "E7")] == -45
    assert v[(h, "E10")] == 0
    assert v[(h, "E12")] == 0
    assert sum(v[(h, f"E{f}")] for f in range(6, 12)) == pytest.approx(v[(h, "E12")])


def test_credito_calcula_plazo_y_peso_en_la_deuda(calculado):
    v, _ = calculado
    assert v[("CRÉDITOS", "M5")] == pytest.approx(2.0, abs=0.01)
    assert v[("CRÉDITOS", "N5")] == pytest.approx(65 / 650)


def test_un_descuadre_se_marca(tmp_path):
    vacio = tmp_path / "v.xlsx"
    p.construir(vacio, {}, MERCADO)
    wb = openpyxl.load_workbook(vacio)
    ws = wb["YPF"]
    for clave, valor in EJEMPLO.items():
        ws[f"F{p.FILAS[clave]}"] = valor
    ws[f"F{p.FILAS['patrimonio']}"] = 950
    lleno = tmp_path / "l.xlsx"
    wb.save(lleno)
    assert p.calcular(lleno)[("YPF", f"F{p.FILAS['estado']}")] == "Revisar"


def test_sin_cargar_no_da_errores(calculado):
    _, vacio = calculado
    assert not [k for k, x in vacio.items() if str(x).startswith("#")]
    assert vacio[("YPF", f"F{p.FILAS['estado']}")] == "Sin cargar"
    assert vacio[("RATIOS", bloque(indice("Deuda / EBITDA"), "F", 1))] == "n.d."
