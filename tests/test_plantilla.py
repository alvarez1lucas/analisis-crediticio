from datetime import datetime

import openpyxl
import pandas as pd
import pytest

from empresas import plantilla as p

# Empresa inventada que cuadra: sólo prueba las fórmulas, no es un dato real.
EJEMPLO = {"ventas": 1000, "costo": 600, "comercial": 50, "admin": 40, "explor": 10, "otros_op": 5, "deterioro": 5, "no_rec": 0, "amort": 100,
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
    assert v[("RATIOS", bloque(indice("Deuda / EBITDA"), "D"))] == pytest.approx(deuda / 405)
    assert v[("RATIOS", bloque(indice("Deuda neta / EBITDA"), "D"))] == pytest.approx(neta / 405)
    assert v[("RATIOS", bloque(indice("EBITDA / intereses"), "D"))] == pytest.approx(405 / 50)
    assert v[("RATIOS", bloque(indice("Cobertura de servicio de deuda"), "D"))] == pytest.approx(350 / (45 + 110))
    assert v[("RATIOS", bloque(indice("Flujo libre / deuda"), "D"))] == pytest.approx((350 - 45 - 200) / 650)
    assert v[("RATIOS", bloque(indice("Liquidez corriente"), "D"))] == pytest.approx(400 / 300)


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
    assert v[("RATIOS", bloque(indice("ON / deuda financiera"), "D"))] == pytest.approx(500 / 600)
    assert v[("RATIOS", bloque(indice("Caja / vencimientos a 1 año"), "D"))] == pytest.approx(150 / 100)
    assert v[("RATIOS", bloque(indice("Costo medio de la deuda"), "D"))] == pytest.approx(50 / 650)


def test_balance_a_fondo_a_mano(calculado):
    v, _ = calculado
    assert v[("RATIOS", bloque(indice("Ciclo de caja (días)"), "D"))] == pytest.approx((150 / 1000 + 80 / 600 - 120 / 600) * 365)
    assert v[("RATIOS", bloque(indice("Rotación de activos"), "D"))] == pytest.approx(1000 / 1900)
    altman = 3.25 + 6.56 * 100 / 1900 + 3.26 * 300 / 1900 + 6.72 * 300 / 1900 + 1.05 * 900 / 1000
    assert v[("RATIOS", bloque(indice("Altman Z'' mercados emergentes"), "D"))] == pytest.approx(altman)
    assert v[("ESTRUCTURA", "D7")] == pytest.approx(405 / 1000)


def test_fuentes_y_usos_cierra_con_la_variacion_de_caja(calculado):
    v, _ = calculado
    h = "FUENTES Y USOS"
    assert v[(h, "C6")] == 350
    assert v[(h, "C7")] == -45
    assert v[(h, "C10")] == 0
    assert v[(h, "C12")] == 0
    assert sum(v[(h, f"C{f}")] for f in range(6, 12)) == pytest.approx(v[(h, "C12")])


def test_credito_calcula_plazo_y_peso_en_la_deuda(calculado):
    v, _ = calculado
    assert v[("CRÉDITOS", "M5")] == pytest.approx(2.0, abs=0.01)
    assert v[("CRÉDITOS", "N5")] == pytest.approx(65 / 650)


def test_el_ebitda_excluye_lo_no_recurrente_pero_el_resultado_lo_conserva(tmp_path):
    vacio = tmp_path / "v.xlsx"
    p.construir(vacio, {}, MERCADO, {("Vista", 2025, clave): float(valor) for clave, valor in {**EJEMPLO, "no_rec": 50}.items()})
    v = p.calcular(vacio)
    assert v[("VISTA", f"D{p.FILAS['ebitda']}")] == 355
    assert v[("VISTA", f"D{p.FILAS['ebit']}")] == 300


def test_vista_cierra_contra_sus_estados_contables():
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    h = wb["Vista"]
    assert [h[f"{c}{p.FILAS['estado']}"].value for c in "BCD"] == ["OK"] * 3
    informado = {"B": (631.487, 396.955), "C": (625.390, 477.521), "D": (1250.830, 719.063)}
    for col, (ebit, neto) in informado.items():
        assert h[f"{col}{p.FILAS['ebit']}"].value == pytest.approx(ebit, abs=0.002)
        assert h[f"{col}{p.FILAS['neto']}"].value == pytest.approx(neto, abs=0.002)
    assert h[f"D{p.FILAS['ebitda']}"].value == pytest.approx(1537.455, abs=0.002)


def test_pampa_cierra_contra_su_20f():
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    h = wb["Pampa"]
    assert [h[f"{c}{p.FILAS['estado']}"].value for c in "BCD"] == ["OK"] * 3
    for col, (operativo, neto, ac) in {"B": (424, 305, 1335), "C": (440, 619, 2380), "D": (503, 378, 1988)}.items():
        assert h[f"{col}{p.FILAS['ebit']}"].value + h[f"{col}{p.FILAS['asociadas']}"].value == operativo
        assert h[f"{col}{p.FILAS['neto']}"].value == neto
        assert h[f"{col}{p.FILAS['ac']}"].value == ac


def test_ypf_cierra_contra_su_20f():
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    h = wb["YPF"]
    assert [h[f"{c}{p.FILAS['estado']}"].value for c in "BCD"] == ["OK"] * 3
    for col, (operativo, neto, deuda) in {"B": (-1248, -1277, 8190), "C": (1480, 2393, 8942), "D": (1740, -799, 10581)}.items():
        assert h[f"{col}{p.FILAS['ebit']}"].value == operativo
        assert h[f"{col}{p.FILAS['neto']}"].value == neto
        assert h[f"{col}{p.FILAS['deuda_cp']}"].value + h[f"{col}{p.FILAS['deuda_lp']}"].value == deuda


def test_las_on_de_ypf_suman_su_nota_22():
    e = pd.read_csv(p.EMISIONES, encoding="utf-8")
    anteriores_a_2021 = 15 + 648 + 19 + 530 + 2 + 397  # 1998, LIII, LIV y Clase I: fuera del período con riesgo país
    assert e[e.empresa == "YPF"].saldo_2025.sum() + anteriores_a_2021 == pytest.approx(8952)
    c = pd.read_csv(p.CREDITOS, encoding="utf-8")
    assert c[c.empresa == "YPF"].saldo.sum() == pytest.approx(10581 - 8952)


def test_el_detalle_de_vista_suma_el_total_de_deuda():
    e = pd.read_csv(p.EMISIONES, encoding="utf-8")
    c = pd.read_csv(p.CREDITOS, encoding="utf-8")
    assert e[e.empresa == "Vista"].saldo_2025.sum() == pytest.approx(2440.344, abs=0.002)
    assert c[c.empresa == "Vista"].saldo.sum() == pytest.approx(713.733, abs=0.002)


def test_el_rescate_guarda_lo_cargado_a_mano(tmp_path, monkeypatch):
    archivo, balances = tmp_path / "a.xlsx", tmp_path / "b.csv"
    pd.DataFrame([("Vista", 2025, "ventas", 7)], columns=["empresa", "anio", "clave", "valor"]).to_csv(balances, index=False)
    p.construir(archivo, {}, MERCADO)
    wb = openpyxl.load_workbook(archivo)
    wb["Pampa"][f"D{p.FILAS['ventas']}"] = 123
    wb.save(archivo)
    monkeypatch.setattr(p, "ARCHIVO", archivo)
    monkeypatch.setattr(p, "BALANCES", balances)
    monkeypatch.setattr(p, "CARGA_MANUAL", ["Pampa"])
    p.rescatar_manual()
    d = p.leer_balances()
    assert d[("Pampa", 2025, "ventas")] == 123
    assert d[("Vista", 2025, "ventas")] == 7


def test_emision_se_compara_con_el_credito_del_mes(tmp_path):
    archivo = tmp_path / "t.xlsx"
    tasas = pd.DataFrame({"mes": ["2024-01-01"], "doc_usd": [7.0], "doc_usd_corto": [6.0], "doc_usd_largo": [7.5], "doc_ars": [100.0], "pf_usd": [1.0], "pf_ars": [90.0]})
    p.construir(archivo, {}, MERCADO, tasas=tasas)
    wb = openpyxl.load_workbook(archivo)
    for col, valor in dict(A="Vista", C=datetime(2024, 1, 3), D=datetime(2030, 1, 3), E="Dólar duro", H=0.08).items():
        wb["Emisiones"][f"{col}5"] = valor
    wb.save(archivo)
    v = p.calcular(archivo)
    assert v[("EMISIONES", "R5")] == 7.5
    assert v[("EMISIONES", "S5")] == pytest.approx(50)
    assert v[("TASAS", "H5")] == 6


def test_las_tasas_no_cruzan_dolares_con_pesos():
    t = pd.read_csv(p.TASAS_SISTEMA)
    assert t.mes.iloc[0] == "2020-01-01"
    assert t.doc_usd.median() < 15 < t.doc_ars.median()
    assert t.pf_usd.median() < 5 < t.pf_ars.median()


def test_proyeccion_repite_la_mecanica_a_mano():
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    pr, h = wb["Proyección"], wb["Pampa"]
    ventas, ebitda, deuda, amort = (h[f"D{p.FILAS[k]}"].value for k in ("ventas", "ebitda", "deuda", "amort"))
    caja = h[f"D{p.FILAS['caja']}"].value + h[f"D{p.FILAS['inv_cp']}"].value
    conv = sum(h[f"{c}{p.FILAS['cfo']}"].value for c in "BCD") / sum(h[f"{c}{p.FILAS['ebitda']}"].value for c in "BCD")
    costo = pr["B16"].value
    s = 20  # Pampa, primer bloque
    v_ventas = ventas * (1 - 0.20)
    v_ebitda = v_ventas * (ebitda / ventas - 0.05)
    interes = deuda * costo
    necesidad = amort + interes - v_ebitda * conv
    nueva_deuda = deuda + necesidad
    assert pr[f"G{s + 2}"].value == pytest.approx(v_ventas)
    assert pr[f"G{s + 3}"].value == pytest.approx(v_ebitda)
    assert pr[f"G{s + 10}"].value == pytest.approx(nueva_deuda)
    assert pr[f"G{s + 15}"].value == pytest.approx((nueva_deuda - caja) / v_ebitda)


def test_rating_y_covenants_dan_resultado_para_las_tres():
    wb = openpyxl.load_workbook(p.ARCHIVO, data_only=True)
    rt, cv = wb["Rating"], wb["Covenants"]
    fila_nota = 7 + len(p.FACTORES)
    notas = [rt[f"{c}{fila_nota}"].value for c in "HIJKLM"]
    assert all(n in ("A", "BBB", "BB", "B", "CCC") for n in notas)
    assert rt[f"B{fila_nota - 1}"].value == pytest.approx(1)
    estados = [cv[f"G{f}"].value for f in range(8, 22) if cv[f"G{f}"].value]
    assert len(estados) == 12 and set(estados) <= {"Cumple", "Incumple"}


def test_un_descuadre_se_marca(tmp_path):
    vacio = tmp_path / "v.xlsx"
    p.construir(vacio, {}, MERCADO)
    wb = openpyxl.load_workbook(vacio)
    ws = wb["YPF"]
    for clave, valor in EJEMPLO.items():
        ws[f"D{p.FILAS[clave]}"] = valor
    ws[f"D{p.FILAS['patrimonio']}"] = 950
    lleno = tmp_path / "l.xlsx"
    wb.save(lleno)
    assert p.calcular(lleno)[("YPF", f"D{p.FILAS['estado']}")] == "Revisar"


def test_sin_cargar_no_da_errores(calculado):
    _, vacio = calculado
    assert not [k for k, x in vacio.items() if str(x).startswith("#")]
    assert vacio[("YPF", f"D{p.FILAS['estado']}")] == "Sin cargar"
    assert vacio[("RATIOS", bloque(indice("Deuda / EBITDA"), "D", 1))] == "n.d."


def test_el_memo_sale_del_excel_y_no_deja_huecos():
    from empresas import memo
    texto = memo.escribir(memo.leer())
    assert "None" not in texto and "n.d." not in texto
    assert "Rating interno **" in texto and "zona gris" in texto
    assert memo.zona_altman(4.8) == "gris" and memo.zona_altman(7.6) == "segura"
