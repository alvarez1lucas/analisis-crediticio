import openpyxl
import pytest

from empresas import plantilla as p

# Empresa inventada que cuadra: sólo prueba las fórmulas, no es un dato real.
EJEMPLO = {"ventas": 1000, "costo": 600, "comercial": 50, "admin": 40, "explor": 10, "otros_op": 5, "deterioro": 5, "amort": 100,
           "asociadas": 10, "ing_fin": 20, "gasto_int": 50, "otros_fin": -10, "impuesto": 60, "neto_inf": 210,
           "caja": 100, "inv_cp": 50, "creditos": 150, "inventarios": 80, "otros_ac": 20, "ppe": 1200, "intang": 100, "inv_asoc": 150, "otros_anc": 50,
           "activo_inf": 1900, "deuda_cp": 100, "arr_cp": 10, "proveedores": 120, "otros_pc": 70, "deuda_lp": 500, "arr_lp": 40, "otros_pnc": 160,
           "patrimonio": 900, "cfo": 350, "capex": 200, "int_pag": 45, "dividendos": 30}


@pytest.fixture(scope="module")
def calculado(tmp_path_factory):
    base = tmp_path_factory.mktemp("xl")
    vacio = base / "vacio.xlsx"
    p.construir(vacio, {})
    wb = openpyxl.load_workbook(vacio)
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


def test_un_descuadre_se_marca(tmp_path):
    vacio = tmp_path / "v.xlsx"
    p.construir(vacio, {})
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
