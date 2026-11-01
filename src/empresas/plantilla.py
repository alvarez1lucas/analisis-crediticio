"""Plantilla de análisis crediticio de empresas en Excel.

`python -m empresas.plantilla` escribe excel/analisis_crediticio.xlsx con una hoja de carga por empresa, controles
de cuadre y ratios comparados. Los balances se cargan a mano desde los 20-F: el archivo es la fuente de esos datos,
así que el comando no pisa uno que ya exista salvo que se le pase --forzar.
Usa funciones de Excel 2016.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import formulas
import numpy as np
import xlsxwriter

ROOT = Path(__file__).resolve().parents[2]
ARCHIVO = ROOT / "excel" / "analisis_crediticio.xlsx"
EMPRESAS = ["Pampa", "YPF", "Vista"]
ANIOS = [2021, 2022, 2023, 2024, 2025]
COLUMNAS = "BCDEF"
AZUL, VERDE, GRIS = "#0000FF", "#008000", "#52514E"
ENCABEZADO, LINEA, RESALTE = "#2A78D6", "#E1E0D9", "#FFFF00"
COLOR_EMPRESA = {"Pampa": "#2A78D6", "YPF": "#1BAF7A", "Vista": "#EDA100"}
MIL = '#,##0;(#,##0);"-"'
FILA_INICIO = 7

# (clave, etiqueta, fórmula o None si se carga, ayuda). Clave None: subtítulo.
LINEAS = [
    (None, "Estado de resultados", None, None),
    ("ventas", "Ventas", None, None),
    ("costo", "Costo de ventas", None, "Positivo"),
    ("bruto", "Resultado bruto", "{ventas}-{costo}", None),
    ("comercial", "Gastos de comercialización", None, "Positivo"),
    ("admin", "Gastos de administración", None, "Positivo"),
    ("explor", "Gastos de exploración", None, "Positivo"),
    ("otros_op", "Otros resultados operativos netos", None, "Con signo: ganancia positiva"),
    ("deterioro", "Deterioro de activos", None, "Cargo positivo, recupero negativo"),
    ("ebit", "Resultado operativo", "{bruto}-{comercial}-{admin}-{explor}+{otros_op}-{deterioro}", "Sin asociadas, aunque la empresa las incluya"),
    ("amort", "Depreciación y amortización", None, "Del estado de flujo; ya está dentro de costos y gastos"),
    ("ebitda", "EBITDA", "{ebit}+{amort}+{deterioro}", "Definición propia y común a las tres empresas"),
    ("asociadas", "Resultado de asociadas", None, "Con signo"),
    ("ing_fin", "Ingresos financieros", None, "Positivo"),
    ("gasto_int", "Intereses devengados", None, "Positivo"),
    ("otros_fin", "Otros resultados financieros", None, "Con signo"),
    ("rai", "Resultado antes de impuestos", "{ebit}+{asociadas}+{ing_fin}-{gasto_int}+{otros_fin}", None),
    ("impuesto", "Impuesto a las ganancias", None, "Cargo positivo, beneficio negativo"),
    ("neto", "Resultado neto", "{rai}-{impuesto}", None),
    ("neto_inf", "Resultado neto informado", None, "Copiar del 20-F: es el control"),
    (None, "Balance", None, None),
    ("caja", "Caja y equivalentes", None, None),
    ("inv_cp", "Inversiones corrientes", None, None),
    ("creditos", "Créditos comerciales", None, None),
    ("inventarios", "Inventarios", None, None),
    ("otros_ac", "Otros activos corrientes", None, None),
    ("ac", "Activo corriente", "SUM({caja}:{otros_ac})", None),
    ("ppe", "Propiedad, planta y equipo", None, None),
    ("intang", "Activos intangibles", None, None),
    ("inv_asoc", "Inversiones en asociadas", None, None),
    ("otros_anc", "Otros activos no corrientes", None, None),
    ("activo", "Activo total", "{ac}+SUM({ppe}:{otros_anc})", None),
    ("activo_inf", "Activo total informado", None, "Copiar del 20-F: es el control"),
    ("deuda_cp", "Deuda financiera corriente", None, None),
    ("arr_cp", "Arrendamientos corrientes", None, None),
    ("proveedores", "Cuentas por pagar", None, None),
    ("otros_pc", "Otros pasivos corrientes", None, None),
    ("pc", "Pasivo corriente", "SUM({deuda_cp}:{otros_pc})", None),
    ("deuda_lp", "Deuda financiera no corriente", None, None),
    ("arr_lp", "Arrendamientos no corrientes", None, None),
    ("otros_pnc", "Provisiones y otros no corrientes", None, None),
    ("pasivo", "Pasivo total", "{pc}+SUM({deuda_lp}:{otros_pnc})", None),
    ("patrimonio", "Patrimonio", None, "Incluye participación minoritaria"),
    ("deuda", "Deuda total", "{deuda_cp}+{arr_cp}+{deuda_lp}+{arr_lp}", "Incluye arrendamientos"),
    ("deuda_neta", "Deuda neta", "{deuda}-{caja}-{inv_cp}", None),
    (None, "Flujo de fondos", None, None),
    ("cfo", "Flujo operativo antes de intereses", None, "Si el 20-F ya restó los intereses pagados, volvé a sumarlos"),
    ("capex", "Inversiones en activos fijos", None, "Positivo"),
    ("int_pag", "Intereses pagados", None, "Positivo"),
    ("dividendos", "Dividendos y recompras", None, "Positivo"),
    ("fcf", "Flujo libre", "{cfo}-{int_pag}-{capex}", "Después de intereses y antes de dividendos"),
    (None, "Controles", None, None),
    ("dif_balance", "Activo - pasivo - patrimonio", "{activo}-{pasivo}-{patrimonio}", None),
    ("dif_activo", "Activo calculado - informado", "{activo}-{activo_inf}", None),
    ("dif_neto", "Resultado calculado - informado", "{neto}-{neto_inf}", None),
    ("estado", "Estado", 'IF({ventas}=0,"Sin cargar",IF(AND(ABS({dif_balance})<=$B$4,ABS({dif_activo})<=$B$4,ABS({dif_neto})<=$B$4),"OK","Revisar"))', None),
]
FILAS = {clave: FILA_INICIO + i for i, (clave, *_) in enumerate(LINEAS) if clave}
CARGA = [clave for clave, _, formula, _ in LINEAS if clave and formula is None]

# (etiqueta, fórmula, formato, descripción). {x} es el año, {x_ant} el año anterior; las de saldo anterior no existen en 2021.
X = '0.0"x";(0.0"x");"-"'
P = '0.0%;(0.0%);"-"'
RATIOS = [
    ("Crecimiento de ventas", "{ventas}/{ventas_ant}-1", P, "ventas_ant"),
    ("Margen EBITDA", "{ebitda}/{ventas}", P, "ventas"),
    ("Margen neto", "{neto}/{ventas}", P, "ventas"),
    ("Deuda / EBITDA", "{deuda}/{ebitda}", X, "ebitda"),
    ("Deuda neta / EBITDA", "{deuda_neta}/{ebitda}", X, "ebitda"),
    ("EBITDA / intereses", "{ebitda}/{gasto_int}", X, "gasto_int"),
    ("(EBITDA - capex) / intereses", "({ebitda}-{capex})/{gasto_int}", X, "gasto_int"),
    ("Cobertura de servicio de deuda", "{cfo}/({int_pag}+{deuda_cp_ant}+{arr_cp_ant})", X, "int_pag"),
    ("Liquidez corriente", "{ac}/{pc}", X, "pc"),
    ("Prueba ácida", "({ac}-{inventarios})/{pc}", X, "pc"),
    ("Deuda / patrimonio", "{deuda}/{patrimonio}", X, "patrimonio"),
    ("Flujo libre / deuda", "{fcf}/{deuda}", P, "deuda"),
    ("Rentabilidad sobre patrimonio", "{neto}/{patrimonio}", P, "patrimonio"),
    ("Días de cobro", "{creditos}/{ventas}*365", '0;(0);"-"', "ventas"),
]
FILA_COMPARA = 7
FILA_BLOQUES = FILA_COMPARA + len(RATIOS) + 2


def ref(hoja: str, clave: str, col: str) -> str:
    return f"'{hoja}'!{col}${FILAS[clave]}"


def formula_ratio(hoja: str, plantilla: str, i: int) -> str | None:
    """Fórmula del ratio para el año i, o None si necesita un año anterior que no está."""
    faltan = False

    def valor(m: re.Match) -> str:
        nonlocal faltan
        clave = m.group(1)
        if clave.endswith("_ant"):
            if i == 0:
                faltan = True
                return "0"
            return ref(hoja, clave[:-4], COLUMNAS[i - 1])
        return ref(hoja, clave, COLUMNAS[i])

    f = re.sub(r"\{(\w+)\}", valor, plantilla)
    return None if faltan else f


def construir(archivo: Path, valores: dict) -> None:
    wb = xlsxwriter.Workbook(archivo)
    cache: dict = {}

    def fmt(**k):
        clave = tuple(sorted(k.items()))
        if clave not in cache:
            cache[clave] = wb.add_format({"font_name": "Arial", "font_size": 10, "valign": "vcenter", **k})
        return cache[clave]

    def formula(ws, celda: str, f: str, **k) -> None:
        ws.write_formula(celda, "=" + f, fmt(**k), valores.get((ws.get_name().upper(), celda), 0))

    def nueva_hoja(nombre: str, titulo: str, bajada: str, color: str = ENCABEZADO):
        ws = wb.add_worksheet(nombre)
        ws.hide_gridlines(2)
        ws.set_tab_color(color)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_footer("&L&F&C&A&R&P / &N")
        ws.set_row(0, 28)
        ws.write("A1", titulo, fmt(bold=True, font_size=16))
        ws.write("A2", bajada, fmt(font_color=GRIS))
        return ws

    encabezado = dict(bold=True, font_color="#FFFFFF", bg_color=ENCABEZADO, align="center", border=1, border_color=ENCABEZADO)

    # --- Guía
    g = nueva_hoja("Guía", "Análisis crediticio de empresas", "¿Le prestaría a esta empresa, cuánto y con qué condiciones?")
    g.set_column("A:A", 30)
    g.set_column("B:B", 100)
    fila = 3
    bloques = [("Cómo cargar", [
        ("1. Unidad", "Millones de dólares, tal como reporta cada empresa en su 20-F. Vista informa en miles: dividir por 1.000."),
        ("2. Celdas", "Completá sólo las celdas azules sobre amarillo. Las negras se calculan solas."),
        ("3. Signos", "Costos, gastos e intereses en positivo. La columna Ayuda aclara las líneas con signo."),
        ("4. Agrupar", "Cada línea del 20-F va en una sola línea de acá. Lo que no tenga lugar va en la línea Otros que corresponda."),
        ("5. Fuente", "En la fila Fuente anotá el 20-F y la página. Usá el dato más nuevo si un año fue reexpresado."),
        ("6. Control", "Copiá el activo y el resultado neto informados. Un año está cargado cuando su Estado dice OK."),
    ]), ("Ejemplo de la línea Costo de ventas", [
        ("Dato del 20-F", "Si el informe muestra (1.234,5), cargá 1234,5 en positivo. Es un valor de ejemplo, no un dato de la empresa."),
    ]), ("Definiciones", [
        ("EBITDA", "Resultado operativo sin asociadas + depreciación y amortización + deterioro. No es el EBITDA ajustado que publica cada empresa."),
        ("Deuda", "Deuda financiera corriente y no corriente, con arrendamientos."),
        ("Cobertura de servicio de deuda", "Flujo operativo antes de intereses / (intereses pagados + deuda y arrendamientos corrientes del año anterior)."),
    ])]
    for titulo, filas in bloques:
        g.write(fila, 0, titulo, fmt(bold=True, font_size=12, font_color=ENCABEZADO))
        fila += 1
        for a, b in filas:
            g.write(fila, 0, a, fmt(bold=True, valign="top"))
            g.write(fila, 1, b, fmt(text_wrap=True))
            fila += 1
        fila += 1
    g.write(fila, 0, "Años cargados", fmt(bold=True, font_size=12, font_color=ENCABEZADO))
    for i, e in enumerate(EMPRESAS):
        g.write(fila + 1 + i, 0, e, fmt(bold=True))
        formula(g, f"B{fila + 2 + i}", f"COUNTIF('{e}'!$B${FILAS['estado']}:$F${FILAS['estado']},\"OK\")&\" de {len(ANIOS)} con Estado OK\"", font_color=VERDE)

    # --- Empresas
    celda = dict(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format="#,##0.0;(#,##0.0);0.0")
    fuertes = {"bruto", "ebit", "ebitda", "neto", "ac", "activo", "pc", "pasivo", "deuda", "deuda_neta", "fcf"}
    for e in EMPRESAS:
        h = nueva_hoja(e, f"{e}: carga de estados contables", "Millones de dólares. Completá sólo las celdas azules sobre amarillo.", COLOR_EMPRESA[e])
        h.set_column("A:A", 38)
        h.set_column("B:F", 12)
        h.set_column("G:G", 62)
        h.write("A4", "Tolerancia del control (US$ M)", fmt())
        h.write("B4", 2, fmt(**celda))
        h.write("A5", "Fuente (20-F y página)", fmt())
        h.write_row("B5", [None] * 5, fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, text_wrap=True, font_size=8))
        h.write("A6", "US$ millones", fmt(**encabezado))
        for j, a in enumerate(ANIOS):
            h.write(5, 1 + j, a, fmt(**encabezado))
        h.write("G6", "Ayuda", fmt(**encabezado))
        h.freeze_panes(6, 1)
        for i, (clave, texto, plantilla, ayuda) in enumerate(LINEAS):
            f = FILA_INICIO + i
            if clave is None:
                h.write(f"A{f}", texto, fmt(bold=True, font_color=ENCABEZADO, bottom=1, bottom_color=ENCABEZADO))
                continue
            fuerte = dict(bold=True, top=1, top_color=LINEA) if clave in fuertes else {}
            h.write(f"A{f}", texto, fmt(**fuerte))
            if ayuda:
                h.write(f"G{f}", ayuda, fmt(font_color=GRIS, font_size=8))
            for col in COLUMNAS:
                if plantilla is None:
                    h.write_blank(f"{col}{f}", None, fmt(**celda))
                else:
                    texto_f = re.sub(r"\{(\w+)\}", lambda m: f"{col}{FILAS[m.group(1)]}", plantilla)
                    extra = dict(align="center") if clave == "estado" else dict(num_format=MIL)
                    formula(h, f"{col}{f}", texto_f, **fuerte, **extra)
        rango = f"B{FILAS['estado']}:F{FILAS['estado']}"
        h.conditional_format(rango, {"type": "cell", "criteria": "==", "value": '"OK"', "format": wb.add_format({"font_color": "#006100", "bg_color": "#C6EFCE"})})
        h.conditional_format(rango, {"type": "cell", "criteria": "==", "value": '"Revisar"', "format": wb.add_format({"font_color": "#9C0006", "bg_color": "#FFC7CE"})})

    # --- Ratios
    r = nueva_hoja("Ratios", "Ratios comparados", "Se calculan solos. n.d.: falta el dato. n.s.: no significativo, porque el denominador es cero o negativo.", "#6250D6")
    r.set_column("A:A", 34)
    r.set_column("B:F", 11)
    r.write("A4", "Año a comparar", fmt(bold=True))
    r.write("B4", ANIOS[-1], fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, align="center"))
    r.data_validation("B4", {"validate": "list", "source": ANIOS})
    r.write(FILA_COMPARA - 2, 0, "Comparación del año elegido", fmt(bold=True, font_color=ENCABEZADO))
    r.write(FILA_COMPARA - 1, 0, "Ratio", fmt(**encabezado))
    for k, e in enumerate(EMPRESAS):
        r.write(FILA_COMPARA - 1, 1 + k, e, fmt(**encabezado))
    alto = len(RATIOS) + 3
    inicio = {e: FILA_BLOQUES + 1 + k * alto for k, e in enumerate(EMPRESAS)}
    for k, e in enumerate(EMPRESAS):
        ini = inicio[e]
        r.write(f"A{ini}", e, fmt(bold=True, font_color=ENCABEZADO))
        r.write(f"A{ini + 1}", "Ratio", fmt(**encabezado))
        for j, a in enumerate(ANIOS):
            r.write(f"{COLUMNAS[j]}{ini + 1}", a, fmt(**encabezado))
        for i, (texto, plantilla, formato, denominador) in enumerate(RATIOS):
            f = ini + 2 + i
            r.write(f"A{f}", texto, fmt())
            for j, col in enumerate(COLUMNAS):
                cuerpo = formula_ratio(e, plantilla, j)
                if cuerpo is None:
                    r.write(f"{col}{f}", "n.d.", fmt(font_color=GRIS, align="right"))
                    continue
                den = formula_ratio(e, "{" + denominador + "}", j)
                ventas = ref(e, "ventas", col)
                formula(r, f"{col}{f}", f'IF({ventas}=0,"n.d.",IF({den}<=0,"n.s.",{cuerpo}))', font_color=VERDE, num_format=formato, align="right")
    for i, (texto, _, formato, _) in enumerate(RATIOS):
        f = FILA_COMPARA + 1 + i
        r.write(f"A{f}", texto, fmt())
        for k, e in enumerate(EMPRESAS):
            ini = inicio[e]
            formula(r, f"{'BCD'[k]}{f}", f"INDEX($B${ini + 2 + i}:$F${ini + 2 + i},MATCH($B$4,$B${ini + 1}:$F${ini + 1},0))", num_format=formato, align="right")
    fila_apal = 4  # Deuda neta / EBITDA dentro de RATIOS
    g_ap = wb.add_chart({"type": "line"})
    for e in EMPRESAS:
        ini = inicio[e]
        g_ap.add_series({"name": e, "categories": f"=Ratios!$B${ini + 1}:$F${ini + 1}", "values": f"=Ratios!$B${ini + 2 + fila_apal}:$F${ini + 2 + fila_apal}",
                         "line": {"color": COLOR_EMPRESA[e], "width": 2.25}, "marker": {"type": "circle", "size": 7, "fill": {"color": COLOR_EMPRESA[e]}, "border": {"color": COLOR_EMPRESA[e]}}})
    g_ap.set_title({"name": "Deuda neta / EBITDA", "name_font": {"name": "Arial", "size": 11}})
    g_ap.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    g_ap.set_y_axis({"num_format": '0.0"x"', "major_gridlines": {"visible": True, "line": {"color": LINEA}}, "num_font": {"name": "Arial", "size": 8}})
    g_ap.set_x_axis({"label_position": "low", "num_font": {"name": "Arial", "size": 9}})
    g_ap.set_size({"width": 500, "height": 300})
    r.insert_chart("H6", g_ap)

    wb.set_properties({"title": "Análisis crediticio de empresas", "author": "Lucas Alvarez"})
    wb.close()


def calcular(archivo: Path) -> dict:
    """Calcula el libro con un motor independiente de Excel y devuelve {(HOJA, celda): valor}."""
    solucion = formulas.ExcelModel().loads(str(archivo)).finish().calculate()
    valores = {}
    for clave, rango in solucion.items():
        texto = str(clave)
        if ":" in texto.rsplit("!", 1)[-1] or "]" not in texto or not hasattr(rango, "value"):
            continue
        hoja = texto.split("]", 1)[1].rsplit("'!", 1)[0]
        v = rango.value[0, 0]
        if isinstance(v, np.generic):
            v = v.item()
        valores[(hoja.upper(), texto.rsplit("!", 1)[1])] = v
    return valores


def main() -> None:
    if ARCHIVO.exists() and "--forzar" not in sys.argv:
        raise SystemExit(f"{ARCHIVO.name} ya existe y puede tener datos cargados. Usá --forzar para reemplazarlo.")
    ARCHIVO.parent.mkdir(exist_ok=True)
    construir(ARCHIVO, {})
    valores = calcular(ARCHIVO)
    construir(ARCHIVO, valores)
    print(f"{ARCHIVO.name}: {len(valores)} celdas calculadas")


if __name__ == "__main__":
    main()
