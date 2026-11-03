"""Plantilla de análisis crediticio de empresas en Excel.

`python -m empresas.plantilla` escribe excel/analisis_crediticio.xlsx con una hoja de carga por empresa, controles
de cuadre y ratios comparados. Los balances se cargan a mano desde los 20-F: el archivo es la fuente de esos datos,
así que el comando rescata primero lo cargado en las hojas manuales y lo guarda en data/balances.csv.
Usa funciones de Excel 2016.
"""

from __future__ import annotations

import re
from pathlib import Path

import formulas
import numpy as np
import openpyxl
import pandas as pd
import xlsxwriter

ROOT = Path(__file__).resolve().parents[2]
ARCHIVO = ROOT / "excel" / "analisis_crediticio.xlsx"
MERCADO = ROOT / "data" / "mercado.csv"
BALANCES = ROOT / "data" / "balances.csv"
EMISIONES = ROOT / "data" / "emisiones.csv"
CREDITOS = ROOT / "data" / "creditos.csv"
TASAS_SISTEMA = ROOT / "data" / "tasas.csv"
VENCIMIENTOS = ROOT / "data" / "deuda_vigente.csv"
CARGA_MANUAL: list[str] = []  # hojas que se cargan a mano en el Excel; hoy todas salen de data/balances.csv
PLAZOS_TESORO = [1, 2, 3, 5, 7, 10]
INSTRUMENTOS = ["Dólar duro", "Dólar linked", "Pesos tasa variable", "Pesos tasa fija", "UVA"]
LEYES = ["Nueva York", "Argentina"]
TIPOS_CREDITO = ["Préstamo bancario", "Prefinanciación", "Sindicado", "Pagaré", "Otro"]
MONEDAS = ["Dólar", "Dólar linked", "Pesos"]
TASAS = ["Fija", "Variable"]
REFERENCIAS = ["SOFR", "BADLAR", "TAMAR", "Otra", "Ninguna"]
FILAS_CREDITOS = 100
# Proyección. SUPUESTOS: vida media ponderada por saldo al 31-12-2025 (Emisiones, Créditos y nota 12.5.2 de Pampa) y rendimiento
# de la última ON larga en dólar duro de cada empresa (Pampa Clase 26, YPF Clase XXXIX, Vista ON XXIX).
VIDA_MEDIA = {"Pampa": 7.7, "YPF": 4.8, "Vista": 5.6}
TASA_NUEVA = {"Pampa": 0.08125, "YPF": 0.0875, "Vista": 0.085}
PROYECCION = [2026, 2027, 2028]
# Rating interno: (factor, clave de ratio o línea, peso, límites, mayor es mejor)
FACTORES = [("Deuda neta / EBITDA", "Deuda neta / EBITDA", 0.25, (1, 2, 3, 4), False),
            ("EBITDA / intereses", "EBITDA / intereses", 0.20, (1.5, 3, 5, 8), True),
            ("Liquidez corriente", "Liquidez corriente", 0.15, (0.75, 1, 1.5, 2), True),
            ("Flujo libre / deuda", "Flujo libre / deuda", 0.15, (-0.1, 0, 0.05, 0.15), True),
            ("Margen EBITDA", "Margen EBITDA", 0.10, (0.1, 0.2, 0.3, 0.4), True),
            ("Escala: EBITDA (US$ M)", "ebitda", 0.15, (200, 500, 1000, 3000), True)]
ESCALA = [(4.25, "A"), (3.5, "BBB"), (2.75, "BB"), (2.0, "B")]  # puntaje mínimo de cada nota; debajo, CCC
FILAS_EMISIONES = 150
EMPRESAS = ["Pampa", "YPF", "Vista"]
ANIOS = [2023, 2024, 2025]  # los tres años que los 20-F 2025 informan en dólares y auditados
COLUMNAS = "BCD"
AYUDA = "E"
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
    ("no_rec", "Resultados no recurrentes", None, "Ya incluidos en las líneas de arriba. Ganancia positiva. Sólo se restan del EBITDA"),
    ("ebit", "Resultado operativo", "{bruto}-{comercial}-{admin}-{explor}+{otros_op}-{deterioro}", "Sin asociadas, aunque la empresa las incluya"),
    ("amort", "Depreciación y amortización", None, "Del estado de flujo; ya está dentro de costos y gastos"),
    ("ebitda", "EBITDA", "{ebit}+{amort}+{deterioro}-{no_rec}", "Definición propia y común a las tres empresas"),
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
    ("res_acum", "Resultados acumulados", None, "Resultados no asignados + reservas legal y voluntaria, para Altman. Con signo"),
    ("deuda", "Deuda total", "{deuda_cp}+{arr_cp}+{deuda_lp}+{arr_lp}", "Incluye arrendamientos"),
    ("deuda_neta", "Deuda neta", "{deuda}-{caja}-{inv_cp}", None),
    (None, "Composición de la deuda financiera", None, None),
    ("d_on", "Obligaciones negociables", None, "Sin arrendamientos, como todo este bloque"),
    ("d_bancos", "Préstamos bancarios", None, None),
    ("d_prefin", "Prefinanciación de exportaciones", None, None),
    ("d_otra", "Otra deuda financiera", None, None),
    ("m_dolar", "En dólares", None, "Total por moneda"),
    ("m_linked", "Dólar linked", None, None),
    ("m_pesos", "En pesos", None, "Incluye UVA"),
    ("m_otra", "Otras monedas", None, None),
    ("v_1", "Vence en menos de 1 año", None, "Total por vencimiento"),
    ("v_2", "Vence de 1 a 2 años", None, None),
    ("v_3", "Vence de 2 a 5 años", None, None),
    ("v_mas", "Vence en más de 5 años", None, None),
    ("tasa_fija", "Deuda a tasa fija", None, None),
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
    ("dif_tipo", "Deuda por tipo - balance", "SUM({d_on}:{d_otra})-{deuda_cp}-{deuda_lp}", None),
    ("dif_moneda", "Deuda por moneda - balance", "SUM({m_dolar}:{m_otra})-{deuda_cp}-{deuda_lp}", None),
    ("dif_vto", "Deuda por vencimiento - balance", "SUM({v_1}:{v_mas})-{deuda_cp}-{deuda_lp}", "Puede diferir si el 20-F abre vencimientos contractuales con intereses: anotalo en Fuente"),
    ("estado", "Estado", 'IF({ventas}=0,"Sin cargar",IF(AND(ABS({dif_balance})<=$B$4,ABS({dif_activo})<=$B$4,ABS({dif_neto})<=$B$4,OR(SUM({d_on}:{d_otra})=0,ABS({dif_tipo})<=$B$4),OR(SUM({m_dolar}:{m_otra})=0,ABS({dif_moneda})<=$B$4),OR(SUM({v_1}:{v_mas})=0,ABS({dif_vto})<=$B$4)),"OK","Revisar"))', "Los bloques de deuda vacíos no se controlan: el detalle no siempre se publica"),
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
    ("ON / deuda financiera", "{d_on}/({deuda_cp}+{deuda_lp})", P, "deuda"),
    ("Deuda en dólares / deuda financiera", "{m_dolar}/({deuda_cp}+{deuda_lp})", P, "deuda"),
    ("Tasa fija / deuda financiera", "{tasa_fija}/({deuda_cp}+{deuda_lp})", P, "deuda"),
    ("Caja / vencimientos a 1 año", "({caja}+{inv_cp})/{v_1}", X, "v_1"),
    ("Costo medio de la deuda", "{gasto_int}/(({deuda}+{deuda_ant})/2)", P, "deuda_ant"),
    ("Días de cobro", "{creditos}/{ventas}*365", '0;(0);"-"', "ventas"),
    ("Días de inventario", "{inventarios}/{costo}*365", '0;(0);"-"', "costo"),
    ("Días de pago", "{proveedores}/{costo}*365", '0;(0);"-"', "costo"),
    ("Ciclo de caja (días)", "({creditos}/{ventas}+{inventarios}/{costo}-{proveedores}/{costo})*365", '0;(0);0', "costo"),
    ("Rotación de activos", "{ventas}/{activo}", '0.00"x"', "activo"),
    ("Activo / patrimonio", "{activo}/{patrimonio}", X, "patrimonio"),
    ("Capex / ventas", "{capex}/{ventas}", P, "ventas"),
    ("Capex / depreciación", "{capex}/{amort}", X, "amort"),
    ("Altman Z'' mercados emergentes", "3.25+6.56*({ac}-{pc})/{activo}+3.26*{res_acum}/{activo}+6.72*{ebit}/{activo}+1.05*{patrimonio}/{pasivo}", "0.00;(0.00)", "activo"),
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


def construir(archivo: Path, valores: dict, mercado: pd.DataFrame, balances: dict | None = None, emisiones: pd.DataFrame | None = None, creditos: pd.DataFrame | None = None, tasas: pd.DataFrame | None = None, vencimientos: pd.DataFrame | None = None) -> None:
    balances = balances or {}
    tasas = pd.DataFrame(columns=["mes", "doc_usd", "doc_usd_corto", "doc_usd_largo", "doc_ars", "pf_usd", "pf_ars"]) if tasas is None else tasas
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

    def poner(ws, celda: str, valor, formato) -> None:
        if valor is None or (isinstance(valor, float) and np.isnan(valor)):
            ws.write_blank(celda, None, formato)
        elif isinstance(valor, pd.Timestamp):
            ws.write_datetime(celda, valor.to_pydatetime(), formato)
        else:
            ws.write(celda, valor, formato)

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
        ("7. Emisiones", "En Emisiones cargá una fila por ON o reapertura, desde el aviso de resultados. Las columnas de la derecha se calculan solas."),
        ("8. Créditos", "En Créditos cargá una fila por préstamo o prefinanciación que detalle el 20-F, con su saldo al cierre."),
    ]), ("Ejemplo de la línea Costo de ventas", [
        ("Dato del 20-F", "Si el informe muestra (1.234,5), cargá 1234,5 en positivo. Es un valor de ejemplo, no un dato de la empresa."),
        ("Fila de Emisiones", "Empresa, clase, fecha de emisión, vencimiento, instrumento, ley, monto en US$ M, cupón, rendimiento, colocadores y fuente. Si se colocó a la par, el rendimiento es el cupón."),
    ]), ("Definiciones", [
        ("EBITDA", "Resultado operativo sin asociadas + depreciación y amortización + deterioro - resultados no recurrentes. No es el EBITDA ajustado que publica cada empresa."),
        ("Deuda", "Deuda financiera corriente y no corriente, con arrendamientos."),
        ("Cobertura de servicio de deuda", "Flujo operativo antes de intereses / (intereses pagados + deuda y arrendamientos corrientes del año anterior)."),
        ("Altman Z'' EM", "Versión para mercados emergentes: 3,25 + 6,56 capital de trabajo/activo + 3,26 resultados acumulados/activo + 6,72 resultado operativo/activo + 1,05 patrimonio/pasivo. Es una referencia, no un rating."),
        ("Spread", "Rendimiento de la emisión menos la tasa del Tesoro de EE.UU. del plazo inmediato inferior, en puntos básicos. Sólo para dólar duro."),
        ("Prima contra riesgo país", "Spread menos riesgo país de ese día. Negativa: la empresa se financió más barato que el Estado."),
        ("Crédito del sistema", "Tasa de documentos a sola firma en dólares a 90 días o más, de empresas que no son pymes. Promedio del mes. No incluye tasa variable."),
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
        formula(g, f"B{fila + 2 + i}", f"COUNTIF('{e}'!$B${FILAS['estado']}:${COLUMNAS[-1]}${FILAS['estado']},\"OK\")&\" de {len(ANIOS)} con Estado OK\"", font_color=VERDE)

    # --- Empresas
    celda = dict(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format="#,##0.0;(#,##0.0);0.0")
    fuertes = {"bruto", "ebit", "ebitda", "neto", "ac", "activo", "pc", "pasivo", "deuda", "deuda_neta", "fcf"}
    for e in EMPRESAS:
        h = nueva_hoja(e, f"{e}: carga de estados contables", "Millones de dólares. Completá sólo las celdas azules sobre amarillo.", COLOR_EMPRESA[e])
        h.set_column("A:A", 38)
        h.set_column(f"B:{COLUMNAS[-1]}", 12)
        h.set_column(f"{AYUDA}:{AYUDA}", 62)
        h.write("A4", "Tolerancia del control (US$ M)", fmt())
        h.write("B4", 2, fmt(**celda))
        h.write("A5", "Fuente (20-F y página)", fmt())
        h.set_row(4, 34)
        for col, a in zip(COLUMNAS, ANIOS):
            poner(h, f"{col}5", balances.get((e, a, "fuente")), fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, text_wrap=True, font_size=8))
        h.write("A6", "US$ millones", fmt(**encabezado))
        for j, a in enumerate(ANIOS):
            h.write(5, 1 + j, a, fmt(**encabezado))
        h.write(f"{AYUDA}6", "Ayuda", fmt(**encabezado))
        h.freeze_panes(6, 1)
        for i, (clave, texto, plantilla, ayuda) in enumerate(LINEAS):
            f = FILA_INICIO + i
            if clave is None:
                h.write(f"A{f}", texto, fmt(bold=True, font_color=ENCABEZADO, bottom=1, bottom_color=ENCABEZADO))
                continue
            fuerte = dict(bold=True, top=1, top_color=LINEA) if clave in fuertes else {}
            h.write(f"A{f}", texto, fmt(**fuerte))
            if ayuda:
                h.write(f"{AYUDA}{f}", ayuda, fmt(font_color=GRIS, font_size=8))
            for col, a in zip(COLUMNAS, ANIOS):
                if plantilla is None:
                    poner(h, f"{col}{f}", balances.get((e, a, clave)), fmt(**celda))
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
                falta = f"OR({ventas}=0,{ref(e, 'ventas', COLUMNAS[j - 1])}=0)" if "_ant" in plantilla else f"{ventas}=0"
                formula(r, f"{col}{f}", f'IF({falta},"n.d.",IF({den}<=0,"n.s.",{cuerpo}))', font_color=VERDE, num_format=formato, align="right")
    for i, (texto, _, formato, _) in enumerate(RATIOS):
        f = FILA_COMPARA + 1 + i
        r.write(f"A{f}", texto, fmt())
        for k, e in enumerate(EMPRESAS):
            ini = inicio[e]
            formula(r, f"{'BCD'[k]}{f}", f"INDEX($B${ini + 2 + i}:${COLUMNAS[-1]}${ini + 2 + i},MATCH($B$4,$B${ini + 1}:${COLUMNAS[-1]}${ini + 1},0))", num_format=formato, align="right")
    fila_apal = 4  # Deuda neta / EBITDA dentro de RATIOS
    g_ap = wb.add_chart({"type": "line"})
    for e in EMPRESAS:
        ini = inicio[e]
        g_ap.add_series({"name": e, "categories": f"=Ratios!$B${ini + 1}:${COLUMNAS[-1]}${ini + 1}", "values": f"=Ratios!$B${ini + 2 + fila_apal}:${COLUMNAS[-1]}${ini + 2 + fila_apal}",
                         "line": {"color": COLOR_EMPRESA[e], "width": 2.25}, "marker": {"type": "circle", "size": 7, "fill": {"color": COLOR_EMPRESA[e]}, "border": {"color": COLOR_EMPRESA[e]}}})
    g_ap.set_title({"name": "Deuda neta / EBITDA", "name_font": {"name": "Arial", "size": 11}})
    g_ap.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    g_ap.set_y_axis({"num_format": '0.0"x"', "major_gridlines": {"visible": True, "line": {"color": LINEA}}, "num_font": {"name": "Arial", "size": 8}})
    g_ap.set_x_axis({"label_position": "low", "num_font": {"name": "Arial", "size": 9}})
    g_ap.set_size({"width": 500, "height": 300})
    r.insert_chart("H6", g_ap)

    # --- Estructura
    v = nueva_hoja("Estructura", "Análisis vertical", "Cada línea como porcentaje de las ventas o del activo.", "#6250D6")
    v.set_column("A:A", 34)
    v.set_column("B:F", 11)
    vertical = [("Costo de ventas / ventas", "costo", "ventas"), ("EBITDA / ventas", "ebitda", "ventas"), ("Resultado operativo / ventas", "ebit", "ventas"), ("Resultado neto / ventas", "neto", "ventas"),
                ("Activo corriente / activo", "ac", "activo"), ("Propiedad, planta y equipo / activo", "ppe", "activo"), ("Deuda total / activo", "deuda", "activo"),
                ("Pasivo total / activo", "pasivo", "activo"), ("Patrimonio / activo", "patrimonio", "activo")]
    for k, e in enumerate(EMPRESAS):
        ini = 4 + k * (len(vertical) + 3)
        v.write(f"A{ini}", e, fmt(bold=True, font_color=ENCABEZADO))
        v.write(f"A{ini + 1}", "% de", fmt(**encabezado))
        for j, a in enumerate(ANIOS):
            v.write(f"{COLUMNAS[j]}{ini + 1}", a, fmt(**encabezado))
        for i, (texto, num, den) in enumerate(vertical):
            f = ini + 2 + i
            v.write(f"A{f}", texto, fmt())
            for col in COLUMNAS:
                formula(v, f"{col}{f}", f'IF({ref(e, den, col)}=0,"n.d.",{ref(e, num, col)}/{ref(e, den, col)})', font_color=VERDE, num_format=P, align="right")

    # --- Fuentes y usos
    u = nueva_hoja("Fuentes y usos", "Fuentes y usos de fondos", "De dónde salió la plata cada año y en qué se usó. US$ millones.", "#6250D6")
    u.set_column("A:A", 38)
    u.set_column(f"B:{COLUMNAS[-2]}", 11)
    u.set_column(f"{COLUMNAS[-1]}:{COLUMNAS[-1]}", 60)
    filas_u = [("Flujo operativo antes de intereses", lambda e, c, a: ref(e, "cfo", c), None), ("Intereses pagados", lambda e, c, a: "-" + ref(e, "int_pag", c), None),
               ("Capex", lambda e, c, a: "-" + ref(e, "capex", c), None), ("Dividendos y recompras", lambda e, c, a: "-" + ref(e, "dividendos", c), None),
               ("Variación de deuda", lambda e, c, a: f"{ref(e, 'deuda', c)}-{ref(e, 'deuda', a)}", "Del balance: incluye efectos sin movimiento de caja"),
               ("Otros", None, "Diferencia hasta la variación de caja"),
               ("Variación de caja e inversiones", lambda e, c, a: f"{ref(e, 'caja', c)}+{ref(e, 'inv_cp', c)}-{ref(e, 'caja', a)}-{ref(e, 'inv_cp', a)}", None)]
    for k, e in enumerate(EMPRESAS):
        ini = 4 + k * (len(filas_u) + 3)
        u.write(f"A{ini}", e, fmt(bold=True, font_color=ENCABEZADO))
        u.write(f"A{ini + 1}", "US$ millones", fmt(**encabezado))
        for j, a in enumerate(ANIOS[1:]):
            u.write(f"{COLUMNAS[j]}{ini + 1}", a, fmt(**encabezado))
        for i, (texto, calculo, nota) in enumerate(filas_u):
            f = ini + 2 + i
            u.write(f"A{f}", texto, fmt(**(dict(bold=True, top=1, top_color=LINEA) if i == len(filas_u) - 1 else {})))
            if nota:
                u.write(f"{COLUMNAS[-1]}{f}", nota, fmt(font_color=GRIS, font_size=8))
            for j, col in enumerate(COLUMNAS[:-1]):
                actual, anterior = COLUMNAS[j + 1], COLUMNAS[j]
                cuerpo = f"{col}{ini + 8}-SUM({col}{ini + 2}:{col}{ini + 6})" if calculo is None else calculo(e, actual, anterior)
                cargado = f"OR({ref(e, 'ventas', actual)}=0,{ref(e, 'ventas', anterior)}=0)"
                extra = dict(font_color=VERDE) if calculo else {}
                formula(u, f"{col}{f}", f'IF({cargado},"n.d.",{cuerpo})', num_format=MIL, align="right", **extra)

    # --- Vencimientos
    venc = pd.DataFrame(columns=["empresa", "tipo", "instrumento", "moneda", "anio", "monto", "supuesto"]) if vencimientos is None else vencimientos
    vt = nueva_hoja("Vencimientos", "Muro de vencimientos", "Deuda financiera al 31-12-2025 por año de vencimiento, sin arrendamientos. US$ millones. Columnas A a G: datos de las notas de deuda.", "#E34948")
    vt.set_column("A:B", 14)
    vt.set_column("C:C", 28)
    vt.set_column("D:F", 11)
    vt.set_column("G:G", 40)
    vt.set_column("I:L", 12)
    vt.write_row(3, 0, ["Empresa", "Tipo", "Instrumento", "Moneda", "Año", "US$ M", "Supuesto"], fmt(**encabezado))
    for i, fila_v in enumerate(venc.itertuples(index=False)):
        f = 5 + i
        for c, valor in zip("ABCDEFG", (fila_v.empresa, fila_v.tipo, fila_v.instrumento, fila_v.moneda, int(fila_v.anio), float(fila_v.monto), fila_v.supuesto)):
            poner(vt, f"{c}{f}", valor, fmt(font_color=AZUL, num_format="#,##0.0" if c == "F" else "General", font_size=9 if c == "G" else 10))
    fin_v = 4 + max(len(venc), 1)
    rango_v = lambda c: f"${c}$5:${c}${fin_v}"
    vt.write_row(3, 8, ["Año"] + EMPRESAS, fmt(**encabezado))
    fila_venc = {}
    anios_v = list(range(PROYECCION[0], PROYECCION[0] + 10))
    for j, a in enumerate(anios_v + ["Después"]):
        f = 5 + j
        fila_venc[a] = f
        vt.write(f"I{f}", a if isinstance(a, int) else f"{anios_v[-1] + 1} y después", fmt(align="left"))
        for k, e in enumerate(EMPRESAS):
            criterio = f'">={anios_v[-1] + 1}"' if isinstance(a, str) else str(a)
            formula(vt, f"{'JKL'[k]}{f}", f'SUMIFS({rango_v("F")},{rango_v("A")},"{e}",{rango_v("E")},{criterio})', num_format=MIL, font_color="#0B0B0B")
    f_total = 5 + len(anios_v) + 1
    vt.write(f"I{f_total}", "Total", fmt(bold=True, top=1))
    vt.write(f"I{f_total + 1}", "Deuda financiera del balance", fmt())
    vt.write(f"I{f_total + 2}", "Caja e inversiones 2025", fmt())
    vt.write(f"I{f_total + 3}", f"Vence hasta {PROYECCION[-1]} / deuda", fmt())
    vt.write(f"I{f_total + 4}", f"Caja / vencimientos {PROYECCION[0]}", fmt())
    for k, e in enumerate(EMPRESAS):
        c = "JKL"[k]
        formula(vt, f"{c}{f_total}", f"SUM({c}5:{c}{f_total - 1})", num_format=MIL, bold=True, top=1)
        formula(vt, f"{c}{f_total + 1}", f"{ref(e, 'deuda_cp', COLUMNAS[-1])}+{ref(e, 'deuda_lp', COLUMNAS[-1])}", num_format=MIL, font_color=VERDE)
        formula(vt, f"{c}{f_total + 2}", f"{ref(e, 'caja', COLUMNAS[-1])}+{ref(e, 'inv_cp', COLUMNAS[-1])}", num_format=MIL, font_color=VERDE)
        formula(vt, f"{c}{f_total + 3}", f'IF({c}{f_total}=0,"n.d.",SUM({c}5:{c}{4 + len(PROYECCION)})/{c}{f_total})', num_format=P, align="right")
        formula(vt, f"{c}{f_total + 4}", f'IF({c}5=0,"n.d.",{c}{f_total + 2}/{c}5)', num_format=X, align="right")
    vt.write(f"I{f_total + 6}", "Control: el total por año tiene que coincidir con la deuda del balance (diferencias por intereses devengados).", fmt(font_color=GRIS, font_size=8))
    g_v = wb.add_chart({"type": "column"})
    for k, e in enumerate(EMPRESAS):
        g_v.add_series({"name": e, "categories": f"=Vencimientos!$I$5:$I${f_total - 1}", "values": f"=Vencimientos!${'JKL'[k]}$5:${'JKL'[k]}${f_total - 1}",
                        "fill": {"color": COLOR_EMPRESA[e]}, "gap": 80})
    g_v.set_title({"name": "Vencimientos por año (US$ M)", "name_font": {"name": "Arial", "size": 11}})
    g_v.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    g_v.set_y_axis({"num_format": "#,##0", "major_gridlines": {"visible": True, "line": {"color": LINEA}}, "num_font": {"name": "Arial", "size": 8}})
    g_v.set_x_axis({"num_font": {"name": "Arial", "size": 8}})
    g_v.set_size({"width": 620, "height": 300})
    vt.insert_chart("N4", g_v)

    # --- Rating
    fp = lambda ws, celda, cuerpo, **k: formula(ws, celda, f'IFERROR({cuerpo},"n.d.")', **k)  # sin datos, sin errores
    nombres = [x[0] for x in RATIOS]
    real = lambda e, nombre: f"Ratios!${COLUMNAS[-1]}${inicio[e] + 2 + nombres.index(nombre)}"
    rt = nueva_hoja("Rating", "Rating interno", "Puntaje de 1 a 5 por factor con los datos de 2025, y el mismo puntaje con el peor año del escenario de estrés.", "#1BAF7A")
    rt.set_column("A:A", 30)
    rt.set_column("B:G", 10)
    rt.set_column("H:N", 11)
    rt.set_row(4, 30)
    rt.write_row(4, 0, ["Factor", "Peso", "Límite 1", "Límite 2", "Límite 3", "Límite 4", "Mejor si"] + [f"{e} 2025" for e in EMPRESAS] + [f"{e} estrés" for e in EMPRESAS], fmt(**{**encabezado, "text_wrap": True}))
    fila_factor = {}
    for i, (texto, _, peso, limites, mayor) in enumerate(FACTORES):
        f = 6 + i
        fila_factor[texto] = f
        rt.write(f"A{f}", texto, fmt())
        rt.write(f"B{f}", peso, fmt(num_format="0%", font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA))
        for j, x in enumerate(limites):
            rt.write(f"{'CDEF'[j]}{f}", x, fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format="0.00" if abs(x) < 10 else "#,##0"))
        rt.write(f"G{f}", "mayor" if mayor else "menor", fmt(align="center", font_color=GRIS))
    fin_f = 5 + len(FACTORES)
    rt.write(f"A{fin_f + 1}", "Puntaje ponderado", fmt(bold=True, top=1))
    rt.write(f"A{fin_f + 2}", "Rating interno", fmt(bold=True))
    formula(rt, f"B{fin_f + 1}", f"SUM(B6:B{fin_f})", num_format="0%", bold=True, top=1)
    rt.write(f"A{fin_f + 4}", "Puntaje por factor: 1 debajo del límite 1, 5 sobre el límite 4 (al revés si es mejor el menor). Texto (n.s., n.d.): 1.", fmt(font_color=GRIS, font_size=8))
    rt.write(f"A{fin_f + 5}", "Escala interna: A desde 4,25; BBB desde 3,5; BB desde 2,75; B desde 2; debajo, CCC. No es comparable con una calificadora ni aplica techo soberano.", fmt(font_color=GRIS, font_size=8))
    rt.write(f"A{fin_f + 6}", "Estrés: deuda neta / EBITDA y EBITDA / intereses del peor año proyectado en Proyección. El resto queda como en 2025.", fmt(font_color=GRIS, font_size=8))

    # --- Proyección
    pr = nueva_hoja("Proyección", "Proyección 2026-2028", "¿Puede pagar si deja de crecer y se le encarece la deuda? Base: ventas y margen de 2025, capex de mantenimiento. Estrés: caen las ventas y sube la tasa.", "#E34948")
    pr.set_column("A:A", 40)
    pr.set_column("B:I", 11)
    supuesto = fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format="0.0%")
    pr.write("A4", "Supuestos generales", fmt(bold=True, font_color=ENCABEZADO))
    for f, texto, valor in ((5, "Crecimiento anual de ventas", 0.0), (6, "Estrés: caída de ventas en 2026", -0.20), (7, "Estrés: baja del margen EBITDA (pp)", 0.05), (8, "Estrés: suba de la tasa de deuda nueva (pp)", 0.04)):
        pr.write(f"A{f}", texto, fmt())
        pr.write(f"B{f}", valor, supuesto)
    pr.write("C6", "Supuesto: un año de precios bajos del petróleo y el gas.", fmt(font_color=GRIS, font_size=8))
    pr.write("C8", "Supuesto: un salto de riesgo país como el de 2023 encarece toda la deuda que se renueva.", fmt(font_color=GRIS, font_size=8))
    pr.write("A10", "Supuestos por empresa", fmt(bold=True, font_color=ENCABEZADO))
    pr.write_row(10, 0, ["Supuesto"] + EMPRESAS, fmt(**encabezado))
    fila_sup = {"conv": 12, "vida": 13, "tasa": 14, "margen": 15, "costo": 16}
    textos_sup = {"conv": "Flujo operativo / EBITDA (promedio 2023-2025)", "vida": "Vida media de la deuda (años)", "tasa": "Tasa de deuda nueva (base)",
                  "margen": "Margen EBITDA 2025", "costo": "Costo medio de la deuda 2025"}
    for k, f in fila_sup.items():
        pr.write(f"A{f}", textos_sup[k], fmt())
    ultimo = COLUMNAS[-1]
    for j, e in enumerate(EMPRESAS):
        col = "BCD"[j]
        conv = "+".join(ref(e, "cfo", c) for c in COLUMNAS) + ")/(" + "+".join(ref(e, "ebitda", c) for c in COLUMNAS)
        fp(pr, f"{col}12", f"({conv})", num_format="0.00", font_color=VERDE)
        pr.write(f"{col}13", VIDA_MEDIA[e], fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format="0.0"))
        pr.write(f"{col}14", TASA_NUEVA[e], supuesto)
        fp(pr, f"{col}15", f"{ref(e, 'ebitda', ultimo)}/{ref(e, 'ventas', ultimo)}", num_format="0.0%", font_color=VERDE)
        fp(pr, f"{col}16", real(e, "Costo medio de la deuda"), num_format="0.0%", font_color=VERDE)
    pr.write("A17", "Vida media: ponderada por saldo al 31-12-2025. Tasa nueva: rendimiento de la última ON larga en dólar duro.", fmt(font_color=GRIS, font_size=8))
    lineas = ["Ventas", "EBITDA", "Flujo operativo antes de intereses", "Capex", "Intereses", "Dividendos", "Vencimientos", "Necesidad de fondos", "Deuda",
              "Deuda tomada a tasa nueva", "Tasa de deuda nueva", "Costo medio de la deuda", "Caja e inversiones", "Deuda neta / EBITDA", "EBITDA / intereses", "Cobertura de servicio de deuda"]
    L = {t: i for i, t in enumerate(lineas)}
    alto_p = len(lineas) + 4
    salida = {}
    escenarios = (("Base", "CDE"), ("Estrés", "GHI"))
    for j, e in enumerate(EMPRESAS):
        s = 20 + j * alto_p
        c_e = "BCD"[j]
        sup = lambda k: f"${c_e}${fila_sup[k]}"
        fila = lambda t: s + 2 + L[t]
        pr.write(f"A{s}", e, fmt(bold=True, font_size=12, font_color=ENCABEZADO))
        pr.write(f"C{s}", "Base", fmt(bold=True, align="center"))
        pr.write(f"G{s}", "Estrés", fmt(bold=True, align="center"))
        pr.write_row(s, 0, ["US$ millones", "2025 real"] + PROYECCION + [""] + PROYECCION, fmt(**encabezado))
        for t in lineas:
            fuerte = dict(bold=True, top=1, top_color=LINEA) if t in ("Deuda neta / EBITDA",) else {}
            pr.write(f"A{fila(t)}", t, fmt(**fuerte))
        r = lambda clave: ref(e, clave, ultimo)
        reales = {"Ventas": r("ventas"), "EBITDA": r("ebitda"), "Flujo operativo antes de intereses": r("cfo"), "Capex": r("capex"), "Intereses": r("gasto_int"),
                  "Dividendos": r("dividendos"), "Deuda": r("deuda"), "Deuda tomada a tasa nueva": "0", "Costo medio de la deuda": sup("costo"),
                  "Caja e inversiones": f"{r('caja')}+{r('inv_cp')}"}
        for t, x in reales.items():
            fp(pr, f"B{fila(t)}", x, font_color=VERDE, num_format=P if "Costo" in t else MIL)
        fp(pr, f"B{fila('Deuda neta / EBITDA')}", f"(B{fila('Deuda')}-B{fila('Caja e inversiones')})/B{fila('EBITDA')}", num_format=X)
        fp(pr, f"B{fila('EBITDA / intereses')}", f"B{fila('EBITDA')}/B{fila('Intereses')}", num_format=X)
        for nombre, cols in escenarios:
            estres = nombre == "Estrés"
            for i, c in enumerate(cols):
                a = "B" if i == 0 else cols[i - 1]
                x = lambda t, col=c: f"{col}{fila(t)}"
                ant = lambda t: f"{a}{fila(t)}"
                caida = "+$B$6" if estres and i == 0 else ""
                v = {"Ventas": f"{ant('Ventas')}*(1+$B$5{caida})",
                     "EBITDA": f"{x('Ventas')}*({sup('margen')}{'-$B$7' if estres else ''})",
                     "Flujo operativo antes de intereses": f"{x('EBITDA')}*{sup('conv')}",
                     "Capex": r("amort"),
                     "Intereses": f"{ant('Deuda')}*{ant('Costo medio de la deuda')}",
                     "Dividendos": "0",
                     "Vencimientos": f"Vencimientos!${'JKL'[j]}${fila_venc[PROYECCION[i]]}+{ant('Deuda tomada a tasa nueva')}/{sup('vida')}",
                     "Necesidad de fondos": f"{x('Capex')}+{x('Intereses')}+{x('Dividendos')}-{x('Flujo operativo antes de intereses')}",
                     "Deuda": f"MAX(0,{ant('Deuda')}+{x('Necesidad de fondos')})",
                     "Deuda tomada a tasa nueva": f"MIN({x('Deuda')},{ant('Deuda tomada a tasa nueva')}+{x('Vencimientos')}+MAX(0,{x('Necesidad de fondos')}))",
                     "Tasa de deuda nueva": f"{sup('tasa')}{'+$B$8' if estres else ''}",
                     "Costo medio de la deuda": f"IF({x('Deuda')}=0,0,({x('Deuda tomada a tasa nueva')}*{x('Tasa de deuda nueva')}+({x('Deuda')}-{x('Deuda tomada a tasa nueva')})*{sup('costo')})/{x('Deuda')})",
                     "Caja e inversiones": f"$B{fila('Caja e inversiones')}",
                     "Deuda neta / EBITDA": f'IF({x("EBITDA")}<=0,"n.s.",({x("Deuda")}-{x("Caja e inversiones")})/{x("EBITDA")})',
                     "EBITDA / intereses": f'IF({x("Intereses")}<=0,"n.s.",{x("EBITDA")}/{x("Intereses")})',
                     "Cobertura de servicio de deuda": f"{x('Flujo operativo antes de intereses')}/({x('Intereses')}+{x('Vencimientos')})"}
                for t, cuerpo in v.items():
                    formato = P if t in ("Tasa de deuda nueva", "Costo medio de la deuda") else X if t in ("Deuda neta / EBITDA", "EBITDA / intereses", "Cobertura de servicio de deuda") else MIL
                    fp(pr, x(t), cuerpo, num_format=formato, font_color=VERDE if t == "Capex" else "#0B0B0B", align="right")
        salida[e] = {t: fila(t) for t in lineas}
    pr.write(f"A{20 + len(EMPRESAS) * alto_p}", "Mecánica: la necesidad de fondos se cubre con deuda nueva y la caja queda constante. Vencen los montos de la hoja Vencimientos, más la deuda nueva según la vida media; lo que vence se renueva a la tasa nueva. Sin dividendos.", fmt(font_color=GRIS, font_size=8))

    # Puntajes del rating, ahora que existe la proyección
    for j, e in enumerate(EMPRESAS):
        for estres, col in ((False, "HIJ"[j]), (True, "KLM"[j])):
            celdas_score = []
            for texto, clave, _, _, mayor in FACTORES:
                f = fila_factor[texto]
                if clave == "ebitda":
                    valor = ref(e, "ebitda", ultimo)
                elif estres and texto == "Deuda neta / EBITDA":
                    valor = f"MAX(Proyección!$G${salida[e][texto]}:$I${salida[e][texto]})"
                elif estres and texto == "EBITDA / intereses":
                    valor = f"MIN(Proyección!$G${salida[e][texto]}:$I${salida[e][texto]})"
                else:
                    valor = real(e, clave)
                formula(rt, f"{col}{f}", valor, font_color=VERDE, num_format="0.00", align="right")
                v = f"{col}{f}"
                if mayor:
                    puntos = f"1+({v}>=$C{f})+({v}>=$D{f})+({v}>=$E{f})+({v}>=$F{f})"
                else:
                    puntos = f"5-({v}>$C{f})-({v}>$D{f})-({v}>$E{f})-({v}>$F{f})"
                celdas_score.append(f"IF(ISNUMBER({v}),{puntos},1)*$B{f}")
            formula(rt, f"{col}{fin_f + 1}", "+".join(celdas_score), num_format="0.00", bold=True, top=1, align="right")
            nota = f"{col}{fin_f + 1}"
            letras = "".join(f'IF({nota}>={m},"{n}",' for m, n in ESCALA) + '"CCC"' + ")" * len(ESCALA)
            formula(rt, f"{col}{fin_f + 2}", letras, bold=True, align="center", font_color=ENCABEZADO)

    # --- Covenants
    cv = nueva_hoja("Covenants", "Covenants", "Cuánto margen tiene cada empresa antes de incumplir. Los límites son supuestos: los 20-F no publican los de cada contrato.", "#EDA100")
    cv.set_column("A:A", 44)
    cv.set_column("B:H", 11)
    cv.write("A4", "Máximo de deuda neta / EBITDA", fmt())
    cv.write("B4", 3.5, fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format=X))
    cv.write("A5", "Mínimo de EBITDA / intereses", fmt())
    cv.write("B5", 2.0, fmt(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA, num_format=X))
    cv.write("C4", "Supuesto: límites habituales en deuda de energía de la región. Vista declara covenants financieros y dice cumplirlos.", fmt(font_color=GRIS, font_size=8))
    cv.write_row(6, 0, ["Empresa y escenario", "2025 real"] + PROYECCION + ["Peor año", "Estado", "EBITDA puede caer"], fmt(**{**encabezado, "text_wrap": True}))
    cv.set_row(6, 30)
    f = 8
    for e in EMPRESAS:
        for nombre, cols in escenarios:
            for medida, limite, peor, mayor in (("Deuda neta / EBITDA", "$B$4", "MAX", False), ("EBITDA / intereses", "$B$5", "MIN", True)):
                cv.write(f"A{f}", f"{e} {nombre.lower()}: {medida.lower()}", fmt())
                filap = salida[e][medida]
                fp(cv, f"B{f}", f"Proyección!$B${filap}", num_format=X, font_color=VERDE, align="right")
                for i, c in enumerate(cols):
                    fp(cv, f"{'CDE'[i]}{f}", f"Proyección!${c}${filap}", num_format=X, font_color=VERDE, align="right")
                fp(cv, f"F{f}", f"{peor}(C{f}:E{f})", num_format=X, align="right")
                condicion = f"F{f}>={limite}" if mayor else f"F{f}<={limite}"
                fp(cv, f"G{f}", f'IF({condicion},"Cumple","Incumple")', align="center")
                ebitda, deuda_neta = f"Proyección!$B${salida[e]['EBITDA']}", f"(Proyección!$B${salida[e]['Deuda']}-Proyección!$B${salida[e]['Caja e inversiones']})"
                caida = f"1-{deuda_neta}/{limite}/{ebitda}" if not mayor else f"1-{limite}*Proyección!$B${salida[e]['Intereses']}/{ebitda}"
                fp(cv, f"H{f}", f"MAX(0,{caida})" if nombre == "Base" else '""', num_format="0%", align="right")
                f += 1
        f += 1
    cv.conditional_format(f"G8:G{f}", {"type": "cell", "criteria": "==", "value": '"Incumple"', "format": wb.add_format({"font_color": "#9C0006", "bg_color": "#FFC7CE"})})
    cv.conditional_format(f"G8:G{f}", {"type": "cell", "criteria": "==", "value": '"Cumple"', "format": wb.add_format({"font_color": "#006100", "bg_color": "#C6EFCE"})})
    cv.write(f"A{f + 1}", "EBITDA puede caer: cuánto podría bajar el EBITDA de 2025, con la deuda y los intereses de 2025, antes de llegar al límite.", fmt(font_color=GRIS, font_size=8))

    # --- Mercado
    k = nueva_hoja("Mercado", "Series de mercado", "Riesgo país: ArgentinaDatos. Tesoro de EE.UU.: FRED. BADLAR y adelantos a empresas: BCRA. Se actualiza con python -m empresas.mercado.", "#898781")
    k.set_column("A:A", 12)
    k.set_column("B:J", 11)
    k.freeze_panes(4, 1)
    columnas = ["Fecha", "Riesgo país (pb)"] + PLAZOS_TESORO + ["BADLAR (%)", "Adelantos (%)"]
    alto_m = fmt(**{**encabezado, "text_wrap": True})
    k.set_row(3, 30)
    k.write_row(3, 0, columnas, alto_m)
    k.write("C3", "Tesoro de EE.UU. por plazo en años (%)", fmt(font_color=GRIS))
    fecha = fmt(num_format="dd/mm/yyyy", font_color=AZUL)
    dato = fmt(num_format="0.00", font_color=AZUL)
    for i, fila_m in enumerate(mercado.itertuples(index=False)):
        k.write_datetime(4 + i, 0, pd.Timestamp(fila_m.fecha).to_pydatetime(), fecha)
        k.write_number(4 + i, 1, fila_m.riesgo_pais, fmt(num_format="0", font_color=AZUL))
        for j, p in enumerate(PLAZOS_TESORO):
            k.write_number(4 + i, 2 + j, getattr(fila_m, f"ust_{p}"), dato)
        k.write_number(4 + i, 8, fila_m.badlar, dato)
        k.write_number(4 + i, 9, fila_m.adelantos, dato)
    ult = 4 + len(mercado)
    M_FECHA = f"Mercado!$A$5:$A${ult}"

    # --- Tasas
    a = nueva_hoja("Tasas", "Tasas del sistema financiero", "Promedio mensual. Crédito: BCRA, cuadro VIII, empresas que no son pymes, tasa fija y repactable. Plazo fijo: datos.gob.ar. Vacío: sin operaciones o sin dato.", "#898781")
    a.set_column("A:A", 10)
    a.set_column("B:I", 12)
    a.freeze_panes(4, 1)
    a.set_row(3, 56)
    a.write_row(3, 0, ["Mes", "Crédito USD total (%)", "Crédito USD hasta 89 días (%)", "Crédito USD 90 días o más (%)", "Crédito pesos (%)", "Plazo fijo USD (%)", "Plazo fijo pesos (%)", "Margen USD (pp)", "Margen pesos (pp)"], fmt(**{**encabezado, "text_wrap": True}))
    fin_a = 4 + max(len(tasas), 1)
    azul_tasa = fmt(num_format="0.00", font_color=AZUL)
    for i, fila_t in enumerate(tasas.itertuples(index=False)):
        f = 5 + i
        a.write_datetime(f"A{f}", pd.Timestamp(fila_t.mes).to_pydatetime(), fmt(num_format="mmm-yy", font_color=AZUL))
        for c, k in zip("BCDEFG", ("doc_usd", "doc_usd_corto", "doc_usd_largo", "doc_ars", "pf_usd", "pf_ars")):
            poner(a, f"{c}{f}", float(getattr(fila_t, k)), azul_tasa)
        formula(a, f"H{f}", f'IF(OR(B{f}="",F{f}=""),"",B{f}-F{f})', num_format="0.00")
        formula(a, f"I{f}", f'IF(OR(E{f}="",G{f}=""),"",E{f}-G{f})', num_format="0.00")
    g_t = wb.add_chart({"type": "line"})
    for col, nombre, color in (("B", "Crédito en dólares", "#2A78D6"), ("F", "Plazo fijo en dólares", "#EDA100")):
        g_t.add_series({"name": nombre, "categories": f"=Tasas!$A$5:$A${fin_a}", "values": f"=Tasas!${col}$5:${col}${fin_a}", "line": {"color": color, "width": 2.25}})
    g_t.show_blanks_as("gap")
    g_t.set_title({"name": "Tasa activa y pasiva en dólares", "name_font": {"name": "Arial", "size": 11}})
    g_t.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 9}})
    g_t.set_y_axis({"num_format": '0"%"', "major_gridlines": {"visible": True, "line": {"color": LINEA}}, "num_font": {"name": "Arial", "size": 8}})
    g_t.set_x_axis({"date_axis": True, "num_format": "mmm-yy", "major_unit": 12, "major_unit_type": "months", "num_font": {"name": "Arial", "size": 8}})
    g_t.set_size({"width": 560, "height": 300})
    a.insert_chart("K4", g_t)

    # --- Emisiones
    n = nueva_hoja("Emisiones", "Emisiones de obligaciones negociables", "Una fila por ON o reapertura. Columnas A a K: cargá desde el aviso de resultados. L a S: se calculan solas.", "#E34948")
    titulos = ["Empresa", "Clase", "Fecha de emisión", "Vencimiento", "Instrumento", "Ley", "Monto (US$ M)", "Cupón", "Rendimiento", "Colocadores", "Fuente",
               "Plazo (años)", "Riesgo país (pb)", "Tesoro (%)", "Spread (pb)", "Prima contra riesgo país (pb)", "Año", "Crédito del sistema USD (%)", "ON menos crédito (pb)"]
    n.set_row(3, 42)
    n.write_row(3, 0, titulos, fmt(**{**encabezado, "text_wrap": True}))
    n.set_column("A:A", 9)
    n.set_column("B:B", 10)
    n.set_column("C:D", 11)
    n.set_column("E:F", 16)
    n.set_column("G:I", 11)
    n.set_column("J:K", 26)
    n.set_column("L:S", 11)
    n.freeze_panes(4, 2)
    n.autofilter(3, 0, 3 + FILAS_EMISIONES, len(titulos) - 1)
    ed = dict(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA)
    formatos = ["", "", "dd/mm/yyyy", "dd/mm/yyyy", "", "", "#,##0.0", "0.00%", "0.00%", "", ""]
    fin_e = 4 + FILAS_EMISIONES
    n.data_validation(f"A5:A{fin_e}", {"validate": "list", "source": EMPRESAS})
    n.data_validation(f"E5:E{fin_e}", {"validate": "list", "source": INSTRUMENTOS})
    n.data_validation(f"F5:F{fin_e}", {"validate": "list", "source": LEYES})
    for f in range(5, fin_e + 1):
        for j, (c, nf) in enumerate(zip("ABCDEFGHIJK", formatos)):
            valor = emisiones.iat[f - 5, j] if emisiones is not None and f - 5 < len(emisiones) else None
            poner(n, f"{c}{f}", valor, fmt(**ed, **({"num_format": nf} if nf else {})))
        formula(n, f"L{f}", f'IF($C{f}="","",($D{f}-$C{f})/365.25)', num_format="0.0")
        formula(n, f"M{f}", f'IF($C{f}="","",INDEX(Mercado!$B$5:$B${ult},MATCH($C{f},{M_FECHA},1)))', font_color=VERDE, num_format="0")
        formula(n, f"N{f}", f'IF($C{f}="","",INDEX(Mercado!$C$5:$H${ult},MATCH($C{f},{M_FECHA},1),MATCH(MAX($L{f},1),Mercado!$C$4:$H$4,1)))', font_color=VERDE, num_format="0.00")
        formula(n, f"O{f}", f'IF($C{f}="","",IF($E{f}="{INSTRUMENTOS[0]}",(IF($I{f}="",$H{f},$I{f})*100-$N{f})*100,"n.a."))', num_format="0")
        formula(n, f"P{f}", f'IF($C{f}="","",IF($E{f}="{INSTRUMENTOS[0]}",$O{f}-$M{f},"n.a."))', num_format='0;(0);0')
        formula(n, f"Q{f}", f'IF($C{f}="","",YEAR($C{f}))', num_format="0")
        formula(n, f"R{f}", f'IF($C{f}="","",IFERROR(IF(INDEX(Tasas!$D$5:$D${fin_a},MATCH($C{f},Tasas!$A$5:$A${fin_a},1))="","n.d.",INDEX(Tasas!$D$5:$D${fin_a},MATCH($C{f},Tasas!$A$5:$A${fin_a},1))),"n.d."))', font_color=VERDE, num_format="0.00", align="right")
        formula(n, f"S{f}", f'IF($C{f}="","",IF(AND($E{f}="{INSTRUMENTOS[0]}",ISNUMBER($R{f})),(IF($I{f}="",$H{f},$I{f})*100-$R{f})*100,"n.a."))', num_format="0;(0);0", align="right")

    # --- Créditos
    t = nueva_hoja("Créditos", "Préstamos y prefinanciaciones", "Una fila por crédito que detalle el 20-F. Columnas A a L: cargá. M y N: se calculan solas.", "#E34948")
    titulos_c = ["Empresa", "Año de cierre", "Tipo", "Acreedor", "Moneda", "Tasa", "Referencia", "Tasa o margen", "Inicio", "Vencimiento", "Saldo (US$ M)", "Fuente", "Plazo (años)", "Peso en la deuda"]
    t.set_row(3, 30)
    t.write_row(3, 0, titulos_c, fmt(**{**encabezado, "text_wrap": True}))
    t.set_column("A:C", 12)
    t.set_column("D:D", 24)
    t.set_column("E:K", 11)
    t.set_column("L:L", 26)
    t.set_column("M:N", 11)
    t.freeze_panes(4, 1)
    t.autofilter(3, 0, 3 + FILAS_CREDITOS, len(titulos_c) - 1)
    fin_t = 4 + FILAS_CREDITOS
    for rango, lista in (("A", EMPRESAS), ("B", ANIOS), ("C", TIPOS_CREDITO), ("E", MONEDAS), ("F", TASAS), ("G", REFERENCIAS)):
        t.data_validation(f"{rango}5:{rango}{fin_t}", {"validate": "list", "source": lista})
    formatos_c = ["", "0", "", "", "", "", "", "0.00%", "dd/mm/yyyy", "dd/mm/yyyy", "#,##0.0", ""]
    for f in range(5, fin_t + 1):
        for j, (c, nf) in enumerate(zip("ABCDEFGHIJKL", formatos_c)):
            valor = creditos.iat[f - 5, j] if creditos is not None and f - 5 < len(creditos) else None
            poner(t, f"{c}{f}", valor, fmt(**ed, **({"num_format": nf} if nf else {})))
        formula(t, f"M{f}", f'IF(OR($I{f}="",$J{f}=""),"",($J{f}-$I{f})/365.25)', num_format="0.0")
        deuda_e = "+".join(f'($A{f}="{e}")*INDEX(\'{e}\'!$B${FILAS["deuda"]}:${COLUMNAS[-1]}${FILAS["deuda"]},MATCH($B{f},\'{e}\'!$B$6:${COLUMNAS[-1]}$6,0))' for e in EMPRESAS)
        formula(t, f"N{f}", f'IF(OR($A{f}="",$B{f}="",$K{f}=""),"",IFERROR($K{f}/({deuda_e}),"n.d."))', font_color=VERDE, num_format=P)

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


def rescatar_manual() -> None:
    """Pasa a data/balances.csv lo cargado a mano en el Excel, para que regenerar el archivo no lo pierda."""
    if not ARCHIVO.exists():
        return
    wb = openpyxl.load_workbook(ARCHIVO, data_only=True)
    filas = []
    for e in CARGA_MANUAL:
        for col, a in zip(COLUMNAS, ANIOS):
            filas += [(e, a, clave, wb[e][f"{col}{FILAS[clave]}"].value) for clave in CARGA]
            filas.append((e, a, "fuente", wb[e][f"{col}5"].value))
    nuevas = pd.DataFrame(filas, columns=["empresa", "anio", "clave", "valor"]).dropna()
    viejas = pd.read_csv(BALANCES, encoding="utf-8")
    viejas = viejas[~viejas.empresa.isin(CARGA_MANUAL)]
    pd.concat([viejas, nuevas]).to_csv(BALANCES, index=False, encoding="utf-8")


def leer_balances() -> dict:
    d = pd.read_csv(BALANCES, encoding="utf-8")
    return {(r.empresa, int(r.anio), r.clave): (r.valor if r.clave == "fuente" else float(r.valor)) for r in d.itertuples(index=False)}


def leer_detalle(ruta: Path, fechas: list[str]) -> pd.DataFrame:
    d = pd.read_csv(ruta, encoding="utf-8")
    for c in fechas:
        d[c] = pd.to_datetime(d[c])
    return d.astype(object).where(d.notna(), None)


def main() -> None:
    ARCHIVO.parent.mkdir(exist_ok=True)
    rescatar_manual()
    datos = (pd.read_csv(MERCADO), leer_balances(), leer_detalle(EMISIONES, ["fecha", "vencimiento"]), leer_detalle(CREDITOS, ["inicio", "vencimiento"]), pd.read_csv(TASAS_SISTEMA), pd.read_csv(VENCIMIENTOS))
    construir(ARCHIVO, {}, *datos)
    valores = calcular(ARCHIVO)
    construir(ARCHIVO, valores, *datos)
    print(f"{ARCHIVO.name}: {len(valores)} celdas calculadas")


if __name__ == "__main__":
    main()
