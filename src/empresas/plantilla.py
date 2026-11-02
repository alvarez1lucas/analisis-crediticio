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
import pandas as pd
import xlsxwriter

ROOT = Path(__file__).resolve().parents[2]
ARCHIVO = ROOT / "excel" / "analisis_crediticio.xlsx"
MERCADO = ROOT / "data" / "mercado.csv"
PLAZOS_TESORO = [1, 2, 3, 5, 7, 10]
INSTRUMENTOS = ["Dólar duro", "Dólar linked", "Pesos tasa variable", "Pesos tasa fija", "UVA"]
LEYES = ["Nueva York", "Argentina"]
TIPOS_CREDITO = ["Préstamo bancario", "Prefinanciación", "Sindicado", "Pagaré", "Otro"]
MONEDAS = ["Dólar", "Pesos"]
TASAS = ["Fija", "Variable"]
REFERENCIAS = ["SOFR", "BADLAR", "TAMAR", "Otra", "Ninguna"]
FILAS_CREDITOS = 100
FILAS_EMISIONES = 150
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
    ("res_acum", "Resultados acumulados", None, "Parte del patrimonio, para Altman. Con signo"),
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
    ("v_1", "Vence en 1 año", None, "Total por vencimiento"),
    ("v_2", "Vence en 2 años", None, None),
    ("v_3", "Vence en 3 años", None, None),
    ("v_mas", "Vence en más de 3 años", None, None),
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
    ("estado", "Estado", 'IF({ventas}=0,"Sin cargar",IF(AND(ABS({dif_balance})<=$B$4,ABS({dif_activo})<=$B$4,ABS({dif_neto})<=$B$4,ABS({dif_tipo})<=$B$4,ABS({dif_moneda})<=$B$4,ABS({dif_vto})<=$B$4),"OK","Revisar"))', None),
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


def construir(archivo: Path, valores: dict, mercado: pd.DataFrame) -> None:
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
        ("7. Emisiones", "En Emisiones cargá una fila por ON o reapertura, desde el aviso de resultados. Las columnas de la derecha se calculan solas."),
        ("8. Créditos", "En Créditos cargá una fila por préstamo o prefinanciación que detalle el 20-F, con su saldo al cierre."),
    ]), ("Ejemplo de la línea Costo de ventas", [
        ("Dato del 20-F", "Si el informe muestra (1.234,5), cargá 1234,5 en positivo. Es un valor de ejemplo, no un dato de la empresa."),
        ("Fila de Emisiones", "Empresa, clase, fecha de emisión, vencimiento, instrumento, ley, monto en US$ M, cupón, rendimiento, colocadores y fuente. Si se colocó a la par, el rendimiento es el cupón."),
    ]), ("Definiciones", [
        ("EBITDA", "Resultado operativo sin asociadas + depreciación y amortización + deterioro. No es el EBITDA ajustado que publica cada empresa."),
        ("Deuda", "Deuda financiera corriente y no corriente, con arrendamientos."),
        ("Cobertura de servicio de deuda", "Flujo operativo antes de intereses / (intereses pagados + deuda y arrendamientos corrientes del año anterior)."),
        ("Altman Z'' EM", "Versión para mercados emergentes: 3,25 + 6,56 capital de trabajo/activo + 3,26 resultados acumulados/activo + 6,72 resultado operativo/activo + 1,05 patrimonio/pasivo. Es una referencia, no un rating."),
        ("Spread", "Rendimiento de la emisión menos la tasa del Tesoro de EE.UU. del plazo inmediato inferior, en puntos básicos. Sólo para dólar duro."),
        ("Prima contra riesgo país", "Spread menos riesgo país de ese día. Negativa: la empresa se financió más barato que el Estado."),
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
    u.set_column("B:E", 11)
    u.set_column("F:F", 60)
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
            u.write(f"{'BCDE'[j]}{ini + 1}", a, fmt(**encabezado))
        for i, (texto, calculo, nota) in enumerate(filas_u):
            f = ini + 2 + i
            u.write(f"A{f}", texto, fmt(**(dict(bold=True, top=1, top_color=LINEA) if i == len(filas_u) - 1 else {})))
            if nota:
                u.write(f"F{f}", nota, fmt(font_color=GRIS, font_size=8))
            for j, col in enumerate("BCDE"):
                actual, anterior = COLUMNAS[j + 1], COLUMNAS[j]
                cuerpo = f"{col}{ini + 8}-SUM({col}{ini + 2}:{col}{ini + 6})" if calculo is None else calculo(e, actual, anterior)
                cargado = f"OR({ref(e, 'ventas', actual)}=0,{ref(e, 'ventas', anterior)}=0)"
                extra = dict(font_color=VERDE) if calculo else {}
                formula(u, f"{col}{f}", f'IF({cargado},"n.d.",{cuerpo})', num_format=MIL, align="right", **extra)

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

    # --- Emisiones
    n = nueva_hoja("Emisiones", "Emisiones de obligaciones negociables", "Una fila por ON o reapertura. Columnas A a K: cargá desde el aviso de resultados. L a Q: se calculan solas.", "#E34948")
    titulos = ["Empresa", "Clase", "Fecha de emisión", "Vencimiento", "Instrumento", "Ley", "Monto (US$ M)", "Cupón", "Rendimiento", "Colocadores", "Fuente",
               "Plazo (años)", "Riesgo país (pb)", "Tesoro (%)", "Spread (pb)", "Prima contra riesgo país (pb)", "Año"]
    n.set_row(3, 42)
    n.write_row(3, 0, titulos, fmt(**{**encabezado, "text_wrap": True}))
    n.set_column("A:A", 9)
    n.set_column("B:B", 10)
    n.set_column("C:D", 11)
    n.set_column("E:F", 16)
    n.set_column("G:I", 11)
    n.set_column("J:K", 26)
    n.set_column("L:Q", 11)
    n.freeze_panes(4, 2)
    n.autofilter(3, 0, 3 + FILAS_EMISIONES, len(titulos) - 1)
    ed = dict(font_color=AZUL, bg_color=RESALTE, border=1, border_color=LINEA)
    formatos = ["", "", "dd/mm/yyyy", "dd/mm/yyyy", "", "", "#,##0.0", "0.00%", "0.00%", "", ""]
    fin_e = 4 + FILAS_EMISIONES
    n.data_validation(f"A5:A{fin_e}", {"validate": "list", "source": EMPRESAS})
    n.data_validation(f"E5:E{fin_e}", {"validate": "list", "source": INSTRUMENTOS})
    n.data_validation(f"F5:F{fin_e}", {"validate": "list", "source": LEYES})
    for f in range(5, fin_e + 1):
        for c, nf in zip("ABCDEFGHIJK", formatos):
            n.write_blank(f"{c}{f}", None, fmt(**ed, **({"num_format": nf} if nf else {})))
        formula(n, f"L{f}", f'IF($C{f}="","",($D{f}-$C{f})/365.25)', num_format="0.0")
        formula(n, f"M{f}", f'IF($C{f}="","",INDEX(Mercado!$B$5:$B${ult},MATCH($C{f},{M_FECHA},1)))', font_color=VERDE, num_format="0")
        formula(n, f"N{f}", f'IF($C{f}="","",INDEX(Mercado!$C$5:$H${ult},MATCH($C{f},{M_FECHA},1),MATCH(MAX($L{f},1),Mercado!$C$4:$H$4,1)))', font_color=VERDE, num_format="0.00")
        formula(n, f"O{f}", f'IF($C{f}="","",IF($E{f}="{INSTRUMENTOS[0]}",($I{f}*100-$N{f})*100,"n.a."))', num_format="0")
        formula(n, f"P{f}", f'IF($C{f}="","",IF($E{f}="{INSTRUMENTOS[0]}",$O{f}-$M{f},"n.a."))', num_format='0;(0);0')
        formula(n, f"Q{f}", f'IF($C{f}="","",YEAR($C{f}))', num_format="0")

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
        for c, nf in zip("ABCDEFGHIJKL", formatos_c):
            t.write_blank(f"{c}{f}", None, fmt(**ed, **({"num_format": nf} if nf else {})))
        formula(t, f"M{f}", f'IF(OR($I{f}="",$J{f}=""),"",($J{f}-$I{f})/365.25)', num_format="0.0")
        deuda_e = "+".join(f'($A{f}="{e}")*INDEX(\'{e}\'!$B${FILAS["deuda"]}:$F${FILAS["deuda"]},MATCH($B{f},\'{e}\'!$B$6:$F$6,0))' for e in EMPRESAS)
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


def main() -> None:
    if ARCHIVO.exists() and "--forzar" not in sys.argv:
        raise SystemExit(f"{ARCHIVO.name} ya existe y puede tener datos cargados. Usá --forzar para reemplazarlo.")
    ARCHIVO.parent.mkdir(exist_ok=True)
    mercado = pd.read_csv(MERCADO)
    construir(ARCHIVO, {}, mercado)
    valores = calcular(ARCHIVO)
    construir(ARCHIVO, valores, mercado)
    print(f"{ARCHIVO.name}: {len(valores)} celdas calculadas")


if __name__ == "__main__":
    main()
