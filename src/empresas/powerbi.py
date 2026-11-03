"""Tablero de Power BI generado desde outputs/, como proyecto de texto (TMDL y PBIR).

`python -m empresas.powerbi` escribe powerbi/analisis_crediticio.pbip: modelo, medidas, tema y cuatro páginas. Al abrirlo,
Power BI lee los CSV del proyecto y dibuja el tablero. El esqueleto (identidad del reporte, versión y tema base) se tomó de
un proyecto que creó Power BI Desktop; el resto se escribe acá. Mismo método que el tablero de umbrales-credito.
"""

from __future__ import annotations

import json
import re
import shutil
import uuid
from pathlib import Path

import pandas as pd

from empresas import memo, plantilla as p
from empresas.sql import OUTPUTS

NOMBRE = "analisis_crediticio"
SALIDA = p.ROOT / "powerbi"
SEMILLA = uuid.UUID("4c6f1a2e-8a51-4d0e-9b7a-2f3d5c1e9a07")
VC = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.13.0/schema.json"
PG = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
PGS = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json"
FONDO, SUPERFICIE, LINEA = "#f9f9f7", "#fcfcfb", "#e1e0d9"
TEXTO, SECUNDARIO, MUTED = "#0b0b0b", "#52514e", "#6f6e69"
COLOR_EMPRESA = {e: c.lower() for e, c in p.COLOR_EMPRESA.items()}
COLOR_INSTRUMENTO = {"Dólar duro": "#2a78d6", "Dólar linked": "#e34948"}
COLOR_ESCENARIO = {"Real": "#52514e", "Base": "#2a78d6", "Estrés": "#e34948"}
TABLA_MEDIDAS = "Medidas"
VECES = "#,0.0x"

# nombre de tabla -> (archivo relativo a la raíz del proyecto, columnas de fecha)
TABLAS = {
    "dim_empresa": ("outputs/empresas.csv", ()),
    "indicadores": ("outputs/indicadores.csv", ()),
    "emisiones": ("outputs/emisiones.csv", ("fecha", "vencimiento")),
    "tendencias": ("outputs/tendencias.csv", ()),
    "proyeccion": ("outputs/proyeccion.csv", ()),
    "rating": ("outputs/rating.csv", ()),
    "mercado": ("data/mercado.csv", ("fecha",)),
    "vencimientos": ("outputs/vencimientos.csv", ()),
    "puente": ("outputs/puente.csv", ()),
    "colocadores_entidad": ("outputs/colocadores_entidad.csv", ()),
    "colocadores_emision": ("outputs/colocadores_emision.csv", ("fecha",)),
}
RELACIONES = [(t, "empresa", "dim_empresa", "empresa") for t in ("indicadores", "emisiones", "tendencias", "proyeccion", "rating", "vencimientos", "puente", "colocadores_emision")]
ORDEN = {("dim_empresa", "empresa"): "orden", ("puente", "concepto"): "orden"}
FORMATO = {("emisiones", "tasa"): "0.00%", ("emisiones", "cupon"): "0.00%", ("emisiones", "monto"): "#,0", ("emisiones", "riesgo_pais"): "#,0",
           ("emisiones", "prima_riesgo_pais_pb"): "#,0", ("emisiones", "spread_pb"): "#,0", ("emisiones", "plazo"): "0.0",
           ("rating", "on_deuda"): "0%", ("rating", "caja_2025"): "#,0", ("rating", "vence_primer_anio"): "#,0", ("rating", "caja_sobre_vencimientos"): "#,0.0x",
           ("rating", "anios_cubiertos"): "0", ("rating", "ebitda_puede_caer"): "0%", ("rating", "puntaje_2025"): "0.00", ("rating", "puntaje_estres"): "0.00",
           **{("rating", c): "#,0.0x" for c in ("deuda_neta_ebitda", "ebitda_intereses", "liquidez_corriente", "peor_apalancamiento_estres",
                                                 "peor_cobertura_estres", "peor_servicio_estres")}}
INDICADORES = {"Margen EBITDA": ("margen_ebitda", "0.0%"), "Deuda neta / EBITDA": ("deuda_neta_ebitda", VECES),
               "EBITDA / intereses": ("ebitda_intereses", VECES), "Liquidez corriente": ("liquidez_corriente", VECES),
               "Flujo libre / deuda": ("flujo_libre_deuda", "0.0%"), "ON / deuda": ("on_deuda", "0.0%"),
               "Deuda en dólares": ("dolares_deuda", "0.0%"), "Capex / ventas": ("capex_ventas", "0.0%"),
               "Rentabilidad sobre patrimonio": ("roe", "0.0%"), "Margen neto": ("margen_neto", "0.0%"),
               "Rotación de activos": ("rotacion_activos", "0.00"), "Activo / patrimonio": ("activo_patrimonio", "#,0.0x")}
MEDIDAS = [(n, f'CALCULATE(SUM(indicadores[valor]), indicadores[indicador] = "{c}")', f) for n, (c, f) in INDICADORES.items()] + [
    ("Monto emitido", "SUM(emisiones[monto])", "#,0"),
    ("Tasa media ponderada", "DIVIDE(SUMX(emisiones, emisiones[monto] * emisiones[tasa]), SUM(emisiones[monto]))", "0.00%"),
    ("Riesgo país", "AVERAGE(mercado[riesgo_pais])", "#,0"),
    ("Riesgo país al emitir", "AVERAGE(emisiones[riesgo_pais])", "#,0"),
    ("Proyección deuda neta / EBITDA", 'CALCULATE(SUM(proyeccion[valor]), proyeccion[indicador] = "deuda_neta_ebitda")', VECES),
    ("Proyección cobertura de servicio", 'CALCULATE(SUM(proyeccion[valor]), proyeccion[indicador] = "cobertura_servicio")', VECES),
    ("% de la deuda", "DIVIDE(SUM(vencimientos[monto]), CALCULATE(SUM(vencimientos[monto]), REMOVEFILTERS(vencimientos[anio], vencimientos[tipo], vencimientos[moneda])))", "0%"),
    ("Monto del puente", "SUM(puente[monto])", "#,0"),
    ("Participaciones", "SUM(colocadores_entidad[emisiones])", "0"),
]


def tag(*partes: object) -> str:
    return str(uuid.uuid5(SEMILLA, "|".join(map(str, partes))))


def q(nombre: str) -> str:
    return nombre if re.fullmatch(r"[A-Za-z_]\w*", nombre, re.ASCII) else "'" + nombre.replace("'", "''") + "'"


def sangrar(texto: str, n: int) -> str:
    return "\n".join("\t" * n + linea for linea in texto.splitlines())


# --- Modelo ---

def tipos(df: pd.DataFrame, fechas: tuple) -> dict[str, tuple[str, str]]:
    salida = {}
    for c in df.columns:
        if c in fechas:
            salida[c] = ("dateTime", "type date")
        elif pd.api.types.is_integer_dtype(df[c]):
            salida[c] = ("int64", "Int64.Type")
        elif pd.api.types.is_numeric_dtype(df[c]):
            salida[c] = ("double", "type number")
        else:
            salida[c] = ("string", "type text")
    return salida


def consulta_m(archivo: str, tp: dict) -> str:
    cols = ", ".join(f'{{"{c}", {t[1]}}}' for c, t in tp.items())
    ruta = archivo.replace("/", "\\")
    return ("let\n"
            f'    Origen = Csv.Document(File.Contents(RutaDatos & "\\{ruta}"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),\n'
            "    Encabezados = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),\n"
            f'    Resultado = Table.TransformColumnTypes(Encabezados, {{{cols}}}, "en-US")\n'
            "in\n    Resultado")


def tabla_tmdl(nombre: str, archivo: str, fechas: tuple) -> str:
    tp = tipos(pd.read_csv(p.ROOT / archivo, encoding="utf-8"), fechas)
    partes = [f"table {q(nombre)}", f"\tlineageTag: {tag(nombre)}", ""]
    for c, (dt, _) in tp.items():
        partes += [f"\tcolumn {q(c)}", f"\t\tdataType: {dt}"]
        if dt == "dateTime":
            partes.append("\t\tformatString: yyyy-mm-dd")
        elif (nombre, c) in FORMATO:
            partes.append(f"\t\tformatString: {FORMATO[(nombre, c)]}")
        partes += [f"\t\tlineageTag: {tag(nombre, c)}", "\t\tsummarizeBy: none", f"\t\tsourceColumn: {c}"]
        if (nombre, c) in ORDEN:
            partes.append(f"\t\tsortByColumn: {ORDEN[(nombre, c)]}")
        partes += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
    partes += [f"\tpartition {q(nombre)} = m", "\t\tmode: import", "\t\tsource ="]
    partes.append(sangrar(consulta_m(archivo, tp), 4))
    return "\n".join(partes) + "\n"


def medidas_tmdl() -> str:
    partes = [f"table {TABLA_MEDIDAS}", f"\tlineageTag: {tag('Medidas')}", ""]
    for nombre, dax, formato in MEDIDAS:
        partes += [f"\tmeasure {q(nombre)} = {dax}", f"\t\tformatString: {formato}", f"\t\tlineageTag: {tag('Medidas', nombre)}", ""]
    partes += ["\tcolumn _", "\t\tdataType: string", "\t\tisHidden", f"\t\tlineageTag: {tag('Medidas', '_')}",
               "\t\tsummarizeBy: none", "\t\tsourceColumn: _", "", "\t\tannotation SummarizationSetBy = Automatic", "",
               f"\tpartition {TABLA_MEDIDAS} = m", "\t\tmode: import", "\t\tsource ="]
    partes.append(sangrar('let\n    Resultado = #table(type table [_ = text], {})\nin\n    Resultado', 4))
    return "\n".join(partes) + "\n"


def escribir_modelo(raiz: Path) -> None:
    d = raiz / "definition"
    shutil.rmtree(d / "tables", ignore_errors=True)
    (d / "tables").mkdir(parents=True)
    for nombre, (archivo, fechas) in TABLAS.items():
        (d / "tables" / f"{nombre}.tmdl").write_text(tabla_tmdl(nombre, archivo, fechas), encoding="utf-8")
    (d / "tables" / f"{TABLA_MEDIDAS}.tmdl").write_text(medidas_tmdl(), encoding="utf-8")
    relaciones = "".join(f"relationship {tag('rel', a, b)}\n\tfromColumn: {a}.{b}\n\ttoColumn: {c}.{e}\n\n" for a, b, c, e in RELACIONES)
    (d / "relationships.tmdl").write_text(relaciones, encoding="utf-8")
    ruta = str(p.ROOT).replace('"', '""')
    (d / "expressions.tmdl").write_text(
        f'expression RutaDatos = "{ruta}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]\n'
        f"\tlineageTag: {tag('RutaDatos')}\n\n\tannotation PBI_ResultType = Text\n", encoding="utf-8")
    refs = "\n".join(f"ref table {q(t)}" for t in [*TABLAS, TABLA_MEDIDAS])
    (d / "model.tmdl").write_text(f"""model Model
\tculture: en-US
\tdefaultPowerBIDataSourceVersion: powerBI_V3
\tsourceQueryCulture: en-US
\tvalueFilterBehavior: independent
\tdataAccessOptions
\t\tlegacyRedirects
\t\treturnErrorValuesAsNull

annotation __PBI_TimeIntelligenceEnabled = 0

annotation PBI_ProTooling = ["DevMode"]

{refs}

ref cultureInfo en-US
""", encoding="utf-8")


# --- Informe ---

def lit(v: object) -> dict:
    if isinstance(v, bool):
        valor = "true" if v else "false"
    elif isinstance(v, int):
        valor = f"{v}L"
    elif isinstance(v, float):
        valor = f"{v}D"
    else:
        valor = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": valor}}}


def color(c: str) -> dict:
    return {"solid": {"color": lit(c)}}


def col(t: str, c: str, titulo: str | None = None) -> dict:
    return {"campo": {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": c}}, "ref": f"{t}.{c}", "titulo": titulo}


def med(m: str, titulo: str | None = None) -> dict:
    return {"campo": {"Measure": {"Expression": {"SourceRef": {"Entity": TABLA_MEDIDAS}}, "Property": m}}, "ref": f"{TABLA_MEDIDAS}.{m}", "titulo": titulo}


def proyeccion(f: dict) -> dict:
    salida = {"field": f["campo"], "queryRef": f["ref"], "nativeQueryRef": f["ref"].replace(".", "_")}
    if f["titulo"]:
        salida["displayName"] = f["titulo"]
    return salida


def filtro_lista(nombre: str, t: str, c: str, valores: list) -> dict:
    return {"name": nombre, "field": {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": c}}, "type": "Categorical",
            "filter": {"Version": 2, "From": [{"Name": "t", "Entity": t, "Type": 0}],
                       "Where": [{"Condition": {"In": {
                           "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "t"}}, "Property": c}}],
                           "Values": [[{"Literal": lit(v)["expr"]["Literal"]}] for v in valores]}}}]},
            "howCreated": "User"}


def colores_serie(t: str, c: str, mapa: dict) -> list:
    return [{"properties": {"fill": color(h)},
             "selector": {"data": [{"scopeId": {"Comparison": {
                 "ComparisonKind": 0, "Left": {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": c}},
                 "Right": lit(v)["expr"]}}}]}} for v, h in mapa.items()]


def tabla_legible(campos: list, ancho: int, fuente: float = 11.0, aire: int = 4, primera: int | None = None) -> dict:
    util = ancho - 44
    resto = (util - primera) // (len(campos) - 1) if primera else util // len(campos)
    anchos = [primera or resto] + [resto] * (len(campos) - 1)
    return {"total": [{"properties": {"totals": lit(False)}}], "columnHeaders": [{"properties": {"fontSize": lit(fuente), "wordWrap": lit(True)}}],
            "values": [{"properties": {"fontSize": lit(fuente)}}], "grid": [{"properties": {"rowPadding": lit(aire), "textSize": lit(fuente)}}],
            "columnWidth": [{"properties": {"value": lit(float(a))}, "selector": {"metadata": f["ref"]}} for f, a in zip(campos, anchos)]}


LINEA_CON_MARCAS = {"lineStyles": [{"properties": {"lineChartType": lit("straight"), "showMarker": lit(True)}}]}
LINEA_RECTA = {"lineStyles": [{"properties": {"lineChartType": lit("straight")}}]}


def ejes(eje_x: str | None = None, eje_y: str | None = None, leyenda: bool = True) -> dict:
    obj: dict = {"valueAxis": [{"properties": {"show": lit(True), "showAxisTitle": lit(bool(eje_y)), "gridlineColor": color(LINEA)}}],
                 "categoryAxis": [{"properties": {"showAxisTitle": lit(bool(eje_x))}}]}
    if eje_x:
        obj["categoryAxis"][0]["properties"]["titleText"] = lit(eje_x)
    if eje_y:
        obj["valueAxis"][0]["properties"]["titleText"] = lit(eje_y)
    if leyenda:
        obj["legend"] = [{"properties": {"show": lit(True), "position": lit("Top"), "showTitle": lit(False)}}]
    return obj


class Pagina:
    def __init__(self, nombre: str, titulo: str) -> None:
        self.nombre, self.titulo, self.visuales = nombre, titulo, []

    def agregar(self, tipo: str, x: int, y: int, w: int, h: int, roles: dict | None = None, titulo: str | None = None,
                objetos: dict | None = None, filtros: list | None = None, orden: list | None = None, contenedor: dict | None = None) -> None:
        n = len(self.visuales)
        v: dict = {"visualType": tipo, "drillFilterOtherVisuals": True}
        if roles:
            v["query"] = {"queryState": {r: {"projections": [proyeccion(f) for f in fs]} for r, fs in roles.items()}}
            if orden:
                v["query"]["sortDefinition"] = {"sort": [{"field": f["campo"], "direction": d} for f, d in orden], "isDefaultSort": False}
        if objetos:
            v["objects"] = objetos
        cont = dict(contenedor or {})
        if tipo != "textbox":
            cont["subTitle"] = [{"properties": {"show": lit(False)}}]
        if titulo:
            cont["title"] = [{"properties": {"show": lit(True), "text": lit(titulo)}}]
        elif tipo != "textbox":
            cont.setdefault("title", [{"properties": {"show": lit(False)}}])
        if cont:
            v["visualContainerObjects"] = cont
        cuerpo = {"$schema": VC, "name": f"{self.nombre}_{n:02d}", "position": {"x": x, "y": y, "z": n * 100, "height": h, "width": w, "tabOrder": n * 100}, "visual": v}
        if filtros:
            cuerpo["filterConfig"] = {"filters": filtros}
        self.visuales.append(cuerpo)

    def texto(self, x: int, y: int, w: int, h: int, parrafos: list, tarjeta: bool = False) -> None:
        cont = {} if tarjeta else {"background": [{"properties": {"show": lit(False)}}], "border": [{"properties": {"show": lit(False)}}]}
        self.agregar("textbox", x, y, w, h, objetos={"general": [{"properties": {"paragraphs": parrafos}}]}, contenedor=cont)


def corrida(texto: str, pt: int, c: str = TEXTO, negrita: bool = False) -> dict:
    return {"value": texto, "textStyle": {"fontSize": f"{pt}pt", "color": c, "fontFamily": "Segoe UI Semibold" if negrita else "Segoe UI"}}


def encabezado(pg: Pagina, titulo: str, bajada: str, ancho: int = 1000) -> None:
    pg.texto(24, 12, ancho, 66, [{"textRuns": [corrida(titulo, 20, negrita=True)]}, {"textRuns": [corrida(bajada, 11, SECUNDARIO)]}])


def pie(pg: Pagina, nota: str) -> None:
    pg.texto(24, 692, 1232, 24, [{"textRuns": [corrida(nota, 9, MUTED)]}])


def selector_empresa(pg: Pagina) -> None:
    pg.agregar("slicer", 1036, 16, 220, 56, {"Values": [col("dim_empresa", "empresa", "Empresa")]},
               objetos={"data": [{"properties": {"mode": lit("Dropdown")}}], "header": [{"properties": {"show": lit(True), "text": lit("Empresa")}}]},
               contenedor={"background": [{"properties": {"show": lit(False)}}], "border": [{"properties": {"show": lit(False)}}]})


def linea_indicador(pg: Pagina, x: int, y: int, w: int, h: int, medida: str, titulo: str) -> None:
    t = "indicadores"
    pg.agregar("lineChart", x, y, w, h, {"Category": [col(t, "anio", "Año")], "Y": [med(medida)], "Series": [col("dim_empresa", "empresa", "Empresa")]},
               titulo, {**ejes(), **LINEA_CON_MARCAS, "dataPoint": colores_serie("dim_empresa", "empresa", COLOR_EMPRESA)},
               orden=[(col(t, "anio"), "Ascending")])


def resumen(d: dict, modelo: pd.Series) -> Pagina:
    pg = Pagina("resumen", "Resumen")
    encabezado(pg, "¿Cuándo conviene una ON y cuándo un crédito?", "Pampa Energía, YPF y Vista, 2023-2025: balances, deuda, emisiones y estrés", 1232)
    t = "rating"
    campos = [col(t, "empresa", "Empresa"), col(t, "rating_2025", "Rating 2025"), col(t, "rating_estres", "Rating bajo estrés"),
              col(t, "deuda_neta_ebitda", "Deuda neta / EBITDA"), col(t, "ebitda_intereses", "EBITDA / intereses"),
              col(t, "liquidez_corriente", "Liquidez"), col(t, "on_deuda", "ON / deuda")]
    pg.agregar("tableEx", 24, 86, 1232, 190, {"Values": campos}, "Rating interno y ratios de 2025", objetos=tabla_legible(campos, 1232, 13.0, 8))
    pa = d["Pampa"]
    linea = 10 * int(memo.PARTE_DEL_BANCO * memo.capacidad(pa) // 10)
    descubiertas = [e for e in p.EMPRESAS if min(d[e]["estres"]["dscr"]) < 1]
    hallazgos = [
        ("Las tres se financian en el mercado", f"Entre {memo.pct(min(d[e]['ratio']['ON / deuda financiera'] for e in p.EMPRESAS))} y "
         f"{memo.pct(max(d[e]['ratio']['ON / deuda financiera'] for e in p.EMPRESAS))} de su deuda son ON. El banco aporta corto plazo y prefinanciación."),
        ("El riesgo país decide el instrumento", f"Con 2.000 puntos, la probabilidad de emitir dólar linked es {memo.pct(modelo['probabilidad_con_2000'])}; "
         f"con 700, {memo.pct(modelo['probabilidad_con_700'])}."),
        ("El riesgo es refinanciar", f"Bajo estrés, {' y '.join(descubiertas)} no cubren sus vencimientos con caja propia. "
         f"Recomendación: hasta US$ {memo.numero(linea, 0)} M a 3 años para Pampa."),
    ]
    for i, (titulo, texto) in enumerate(hallazgos):
        pg.texto(24 + i * 416, 292, 400, 150, [{"textRuns": [corrida(titulo, 15, negrita=True)]}, {"textRuns": [corrida(texto, 13, SECUNDARIO)]}], tarjeta=True)
    linea_indicador(pg, 24, 458, 610, 226, "Deuda neta / EBITDA", "Deuda neta / EBITDA")
    linea_indicador(pg, 646, 458, 610, 226, "Flujo libre / deuda", "Flujo libre / deuda: las tres invierten más de lo que generan")
    pie(pg, "Fuentes: 20-F 2025 de Pampa e YPF, estados contables de Vista, ArgentinaDatos, FRED, BCRA y datos.gob.ar. Ratings: escala interna, sin techo soberano.")
    return pg


def balances() -> Pagina:
    pg = Pagina("balances", "Balances")
    encabezado(pg, "Rentabilidad, endeudamiento y liquidez", "Indicadores calculados en SQL sobre los estados contables, con la misma definición para las tres")
    selector_empresa(pg)
    graficos = [("Margen EBITDA", "Margen EBITDA"), ("EBITDA / intereses", "EBITDA / intereses"),
                ("Liquidez corriente", "Liquidez corriente"), ("Capex / ventas", "Capex / ventas"),
                ("ON / deuda", "ON / deuda financiera"), ("Deuda en dólares", "Deuda en dólares / deuda financiera")]
    for i, (m, titulo) in enumerate(graficos):
        linea_indicador(pg, 24 + (i % 3) * 416, 84 + (i // 3) * 304, 400, 292, m, titulo)
    pie(pg, "EBITDA propio: sin asociadas ni resultados no recurrentes. Pampa y Vista no publican el detalle de deuda de 2023.")
    return pg


def financiamiento(modelo: pd.Series) -> Pagina:
    pg = Pagina("financiamiento", "Financiamiento")
    encabezado(pg, "Con riesgo país alto, emitieron dólar linked", "42 ON de 2021 a 2025, con el riesgo país del día de emisión")
    selector_empresa(pg)
    pg.agregar("lineChart", 24, 84, 760, 300, {"Category": [col("mercado", "fecha", "Fecha")], "Y": [med("Riesgo país")]},
               "Riesgo país diario (puntos básicos)", {**ejes(leyenda=False), **LINEA_RECTA, "dataPoint": [{"properties": {"fill": color("#898781")}}]},
               orden=[(col("mercado", "fecha"), "Ascending")])
    t = "emisiones"
    pg.agregar("scatterChart", 796, 84, 460, 300, {"Category": [col(t, "clase", "Clase")], "Series": [col(t, "instrumento", "Instrumento")],
                                                    "X": [med("Riesgo país al emitir", "Riesgo país")], "Y": [med("Tasa media ponderada", "Tasa")]},
               f"Ninguna dólar linked bajo 1.500 puntos; probabilidad con 2.000: {memo.pct(modelo['probabilidad_con_2000'])}",
               {**ejes("Riesgo país el día de emisión", "Tasa"), "dataPoint": colores_serie(t, "instrumento", COLOR_INSTRUMENTO)})
    pg.agregar("clusteredColumnChart", 24, 396, 500, 290, {"Category": [col(t, "anio", "Año")], "Y": [med("Monto emitido", "US$ M")], "Series": [col(t, "instrumento", "Instrumento")]},
               "Monto emitido por año e instrumento (US$ M)", {**ejes(), "dataPoint": colores_serie(t, "instrumento", COLOR_INSTRUMENTO)},
               orden=[(col(t, "anio"), "Ascending")])
    campos = [col(t, "empresa", "Empresa"), col(t, "clase", "Clase"), col(t, "fecha", "Fecha"), col(t, "instrumento", "Instrumento"),
              col(t, "monto", "US$ M"), col(t, "tasa", "Tasa"), col(t, "riesgo_pais", "Riesgo país"), col(t, "prima_riesgo_pais_pb", "Prima vs. riesgo país (pb)")]
    pg.agregar("tableEx", 536, 396, 720, 290, {"Values": campos}, "Detalle de emisiones", objetos=tabla_legible(campos, 720, 10.0, 2, 70),
               orden=[(col(t, "fecha"), "Descending")])
    pie(pg, "Tasa: rendimiento de emisión si se informa; si no, cupón. Fechas de Vista e YPF: mes informado. Prima negativa: la empresa se financió más barato que el Estado.")
    return pg


def estres(d: dict) -> Pagina:
    pg = Pagina("estres", "Estrés")
    s = d["supuestos"]
    encabezado(pg, "¿Puede pagar si deja de crecer y se le encarece la deuda?",
               f"Estrés: ventas {memo.pct(s['caida'])} en 2026, margen {memo.numero(100 * s['margen'], 0)} pp menor y deuda nueva {memo.numero(100 * s['tasa'], 0)} pp más cara", 1232)
    t = "proyeccion"
    for i, (m, titulo) in enumerate((("Proyección deuda neta / EBITDA", "Pampa: deuda neta / EBITDA, con covenant supuesto en 3,5x"),
                                     ("Proyección cobertura de servicio", "Pampa: cobertura de servicio de deuda (debajo de 1x depende de refinanciar)"))):
        pg.agregar("lineChart", 24 + i * 624, 84, 608, 300, {"Category": [col(t, "anio", "Año")], "Y": [med(m)], "Series": [col(t, "escenario", "Escenario")]},
                   titulo, {**ejes(), **LINEA_CON_MARCAS, "dataPoint": colores_serie(t, "escenario", COLOR_ESCENARIO)},
                   filtros=[filtro_lista("empresa", "dim_empresa", "empresa", ["Pampa"])], orden=[(col(t, "anio"), "Ascending")])
    campos = [col("rating", "empresa", "Empresa"), col("rating", "peor_apalancamiento_estres", "Peor deuda neta / EBITDA"),
              col("rating", "peor_cobertura_estres", "Peor EBITDA / intereses"), col("rating", "peor_servicio_estres", "Peor cobertura de servicio"),
              col("rating", "ebitda_puede_caer", "EBITDA puede caer (base)"), col("rating", "rating_estres", "Rating bajo estrés")]
    pg.agregar("tableEx", 24, 396, 1232, 290, {"Values": campos}, "Peor año del escenario de estrés, por empresa", objetos=tabla_legible(campos, 1232, 13.0, 8))
    pie(pg, "Base: ventas y margen de 2025 y capex igual a la depreciación. La necesidad de fondos se cubre con deuda nueva. Los gráficos muestran Pampa, la empresa del memo.")
    return pg


def deuda(d: dict) -> Pagina:
    pg = Pagina("deuda", "Deuda y caja")
    venc = {e: d["vencimientos"][e] for e in p.EMPRESAS}
    pa, yp = venc["Pampa"], venc["YPF"]
    encabezado(pg, "Cuándo vence la deuda y adónde se va la plata",
               f"Con su caja, Pampa paga todo lo que vence hasta {pa['primer_hueco'] - 1}; a YPF le vence en {p.PROYECCION[0]} {memo.veces(yp['primer_anio'] / yp['caja'])} su caja")
    selector_empresa(pg)
    t = "vencimientos"
    pg.agregar("clusteredColumnChart", 24, 84, 760, 300, {"Category": [col(t, "anio", "Año")], "Y": [med("% de la deuda", "% de la deuda")], "Series": [col("dim_empresa", "empresa", "Empresa")]},
               "Qué parte de la deuda de 2025 vence cada año", {**ejes(), "dataPoint": colores_serie("dim_empresa", "empresa", COLOR_EMPRESA)},
               orden=[(col(t, "anio"), "Ascending")])
    campos = [col("dim_empresa", "empresa", "Empresa"), med("Rentabilidad sobre patrimonio", "ROE"), med("Margen neto", "Margen neto"),
              med("Rotación de activos", "Rotación"), med("Activo / patrimonio", "Apalancamiento")]
    pg.agregar("tableEx", 796, 84, 460, 300, {"Values": campos}, "DuPont 2025: ROE = margen neto x rotación x apalancamiento",
               objetos=tabla_legible(campos, 460, 11.0, 8, 90), filtros=[filtro_lista("anio", "indicadores", "anio", [p.ANIOS[-1]])])
    t = "puente"
    pg.agregar("waterfallChart", 24, 396, 610, 290, {"Category": [col(t, "concepto", "Concepto")], "Y": [med("Monto del puente", "US$ M")]},
               "Del EBITDA al flujo libre en 2025 (US$ M; elegí una empresa)", {**ejes(leyenda=False)},
               filtros=[filtro_lista("anio", t, "anio", [p.ANIOS[-1]])], orden=[(col(t, "concepto"), "Ascending")])
    campos = [col("rating", "empresa", "Empresa"), col("rating", "caja_2025", "Caja e inversiones"), col("rating", "vence_primer_anio", f"Vence en {p.PROYECCION[0]}"),
              col("rating", "caja_sobre_vencimientos", f"Caja / vencimientos {p.PROYECCION[0]}"), col("rating", "anios_cubiertos", "Años de vencimientos que cubre la caja")]
    pg.agregar("tableEx", 646, 396, 610, 290, {"Values": campos}, "Si el mercado se cierra: hasta cuándo alcanza la caja", objetos=tabla_legible(campos, 610, 12.0, 8))
    pie(pg, "Vencimientos: notas de deuda de los 20-F y de Vista. YPF informa sólo el año de vencimiento de cada ON y reparte sus préstamos por rango de años.")
    return pg


def bancos() -> Pagina:
    from empresas import guia
    x = guia.cifras()
    co, si = x["colocacion"], x["sistema"]
    pg = Pagina("bancos", "Bancos y fintech")
    encabezado(pg, "Qué lugar le queda al banco y a la fintech",
               f"Colocadores de {co['emisiones']} ON con aviso disponible: el banco coloca y distribuye; presta sobre todo a corto plazo", 1232)
    t = "colocadores_entidad"
    tipos = {"Banco": "#2a78d6", "ALyC": "#898781", "ALyC de un banco": "#6f6e69", "Plataforma digital": "#1baf7a", "Banco de inversión del exterior": "#eda100"}
    pg.agregar("clusteredBarChart", 24, 84, 610, 450, {"Category": [col(t, "entidad", "Entidad")], "Y": [med("Participaciones", "Emisiones")], "Series": [col(t, "tipo_entidad", "Tipo")]},
               f"En cuántas de las {co['emisiones']} emisiones participó cada entidad", {**ejes(), "dataPoint": colores_serie(t, "tipo_entidad", tipos)},
               orden=[(med("Participaciones"), "Descending")])
    t = "colocadores_emision"
    campos = [col(t, "empresa", "Empresa"), col(t, "clase", "Clase"), col(t, "monto", "US$ M"), col(t, "bancos", "Bancos"), col(t, "alycs", "ALyC"),
              col(t, "plataformas_digitales", "Plataformas"), col(t, "bancos_exterior", "Exterior")]
    pg.agregar("tableEx", 646, 84, 610, 230, {"Values": campos}, "Entidades por tipo en cada emisión", objetos=tabla_legible(campos, 610, 10.0, 3))
    textos = [("Empresa", f"Dólar duro desde {memo.numero(x['duro']['rp_min'], 0)} puntos de riesgo país; dólar linked sólo arriba de {memo.numero(x['linked']['rp_min'], 0)}. "
                          "El banco, para plazos cortos y prefinanciación."),
              ("Banco", f"Coloca sin asumir el riesgo: {' y '.join(co['siempre'])} estuvieron en todas. Prestando en dólares ganó unos "
                        f"{memo.numero(si['credito'] - si['pf'], 1)} puntos sobre el plazo fijo en {p.ANIOS[-1]}."),
              ("Fintech", f"No les presta: distribuye. Las plataformas digitales estuvieron en {co['con_plataforma']} de las {co['locales']} ON locales de la muestra.")]
    for i, (titulo, texto) in enumerate(textos):
        pg.texto(646, 326 + i * 70, 610, 64, [{"textRuns": [corrida(titulo, 12, negrita=True)]}, {"textRuns": [corrida(texto, 10, SECUNDARIO)]}], tarjeta=True)
    pie(pg, f"Muestra: {co['emisiones']} de las 42 ON, las que tienen aviso de resultados o suplemento disponible. Comprador inicial: banco de inversión que compra la ON para revenderla afuera.")
    return pg


def tema() -> dict:
    return {
        "name": "Análisis crediticio", "dataColors": list(COLOR_EMPRESA.values()) + ["#e34948", "#6250d6"], "background": SUPERFICIE, "foreground": TEXTO,
        "tableAccent": "#2a78d6",
        "textClasses": {"title": {"fontFace": "Segoe UI Semibold", "fontSize": 12, "color": TEXTO},
                        "label": {"fontFace": "Segoe UI", "fontSize": 10, "color": SECUNDARIO},
                        "callout": {"fontFace": "Segoe UI Semibold", "fontSize": 26, "color": TEXTO}},
        "visualStyles": {
            "*": {"*": {"background": [{"show": True, "color": {"solid": {"color": SUPERFICIE}}, "transparency": 0}],
                        "border": [{"show": True, "color": {"solid": {"color": LINEA}}, "radius": 8}],
                        "title": [{"titleWrap": True, "alignment": "left", "fontSize": 12, "fontFamily": "Segoe UI Semibold", "fontColor": {"solid": {"color": TEXTO}}}]}},
            "page": {"*": {"background": [{"color": {"solid": {"color": FONDO}}, "transparency": 0}],
                           "outspace": [{"color": {"solid": {"color": FONDO}}, "transparency": 0}]}},
            "tableEx": {"*": {"grid": [{"gridVertical": False, "rowPadding": 4}]}},
        },
    }


def escribir_informe(raiz: Path, paginas: list[Pagina]) -> None:
    recursos = raiz / "StaticResources" / "RegisteredResources"
    recursos.mkdir(parents=True, exist_ok=True)
    (recursos / f"{NOMBRE}.json").write_text(json.dumps(tema(), ensure_ascii=False, indent=2), encoding="utf-8")
    f = raiz / "definition" / "report.json"
    rep = json.loads(f.read_text(encoding="utf-8"))
    version = rep["themeCollection"]["baseTheme"]["reportVersionAtImport"]
    rep["themeCollection"]["customTheme"] = {"name": f"{NOMBRE}.json", "reportVersionAtImport": version, "type": "RegisteredResources"}
    rep["resourcePackages"] = [r for r in rep["resourcePackages"] if r["name"] != "RegisteredResources"]
    rep["resourcePackages"].append({"name": "RegisteredResources", "type": "RegisteredResources",
                                    "items": [{"name": f"{NOMBRE}.json", "path": f"{NOMBRE}.json", "type": "CustomTheme"}]})
    f.write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    d = raiz / "definition" / "pages"
    shutil.rmtree(d, ignore_errors=True)
    d.mkdir()
    (d / "pages.json").write_text(json.dumps({"$schema": PGS, "pageOrder": [x.nombre for x in paginas], "activePageName": paginas[0].nombre}, indent=2), encoding="utf-8")
    for x in paginas:
        carpeta = d / x.nombre
        (carpeta / "visuals").mkdir(parents=True)
        (carpeta / "page.json").write_text(json.dumps({"$schema": PG, "name": x.nombre, "displayName": x.titulo, "displayOption": "FitToPage",
                                                       "height": 720, "width": 1280}, ensure_ascii=False, indent=2), encoding="utf-8")
        for v in x.visuales:
            (carpeta / "visuals" / v["name"]).mkdir()
            (carpeta / "visuals" / v["name"] / "visual.json").write_text(json.dumps(v, ensure_ascii=False, indent=2), encoding="utf-8")


def preparar_esqueleto(origen: Path) -> None:
    """Copia, una sola vez, el esqueleto de un proyecto creado por Power BI Desktop, con nombre e identidad propios."""
    if (SALIDA / f"{NOMBRE}.pbip").exists():
        return
    base = origen.name.split(".")[0] if origen.suffix else origen.name
    for tipo in ("Report", "SemanticModel"):
        destino = SALIDA / f"{NOMBRE}.{tipo}"
        shutil.copytree(origen.parent / f"{base}.{tipo}", destino,
                        ignore=shutil.ignore_patterns(".pbi", "pages", "tables", "RegisteredResources", "diagramLayout.json", "cultures"))
        plataforma = json.loads((destino / ".platform").read_text(encoding="utf-8"))
        plataforma["metadata"]["displayName"] = NOMBRE
        plataforma["config"]["logicalId"] = tag(tipo)
        (destino / ".platform").write_text(json.dumps(plataforma, indent=2), encoding="utf-8")
    (SALIDA / f"{NOMBRE}.SemanticModel" / "definition" / "cultures").mkdir()
    (SALIDA / f"{NOMBRE}.SemanticModel" / "definition" / "cultures" / "en-US.tmdl").write_text("cultureInfo en-US\n", encoding="utf-8")
    f = SALIDA / f"{NOMBRE}.Report" / "definition.pbir"
    f.write_text(f.read_text(encoding="utf-8").replace(f"../{base}.SemanticModel", f"../{NOMBRE}.SemanticModel"), encoding="utf-8")
    pbip = (origen.parent / f"{base}.pbip").read_text(encoding="utf-8").replace(f"{base}.Report", f"{NOMBRE}.Report")
    (SALIDA / f"{NOMBRE}.pbip").write_text(pbip, encoding="utf-8")


def main() -> None:
    SALIDA.mkdir(exist_ok=True)
    preparar_esqueleto(Path(r"F:\ciclo-credito\powerbi\umbrales_credito"))
    d = memo.leer()
    modelo = pd.read_csv(OUTPUTS / "modelo_instrumento.csv").set_index("medida").valor
    escribir_modelo(SALIDA / f"{NOMBRE}.SemanticModel")
    paginas = [resumen(d, modelo), balances(), deuda(d), financiamiento(modelo), bancos(), estres(d)]
    escribir_informe(SALIDA / f"{NOMBRE}.Report", paginas)
    print(f"powerbi: {len(TABLAS) + 1} tablas, {len(MEDIDAS)} medidas, {sum(len(x.visuales) for x in paginas)} visuales")


if __name__ == "__main__":
    main()
