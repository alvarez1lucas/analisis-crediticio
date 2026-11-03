import json
import re

import pandas as pd

from empresas import plantilla as p
from empresas import powerbi as pbi

MODELO = pbi.SALIDA / f"{pbi.NOMBRE}.SemanticModel" / "definition"
PAGINAS = pbi.SALIDA / f"{pbi.NOMBRE}.Report" / "definition" / "pages"


def columnas() -> dict[str, set[str]]:
    return {t: set(pd.read_csv(p.ROOT / archivo, nrows=1).columns) for t, (archivo, _) in pbi.TABLAS.items()}


def test_cada_tabla_del_modelo_existe_y_lee_su_csv():
    for t, (archivo, _) in pbi.TABLAS.items():
        texto = (MODELO / "tables" / f"{t}.tmdl").read_text(encoding="utf-8")
        assert archivo.replace("/", "\\") in texto
        assert (p.ROOT / archivo).exists()


def test_las_medidas_usan_columnas_que_existen():
    cols = columnas()
    for nombre, dax, _ in pbi.MEDIDAS:
        for tabla, columna in re.findall(r"(\w+)\[(\w+)\]", dax):
            assert columna in cols[tabla], (nombre, tabla, columna)


def test_cada_campo_de_cada_visual_existe():
    cols = columnas()
    medidas = {m for m, _, _ in pbi.MEDIDAS}
    total = 0
    for f in PAGINAS.glob("*/visuals/*/visual.json"):
        v = json.loads(f.read_text(encoding="utf-8"))
        total += 1
        for rol in v["visual"].get("query", {}).get("queryState", {}).values():
            for pr in rol["projections"]:
                campo = pr["field"]
                if "Measure" in campo:
                    assert campo["Measure"]["Property"] in medidas, (f.parent.name, campo)
                else:
                    c = campo["Column"]
                    assert c["Property"] in cols[c["Expression"]["SourceRef"]["Entity"]], (f.parent.name, c)
    assert total == 36


def test_los_nombres_internos_no_se_repiten_dentro_de_un_visual():
    for f in PAGINAS.glob("*/visuals/*/visual.json"):
        v = json.loads(f.read_text(encoding="utf-8"))
        nativos = [pr["nativeQueryRef"] for rol in v["visual"].get("query", {}).get("queryState", {}).values() for pr in rol["projections"]]
        assert len(nativos) == len(set(nativos)), f.parent.name
