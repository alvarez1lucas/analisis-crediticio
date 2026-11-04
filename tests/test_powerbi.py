import json
import re

import pandas as pd
import pytest

from empresas import plantilla as p
from empresas import powerbi as pbi


def rutas(L: str):
    nombre = pbi.NOMBRES[L]
    return pbi.SALIDA / f"{nombre}.SemanticModel" / "definition", pbi.SALIDA / f"{nombre}.Report" / "definition" / "pages"


def columnas() -> dict[str, set[str]]:
    return {t: set(pd.read_csv(p.ROOT / archivo, nrows=1).columns) for t, (archivo, _) in pbi.TABLAS.items()}


@pytest.mark.parametrize("L", ["es", "en"])
def test_cada_tabla_del_modelo_existe_y_lee_su_csv(L):
    modelo, _ = rutas(L)
    for t, (archivo, _) in pbi.TABLAS.items():
        texto = (modelo / "tables" / f"{t}.tmdl").read_text(encoding="utf-8")
        assert archivo.replace("/", "\\") in texto
        assert (p.ROOT / archivo).exists()
    assert (modelo / "tables" / f"{pbi.TABLAS_MEDIDAS[L]}.tmdl").exists()


def test_las_medidas_usan_columnas_que_existen():
    cols = columnas()
    for nombre, dax, _ in pbi.MEDIDAS:
        for tabla, columna in re.findall(r"(\w+)\[(\w+)\]", dax):
            assert columna in cols[tabla], (nombre, tabla, columna)
    assert set(pbi.MEDIDAS_EN) == {m for m, _, _ in pbi.MEDIDAS}


@pytest.mark.parametrize("L", ["es", "en"])
def test_cada_campo_de_cada_visual_existe(L):
    cols = columnas()
    medidas = {m if L == "es" else pbi.MEDIDAS_EN[m] for m, _, _ in pbi.MEDIDAS}
    _, paginas = rutas(L)
    total = 0
    for f in paginas.glob("*/visuals/*/visual.json"):
        v = json.loads(f.read_text(encoding="utf-8"))
        total += 1
        for rol in v["visual"].get("query", {}).get("queryState", {}).values():
            for pr in rol["projections"]:
                campo = pr["field"]
                if "Measure" in campo:
                    assert campo["Measure"]["Property"] in medidas, (f.parent.name, campo)
                    assert campo["Measure"]["Expression"]["SourceRef"]["Entity"] == pbi.TABLAS_MEDIDAS[L]
                else:
                    c = campo["Column"]
                    assert c["Property"] in cols[c["Expression"]["SourceRef"]["Entity"]], (f.parent.name, c)
    assert total == 43


@pytest.mark.parametrize("L", ["es", "en"])
def test_los_nombres_internos_no_se_repiten_dentro_de_un_visual(L):
    _, paginas = rutas(L)
    for f in paginas.glob("*/visuals/*/visual.json"):
        v = json.loads(f.read_text(encoding="utf-8"))
        nativos = [pr["nativeQueryRef"] for rol in v["visual"].get("query", {}).get("queryState", {}).values() for pr in rol["projections"]]
        assert len(nativos) == len(set(nativos)), f.parent.name


def test_la_version_en_ingles_traduce_los_valores_que_colorea():
    datos = {("emisiones", "instrumento"): pd.read_csv(p.ROOT / "outputs/emisiones.csv").instrumento,
             ("proyeccion", "escenario"): pd.read_csv(p.ROOT / "outputs/proyeccion.csv").escenario,
             ("colocadores_entidad", "tipo_entidad"): pd.read_csv(p.ROOT / "outputs/colocadores_entidad.csv").tipo_entidad}
    for (t, c), valores in datos.items():
        traducidos = {pbi.VALORES_EN.get((t, c), {}).get(v, v) for v in valores}
        assert not any(x in traducidos for x in ("Dólar duro", "Estrés", "Banco")), (t, c, traducidos)
