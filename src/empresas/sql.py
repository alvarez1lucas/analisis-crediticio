"""Capa SQL con DuckDB.

`python -m empresas.sql` corre los scripts de sql/ en orden sobre los CSV de data/ y exporta las tablas finales a outputs/,
que son las que usan el tablero y el sitio.
"""

from __future__ import annotations

import duckdb

from empresas.plantilla import ROOT

SQL = ROOT / "sql"
OUTPUTS = ROOT / "outputs"
TABLAS = ["lineas", "indicadores", "emisiones", "tendencias", "contexto"]


def conectar() -> duckdb.DuckDBPyConnection:
    con = duckdb.connect()
    con.execute(f"SET file_search_path = '{ROOT.as_posix()}'")
    for script in sorted(SQL.glob("*.sql")):
        con.execute(script.read_text(encoding="utf-8"))
    return con


def main() -> None:
    OUTPUTS.mkdir(exist_ok=True)
    con = conectar()
    for tabla in TABLAS:
        destino = OUTPUTS / f"{tabla}.csv"
        con.execute(f"COPY {tabla} TO '{destino.as_posix()}' (HEADER, DELIMITER ',')")
        print(f"{tabla}: {con.execute(f'SELECT count(*) FROM {tabla}').fetchone()[0]} filas")


if __name__ == "__main__":
    main()
