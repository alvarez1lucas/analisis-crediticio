-- Balances en formato ancho: una fila por empresa y año, una columna por línea cargada.
CREATE OR REPLACE TABLE balances AS
PIVOT (
    SELECT empresa, anio, clave, TRY_CAST(valor AS DOUBLE) AS valor
    FROM read_csv('data/balances.csv', all_varchar = true)
    WHERE clave <> 'fuente'
)
ON clave
USING first(valor)
ORDER BY empresa, anio;

-- Líneas calculadas con la misma definición que el Excel.
CREATE OR REPLACE TABLE lineas AS
SELECT
    *,
    ventas - costo - comercial - admin - explor + otros_op - deterioro          AS ebit,
    ventas - costo - comercial - admin - explor + otros_op + amort - no_rec     AS ebitda,
    deuda_cp + arr_cp + deuda_lp + arr_lp                                       AS deuda,
    deuda_cp + arr_cp + deuda_lp + arr_lp - caja - inv_cp                       AS deuda_neta,
    caja + inv_cp + creditos + inventarios + otros_ac                           AS ac,
    deuda_cp + arr_cp + proveedores + otros_pc                                  AS pc,
    cfo - int_pag - capex                                                       AS fcf
FROM balances;
