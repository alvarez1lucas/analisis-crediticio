-- Muro de vencimientos: deuda financiera al 31-12-2025 por año, tipo y moneda.
CREATE OR REPLACE TABLE vencimientos AS
SELECT empresa, anio, tipo, moneda, sum(monto) AS monto
FROM read_csv('data/deuda_vigente.csv')
GROUP BY ALL
ORDER BY empresa, anio, tipo;

-- Puente del EBITDA al flujo libre: adónde se va la plata que genera la operación. En un gráfico de cascada,
-- cada paso es un monto con signo y el flujo libre es el total.
CREATE OR REPLACE TABLE puente AS
SELECT empresa, anio, 1 AS orden, 'EBITDA' AS concepto, ebitda AS monto FROM lineas
UNION ALL SELECT empresa, anio, 2, 'Impuestos y capital de trabajo', cfo - ebitda FROM lineas
UNION ALL SELECT empresa, anio, 3, 'Intereses pagados', -int_pag FROM lineas
UNION ALL SELECT empresa, anio, 4, 'Capex', -capex FROM lineas
ORDER BY empresa, anio, orden;
