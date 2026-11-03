-- Indicadores por empresa y año, en formato largo para el tablero. NULL cuando el denominador es cero o negativo
-- o falta el dato, igual que "n.d." y "n.s." en el Excel.
CREATE OR REPLACE TABLE indicadores AS
WITH base AS (
    SELECT
        empresa,
        anio,
        ebitda / NULLIF(ventas, 0)                                              AS margen_ebitda,
        CASE WHEN ebitda > 0 THEN deuda_neta / ebitda END                       AS deuda_neta_ebitda,
        CASE WHEN gasto_int > 0 THEN ebitda / gasto_int END                     AS ebitda_intereses,
        CASE WHEN pc > 0 THEN ac / pc END                                       AS liquidez_corriente,
        CASE WHEN deuda > 0 THEN fcf / deuda END                                AS flujo_libre_deuda,
        CASE WHEN deuda_cp + deuda_lp > 0 AND d_on + d_bancos + d_prefin + d_otra > 0
             THEN d_on / (deuda_cp + deuda_lp) END                              AS on_deuda,
        CASE WHEN deuda_cp + deuda_lp > 0 AND m_dolar + m_linked + m_pesos + m_otra > 0
             THEN m_dolar / (deuda_cp + deuda_lp) END                           AS dolares_deuda,
        capex / NULLIF(ventas, 0)                                               AS capex_ventas
    FROM lineas
)
UNPIVOT base
ON margen_ebitda, deuda_neta_ebitda, ebitda_intereses, liquidez_corriente, flujo_libre_deuda, on_deuda, dolares_deuda, capex_ventas
INTO NAME indicador VALUE valor
ORDER BY empresa, indicador, anio;
