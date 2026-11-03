-- Cada ON con el mercado de su día: ASOF JOIN toma el último día hábil con datos anterior o igual a la emisión.
CREATE OR REPLACE TABLE mercado AS
SELECT * FROM read_csv('data/mercado.csv');

CREATE OR REPLACE TABLE tasas AS
SELECT * FROM read_csv('data/tasas.csv');

CREATE OR REPLACE TABLE emisiones AS
WITH e AS (
    SELECT
        empresa, clase, instrumento, monto, cupon, rendimiento,
        CAST(fecha AS DATE)                                                     AS fecha,
        CAST(vencimiento AS DATE)                                               AS vencimiento,
        date_diff('day', CAST(fecha AS DATE), CAST(vencimiento AS DATE)) / 365.25 AS plazo,
        coalesce(rendimiento, cupon)                                            AS tasa
    FROM read_csv('data/emisiones.csv')
)
SELECT
    e.*,
    year(e.fecha)                                                               AS anio,
    m.riesgo_pais,
    -- Tesoro del plazo inmediato inferior, como en el Excel
    CASE WHEN e.plazo >= 10 THEN m.ust_10 WHEN e.plazo >= 7 THEN m.ust_7 WHEN e.plazo >= 5 THEN m.ust_5
         WHEN e.plazo >= 3 THEN m.ust_3 WHEN e.plazo >= 2 THEN m.ust_2 ELSE m.ust_1 END AS tesoro,
    t.doc_usd_largo                                                             AS credito_sistema
FROM e
ASOF LEFT JOIN mercado m ON e.fecha >= m.fecha
LEFT JOIN tasas t ON date_trunc('month', e.fecha) = CAST(t.mes AS DATE);

CREATE OR REPLACE TABLE emisiones AS
SELECT
    *,
    CASE WHEN instrumento = 'Dólar duro' THEN (tasa * 100 - tesoro) * 100 END  AS spread_pb,
    CASE WHEN instrumento = 'Dólar duro' THEN (tasa * 100 - tesoro) * 100 - riesgo_pais END AS prima_riesgo_pais_pb,
    CASE WHEN instrumento = 'Dólar duro' THEN (tasa * 100 - credito_sistema) * 100 END AS on_menos_credito_pb
FROM emisiones
ORDER BY empresa, fecha;
