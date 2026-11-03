-- Cómo cambió la forma de endeudarse: por empresa y año de emisión, ponderado por monto.
CREATE OR REPLACE TABLE tendencias AS
SELECT
    empresa,
    anio,
    count(*)                                                                    AS emisiones,
    sum(monto)                                                                  AS monto,
    sum(monto) FILTER (WHERE instrumento = 'Dólar linked') / sum(monto)         AS parte_dolar_linked,
    sum(monto * plazo) / sum(monto)                                             AS plazo_medio,
    sum(monto * tasa) / sum(monto)                                              AS tasa_media,
    sum(monto * tasa) FILTER (WHERE instrumento = 'Dólar duro')
        / sum(monto) FILTER (WHERE instrumento = 'Dólar duro')                  AS tasa_media_dolar_duro,
    sum(monto * riesgo_pais) / sum(monto)                                       AS riesgo_pais_medio
FROM emisiones
GROUP BY empresa, anio
ORDER BY empresa, anio;

-- Riesgo país promedio de cada año, para leer las tendencias contra el mercado.
CREATE OR REPLACE TABLE contexto AS
WITH m AS (
    SELECT year(fecha) AS anio, avg(riesgo_pais) AS riesgo_pais, avg(ust_5) AS tesoro_5
    FROM mercado
    GROUP BY 1
), t AS (
    SELECT year(CAST(mes AS DATE)) AS anio, avg(doc_usd) AS credito_usd, avg(pf_usd) AS plazo_fijo_usd
    FROM tasas
    GROUP BY 1
)
SELECT m.*, t.credito_usd, t.plazo_fijo_usd
FROM m
LEFT JOIN t USING (anio)
ORDER BY anio;
