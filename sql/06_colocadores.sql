-- Quién coloca las ON: muestra de las emisiones con aviso de resultados o suplemento disponible.
CREATE OR REPLACE TABLE colocadores AS
SELECT c.*, e.fecha, e.monto, e.instrumento
FROM read_csv('data/colocadores.csv') c
JOIN emisiones e USING (empresa, clase);

-- Por entidad: en cuántas emisiones participó y por qué monto total de emisión.
CREATE OR REPLACE TABLE colocadores_entidad AS
SELECT
    entidad,
    tipo_entidad,
    count(*)                                                                    AS emisiones,
    count(*) FILTER (WHERE rol = 'Organizador y colocador')                     AS como_organizador,
    sum(monto)                                                                  AS monto_emisiones,
    string_agg(DISTINCT empresa, ', ' ORDER BY empresa)                         AS empresas
FROM colocadores
GROUP BY entidad, tipo_entidad
ORDER BY emisiones DESC, monto_emisiones DESC, entidad;

-- Por emisión: cuántas entidades de cada tipo participaron.
CREATE OR REPLACE TABLE colocadores_emision AS
SELECT
    empresa, clase, fecha, monto, instrumento,
    count(*) FILTER (WHERE tipo_entidad = 'Banco')                              AS bancos,
    count(*) FILTER (WHERE tipo_entidad IN ('ALyC', 'ALyC de un banco'))        AS alycs,
    count(*) FILTER (WHERE tipo_entidad = 'Plataforma digital')                 AS plataformas_digitales,
    count(*) FILTER (WHERE tipo_entidad = 'Banco de inversión del exterior')    AS bancos_exterior,
    count(*)                                                                    AS entidades
FROM colocadores
GROUP BY ALL
ORDER BY fecha;
