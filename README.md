# Análisis crediticio de empresas

¿Cuándo conviene una obligación negociable y cuándo un crédito, y qué lugar le queda al banco? Estudio de Pampa Energía, YPF y Vista sobre sus 20-F, avisos de resultados y series de mercado. El diseño está en [docs/diseno.md](docs/diseno.md) la conclusión, en el [memo de crédito](docs/memo_credito.md), y cuándo conviene cada instrumento, en la [guía de decisión](docs/guia_instrumentos.md).

**Sitio:** [alvarez1lucas.github.io/analisis-crediticio](https://alvarez1lucas.github.io/analisis-crediticio/) · **Tablero de Power BI:** [powerbi/analisis_crediticio.pbix](powerbi/analisis_crediticio.pbix) · **Modelo:** [excel/analisis_crediticio.xlsx](excel/analisis_crediticio.xlsx)

**Stack:** Excel, SQL con DuckDB, Python (pandas, statsmodels), Plotly, Power BI, pytest, GitHub Actions, GitHub Pages.

## Qué tiene

- **Balances:** Pampa, YPF y Vista, 2023 a 2025, cargados desde sus estados contables en dólares, con controles de cuadre contra lo informado.
- **Deuda:** 42 obligaciones negociables y 25 créditos, con riesgo país, Tesoro de EE.UU. y tasa bancaria del momento de cada emisión.
- **Análisis:** rating interno, proyección 2026-2028 con escenario de estrés, covenants y regresión logística sobre la elección de instrumento.
- **Deuda y mercado:** muro de vencimientos, puente del EBITDA al flujo libre, DuPont y colocadores de cada ON (bancos, ALyC y plataformas digitales).
- **Controles:** 34 tests; SQL, Excel y Power BI dan las mismas cifras.

## Cómo correrlo

`pip install -r requirements.txt` y, con `PYTHONPATH=src`: `empresas.mercado` (series), `empresas.plantilla` (Excel), `empresas.sql` (capa SQL en `sql/`),
`empresas.tendencias` (modelo), `empresas.memo` (memo y exportaciones), `empresas.guia` (guía), `empresas.sitio` (sitio) y `empresas.powerbi` (tablero).
GitHub Actions actualiza el mercado cada lunes. Los balances están en `data/balances.csv` y se actualizan cuando sale un nuevo 20-F.
