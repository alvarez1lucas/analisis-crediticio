# Análisis crediticio de empresas: ¿ON o crédito?

¿Cuándo conviene una obligación negociable y cuándo un crédito, y qué lugar les queda al banco y a la fintech? Una carpeta de crédito de Pampa Energía,
YPF y Vista armada con sus estados contables auditados en dólares, avisos de resultados de sus ON y datos de mercado. La conclusión está en el
[memo de crédito](docs/memo_credito.md) y en la [guía de decisión](docs/guia_instrumentos.md); el diseño, en [docs/diseno.md](docs/diseno.md). [English version](README.md).

**Sitio:** [alvarez1lucas.github.io/analisis-crediticio/es](https://alvarez1lucas.github.io/analisis-crediticio/es/) · **Tablero de Power BI:** [powerbi/analisis_crediticio.pbix](powerbi/analisis_crediticio.pbix) · **Modelo:** [excel/analisis_crediticio.xlsx](excel/analisis_crediticio.xlsx)

**Stack:** Excel, SQL con DuckDB, Python (pandas, statsmodels), Plotly, Power BI, pytest, GitHub Actions, GitHub Pages.

## Qué tiene

- **Balances:** Pampa, YPF y Vista, 2023 a 2025, cargados desde sus estados contables con controles de cuadre contra lo informado; ratios, análisis vertical, fuentes y usos, DuPont y Altman Z''.
- **Deuda:** 42 obligaciones negociables y 25 créditos; cada ON con el riesgo país, el Tesoro de EE.UU. y la tasa bancaria del momento; muro de vencimientos a fin de 2025.
- **Crédito:** rating interno, proyección 2026-2028 con escenario de estrés, margen contra covenants y memo de crédito con recomendación de línea.
- **Forma de financiarse:** regresión logística de dólar linked contra dólar duro según el riesgo país, y colocadores de cada ON (bancos, ALyC y plataformas digitales).
- **Controles:** tests con pytest; SQL, Excel y Power BI dan las mismas cifras.

## Cómo correrlo

`pip install -r requirements.txt` y, con `PYTHONPATH=src`: `empresas.mercado` (series), `empresas.plantilla` (Excel), `empresas.sql` (capa SQL en `sql/`),
`empresas.tendencias` (modelo), `empresas.memo` (memo y exportaciones), `empresas.guia` (guía), `empresas.sitio` (sitio) y `empresas.powerbi` (tablero).
GitHub Actions actualiza el mercado cada lunes. Los balances están en `data/balances.csv` y se actualizan cuando sale un nuevo 20-F.
El Excel y el tablero de Power BI están en español; el sitio, el memo y la guía, en inglés y en español.
