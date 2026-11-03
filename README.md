# Corporate credit analysis: bonds or bank loans?

When does a corporate bond beat a bank loan, and what role is left for banks and fintechs? A credit file on Pampa Energía, YPF and Vista built from their
audited US-dollar financial statements, bond results notices and market data. The conclusion is in the [credit memo](docs/credit_memo.md) and the
[decision guide](docs/instrument_guide.md). [Versión en español](README_ES.md).

**Site:** [alvarez1lucas.github.io/analisis-crediticio](https://alvarez1lucas.github.io/analisis-crediticio/) · **Power BI dashboard:** [powerbi/analisis_crediticio.pbix](powerbi/analisis_crediticio.pbix) · **Model:** [excel/analisis_crediticio.xlsx](excel/analisis_crediticio.xlsx)

**Stack:** Excel, SQL with DuckDB, Python (pandas, statsmodels), Plotly, Power BI, pytest, GitHub Actions, GitHub Pages.

## What it covers

- **Balance sheets:** Pampa, YPF and Vista, 2023-2025, loaded from their financial statements with checks against reported totals; ratios, vertical analysis, sources and uses, DuPont and Altman Z''.
- **Debt:** 42 corporate bonds and 25 loans, each bond matched with that day's country risk, US Treasury yield and the monthly bank lending rate; maturity wall at end-2025.
- **Credit:** internal rating, 2026-2028 projection with a stress scenario, covenant headroom and a credit memo with a facility recommendation.
- **Funding choices:** logistic regression of dollar-linked versus hard-dollar issuance on country risk, and placement agents of each bond (banks, brokers and digital platforms).
- **Controls:** a pytest suite; SQL, Excel and Power BI return the same figures.

## How to run it

`pip install -r requirements.txt` and, with `PYTHONPATH=src`: `empresas.mercado` (market data), `empresas.plantilla` (Excel), `empresas.sql` (SQL layer in `sql/`),
`empresas.tendencias` (model), `empresas.memo` (memo and exports), `empresas.guia` (guide), `empresas.sitio` (site) and `empresas.powerbi` (dashboard).
GitHub Actions refreshes market data every Monday. Financial statements live in `data/balances.csv` and are updated when a new 20-F is filed.
The Excel model and the Power BI dashboard are in Spanish; the site, memo and guide are in English and Spanish.
