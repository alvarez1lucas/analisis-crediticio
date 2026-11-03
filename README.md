# Análisis crediticio de empresas

¿Cuándo conviene una obligación negociable y cuándo un crédito, y qué lugar le queda al banco? Estudio de Pampa Energía, YPF y Vista sobre sus 20-F, avisos de resultados y series de mercado. El diseño está en [docs/diseno.md](docs/diseno.md) y la conclusión, en el [memo de crédito](docs/memo_credito.md).

**Estado:** en construcción. `excel/analisis_crediticio.xlsx` tiene la carga de balances con controles de cuadre, los ratios comparados, las emisiones y las series de mercado. Pampa, YPF y Vista tienen 2023 a 2025 cargados desde sus estados contables en dólares, con 42 ON y 25 créditos.

`python -m empresas.mercado` actualiza las series, `python -m empresas.plantilla` genera el Excel y `python -m empresas.memo` escribe el memo. Antes de regenerar guarda en `data/balances.csv` lo cargado a mano en la hoja de Pampa.
