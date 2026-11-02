# Análisis crediticio de empresas

¿Cuándo conviene una obligación negociable y cuándo un crédito, y qué lugar le queda al banco? Estudio de Pampa Energía, YPF y Vista sobre sus 20-F, avisos de resultados y series de mercado. El diseño está en [docs/diseno.md](docs/diseno.md).

**Estado:** en construcción. `excel/analisis_crediticio.xlsx` tiene la carga de balances con controles de cuadre, los ratios comparados, las emisiones y las series de mercado. Vista tiene 2023 a 2025 cargados, con 17 ON y 16 créditos al cierre de 2025. Faltan Pampa, YPF y los años anteriores.

`python -m empresas.mercado` actualiza las series y `python -m empresas.plantilla` genera la plantilla. Antes de regenerar guarda en `data/balances.csv` lo cargado a mano en la hoja de Pampa.
