# Diseño del estudio

## Pregunta

¿Qué tan sólidas son Pampa, YPF y Vista, cómo se financian y cómo cambió esa elección con el tiempo? ¿Qué lugar le queda al banco?

## Tres miradas

| Mirada | Pregunta |
|---|---|
| Empresa | ¿Qué le sale más barato según plazo y moneda, y cuándo conviene emitir? |
| Banco | ¿Conviene prestar, colocar la ON, comprarla o combinarlas? |
| Fintech | Estas empresas no le piden crédito a una fintech: ¿la ON sirve como inversión de su caja y producto para distribuir? |

## Partes

1. **Análisis de balances (foco principal):** cinco años, análisis vertical, rentabilidad, capital de trabajo, cobertura, liquidez, fuentes y usos, y Altman como referencia.
2. **Créditos y deuda (foco siguiente):** ON, préstamos bancarios y prefinanciación de exportaciones, por moneda, tasa fija o variable, plazo y vencimientos.
3. **Emisiones:** cada ON con su fecha, plazo, moneda, ley, rendimiento, spread, riesgo país y colocadores.
4. **Tendencias de financiamiento:** cómo cambió la elección entre ON y crédito, dólares y pesos, y tasa fija y variable.
5. **Rol del banco:** quién coloca, y rendimiento de la ON contra el crédito cuando haya tasas comparables.
6. **Crédito:** rating interno, proyección a tres años, estrés por renovar deuda más cara y covenants.
7. **Memo:** recomendación de línea de crédito para Pampa y guía de decisión por instrumento.

## Datos y modelos

- **Tendencias:** participación de cada instrumento por empresa y trimestre, comparada con riesgo país, Tesoro y tasas locales.
- **Elección de instrumento:** regresión logística sobre las emisiones para ver qué condiciones acompañan elegir dólares o pesos, y tasa fija o variable. Coeficientes interpretables, sin predicción.
- **Validación:** con pocas emisiones se informan rangos de incertidumbre y se comparan contra la participación media.

## Límites

- Son tres empresas: las relaciones son descriptivas, no causales.
- Sólo se comparan instrumentos de igual moneda y plazo parecido.
- Las tasas del BCRA son del sistema y en pesos: dan contexto, no el costo de estas empresas.
- El detalle de tasas de préstamos bancarios depende de lo que informe cada 20-F.
- La API del BCRA no trae tasas de crédito en dólares: esa comparación depende de los 20-F.
- Sólo se conoce el saldo de crédito al cierre de cada año, no cada préstamo tomado durante el año.
- Sólo se ven las emisiones hechas: no se observan las que una empresa decidió no hacer.
