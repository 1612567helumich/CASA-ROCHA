# Para continuar si se corta la sesión

Rama de trabajo: `claude/practical-cori-l2ai0b`.

## Qué se está haciendo

1. **Casa Rocha** (`losas_nervadas/`): aplicar al generador (`losas_nervadas/generador/scripts`) las decisiones de diseño, regenerar la hoja, recalcularla sin errores y verificarla.
2. **Hoja general** (`hojas_generales/`): un libro `Losas_General_NTC2023.xlsx` con hojas INSTRUCCIONES, DATOS, NERVADAS, MACIZAS, DETALLE, CROQUIS y TABLAS, armado por `hojas_generales/generador/gen_general.py`, para cualquier proyecto. **El usuario pidió croquis bien hechos**, como la hoja CROQUIS de Casa Rocha (gráficas de dispersión que se redibujan solas): planta, corte y alzado de nervadas, y planta y corte de macizas, del tablero que se escoja.
3. Al final: subir todo y entregar al usuario los tres resultados (hoja de Casa Rocha, hoja general de nervadas y hoja general de macizas).

## Decisiones tomadas (Casa Rocha, criterio conservador)

- D1 Castillos C1…C8: aceptados, bajan a cimentación; donde no haya muro abajo, columna de 25×25 como mínimo.
- D2 Principales: se aceptan las secciones propuestas. NP-15 y NP-26 (H/L < 0.6): la trabe carga todo el muro repartido, sin efecto de arco; cumplen 4.1.18.3/4 y f_CM ≤ F_R·f'm como estimación. Queda como criterio del Responsable.
- D3 Capa de compresión de 6 cm, casetón 60×60×19 (h = 25) y malla 6x6-4/4.
- D4 Se conservan los +20 +20 kg/m² (5.1.2.1).
- D5 El usuario corre ETABS con la lista de cambios que viene en RESUMEN.
- D6 Las notas para planos van en RESUMEN.
- D7 Tableros chicos: losa sobre vigas (6.7.8.1.9), como ya hace la hoja.
- D8 Fuera de 3.4.2.1 c): distribución de 3.4.3.1; se confirma con ETABS.
- D9 Bordes discontinuos: macizo o dala de borde de 40 cm (nuevo dato `b_borde` en DATOS).
- D10 Volados F80012 y F80019: **siguen nervados** (el entrepiso es nervado), con nervaduras de 20 cm y casetón 50×60×19.
- D11 Azotea F80032 y F80010: **losa maciza de 20 cm** (así están en el modelo, «Slab1»), con f'c = 250.

## Cómo regenerar Casa Rocha

Ver `losas_nervadas/generador/README.md`. Las rutas ya son relativas a la carpeta del generador; se necesita `pip install shapely ezdxf`.

## Estado al 23 de octubre (en pausa a pedido del usuario)

**Casa Rocha, ~90 %.** La hoja de esta rama recalcula sin errores (11 328 fórmulas) con D1 a D12 y la revisión independiente: 30 de 30 tableros y 32 de 32 principales cumplen. Falta pasar al generador actual tres correcciones de la ronda final de verificación, que están en el commit `b3c4bf1` (hecho sobre la versión anterior de `gen_b.py` y `gen_d.py`; no se aplicó), y luego regenerar y volver a revisar:

1. Distribución de 3.4.3.1 sólo entre vecinos recíprocos (el vecino lista a este tablero) y siempre con su borde opuesto; hoy, si el vecino no lista a este tablero se distribuye igual con su borde opuesto, y si lo lista en otro borde se usa ese.
2. Desequilibrio del volado sólo por su borde de empotramiento (`dirc` del volado = dirección del momento en ese borde).
3. Hileras a macizar topadas en el número de casetones del claro (`ncs`).

La revisión Vul/VcRl de `b3c4bf1` no se pasa: el cortante de los respaldos ya lo revisan las columnas RESPALDO de TABLEROS.

**Hoja general, ~60 %.** Está en la rama `claude/respaldo-wip-losas` (commit `de115a3`), no en ésta. Tiene las 7 hojas más CROQUIS_DATOS (11 915 fórmulas, 0 errores) y los 7 ejemplos dan CUMPLE. Falta:

- Croquis: títulos encimados con el armado en corte y alzado; los dibujos vacíos del tipo no escogido muestran números sueltos (102.6, 57.5…); la planta de maciza lleva demasiados letreros.
- Pasar los valores por omisión a D3 (capa de 6 cm, casetón 60×60×19, malla 6x6-4/4 al centro de la capa) y lo que aplique de la revisión de Casa Rocha (dala de borde, ganchos de 180°, respaldo de volados).
- README, revisión independiente contra la NTC, comparar los ejemplos con Casa Rocha y traerla a esta rama.
