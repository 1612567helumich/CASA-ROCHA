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

## Estado al 23 de octubre: terminado

**Casa Rocha (`losas_nervadas/`).** La ronda final de verificación ya está en el generador actual y pasó una revisión independiente:

- 3.4.3.1 y 3.4.2.1 c) sólo entre vecinos recíprocos, con el borde opuesto del vecino, y en tableros irregulares sólo si las orillas se tocan de veras (> 5 cm).
- Hileras a macizar topadas en la mitad de los casetones del claro corto desde cada apoyo (todos en un volado); si hiciera falta macizar todo, NO CUMPLE hasta rediseñarlo como maciza.
- Croquis a escala: el acomodo del área de las gráficas va en `ch.layout` (openpyxl ignoraba `plot_area.layout`).

Recalcula sin errores (11 328 fórmulas): 30 de 30 tableros y 32 de 32 principales cumplen.

**Hoja general (`hojas_generales/`).** Terminada y revisada (fórmulas contra la NTC y croquis). Ver su README. Lo principal:

- Valores de D3, malla al centro de la capa, dala de borde (D9) en el anclaje (gancho de 180°) y en el acomodo de casetones.
- Volados: L_b obligatorio, peso de los macizos, acero hasta el apoyo lejano con gancho y aviso de levantamiento, y revisión del respaldo con nervaduras anchas (D10).
- Cargas vivas cotejadas contra la Tabla 6.1.2.2 (azoteas por pendiente) y W_g ≥ 25 kg/m².
- Croquis a escala, con ganchos de 180°, bastón del volado, apoyo lejano y leyendas que dicen «—» cuando un trazo no aplica.

Recalcula sin errores; los siete ejemplos cumplen y los de nervadas cuadran con Casa Rocha salvo lo que depende de los vecinos y del peso volumétrico.
