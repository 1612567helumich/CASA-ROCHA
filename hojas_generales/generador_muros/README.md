# Generador de `Muros_Mamposteria_NTC2023.xlsx`

Hoja de revisión de muros de mampostería confinada por la NTC-Mampostería 2023 y la NTC-Sismo 2023 (CDMX), para
cualquier proyecto. Trae de ejemplo los 37 muros de Casa Rocha. Las fórmulas quedan vivas en el libro; los scripts
sólo escriben la estructura.

## Cómo regenerar

```bash
cd hojas_generales/generador_muros
python3 gen_muros.py                 # arma, recalcula con LibreOffice y guarda los valores en caché
python3 gen_muros.py --sin-recalculo # sólo arma (Excel calcula al abrir)
python3 render_croquis.py SALIDA_DIR CAD_PA_21 CAD_PB_13   # dibuja el croquis de esos muros a PNG
```

Se necesita `openpyxl`, LibreOffice (`soffice`) y el `recalc.py` del skill xlsx; `render_croquis.py` usa además
`pdftoppm`.

## Qué trae

| Hoja | Qué hace |
|---|---|
| INSTRUCCIONES | Uso en cinco pasos y colores |
| DATOS | Pieza, mortero, f'm y v'm (con su origen), concreto, castillo y dala típicos, factores, sismo y de dónde salen las fuerzas |
| MUROS | Un renglón por muro (hasta 60): sección crítica, compresión, flexocompresión, cortante, refuerzo horizontal, muro sobre viga y estado. Las cuentas intermedias van agrupadas |
| DETALLE | La revisión paso a paso del muro escogido |
| CROQUIS | Alzado armado, alzado acotado, sección crítica en planta y secciones de castillo y dala del muro escogido |
| RESUMEN | Qué cumple, qué no y qué refuerzo horizontal le falta a cada muro |
| ENTREPISO | Cortante resistente de entrepiso (4.1.10, 4.1.12) |
| DISTORSIONES | Seguridad de vida y daño con la tabla Story Drifts de ETABS |
| PIER_FORCES | Donde se pega la tabla Pier Forces de ETABS |
| REFERENCIAS, TABLAS | Incisos y páginas; varillas, mallas, tablas de f'm, Tipo II y Tabla 4.3.3 |
| CROQUIS_DATOS | Coordenadas de los dibujos (fórmulas) |

## Ejemplo de Casa Rocha

`ejemplo_casa_rocha.json` sale de:

- longitudes de los piers del modelo `001_EST_CROCHA_R3.2.e2k`;
- vanos del plano `001_PLANO_CROCHA_R3.0.dxf` (capas 011_Puertas y 012_Ventanas), con alturas típicas (puerta de 2.10 m; ventana con antepecho de 0.90 m y 1.20 m de alto);
- fuerzas de la hoja del usuario `Revision_muros_NTC2023.xlsx` (ETABS R3.2);
- vigas de apoyo con las secciones propuestas en PRINCIPALES de la hoja de losas.

| Archivo | Qué hace |
|---|---|
| `gen_muros.py` | Arma el libro, lo recalcula e inyecta los valores |
| `muros_spec.py` | Columnas de MUROS y rangos compartidos |
| `muros_comun.py` | TABLAS, DATOS, INSTRUCCIONES y REFERENCIAS |
| `muros_tabla.py` | MUROS y PIER_FORCES |
| `muros_detalle.py` | DETALLE |
| `muros_croquis.py` | CROQUIS y CROQUIS_DATOS |
| `muros_otros.py` | ENTREPISO, DISTORSIONES y RESUMEN |
| `xlstyle.py`, `inject_values.py` | Estilos y copia de valores calculados |
| `render_croquis.py` | Dibuja la hoja CROQUIS a PNG para revisarla |
