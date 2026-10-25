# Generador de la hoja de muros de mampostería confinada (NTC 2023)

Revisión de muros de mampostería confinada por la NTC-Mampostería 2023 y la NTC-Sismo 2023 (CDMX). Un solo generador
saca dos libros con las mismas fórmulas:

| Libro | Qué trae |
|---|---|
| `hojas_generales/Muros_Mamposteria_NTC2023.xlsx` | Hoja general para cualquier proyecto: datos en blanco y tres muros de ejemplo (M-1 sin vanos, M-2 con puerta y ventana, M-3 sobre una viga) |
| `muros_casa_rocha/Muros_CasaRocha_NTC2023.xlsx` | Los 37 muros de Casa Rocha con sus fuerzas, entrepisos y distorsiones |

Las fórmulas quedan vivas en el libro; los scripts sólo escriben la estructura.

## Cómo regenerar

```bash
cd hojas_generales/generador_muros
python3 gen_muros.py                      # arma los dos, los recalcula con LibreOffice y guarda los valores en caché
python3 gen_muros.py general              # sólo la hoja general (o: casa_rocha)
python3 gen_muros.py --sin-recalculo      # sólo arma (Excel calcula al abrir)
python3 render_croquis.py SALIDA_DIR M-2                       # croquis de la hoja general a PNG
python3 render_croquis.py SALIDA_DIR --casa_rocha CAD_PA_21    # croquis de Casa Rocha a PNG
```

Se necesita `openpyxl`, LibreOffice (`soffice`) y el `recalc.py` del skill xlsx; `render_croquis.py` usa además
`pdftoppm`.

## Qué revisa y qué no

Revisa muros confinados de bloque, tabique o tabicón (Tipo I o II, Grupo A o B), hasta 60 muros con hasta 4 vanos
cada uno: castillos y dalas, compresión, flexocompresión, cortante con o sin refuerzo horizontal, muro sobre viga
(4.1.18), cortante de entrepiso y distorsiones. Las fuerzas se capturan en MUROS o se toman de la tabla Pier Forces de
ETABS.

No revisa flexión fuera del plano, muros diafragma, mampostería reforzada interiormente ni el diseño de la viga de apoyo.

## Hojas

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

## Datos de Casa Rocha

`ejemplo_casa_rocha.json` sale de:

- longitudes de los piers del modelo `001_EST_CROCHA_R3.2.e2k`;
- vanos del plano `001_PLANO_CROCHA_R3.0.dxf` (capas 011_Puertas y 012_Ventanas), con alturas típicas (puerta de 2.10 m; ventana con antepecho de 0.90 m y 1.20 m de alto);
- fuerzas de la hoja del usuario `Revision_muros_NTC2023.xlsx` (ETABS R3.2);
- vigas de apoyo con las secciones propuestas en PRINCIPALES de la hoja de losas.

## Archivos

| Archivo | Qué hace |
|---|---|
| `gen_muros.py` | Arma los libros (casos `general` y `casa_rocha`), los recalcula e inyecta los valores |
| `muros_spec.py` | Columnas de MUROS y rangos compartidos |
| `muros_comun.py` | TABLAS, DATOS, INSTRUCCIONES y REFERENCIAS; `DEF` trae los valores de la hoja general |
| `muros_tabla.py` | MUROS y PIER_FORCES |
| `muros_detalle.py` | DETALLE |
| `muros_croquis.py` | CROQUIS y CROQUIS_DATOS |
| `muros_otros.py` | ENTREPISO, DISTORSIONES y RESUMEN |
| `xlstyle.py`, `inject_values.py` | Estilos y copia de valores calculados |
| `render_croquis.py` | Dibuja la hoja CROQUIS a PNG para revisarla |
