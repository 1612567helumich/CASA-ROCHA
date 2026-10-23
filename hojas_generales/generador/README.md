# Generador de `Losas_General_NTC2023.xlsx`

Hoja general de losas nervadas (encasetonadas) y macizas por el método de coeficientes de la
NTC-Concreto 2023 (Tabla 3.4.2.1, caso I), para cualquier proyecto. Sale de la hoja de Casa Rocha
(`losas_nervadas/`), sin lo que depende de ese modelo (vecinos, nervaduras principales, áreas
tributarias). Las fórmulas quedan vivas en el libro; los scripts sólo escriben la estructura.

## Cómo regenerar

```bash
cd hojas_generales/generador
python3 gen_general.py                 # arma, recalcula con LibreOffice y guarda los valores en caché
python3 gen_general.py --sin-recalculo # sólo arma (Excel calcula al abrir)
python3 render_croquis.py SALIDA_DIR   # dibuja la hoja CROQUIS de los siete ejemplos a PNG para revisarla
```

Se necesita `openpyxl`, LibreOffice (`soffice`) y el `recalc.py` del skill xlsx; `render_croquis.py`
usa además `pdftoppm` y `pdftotext`.

## Qué trae

| Hoja | Qué hace |
|---|---|
| INSTRUCCIONES | Alcance, límites, colores y uso paso a paso |
| DATOS | Materiales, factores, recubrimientos, sección nervada, malla y capa de compresión, valores de macizas |
| NERVADAS | Hasta 40 tableros nervados (perimetrales o volados), uno por renglón |
| MACIZAS | Hasta 40 tableros macizos, por franja de 1 m |
| DETALLE | Memoria paso a paso del tablero escogido, con fórmula y referencia de cada paso |
| CROQUIS | Planta, corte y alzado a escala del tablero escogido, que se redibujan solos |
| TABLAS | Coeficientes de la Tabla 3.4.2.1, varillas, L_d, mallas y listas |
| CROQUIS_DATOS | Coordenadas de los dibujos (fórmulas) |

Los valores por omisión son los de Casa Rocha: capa de 6 cm, casetón 60×60×19 (h = 25 cm) y malla
6x6-4/4 al centro de la capa (D3); dala de borde armada de 40 × 25 cm en bordes discontinuos, con
gancho de 180° (D9); volados protegidos con impermeabilizante y piso (D12). Los ejemplos de NERVADAS
son tableros de Casa Rocha; dan lo mismo que esa hoja salvo lo que depende de los vecinos
(distribución de 3.4.3.1), el peso volumétrico (2.40 aquí, 2.44 en el modelo de Casa Rocha) y la
media nervadura de borde que en Casa Rocha queda fuera del eje del volado.

En los volados el claro del tablero de respaldo L_b es obligatorio: de él depende hasta dónde entra
el acero superior (si con 0.9·PP el M⁻ no se anula, llega al apoyo lejano, se ancla con gancho y la
nota avisa que ese apoyo levanta). En los bordes discontinuos de los perimetrales el acomodo de
casetones deja lugar a la dala de borde de DATOS.

Un volado nervado con b_w propio mayor que el de DATOS revisa además su tablero de respaldo
(D10 de Casa Rocha): cuántas hileras siguen con las nervaduras anchas hasta donde la nervadura normal,
con la varilla corrida del volado, resiste el M⁻ y el cortante.

| Archivo | Qué hace |
|---|---|
| `gen_general.py` | Arma el libro, lo recalcula e inyecta los valores |
| `common.py` | Constantes, TABLAS, DATOS e INSTRUCCIONES |
| `losas.py` | NERVADAS y MACIZAS (un renglón por tablero) |
| `detalle.py` | DETALLE, CROQUIS y CROQUIS_DATOS |
| `xlstyle.py` | Estilos y ayudantes de celdas |
| `inject_values.py` | Copia los valores calculados por LibreOffice al libro con fórmulas |
| `render_croquis.py` | Dibuja la hoja CROQUIS a PNG para revisarla |
