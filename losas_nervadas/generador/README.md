# Generador de `LosasNervadas_CasaRocha.xlsx`

La hoja se arma con estos scripts (openpyxl) a partir de los datos ya extraídos
del E2K y de los DXF (`work/`). Las fórmulas quedan vivas en el libro; los
scripts sólo escriben la estructura.

Los scripts leen y escriben en una carpeta de trabajo fija (variable `W` al
principio de `gen_a.py`, `propose.py` y `build_data.py`). Para regenerar:

```bash
W=<carpeta de trabajo>           # la misma ruta que tiene W en los scripts
mkdir -p "$W/out" && cp -r scripts work "$W/"
cd "$W/scripts"
python3 propose.py               # sólo si cambian cargas o secciones: itera las secciones de las principales (work/propuesta_np.json)
python3 gen_d.py ../out/LosasNervadas_CasaRocha.xlsx
# recalcular con LibreOffice y guardar los valores en caché:
cp ../out/LosasNervadas_CasaRocha.xlsx ../out/rc.xlsx
python3 <recalc.py del skill xlsx> ../out/rc.xlsx 300
python3 inject_values.py ../out/LosasNervadas_CasaRocha.xlsx ../out/rc.xlsx ../out/final.xlsx
```

`final.xlsx` es el archivo que se copia a `losas_nervadas/`. Las decisiones de diseño D1 a D11 (RESUMEN sección 2) están
aplicadas en los scripts: en TABLEROS cada renglón lleva su tipo de losa (columna «Losa»: NERVADA o MACIZA, por franja
de 1 m) y su b_w (20 cm en las nervaduras de los volados, dato `bw_vol` de DATOS); el ancho del macizo de borde de los
bordes discontinuos es el dato `b_borde`.

`summ.py` imprime un resumen de TABLEROS, PRINCIPALES y NERVADURA desde el libro recalculado.

| Archivo | Qué hace |
|---|---|
| `gen_a.py` | DATOS, CARGAS, TABLAS (coeficientes NTC-Concreto 2023, Tabla 3.4.2.1 caso I; L_d, L_dh, traslapes) |
| `gen_b.py` | TABLEROS, NERVADURA, PRINCIPALES, TRIB, TRANSFER, MUROS |
| `gen_c.py` | CROQUIS, CROQUIS_DATOS, CUANTIFICACION |
| `gen_d.py` | RESUMEN, REFERENCIAS y armado del libro |
| `propose.py` | Propuesta iterativa de b × h de cada nervadura principal |
| `build_data.py`, `parse_e2k.py`, `geo*.py`, `lines.py` | Extracción de geometría y cargas del E2K y de los DXF hacia `work/` |
