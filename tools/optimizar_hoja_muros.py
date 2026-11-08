#!/usr/bin/env python3
"""Optimiza la hoja de revision de muros (NTC 2023) sin perder croquis,
listas desplegables ni formulas: edita el XML del .xlsx directamente.

1. Un renglon por pier de ETABS en MUROS (nombre, piso, direccion, L del
   modelo, viga de apoyo y datos de la viga del muro original).
2. Pega la tabla Pier Forces de ETABS en PIER_FORCES (Comb4.x y Comb2.2) y
   pone DATOS en modo ETABS.
3. Textos mas cortos; columnas de explicacion y calculo agrupadas con [+].
4. Alturas de renglon y ajuste de texto para que nada se corte.
5. Croquis: leyendas sin series vacias, sin encimarse con el dibujo.
6. Recalcula con LibreOffice y guarda los resultados como valores en cache
   (para que se vean bien aun en vista protegida); Excel recalcula al abrir.

Uso: python3 optimizar_hoja_muros.py HOJA.xlsx PIER_FORCES.xlsx MODELO.e2k SALIDA.xlsx
"""
import copy
import math
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import openpyxl
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from redibujar_muros import Model, TOL  # noqa: E402

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSC = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NSA = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
q = lambda t: f"{{{NS}}}{t}"
qc = lambda t: f"{{{NSC}}}{t}"
qa = lambda t: f"{{{NSA}}}{t}"

COMBOS = {f"Comb4.{i}" for i in range(1, 9)} | {"Comb2.2"}

# ---------------------------------------------------------------- textos
TEXTOS = {
    # INSTRUCCIONES
    "REVISIÓN DE MUROS DE MAMPOSTERÍA CONFINADA — NTC 2023 (CDMX)": "MUROS DE MAMPOSTERÍA CONFINADA — NTC 2023 (CDMX)",
    "1. DATOS: pieza, mortero, f'm, v'm, concreto, castillo y dala típicos.": "1. DATOS: materiales, castillo y dala típicos.",
    "2. MUROS: un renglón por muro (longitud, vanos de izquierda a derecha, viga de apoyo si la hay y fuerzas).":
        "2. MUROS: un renglón por pier de ETABS (nombre, piso, dirección, L y viga de apoyo).",
    "3. Fuerzas: captúralas en MUROS, o pega la tabla «Pier Forces» de ETABS en PIER_FORCES y en DATOS cambia «Fuerzas de los muros» a ETABS.":
        "3. PIER_FORCES: aquí va la tabla «Pier Forces» de ETABS (tonf, m). Ya trae la del modelo R3.2.",
    "4. DETALLE: escoge un muro y sale su revisión paso a paso. CROQUIS lo dibuja.":
        "4. DETALLE: escoge un muro y ve su revisión paso a paso. CROQUIS lo dibuja.",
    "5. RESUMEN: qué cumple, qué no y qué le falta a cada muro. ENTREPISO y DISTORSIONES: revisión global.":
        "5. RESUMEN: qué cumple y qué no. ENTREPISO y DISTORSIONES: revisión global.",
    "Amarillo = se captura · Verde = CUMPLE · Ámbar = PENDIENTE (por ejemplo, falta el análisis muro–viga) · Rojo = NO CUMPLE.":
        "Amarillo = captura · Verde = cumple · Ámbar = pendiente · Rojo = no cumple.",
    "En MUROS las columnas de cálculo intermedio están agrupadas: ábrelas con el botón [+] de arriba si quieres ver las cuentas.":
        "Las cuentas intermedias (MUROS) y las explicaciones (DATOS y DETALLE) están agrupadas: ábrelas con [+].",
    "Flexión fuera del plano, muros diafragma, mampostería reforzada interiormente y el diseño de la viga de apoyo (va en la hoja de losas).":
        "Flexión fuera del plano, muros diafragma, refuerzo interior y diseño de la viga de apoyo.",
    "37 muros del modelo R3.2 con los vanos del plano arquitectónico y las fuerzas de tu hoja de revisión. Alturas de vanos típicas: confírmalas.":
        "51 piers del modelo R3.2 con sus fuerzas de ETABS. Al correr la R3.4, pega su tabla en PIER_FORCES.",
    # MUROS
    "MUROS — UN RENGLÓN POR MURO (NTC-MAMPOSTERÍA 2023)": "MUROS — UN RENGLÓN POR PIER DE ETABS (NTC-M 2023)",
    "Amarillo = captura. Las cuentas intermedias están agrupadas: ábrelas con [+] arriba de las columnas. DETALLE explica un muro paso a paso y CROQUIS lo dibuja.":
        "Amarillo = captura. Cuentas agrupadas: ábrelas con [+]. DETALLE explica un muro y CROQUIS lo dibuja.",
    "Muro y Story iguales que en ETABS. H y t vacíos = los de DATOS.": "Nombre y Story como en ETABS. H y t vacíos = DATOS.",
    "De izquierda a derecha. x = del extremo izquierdo al vano. Puerta: antepecho 0.": "Sólo si el muro tiene vanos.",
    "Opcional. Refuerzo horizontal: varilla en juntas o malla.": "Opcional.",
    "Sólo si el muro desplanta sobre una viga (4.1.18).": "Sólo si desplanta sobre viga (4.1.18).",
    "Sólo si DATOS dice CAPTURA. P = la mínima sin factor.": "Sólo si DATOS dice CAPTURA.",
    "Casa Rocha: 37 muros del modelo R3.2. Lo amarillo se puede cambiar.":
        "Casa Rocha: 51 piers del modelo R3.2 con fuerzas de ETABS.",
    # DATOS
    "DATOS GENERALES — MUROS DE MAMPOSTERÍA CONFINADA, NTC 2023 (CDMX)": "DATOS GENERALES — NTC 2023 (CDMX)",
    "Sólo se tocan las celdas amarillas. Valen para todos los muros; lo de cada muro va en MUROS.":
        "Sólo se capturan las celdas amarillas. Aplican a todos los muros.",
    "¿Hay muros estructurales desplantados sobre vigas o sobre la losa?": "¿Hay muros sobre vigas o losa?",
    "2. MAMPOSTERÍA — PIEZA, MORTERO, f'm Y v'm (NTC-M 2.1, 2.8)": "2. MAMPOSTERÍA (NTC-M 2.1 y 2.8)",
    "f'p — resistencia de diseño a compresión de la pieza": "f'p — resistencia de la pieza",
    "f'm — RESISTENCIA DE DISEÑO A COMPRESIÓN QUE SE USA": "f'm QUE SE USA",
    "v'm — RESISTENCIA DE DISEÑO A COMPRESIÓN DIAGONAL QUE SE USA": "v'm QUE SE USA",
    "4. MURO, CASTILLO Y DALA TÍPICOS (los usan todos los muros; MUROS deja cambiar H, t y los castillos extremos)":
        "4. MURO, CASTILLO Y DALA TÍPICOS",
    "Recubrimiento libre de las varillas de castillos y dalas": "Recubrimiento libre (castillos y dalas)",
    "Varilla o alambre del refuerzo horizontal (si MUROS no dice otro)": "Refuerzo horizontal (si MUROS no dice otro)",
    "7. SISMO — FACTOR DE COMPORTAMIENTO Y DISTORSIONES (NTC-Sismo 1.7 y Tabla 4.3.3)":
        "7. SISMO — Q Y DISTORSIONES (NTC-Sismo 1.7 y Tabla 4.3.3)",
    "Factor de la distorsión de seguridad de vida = Q·R'·FC": "Factor de distorsión de seguridad de vida (Q·R'·FC)",
    "¿Los elementos no estructurales están separados de la estructura?": "¿No estructurales separados de la estructura?",
    "9. DISTORSIONES — CÓMO RECONOCER CADA CASO EN LA TABLA DE ETABS": "9. DISTORSIONES — NOMBRES DE CASOS EN ETABS",
    "Texto de los casos de ocupación inmediata (Grupo A)": "Casos de ocupación inmediata (Grupo A)",
    # DETALLE
    "DETALLE — MEMORIA DE CÁLCULO PASO A PASO DEL MURO ESCOGIDO": "DETALLE — REVISIÓN PASO A PASO",
    "Escoge el muro en la celda amarilla. CROQUIS dibuja este mismo muro.": "Escoge el muro en la celda amarilla. CROQUIS lo dibuja.",
    "7. FLEXOCOMPRESIÓN EN EL PLANO — MÉTODO OPTATIVO (6.4.2.2)": "7. FLEXOCOMPRESIÓN EN EL PLANO (6.4.2.2)",
    # CROQUIS
    "Dibuja el muro escogido en DETALLE, a escala. Se redibuja solo.": "Dibujo a escala del muro escogido en DETALLE.",
    # ENTREPISO
    "V_R = F_R·(0.5·v'm + 0.3·σ)·ΣA_T + ΣV_sR ≥ 0.8·V_u en cada dirección (ec. 4.1.12.1). ΣA_T sale de MUROS.":
        "V_R = F_R·(0.5·v'm + 0.3·σ)·ΣA_T + ΣV_sR ≥ 0.8·V_u (ec. 4.1.12.1).",
    "1. DATOS DE CADA ENTREPISO (de ETABS, V_u con factor de carga)": "1. DATOS POR ENTREPISO (de ETABS, V_u factorizado)",
    # DISTORSIONES
    "DISTORSIONES DE ENTREPISO — NTC-SISMO 1.7 Y TABLA 4.3.3": "DISTORSIONES DE ENTREPISO — NTC-SISMO 1.7",
    "Pega la tabla «Story Drifts» de ETABS desde el renglón 16. Seguridad de vida: × Q·R'·FC ≤ γ_SV. Daño: ≤ γ_LD.":
        "Pega la tabla «Story Drifts» de ETABS desde el renglón 16.",
    "TABLA «STORY DRIFTS» DE ETABS — CAPTURA (pegar valores)": "TABLA «STORY DRIFTS» DE ETABS (pegar valores)",
    # PIER_FORCES
    "PIER_FORCES — TABLA «PIER FORCES» DE ETABS (SE PEGA AQUÍ)": "PIER_FORCES — TABLA DE ETABS",
    # REFERENCIAS
    "REFERENCIAS — QUÉ NORMA, QUÉ INCISO Y QUÉ PÁGINA CHECAR": "REFERENCIAS — NORMA, INCISO Y PÁGINA",
    "Páginas impresas de las Normas Técnicas Complementarias 2023 de la CDMX (Gaceta Oficial). Se cita el inciso, que no cambia entre tirajes.":
        "Páginas de las NTC 2023 de la CDMX (Gaceta Oficial).",
    # TABLAS (solo notas)
    "TABLAS — VARILLAS, MALLAS, RESISTENCIAS DE LA NTC Y LISTAS": "TABLAS — VARILLAS, MALLAS, RESISTENCIAS Y LISTAS",
    "Todas las hojas leen de aquí con INDEX/MATCH. Los valores son los de la NTC-Mampostería 2023 y la NTC-Sismo 2023 (CDMX); no se tocan.":
        "Valores de la NTC 2023. No se tocan.",
    "Varilla corrugada NMX-C-407-ONNCCE y alambre corrugado laminado en frío. Las varillas se escogen por su número (2, 2.5, 3, 4 …); los alambres por su nombre.":
        "Varillas NMX-C-407 (por número) y alambres laminados en frío (por nombre).",
    "A_s horizontal por metro de altura de una cara (alambres @ 15.24 cm). Malla en las dos caras = 2 «caras» en MUROS.":
        "A_s horizontal por metro de altura, una cara (alambres @ 15.24 cm).",
    "f'm sobre área bruta en kg/cm² según f'p de la pieza; valores intermedios se interpolan. f'p de 60 y 75 sólo para piezas huecas.":
        "f'm (kg/cm², área bruta) según f'p; se interpola.",
    "Los valores de Tipo II sólo valen si la estructura es Tipo II (1.6.1) y se cumplen a) a c) de 2.8.1.4.3.":
        "Sólo para estructuras Tipo II (1.6.1) que cumplan 2.8.1.4.3 a) a c).",
    "Para edificios de hasta 6 niveles. «Con refuerzo horizontal» sólo si TODOS los muros estructurales lo llevan [2]. Piezas huecas confinadas: no se permiten en el Grupo A [3].":
        "Hasta 6 niveles. «Con refuerzo horizontal» sólo si todos los muros lo llevan [2]. Huecas confinadas: no en Grupo A [3].",
}
PIER_FORCES_NOTA = ("ETABS: Display › Show Tables › Wall Results › Pier Forces, en tonf y m. Pega desde B9, sin encabezados "
                    "y sin la columna «Step Number». Ya trae la del modelo R3.2 (Comb4.1–4.8 y Comb2.2).")

# pier de ETABS -> muro de la hoja original (para heredar posicion, refuerzo y viga de apoyo)
PIER2CAD = {
    "PB-A-1": "CAD_PB_01", "PB-J-1": "CAD_PB_02", "PB-4-1": "CAD_PB_03", "PB-4-2": "CAD_PB_03", "PB-B-1": "CAD_PB_04",
    "PB-6-1": "CAD_PB_05", "PB-5-1": "CAD_PB_06", "PB-5-2": "CAD_PB_06", "PB-G-1": "CAD_PB_07", "PB-8-1": "CAD_PB_09",
    "PB-E-1": "CAD_PB_10", "PB-E-2": "CAD_PB_10", "PB-E-3": "CAD_PB_10", "PB-11-1": "CAD_PB_11", "PB-12-1": "CAD_PB_12",
    "PB-14-1": "CAD_PB_13", "PB-14-2": "CAD_PB_13", "PB-14-3": "CAD_PB_13", "PB-X-1.60-1": "CAD_PB_14", "PB-H-1": "CAD_PB_16",
    "PB-15-1": "CAD_PB_15", "PB-15-2": "CAD_PB_17", "PB-15-3": "CAD_PB_17",
    "PA-A-1": "CAD_PA_18", "PA-A-2": "CAD_PA_18", "PA-J-1": "CAD_PA_19", "PA-3-1": "CAD_PA_20", "PA-3-2": "CAD_PA_20",
    "PA-G-1": "CAD_PA_21", "PA-X4.33-1": "CAD_PA_22", "PA-X4.33-2": "CAD_PA_22", "PA-5-1": "CAD_PA_23", "PA-7-1": "CAD_PA_25",
    "PA-8-1": "CAD_PA_26", "PA-8-2": "CAD_PA_26", "PA-F-1": "CAD_PA_27", "PA-Y14.25-1": "CAD_PA_28", "PA-Y14.02-1": "CAD_PA_29",
    "PA-C-1": "CAD_PA_30", "PA-Y15.85-1": "CAD_PA_31", "PA-Y15.85-2": "CAD_PA_32", "PA-X3.03-1": "CAD_PA_33",
    "PA-X5.82-1": "CAD_PA_34", "PA-12-1": "CAD_PA_35", "PA-12-2": "CAD_PA_35", "PA-12-3": "CAD_PA_35", "PA-12-4": "CAD_PA_35",
    "PA-D-1": "CAD_PA_36", "PA-14-1": "CAD_PA_37", "PA-14-2": "CAD_PA_37", "PA-14-3": "CAD_PA_37",
}
HEREDA = ["E", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP"]


# ---------------------------------------------------------------- utilidades de celdas
def col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m[1], int(m[2])


def fmt_num(v):
    if isinstance(v, bool):
        return "1" if v else "0"
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    return repr(float(v))


class Sheet:
    def __init__(self, data):
        self.root = etree.fromstring(data)
        self.sd = self.root.find(q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.sd.findall(q("row"))}

    def row(self, n, create=True):
        if n in self.rows:
            return self.rows[n]
        if not create:
            return None
        r = etree.Element(q("row"))
        r.set("r", str(n))
        before = [k for k in self.rows if k < n]
        if before:
            self.rows[max(before)].addnext(r)
        else:
            self.sd.insert(0, r)
        self.rows[n] = r
        return r

    def cell(self, ref, create=True, style_from=None):
        col, rn = split_ref(ref)
        r = self.row(rn, create)
        if r is None:
            return None
        prev = None
        ci = col2num(col)
        for c in r.findall(q("c")):
            cc = col2num(split_ref(c.get("r"))[0])
            if cc == ci:
                return c
            if cc < ci:
                prev = c
        if not create:
            return None
        c = etree.Element(q("c"))
        c.set("r", ref)
        if style_from is not None and style_from.get("s"):
            c.set("s", style_from.get("s"))
        if prev is not None:
            prev.addnext(c)
        else:
            r.insert(0, c)
        return c

    def cells(self):
        for rn, r in self.rows.items():
            for c in r.findall(q("c")):
                yield rn, c

    def bytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


def set_value(c, v):
    if c.find(q("f")) is not None:
        raise RuntimeError(f"La celda {c.get('r')} tiene formula")
    for ch in list(c):
        if ch.tag in (q("v"), q("is")):
            c.remove(ch)
    c.attrib.pop("t", None)
    if v is None or v == "":
        return
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        e = etree.SubElement(c, q("v"))
        e.text = fmt_num(v)
    else:
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, q("is"))
        t = etree.SubElement(is_, q("t"))
        t.text = str(v)
        if str(v) != str(v).strip():
            t.set(XML_SPACE, "preserve")


# ---------------------------------------------------------------- paquete
class Package:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = [i.filename for i in self.z.infolist()]
        self.data = {n: self.z.read(n) for n in self.names}
        wb = etree.fromstring(self.data["xl/workbook.xml"])
        rels = etree.fromstring(self.data["xl/_rels/workbook.xml.rels"])
        rid = {r.get("Id"): r.get("Target") for r in rels}
        self.sheet_file = {}
        for s in wb.find(q("sheets")):
            t = rid[s.get(f"{{{NSR}}}id")]
            self.sheet_file[s.get("name")] = "xl/" + t.lstrip("/").replace("xl/", "") if not t.startswith("/") else t.lstrip("/")
        self.sheets = {n: Sheet(self.data[f]) for n, f in self.sheet_file.items()}
        self.sst_root = etree.fromstring(self.data["xl/sharedStrings.xml"])
        self.sst = ["".join(t.text or "" for t in si.iter(q("t"))) for si in self.sst_root.findall(q("si"))]
        self.styles = etree.fromstring(self.data["xl/styles.xml"])

    def text(self, c):
        t = c.get("t")
        if t == "s":
            v = c.find(q("v"))
            return self.sst[int(v.text)] if v is not None else ""
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter(q("t")))
        if t == "str":
            v = c.find(q("v"))
            return v.text or "" if v is not None else ""
        return None

    def replace_sst(self, old, new):
        n = 0
        for i, si in enumerate(self.sst_root.findall(q("si"))):
            if self.sst[i] == old:
                for ch in list(si):
                    si.remove(ch)
                t = etree.SubElement(si, q("t"))
                t.text = new
                self.sst[i] = new
                n += 1
        return n

    def save(self, path):
        refs = 0
        for n, f in self.sheet_file.items():
            S = self.sheets[n]
            for r in S.rows.values():
                r.attrib.pop("spans", None)        # pista opcional; ya no cuadra con las celdas nuevas
                refs += sum(1 for c in r.findall(q("c")) if c.get("t") == "s")
            self.data[f] = S.bytes()
        self.sst_root.set("count", str(refs))
        self.sst_root.set("uniqueCount", str(len(self.sst_root.findall(q("si")))))
        self.data["xl/sharedStrings.xml"] = etree.tostring(self.sst_root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.data["xl/styles.xml"] = etree.tostring(self.styles, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zo:
            for n in self.names:
                zo.writestr(n, self.data[n])


# ---------------------------------------------------------------- fase A: contenido
def piers_del_modelo(e2k):
    M = Model(open(e2k, encoding="utf-8", newline="").read())
    walls, _ = M.wall_pieces()
    P = {}
    for (story, axis, coord), pcs in walls.items():
        for p in pcs:
            if not p["pier"]:
                continue
            d = P.setdefault(p["pier"], dict(story=story, axis=axis, coord=coord, h0=9e9, h1=-9e9))
            d["h0"] = min(d["h0"], p["h0"])
            d["h1"] = max(d["h1"], p["h1"])
    out = []
    for name, d in P.items():
        out.append(dict(pier=name, story=d["story"], dir="Y" if d["axis"] == "X" else "X",
                        L=round(d["h1"] - d["h0"], 3), key=(M.order.index(d["story"]) * -1, d["axis"], d["coord"], d["h0"])))
    out.sort(key=lambda x: x["key"])
    return out


def fase_contenido(pkg, hoja_xlsx, pier_xlsx, e2k, log):
    ws_old = openpyxl.load_workbook(hoja_xlsx, data_only=True)["MUROS"]
    cad = {}
    for r in range(11, 71):
        n = ws_old[f"B{r}"].value
        if n:
            cad[n] = {c: ws_old[f"{c}{r}"].value for c in HEREDA}
    piers = piers_del_modelo(e2k)
    assert len(piers) == 51, len(piers)
    M = pkg.sheets["MUROS"]
    entrada = [num2col(i) for i in range(2, 50)]          # B..AW = captura
    for r in range(11, 71):
        for col in entrada:
            c = M.cell(f"{col}{r}", create=False)
            if c is not None and c.find(q("f")) is None:
                set_value(c, None)
    notas = {}
    for r in range(11, 71):
        n = ws_old[f"B{r}"].value
        if n:
            notas[n] = ws_old[f"EX{r}"].value or ""
        c = M.cell(f"EX{r}", create=False)
        if c is not None and c.find(q("f")) is None:
            set_value(c, None)
    for k, p in enumerate(piers):
        r = 11 + k
        cad_name = PIER2CAD.get(p["pier"])
        base = cad.get(cad_name, {})
        vals = {"B": p["pier"], "C": p["story"], "D": p["dir"], "F": p["L"]}
        viga = re.search(r"Viga (NP-\d+ \d+×\d+)", notas.get(cad_name, ""))
        vals["EX"] = f"Antes {cad_name}." + (f" Viga {viga[1]} (en el modelo es dala 15×20)." if viga else "")
        for c in HEREDA:
            if base.get(c) not in (None, ""):
                vals[c] = base[c]
        for col, v in vals.items():
            cell = M.cell(f"{col}{r}", style_from=M.cell(f"{col}11", create=False))
            set_value(cell, v)
    log["piers"] = piers
    # DATOS: fuerzas de ETABS
    set_value(pkg.sheets["DATOS"].cell("C102"), "ETABS (PIER_FORCES)")
    # DETALLE: muro que se ve
    set_value(pkg.sheets["DETALLE"].cell("C6"), "PB-A-1")
    # PIER_FORCES: tabla de ETABS
    pf = openpyxl.load_workbook(pier_xlsx, read_only=True, data_only=True)["Pier Forces"]
    rows = []
    for row in pf.iter_rows(min_row=4, values_only=True):
        if row[2] in COMBOS:
            rows.append([row[0], row[1], row[2], row[3], row[4], row[6], row[7], row[8], row[9], row[10], row[11], row[12]])
    assert len(rows) <= 3000
    S = pkg.sheets["PIER_FORCES"]
    tmpl = {col: S.cell(f"{col}9", create=False) for col in "BCDEFGHIJKLM"}
    for r in range(9, 3009):
        for col in "BCDEFGHIJKLM":
            c = S.cell(f"{col}{r}", create=False)
            if c is not None and c.find(q("f")) is None:
                set_value(c, None)
    for i, row in enumerate(rows):
        for j, col in enumerate("BCDEFGHIJKLM"):
            c = S.cell(f"{col}{9 + i}", style_from=tmpl[col])
            set_value(c, row[j])
    log["pier_rows"] = len(rows)
    # textos
    faltan = []
    for old, new in TEXTOS.items():
        if pkg.replace_sst(old, new) == 0:
            faltan.append(old[:60])
    log["textos_no_encontrados"] = faltan
    for i, s in enumerate(pkg.sst):
        if s.startswith("En ETABS: Display → Show Tables"):
            pkg.replace_sst(s, PIER_FORCES_NOTA)


# ---------------------------------------------------------------- fase B: recalculo
def recalcular(xlsx, workdir):
    prof = os.path.join(workdir, "lo_profile")
    os.makedirs(os.path.join(prof, "user"), exist_ok=True)
    with open(os.path.join(prof, "user", "registrymodifications.xcu"), "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
                '<value>0</value></prop></item>\n</oor:items>\n')
    out = os.path.join(workdir, "recalc")
    os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "xlsx",
                    "--outdir", out, xlsx], check=True, capture_output=True, timeout=900)
    return os.path.join(out, os.path.basename(xlsx))


# ---------------------------------------------------------------- fase C: formato y caches
def fase_formato(pkg, vals, log):
    st = pkg.styles
    fonts = st.find(q("fonts")).findall(q("font"))
    finfo = []
    for f in fonts:
        sz = f.find(q("sz"))
        finfo.append((float(sz.get("val")) if sz is not None else 11.0, f.find(q("b")) is not None))
    xfs = st.find(q("cellXfs"))
    xf_list = xfs.findall(q("xf"))
    wrap_cache = {}

    def is_wrap(s):
        al = xf_list[s].find(q("alignment"))
        return al is not None and al.get("wrapText") == "1"

    def wrap_style(s):
        if is_wrap(s):
            return s
        if s in wrap_cache:
            return wrap_cache[s]
        nx = copy.deepcopy(xf_list[s])
        al = nx.find(q("alignment"))
        if al is None:
            al = etree.SubElement(nx, q("alignment"))
        al.set("wrapText", "1")
        if not al.get("vertical"):
            al.set("vertical", "center")
        nx.set("applyAlignment", "1")
        xfs.append(nx)
        xf_list.append(nx)
        xfs.set("count", str(len(xf_list)))
        wrap_cache[s] = len(xf_list) - 1
        return wrap_cache[s]

    # 1) cache de formulas desde el recalculo de LibreOffice
    nupd = 0
    for name, S in pkg.sheets.items():
        V = vals.get(name, {})
        for rn, c in S.cells():
            if c.find(q("f")) is None:
                continue
            col, _ = split_ref(c.get("r"))
            v = V.get((rn, col2num(col)))
            ve = c.find(q("v"))
            if ve is None:
                ve = etree.SubElement(c, q("v"))
            if v is None:
                c.set("t", "str")
                ve.text = ""
            elif isinstance(v, bool):
                c.set("t", "b")
                ve.text = "1" if v else "0"
            elif isinstance(v, (int, float)):
                c.attrib.pop("t", None)
                ve.text = fmt_num(v)
            elif isinstance(v, str) and v in ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!"):
                c.set("t", "e")
                ve.text = v
            else:
                c.set("t", "str")
                ve.text = str(v)
            nupd += 1
    log["caches_actualizados"] = nupd

    # 2) columnas agrupadas
    def set_cols(S, lo, hi, **attrs):
        cols = S.root.find(q("cols"))
        if cols is None:
            cols = etree.Element(q("cols"))
            S.root.find(q("sheetData")).addprevious(cols)
        # partir rangos existentes
        new = []
        covered = set()
        for c in list(cols):
            a, b = int(c.get("min")), int(c.get("max"))
            for x0, x1, inside in ((a, min(b, lo - 1), False), (max(a, lo), min(b, hi), True), (max(a, hi + 1), b, False)):
                if x0 > x1:
                    continue
                cc = copy.deepcopy(c)
                cc.set("min", str(x0))
                cc.set("max", str(x1))
                if inside:
                    for k, v in attrs.items():
                        cc.set(k, v)
                    covered.update(range(x0, x1 + 1))
                new.append((x0, cc))
            cols.remove(c)
        fmt = S.root.find(q("sheetFormatPr"))
        dw = fmt.get("defaultColWidth") or "8.88671875"
        for x in range(lo, hi + 1):
            if x not in covered:
                cc = etree.Element(q("col"))
                cc.set("min", str(x))
                cc.set("max", str(x))
                cc.set("width", dw)
                for k, v in attrs.items():
                    cc.set(k, v)
                new.append((x, cc))
        for _, cc in sorted(new, key=lambda t: t[0]):
            cols.append(cc)
        if "outlineLevel" in attrs:
            fmt.set("outlineLevelCol", "1")

    grp = dict(hidden="1", outlineLevel="1")
    set_cols(pkg.sheets["MUROS"], col2num("I"), col2num("AB"), **grp)      # vanos (los piers no llevan)
    set_cols(pkg.sheets["MUROS"], col2num("AQ"), col2num("AW"), **grp)     # fuerzas capturadas (modo ETABS)
    set_cols(pkg.sheets["MUROS"], col2num("EU"), col2num("EU"), width="44", customWidth="1")
    set_cols(pkg.sheets["MUROS"], col2num("EW"), col2num("EW"), width="34", customWidth="1")
    for sh in ("DATOS", "DETALLE"):
        set_cols(pkg.sheets[sh], 5, 5, **grp)                               # "Como se calcula"

    # 3) alturas de renglon para que el texto no se corte
    def col_widths(S):
        fmt = S.root.find(q("sheetFormatPr"))
        dw = float(fmt.get("defaultColWidth") or 8.88671875)
        w = {}
        hidden = set()
        cols = S.root.find(q("cols"))
        if cols is not None:
            for c in cols:
                for x in range(int(c.get("min")), int(c.get("max")) + 1):
                    w[x] = float(c.get("width", dw))
                    if c.get("hidden") == "1":
                        hidden.add(x)
        return (lambda x: 0.0 if x in hidden else w.get(x, dw)), hidden

    skip_rows = {"PIER_FORCES": 9, "DISTORSIONES": 16, "CROQUIS_DATOS": 1}
    cambios = 0
    for name, S in pkg.sheets.items():
        if name == "CROQUIS_DATOS":
            continue
        width, hidden = col_widths(S)
        merges = {}
        mc = S.root.find(q("mergeCells"))
        if mc is not None:
            for m in mc:
                a, b = m.get("ref").split(":")
                (ca, ra), (cb, rb) = split_ref(a), split_ref(b)
                merges[a] = (col2num(ca), col2num(cb), ra, rb)
        fmt = S.root.find(q("sheetFormatPr"))
        default_h = float(fmt.get("defaultRowHeight") or 14.4)
        for rn, r in S.rows.items():
            if name in skip_rows and rn >= skip_rows[name]:
                continue
            need = 0.0
            for c in r.findall(q("c")):
                txt = pkg.text(c)
                if txt is None:
                    v = c.find(q("v"))
                    txt = v.text if v is not None and c.get("t") not in ("b", "e") else None
                    if txt is not None:
                        try:
                            txt = f"{float(txt):.2f}"
                        except ValueError:
                            pass
                if not txt:
                    continue
                col, _ = split_ref(c.get("r"))
                ci = col2num(col)
                if ci in hidden:
                    continue
                ref = c.get("r")
                if ref in merges:
                    c0, c1, r0, r1 = merges[ref]
                    if r1 > r0:
                        continue
                    wch = sum(width(x) for x in range(c0, c1 + 1))
                else:
                    wch = width(ci)
                s = int(c.get("s", "0"))
                fid = int(xf_list[s].get("fontId", "0"))
                size, bold = finfo[fid]
                px_avail = max(10.0, wch * 7.0 + 5 - 8)
                char_px = size * 96 / 72 * (0.56 if bold else 0.52)
                lines = 0
                for par in str(txt).split("\n"):
                    lines += max(1, math.ceil(len(par) * char_px * 1.08 / px_avail))
                if lines > 1 or len(str(txt)) * char_px > px_avail:
                    c.set("s", str(wrap_style(s)))
                need = max(need, lines * size * 1.25 + 4)
            if need == 0:
                continue
            cur = float(r.get("ht")) if r.get("ht") else default_h
            if need > cur + 0.5:
                r.set("ht", f"{need:.1f}")
                r.set("customHeight", "1")
                cambios += 1
    log["renglones_ajustados"] = cambios


# ---------------------------------------------------------------- croquis
def fase_croquis(pkg, vals, log):
    V = vals.get("CROQUIS_DATOS", {})

    def rng_values(f):
        m = re.match(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", f)
        sh, c0, r0, c1, r1 = m.groups()
        c1, r1 = c1 or c0, r1 or r0
        VV = vals.get(sh, {})
        out = []
        for rr in range(int(r0), int(r1) + 1):
            for cc in range(col2num(c0), col2num(c1) + 1):
                out.append(VV.get((rr, cc)))
        return out

    ocultas = 0
    for n in [x for x in pkg.names if re.match(r"xl/charts/chart\d+\.xml", x)]:
        root = etree.fromstring(pkg.data[n])
        # caches
        for ref in list(root.iter(qc("numRef"), qc("strRef"))):
            fe = ref.find(qc("f"))
            vs = rng_values(fe.text)
            if ref.tag == qc("numRef"):
                old = ref.find(qc("numCache"))
                fc = old.find(qc("formatCode")).text if old is not None and old.find(qc("formatCode")) is not None else "General"
                if old is not None:
                    ref.remove(old)
                cache = etree.SubElement(ref, qc("numCache"))
                etree.SubElement(cache, qc("formatCode")).text = fc
                etree.SubElement(cache, qc("ptCount")).set("val", str(len(vs)))
                for i, v in enumerate(vs):
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        pt = etree.SubElement(cache, qc("pt"))
                        pt.set("idx", str(i))
                        etree.SubElement(pt, qc("v")).text = fmt_num(v)
            else:
                old = ref.find(qc("strCache"))
                if old is not None:
                    ref.remove(old)
                cache = etree.SubElement(ref, qc("strCache"))
                etree.SubElement(cache, qc("ptCount")).set("val", str(len(vs)))
                for i, v in enumerate(vs):
                    if v not in (None, ""):
                        pt = etree.SubElement(cache, qc("pt"))
                        pt.set("idx", str(i))
                        etree.SubElement(pt, qc("v")).text = str(v)
        # leyenda: sin series de construccion ni "—"; abajo y sin encimarse
        legend = root.find(f".//{qc('legend')}")
        if legend is not None:
            for e in legend.findall(qc("legendEntry")):
                legend.remove(e)
            borrar = []
            for ser in root.iter(qc("ser")):
                idx = ser.find(qc("idx")).get("val")
                tx = ser.find(qc("tx"))
                name = None
                if tx is not None:
                    pts = tx.findall(f".//{qc('pt')}/{qc('v')}")
                    name = pts[0].text if pts else None
                    if name is None:
                        vv = tx.find(f".//{qc('v')}")
                        name = vv.text if vv is not None else None
                if name is None or name.strip() in ("", "—", "-"):
                    borrar.append(int(idx))
            pos = legend.find(qc("legendPos"))
            pos.set("val", "b")
            insert_at = list(legend).index(pos) + 1
            for k, i in enumerate(sorted(borrar)):
                le = etree.Element(qc("legendEntry"))
                etree.SubElement(le, qc("idx")).set("val", str(i))
                etree.SubElement(le, qc("delete")).set("val", "1")
                legend.insert(insert_at + k, le)
            ov = legend.find(qc("overlay"))
            if ov is not None:
                ov.set("val", "0")
            for d in legend.iter(qa("defRPr")):
                d.set("sz", "900")
            ocultas += len(borrar)
        pkg.data[n] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    log["entradas_leyenda_ocultas"] = ocultas


def leer_valores(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    out = {}
    for ws in wb.worksheets:
        d = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    d[(c.row, c.column)] = c.value
        out[ws.title] = d
    return out


def main(hoja, pier_xlsx, e2k, salida):
    log = {}
    work = tempfile.mkdtemp(prefix="hoja_muros_")
    pkg = Package(hoja)
    fase_contenido(pkg, hoja, pier_xlsx, e2k, log)
    paso_a = os.path.join(work, "paso_a.xlsx")
    pkg.save(paso_a)
    recalc = recalcular(paso_a, work)
    vals = leer_valores(recalc)
    fase_formato(pkg, vals, log)
    fase_croquis(pkg, vals, log)
    wbx = pkg.data["xl/workbook.xml"].decode("utf-8")
    wbx = re.sub(r"<calcPr([^/]*)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="\d"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wbx)
    pkg.data["xl/workbook.xml"] = wbx.encode("utf-8")
    pkg.save(salida)
    log["recalculo_libreoffice"] = recalc
    return log, vals


if __name__ == "__main__":
    log, vals = main(*sys.argv[1:5])
    for k, v in log.items():
        if k == "piers":
            print("piers:", len(v))
        else:
            print(k, ":", v)
