#!/usr/bin/env python3
"""Dibuja, por entrepiso, la elevacion de cada eje de muro antes y despues.
Uso: python3 elevaciones_muros.py ORIGINAL.e2k NUEVO.e2k CARPETA_SALIDA"""
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL  # noqa: E402

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
f = lambda s, b=False: ImageFont.truetype(BOLD if b else FONT, s)
PX = 52           # px por metro
RED = (255, 102, 102)
EDGE = (255, 255, 255)
BLUE = (30, 60, 210)
GREEN = (0, 150, 70)
GRAY = (150, 150, 150)
DARK = (25, 25, 25)

orig_txt = open(sys.argv[1], encoding="utf-8", newline="").read()
A = Model(orig_txt)
B = Model(open(sys.argv[2], encoding="utf-8", newline="").read())
outdir = sys.argv[3]
os.makedirs(outdir, exist_ok=True)

grids = {"X": [], "Y": []}
for m in re.finditer(r'GRID "[^"]+"\s+LABEL "([^"]+)"\s+DIR "(X|Y)"\s+COORD (\S+)', orig_txt):
    grids[m[2]].append((m[1], float(m[3])))


def grid_name(axis, coord):
    for lab, c in grids[axis]:
        if abs(c - coord) < 0.005:
            return f"Eje {lab}"
    return f"{axis} = {coord:.3f}"


def cross_grids(axis):
    # sobre un muro X=const corren las rejillas Y, y viceversa
    return grids["Y" if axis == "X" else "X"]


ga, _ = A.wall_pieces()
gb, _ = B.wall_pieces()
castA, castB = A.castillos(), B.castillos()


def draw_line(d, ox, oy, pieces, cast_h, h_min, zb, zt, axis, moved=()):
    H = (zt - zb) * PX
    X = lambda h: ox + (h - h_min) * PX
    Y = lambda z: oy + (zt - z) * PX
    for lab, c in cross_grids(axis):
        if h_min - 0.05 <= c <= h_min + 30:
            pass
    for p in pieces:
        d.rectangle([X(p["h0"]), Y(p["z1"]), X(p["h1"]), Y(p["z0"])], fill=RED, outline=EDGE, width=2)
    for h in cast_h:
        d.line([(X(h), oy), (X(h), oy + H)], fill=BLUE, width=3)
    for h in moved:
        d.polygon([(X(h) - 7, oy + H + 4), (X(h) + 7, oy + H + 4), (X(h), oy + H - 8)], fill=GREEN)
    d.line([(ox - 6, oy + H), (X(max(p["h1"] for p in pieces)) + 6, oy + H)], fill=GRAY, width=1)


rep_moves = {}
import json  # noqa: E402
rep_path = os.path.join(outdir, "..", "reporte_muros.json")
if os.path.exists(rep_path):
    for mv in json.load(open(rep_path, encoding="utf-8"))["orillas_movidas"]:
        rep_moves.setdefault((mv["piso"], mv["eje"]), []).append(mv["a"])

files = []
for story in A.order[:-1]:
    keys = sorted({k for k in set(ga) | set(gb) if k[0] == story}, key=lambda k: (k[1], k[2]))
    rows = []
    for k in keys:
        _, axis, coord = k
        pa, pb = ga.get(k, []), gb.get(k, [])
        allp = pa + pb
        h_min = min(p["h0"] for p in allp)
        h_max = max(p["h1"] for p in allp)
        ch = []
        for x, y, pt, ln in castB[story]:
            off, h = (x, y) if axis == "X" else (y, x)
            if abs(off - coord) <= TOL and h_min - TOL <= h <= h_max + TOL:
                ch.append(h)
        rows.append((k, pa, pb, ch, h_min, h_max))
    zt, zb = A.elev[story], A.elev[A.below(story)]
    Hpx = (zt - zb) * PX
    width = int(max(r[5] - r[4] for r in rows) * PX) + 260
    row_h = int(2 * Hpx + 150)
    img = Image.new("RGB", (max(width, 900), 120 + row_h * len(rows)), "white")
    d = ImageDraw.Draw(img)
    d.text((24, 18), f"Muros entrepiso {A.below(story)} a {story}: antes y despues", font=f(30, True), fill=DARK)
    d.text((24, 60), "rojo = tablero de muro    azul = eje de castillo    triangulo verde = orilla llevada al castillo",
           font=f(18), fill=DARK)
    y = 110
    for k, pa, pb, ch, h_min, h_max in rows:
        _, axis, coord = k
        eje = f"{axis}={coord:g}".replace(f"{axis}={coord:g}", f"{axis}={float(coord):g}")
        moved = rep_moves.get((story, f"{axis}={coord:g}"), []) or rep_moves.get((story, f"{axis}={float(coord)}"), [])
        tag = "   (cambia geometria)" if moved else ""
        d.text((24, y), f"{grid_name(axis, coord)}   [antes {len(pa)} pedazo{"s" if len(pa) != 1 else ""}, ahora {len(pb)} tablero{"es" if len(pb) != 1 else ""}]{tag}",
               font=f(20, True), fill=GREEN if moved else DARK)
        d.text((24, y + 34), "antes", font=f(15), fill=GRAY)
        draw_line(d, 120, y + 34, pa, ch, h_min, zb, zt, axis)
        d.text((24, y + 34 + Hpx + 30), "despues", font=f(15), fill=GRAY)
        draw_line(d, 120, y + 34 + Hpx + 30, pb, ch, h_min, zb, zt, axis, moved)
        # rejillas que cruzan
        for lab, c in cross_grids(axis):
            if h_min - TOL <= c <= h_max + TOL:
                xx = 120 + (c - h_min) * PX
                d.text((xx - 5, y + 34 + 2 * Hpx + 36), lab, font=f(14), fill=GRAY)
        y += row_h
    path = os.path.join(outdir, f"muros_{story.replace(' ', '').replace('+', 'mas')}.png")
    img.save(path)
    files.append(path)
print("\n".join(files))
