#!/usr/bin/env python3
"""Dibuja, por entrepiso, la elevacion de cada eje de muro antes y despues.
En "antes" pinta en amarillo lo que se quita y en "despues" en verde lo que
se agrega, para que las tiras de pocos centimetros se vean.
Uso: python3 elevaciones_muros.py ORIGINAL.e2k NUEVO.e2k CARPETA_SALIDA"""
import json
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL, cluster  # noqa: E402

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
f = lambda s, b=False: ImageFont.truetype(BOLD if b else FONT, s)
PX = 52           # px por metro
MAX_H = 7800      # alto maximo de cada imagen
RED = (255, 102, 102)
EDGE = (255, 255, 255)
BLUE = (30, 60, 210)
GREEN = (0, 150, 70)
YELLOW = (255, 205, 0)
ADDED = (0, 190, 90)
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


def diff(P, Q):
    """Rectangulos de P que no estan en Q (en el plano h-z)."""
    hs = cluster([p[k] for p in P + Q for k in ("h0", "h1")])
    zs = cluster([p[k] for p in P + Q for k in ("z0", "z1")])
    inside = lambda S, h, z: any(p["h0"] - TOL < h < p["h1"] + TOL and p["z0"] - TOL < z < p["z1"] + TOL for p in S)
    out = []
    for a, b in zip(hs, hs[1:]):
        for c, d in zip(zs, zs[1:]):
            hm, zm = (a + b) / 2, (c + d) / 2
            if inside(P, hm, zm) and not inside(Q, hm, zm):
                out.append((a, b, c, d))
    return out


def draw_line(d, ox, oy, pieces, cast_h, h_min, zb, zt, marks, mark_col, moved=()):
    H = (zt - zb) * PX
    X = lambda h: ox + (h - h_min) * PX
    Y = lambda z: oy + (zt - z) * PX
    for p in pieces:
        d.rectangle([X(p["h0"]), Y(p["z1"]), X(p["h1"]), Y(p["z0"])], fill=RED, outline=EDGE, width=1)
    for a, b, c, e in marks:
        x0, x1 = X(a), X(b)
        if x1 - x0 < 5:                      # que se vea aunque mida 7.5 cm
            mid = (x0 + x1) / 2
            x0, x1 = mid - 2.5, mid + 2.5
        d.rectangle([x0, Y(e), x1, Y(c)], fill=mark_col)
    for h in cast_h:
        d.line([(X(h), oy), (X(h), oy + H)], fill=BLUE, width=2)
    for h in moved:
        d.polygon([(X(h) - 7, oy + H + 4), (X(h) + 7, oy + H + 4), (X(h), oy + H - 8)], fill=GREEN)
    d.line([(ox - 6, oy + H), (X(max(p["h1"] for p in pieces)) + 6, oy + H)], fill=GRAY, width=1)


rep_moves = {}
rep_path = os.path.join(outdir, "..", "reporte_muros.json")
if os.path.exists(rep_path):
    for mv in json.load(open(rep_path, encoding="utf-8"))["orillas_movidas"]:
        rep_moves.setdefault((mv["piso"], mv["eje"]), []).append(mv["a"])

ga, _ = A.wall_pieces()
gb, _ = B.wall_pieces()
castB = B.castillos()
plural = lambda n, s, pl: f"{n} {s if n == 1 else pl}"

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
    width = max(int(max(r[5] - r[4] for r in rows) * PX) + 260, 1000)
    row_h = int(2 * Hpx + 150)
    per_img = max(1, (MAX_H - 140) // row_h)
    chunks = [rows[i:i + per_img] for i in range(0, len(rows), per_img)]
    for ci, chunk in enumerate(chunks, 1):
        img = Image.new("RGB", (width, 140 + row_h * len(chunk)), "white")
        d = ImageDraw.Draw(img)
        part = f"  ({ci} de {len(chunks)})" if len(chunks) > 1 else ""
        d.text((24, 18), f"Muros entrepiso {A.below(story)} a {story}: antes y despues{part}", font=f(30, True), fill=DARK)
        d.text((24, 60), "rojo = muro    azul = eje de castillo    triangulo verde = orilla llevada al castillo", font=f(18), fill=DARK)
        d.text((24, 86), "amarillo (antes) = tira de muro que se quita    verde (despues) = muro que se agrega", font=f(18), fill=DARK)
        y = 135
        for k, pa, pb, ch, h_min, h_max in chunk:
            _, axis, coord = k
            moved = rep_moves.get((story, f"{axis}={coord:g}"), [])
            removed, added = diff(pa, pb), diff(pb, pa)
            tag = "   (cambia geometria)" if (removed or added) else ""
            d.text((24, y), f"{grid_name(axis, coord)}   [antes {plural(len(pa), 'pedazo', 'pedazos')}, "
                            f"ahora {plural(len(pb), 'tablero', 'tableros')}]{tag}",
                   font=f(20, True), fill=GREEN if tag else DARK)
            d.text((24, y + 34), "antes", font=f(15), fill=GRAY)
            draw_line(d, 120, y + 34, pa, ch, h_min, zb, zt, removed, YELLOW)
            d.text((24, y + 34 + Hpx + 30), "despues", font=f(15), fill=GRAY)
            draw_line(d, 120, y + 34 + Hpx + 30, pb, ch, h_min, zb, zt, added, ADDED, moved)
            for lab, c in cross_grids(axis):
                if h_min - TOL <= c <= h_max + TOL:
                    xx = 120 + (c - h_min) * PX
                    d.text((xx - 5, y + 34 + 2 * Hpx + 36), lab, font=f(14), fill=GRAY)
            y += row_h
        suffix = f"_{ci}" if len(chunks) > 1 else ""
        path = os.path.join(outdir, f"muros_{story.replace(' ', '').replace('+', 'mas')}{suffix}.png")
        img.save(path)
        files.append(path)
print("\n".join(files))
