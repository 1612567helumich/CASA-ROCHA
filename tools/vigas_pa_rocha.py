#!/usr/bin/env python3
"""Planta alta para el cliente: muros de PA, ovalo en cada viga que cambia y su medida (ancho x peralte).
Usa la geometria del modelo, los claros (L_v) de la hoja de muros y la propuesta (JSON).

Uso: python3 vigas_pa_rocha.py MODELO.e2k HOJA.xlsx PROPUESTA.json SALIDA.pdf
     PROPUESTA.json: {"PIER": {"seccion": "30×45"}, ..., "_rocha_extra": ["...", ...]}
"""
import json
import os
import re
import sys

from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from plantas_piers_ratios import F, Model, claro_de, leer_hoja, vigas_del_piso  # noqa: E402

PX, M0 = 78, 150
MURO_PA, GRIS = (200, 112, 75), (175, 175, 175)


def main(e2k, hoja, propuesta, salida):
    txt = open(e2k, encoding="utf-8", newline="").read()
    M = Model(txt)
    prop = json.load(open(propuesta, encoding="utf-8"))
    res, _ = leer_hoja(hoja)
    grids = {"X": [], "Y": []}
    for m in re.finditer(r'GRID "[^"]+"\s+LABEL "([^"]+)"\s+DIR "(X|Y)"\s+COORD (\S+)', txt):
        grids[m[2]].append((m[1], float(m[3])))

    def eje(dir_, c):
        for lab, cc in grids[dir_]:
            if abs(cc - c) < 0.02:
                return lab
        return f"{dir_}{c:.2f}"

    walls, _ = M.wall_pieces()
    xs = [p[0] for p in M.points.values()] + [c for _, c in grids["X"]]
    ys = [p[1] for p in M.points.values()] + [c for _, c in grids["Y"]]
    x0, x1, y0, y1 = min(xs) - 0.3, max(xs) + 0.3, min(ys) - 0.3, max(ys) + 0.3
    W = int((x1 - x0) * PX) + 2 * M0 + 520
    H = int((y1 - y0) * PX) + 2 * M0 + 120
    X = lambda x: M0 + (x - x0) * PX
    Y = lambda y: H - M0 - (y - y0) * PX
    story = M.order[-3]                       # planta alta
    abajo = M.below(story)
    img = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(img)
    d.text((30, 22), "PLANTA ALTA — cómo deben quedar las vigas", font=F(34, True), fill=(0, 0, 0))
    d.text((30, 70), "medidas en cm: ancho × peralte", font=F(20), fill=(70, 70, 70))
    # ejes
    for lab, c in grids["X"]:
        d.line([(X(c), Y(y0) + 10), (X(c), Y(y1) - 10)], fill=(215, 222, 238), width=1)
        for yy in (Y(y1) - 40, Y(y0) + 40):
            d.ellipse([X(c) - 17, yy - 17, X(c) + 17, yy + 17], outline=(120, 140, 180), width=2, fill="white")
            d.text((X(c), yy), lab, font=F(17, True), fill=(60, 80, 130), anchor="mm")
    for lab, c in grids["Y"]:
        d.line([(X(x0) - 10, Y(c)), (X(x1) + 10, Y(c))], fill=(215, 222, 238), width=1)
        xx = X(x0) - 45
        d.ellipse([xx - 17, Y(c) - 17, xx + 17, Y(c) + 17], outline=(120, 140, 180), width=2, fill="white")
        d.text((xx, Y(c)), lab, font=F(15, True), fill=(60, 80, 130), anchor="mm")
    # muros de PA
    piers = {}
    for (s, ax, co), pcs in walls.items():
        if s != story:
            continue
        for p in pcs:
            a = (co, p["h0"]) if ax == "X" else (p["h0"], co)
            b = (co, p["h1"]) if ax == "X" else (p["h1"], co)
            d.line([(X(a[0]), Y(a[1])), (X(b[0]), Y(b[1]))], fill=MURO_PA if p["pier"] else GRIS, width=8 if p["pier"] else 5)
            if p["pier"]:
                e = piers.setdefault(p["pier"], dict(ax=ax, co=co, h0=9e9, h1=-9e9))
                e["h0"], e["h1"] = min(e["h0"], p["h0"]), max(e["h1"], p["h1"])
    # vigas que cambian: un ovalo por claro
    vigas = vigas_del_piso(txt, abajo)
    cast = M.castillos()[abajo]
    grupos = {}
    for n, pp in prop.items():
        if n.startswith("_") or n not in piers or n not in res:
            continue
        e = piers[n]
        a, b = claro_de(e, res[n]["AM"], e["ax"], e["co"], walls, abajo, cast, vigas)
        g = grupos.setdefault((e["ax"], round(e["co"], 3), round(a, 2), round(b, 2)),
                              dict(ax=e["ax"], co=e["co"], a=a, b=b, Lv=res[n]["AM"], sec=pp["seccion"]))
    orden = sorted(grupos.values(), key=lambda g: -g["Lv"])
    puestas = []
    for i, g in enumerate(orden, 1):
        if g["ax"] == "X":
            box = [X(g["co"]) - 36, Y(g["b"]) - 18, X(g["co"]) + 36, Y(g["a"]) + 18]
        else:
            box = [X(g["a"]) - 18, Y(g["co"]) - 36, X(g["b"]) + 18, Y(g["co"]) + 36]
        d.ellipse(box, outline=(0, 0, 0), width=5)
        texto = f"{i}:  {g['sec']}"
        f = F(22, True)
        tw, th = d.textbbox((0, 0), texto, font=f)[2:]
        w, h = tw + 14, th + 12
        cx0, cy0 = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
        cands = [(cx0 - w / 2, box[3] + 6), (cx0 - w / 2, box[1] - h - 6), (box[2] + 8, cy0 - h / 2),
                 (box[0] - w - 8, cy0 - h / 2)]
        cands += [(x_ + dx, y_ + dy) for x_, y_ in cands[:4] for dx, dy in ((0, 34), (0, -34), (50, 0), (-50, 0))]
        for cx, cy in cands:
            bx = (cx - 3, cy - 3, cx + w + 3, cy + h + 3)
            if cx > M0 - 10 and all(bx[2] <= o[0] or bx[0] >= o[2] or bx[3] <= o[1] or bx[1] >= o[3] for o in puestas):
                break
        puestas.append(bx)
        d.rounded_rectangle([cx, cy, cx + w, cy + h], radius=6, fill="white", outline=(0, 0, 0), width=3)
        d.text((cx + 7, cy + 5), texto, font=f, fill=(0, 0, 0))
    # lista a la derecha
    tx, ty = int((x1 - x0) * PX) + 2 * M0 + 20, 200
    d.text((tx, ty), "VIGAS QUE CAMBIAN", font=F(26, True), fill=(0, 0, 0))
    ty += 48
    for i, g in enumerate(orden, 1):
        lin, otro = ("Y", "X") if g["ax"] == "Y" else ("X", "Y")
        d.text((tx, ty), f"{i}.  Eje {eje(lin, g['co'])}:", font=F(22), fill=(0, 0, 0))
        d.text((tx + 250, ty), g["sec"], font=F(22, True), fill=(0, 0, 0))
        ty += 36
    ty += 16
    for ln in prop.get("_rocha_extra", []):
        d.text((tx, ty), ln, font=F(19), fill=(60, 60, 60))
        ty += 30
    img.save(salida, resolution=150)
    print(f"{len(orden)} vigas marcadas")


if __name__ == "__main__":
    main(*sys.argv[1:5])
