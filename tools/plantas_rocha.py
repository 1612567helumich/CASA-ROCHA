#!/usr/bin/env python3
"""Plantas sencillas para el cliente, sin textos a las orillas:
  hoja 1, PLANTA BAJA: el ratio de cada muro a un lado (azul = pasa, rojo = no pasa).
  hoja 2, PLANTA ALTA: cada viga que cambia en negro, con su medida (ancho x peralte, cm).
Ratios igual que plantas_piers_ratios.py (tabique: hoja de muros; concreto: revision NTC-C).

Uso: python3 plantas_rocha.py MODELO.e2k HOJA.xlsx PROPUESTA.json SALIDA.pdf
"""
import json
import math
import os
import re
import sys

from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from plantas_piers_ratios import (F, Model, claro_de, leer_hoja, materiales, revisa_concreto,  # noqa: E402
                                  revisa_tabique, vigas_del_piso)

PX, M0 = 78, 150
TABIQUE, CONCRETO, GRIS = (200, 112, 75), (95, 95, 95), (180, 180, 180)
AZUL, ROJO = (21, 101, 192), (200, 35, 35)


def main(e2k, hoja, propuesta, salida):
    txt = open(e2k, encoding="utf-8", newline="").read()
    M = Model(txt)
    prop = json.load(open(propuesta, encoding="utf-8"))
    res, frfm = leer_hoja(hoja)
    secs = materiales(txt)
    grids = {"X": [], "Y": []}
    for m in re.finditer(r'GRID "[^"]+"\s+LABEL "([^"]+)"\s+DIR "(X|Y)"\s+COORD (\S+)', txt):
        grids[m[2]].append((m[1], float(m[3])))
    walls, _ = M.wall_pieces()
    xs = [p[0] for p in M.points.values()] + [c for _, c in grids["X"]]
    ys = [p[1] for p in M.points.values()] + [c for _, c in grids["Y"]]
    x0, x1, y0, y1 = min(xs) - 0.3, max(xs) + 0.3, min(ys) - 0.3, max(ys) + 0.3
    W = int((x1 - x0) * PX) + 2 * M0 + 80
    H = int((y1 - y0) * PX) + 2 * M0 + 110
    X = lambda x: M0 + (x - x0) * PX
    Y = lambda y: H - M0 - (y - y0) * PX

    def lienzo(titulo):
        img = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(img)
        d.text((30, 22), titulo, font=F(32, True), fill=(0, 0, 0))
        for lab, c in grids["X"]:
            d.line([(X(c), Y(y0) + 10), (X(c), Y(y1) - 10)], fill=(218, 224, 238), width=1)
            for yy in (Y(y1) - 40, Y(y0) + 40):
                d.ellipse([X(c) - 17, yy - 17, X(c) + 17, yy + 17], outline=(120, 140, 180), width=2, fill="white")
                d.text((X(c), yy), lab, font=F(17, True), fill=(60, 80, 130), anchor="mm")
        for lab, c in grids["Y"]:
            d.line([(X(x0) - 10, Y(c)), (X(x1) + 10, Y(c))], fill=(218, 224, 238), width=1)
            xx = X(x0) - 45
            d.ellipse([xx - 17, Y(c) - 17, xx + 17, Y(c) + 17], outline=(120, 140, 180), width=2, fill="white")
            d.text((xx, Y(c)), lab, font=F(15, True), fill=(60, 80, 130), anchor="mm")
        return img, d

    def muros(d, story):
        piers = {}
        for (s, ax, co), pcs in walls.items():
            if s != story:
                continue
            for p in pcs:
                a = (co, p["h0"]) if ax == "X" else (p["h0"], co)
                b = (co, p["h1"]) if ax == "X" else (p["h1"], co)
                conc = secs.get(p["section"], {}).get("tipo") == "Concrete"
                col = (CONCRETO if conc else TABIQUE) if p["pier"] else GRIS
                d.line([(X(a[0]), Y(a[1])), (X(b[0]), Y(b[1]))], fill=col, width=9 if p["pier"] else 5)
                if p["pier"]:
                    e = piers.setdefault(p["pier"], dict(ax=ax, co=co, h0=9e9, h1=-9e9, conc=conc,
                                                         sec=secs.get(p["section"], {})))
                    e["h0"], e["h1"] = min(e["h0"], p["h0"]), max(e["h1"], p["h1"])
        return piers

    def poner(d, puestas, texto, color, cx0, cy0, rot, f, borde=None, vertical=None):
        tw, th = d.textbbox((0, 0), texto, font=f)[2:]
        w, h = (th + 10, tw + 10) if rot else (tw + 10, th + 10)
        if rot if vertical is None else vertical:
            cands = [(cx0 + 16 + s, cy0 - h / 2 + k) for k in (0, -40, 40, -80, 80, -120, 120) for s in (0, -w - 32)]
        else:
            cands = [(cx0 - w / 2 + k, cy0 + s) for k in (0, -50, 50, -100, 100) for s in (-h - 10, 10, -2 * h - 16, h + 16)]
        for cx, cy in cands:
            bx = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
            if cx > M0 - 10 and all(bx[2] <= o[0] or bx[0] >= o[2] or bx[3] <= o[1] or bx[1] >= o[3] for o in puestas):
                break
        puestas.append(bx)
        lab = Image.new("RGBA", (int(tw + 10), int(th + 10)), (255, 255, 255, 0))
        ld = ImageDraw.Draw(lab)
        ld.rounded_rectangle([0, 0, tw + 9, th + 9], radius=5, fill=(255, 255, 255, 235),
                             outline=borde if borde else (255, 255, 255, 0), width=3 if borde else 1)
        ld.text((5, 4), texto, font=f, fill=color)
        if rot:
            lab = lab.rotate(90, expand=True)
        return lab, (int(cx), int(cy))

    # ---------------- hoja 1: planta baja con ratios
    pb = M.order[-2]
    Hm = (M.elev[pb] - M.elev[M.below(pb)]) * 100
    img1, d = lienzo("PLANTA BAJA — ratio de cada muro")
    piers = muros(d, pb)
    puestas = []
    for n, e in piers.items():
        r = res.get(n)
        if r is None:
            continue
        if e["conc"]:
            ratio, no, extra = revisa_concreto(r, e["sec"]["t"], e["sec"]["fc"], Hm)
        else:
            ratio, no, extra = revisa_tabique(r, frfm)
        texto = "col." if extra == "col." else ("no pasa" if ratio == math.inf else f"{ratio:.2f}")
        (a0, b0) = ((X(e["co"]), Y(e["h0"])), (X(e["co"]), Y(e["h1"]))) if e["ax"] == "X" else \
                   ((X(e["h0"]), Y(e["co"])), (X(e["h1"]), Y(e["co"])))
        lab, pos = poner(d, puestas, texto, ROJO if no else AZUL, (a0[0] + b0[0]) / 2, (a0[1] + b0[1]) / 2,
                         e["ax"] == "X" and texto != "col.", F(22, True),
                         vertical=e["ax"] == "X")
        img1.paste(lab, pos, lab)

    # ---------------- hoja 2: planta alta con las vigas que cambian
    pa = M.order[-3]
    abajo = M.below(pa)
    img2, d = lienzo("PLANTA ALTA — medida de las vigas (cm)")
    piers = muros(ImageDraw.Draw(Image.new("RGB", (W, H))), pa)      # solo para tener los piers
    vigas = vigas_del_piso(txt, abajo)
    cast = M.castillos()[abajo]
    grupos = {}
    for n, pp in prop.items():
        if n.startswith("_") or n not in piers or n not in res:
            continue
        e = piers[n]
        a, b = claro_de(e, res[n]["AM"], e["ax"], e["co"], walls, abajo, cast, vigas)
        grupos.setdefault((e["ax"], round(e["co"], 3), round(a, 2), round(b, 2)),
                          dict(ax=e["ax"], co=e["co"], a=a, b=b, Lv=res[n]["AM"], sec=pp["seccion"]))
    muros(d, pa)
    puestas = []
    for g in grupos.values():                        # la viga: linea negra a lo largo de su claro, encima del muro
        if g["ax"] == "X":
            p0, p1 = (X(g["co"]), Y(g["a"])), (X(g["co"]), Y(g["b"]))
        else:
            p0, p1 = (X(g["a"]), Y(g["co"])), (X(g["b"]), Y(g["co"]))
        d.line([p0, p1], fill=(0, 0, 0), width=12)
        puestas.append((min(p0[0], p1[0]) - 9, min(p0[1], p1[1]) - 9, max(p0[0], p1[0]) + 9, max(p0[1], p1[1]) + 9))
    for g in sorted(grupos.values(), key=lambda g: -g["Lv"]):
        if g["ax"] == "X":
            cx0, cy0 = X(g["co"]), (Y(g["a"]) + Y(g["b"])) / 2
        else:
            cx0, cy0 = (X(g["a"]) + X(g["b"])) / 2, Y(g["co"])
        lab, pos = poner(d, puestas, f"{g['sec']} peralte", (0, 0, 0), cx0, cy0, False, F(24, True), borde=(0, 0, 0),
                         vertical=g["ax"] == "X")
        img2.paste(lab, pos, lab)

    img1.save(salida, save_all=True, append_images=[img2], resolution=150)
    print("listo:", salida)


if __name__ == "__main__":
    main(*sys.argv[1:5])
