#!/usr/bin/env python3
"""Dibuja una planta por entrepiso con los ejes y el nombre de cada pier
sobre su muro, para ubicarlos rapido.
Uso: python3 plantas_piers.py MODELO.e2k CARPETA_SALIDA"""
import colorsys
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from redibujar_muros import Model, TOL  # noqa: E402

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F = lambda s, b=False: ImageFont.truetype(BOLD if b else FONT, s)
PX = 78           # px por metro
M0 = 150          # margen


def main(e2k, outdir):
    txt = open(e2k, encoding="utf-8", newline="").read()
    M = Model(txt)
    grids = {"X": [], "Y": []}
    for m in re.finditer(r'GRID "[^"]+"\s+LABEL "([^"]+)"\s+DIR "(X|Y)"\s+COORD (\S+)', txt):
        grids[m[2]].append((m[1], float(m[3])))
    walls, _ = M.wall_pieces()
    xs = [p[0] for p in M.points.values()] + [c for _, c in grids["X"]]
    ys = [p[1] for p in M.points.values()] + [c for _, c in grids["Y"]]
    x0, x1, y0, y1 = min(xs) - 0.3, max(xs) + 0.3, min(ys) - 0.3, max(ys) + 0.3
    W = int((x1 - x0) * PX) + 2 * M0 + 120
    H = int((y1 - y0) * PX) + 2 * M0 + 60
    X = lambda x: M0 + (x - x0) * PX
    Y = lambda y: H - M0 - (y - y0) * PX
    os.makedirs(outdir, exist_ok=True)
    out = []
    for story, titulo in ((M.order[-2], "PLANTA BAJA"), (M.order[-3], "PLANTA ALTA")):
        img = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(img)
        d.text((30, 20), f"{titulo} — piers (muros de {M.below(story)} a {story})", font=F(30, True), fill=(20, 20, 20))
        d.text((30, 62), "Color = un pier con su nombre · gris = muro sin pier (antepechos, dinteles, muretes) · cuadro negro = castillo",
               font=F(17), fill=(60, 60, 60))
        # ejes
        for lab, c in grids["X"]:
            d.line([(X(c), Y(y0) + 10), (X(c), Y(y1) - 10)], fill=(205, 215, 235), width=1)
            for yy in (Y(y1) - 40, Y(y0) + 40):
                d.ellipse([X(c) - 17, yy - 17, X(c) + 17, yy + 17], outline=(120, 140, 180), width=2, fill="white")
                d.text((X(c), yy), lab, font=F(17, True), fill=(60, 80, 130), anchor="mm")
        for lab, c in grids["Y"]:
            d.line([(X(x0) - 10, Y(c)), (X(x1) + 10, Y(c))], fill=(205, 215, 235), width=1)
            xx = X(x0) - 45
            d.ellipse([xx - 17, Y(c) - 17, xx + 17, Y(c) + 17], outline=(120, 140, 180), width=2, fill="white")
            d.text((xx, Y(c)), lab, font=F(15, True), fill=(60, 80, 130), anchor="mm")
        # muros
        piers = {}
        for (s, ax, co), pcs in walls.items():
            if s != story:
                continue
            for p in pcs:
                a = (co, p["h0"]) if ax == "X" else (p["h0"], co)
                b = (co, p["h1"]) if ax == "X" else (p["h1"], co)
                if p["pier"]:
                    e = piers.setdefault(p["pier"], dict(ax=ax, co=co, h0=9e9, h1=-9e9))
                    e["h0"], e["h1"] = min(e["h0"], p["h0"]), max(e["h1"], p["h1"])
                else:
                    d.line([(X(a[0]), Y(a[1])), (X(b[0]), Y(b[1]))], fill=(170, 170, 170), width=5)
        names = sorted(piers)
        for i, name in enumerate(names):
            e = piers[name]
            r, g, b = colorsys.hsv_to_rgb((i * 0.618) % 1, 0.75, 0.80)
            col = (int(r * 255), int(g * 255), int(b * 255))
            if e["ax"] == "X":
                p0, p1 = (X(e["co"]), Y(e["h0"])), (X(e["co"]), Y(e["h1"]))
            else:
                p0, p1 = (X(e["h0"]), Y(e["co"])), (X(e["h1"]), Y(e["co"]))
            d.line([p0, p1], fill=col, width=9)
            e["col"], e["p0"], e["p1"] = col, p0, p1
        # castillos
        for x, y, pt, ln in M.castillos()[story]:
            d.rectangle([X(x) - 5, Y(y) - 5, X(x) + 5, Y(y) + 5], fill=(25, 25, 25))
        # etiquetas (sin encimarse: prueba posiciones a lo largo del muro y de los dos lados)
        puestas = []

        def libre(bx):
            return all(bx[2] <= o[0] or bx[0] >= o[2] or bx[3] <= o[1] or bx[1] >= o[3] for o in puestas)

        for name in names:
            e = piers[name]
            (ax0, ay0), (ax1, ay1) = e["p0"], e["p1"]
            f = F(16, True)
            tw, th = d.textbbox((0, 0), name, font=f)[2:]
            lab = Image.new("RGBA", (tw + 10, th + 8), (255, 255, 255, 0))
            ld = ImageDraw.Draw(lab)
            ld.rounded_rectangle([0, 0, tw + 9, th + 7], radius=4, fill=(255, 255, 255, 235), outline=e["col"], width=2)
            ld.text((5, 3), name, font=f, fill=e["col"])
            mx, my = (ax0 + ax1) / 2, (ay0 + ay1) / 2
            if e["ax"] == "X":          # muro vertical en la planta: etiqueta girada junto al muro
                lab = lab.rotate(90, expand=True)
                w, h = lab.size
                cands = [(mx + 8 + side, my - h / 2 + k) for k in (0, -40, 40, -80, 80, -120, 120) for side in (0, -w - 16)]
            else:                        # muro horizontal: etiqueta arriba o abajo
                w, h = lab.size
                cands = [(mx - w / 2 + k, my + side) for k in (0, -50, 50, -100, 100) for side in (-h - 8, 8)]
            for cx, cy in cands:
                bx = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
                if libre(bx):
                    break
            puestas.append(bx)
            img.paste(lab, (int(cx), int(cy)), lab)
        path = os.path.join(outdir, f"piers_{'PB' if titulo.endswith('BAJA') else 'PA'}.png")
        img.save(path)
        out.append((path, len(names)))
    for p, n in out:
        print(p, n, "piers")


if __name__ == "__main__":
    main(*sys.argv[1:3])
