#!/usr/bin/env python3
"""Plantas sencillas para el cliente, sin textos a las orillas:
  hoja 1, PLANTA BAJA: el ratio de cada muro a un lado (azul = pasa, rojo = no pasa). Donde la norma no
  da ratio (machon o V_R = 0 por tension) va el de cortante con P = 0, en rojo.
  hoja 2, PLANTA ALTA: todas las dalas bajo muros de PA en negro con su medida (ancho x peralte, cm) y
  una flecha de la etiqueta a su dala;
  las que cambian con letra grande y las demas (clave _resto del JSON, 15x30) con letra chica.
  Si el claro pide mas de 30 de peralte por L/h <= 14 (NTC-M 4.1.18), abajo va "min. norma NN".
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

    def lienzo(titulo, extra=0):
        img = Image.new("RGB", (W + extra, H), "white")
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

    def largo_flecha(pos, size, seg):
        """Lo que mide la flecha de la etiqueta a su dala (0 si no cabe)."""
        cx, cy = pos[0] + size[0] / 2, pos[1] + size[1] / 2
        (ax_, ay_), (bx_, by_) = seg
        vx, vy = bx_ - ax_, by_ - ay_
        L2 = vx * vx + vy * vy or 1.0
        t = max(0.0, min(1.0, ((cx - ax_) * vx + (cy - ay_) * vy) / L2))
        dx, dy = ax_ + t * vx - cx, ay_ + t * vy - cy
        dist = math.hypot(dx, dy)
        k = min((size[0] / 2 + 3) / abs(dx) if dx else 9e9, (size[1] / 2 + 3) / abs(dy) if dy else 9e9)
        return 0.0 if dist < 1 or k >= 1 else dist * (1 - k) - 8

    def poner(d, puestas, texto, color, cx0, cy0, rot, f, borde=None, vertical=None, gap=None, seg=None):
        tw, th = d.textbbox((0, 0), texto, font=f)[2:]
        w, h = (th + 10, tw + 10) if rot else (tw + 10, th + 10)
        if rot if vertical is None else vertical:
            g_ = 16 if gap is None else gap
            cands = [(cx0 + g_ + s, cy0 - h / 2 + k) for k in (0, -40, 40, -80, 80, -120, 120, -170, 170, -220, 220)
                     for s in (0, -w - 2 * g_)]
        else:
            ss = (-h - 10, 10, -2 * h - 16, h + 16) if gap is None else (-h - gap, gap, -h - gap - 35, gap + 35)
            cands = [(cx0 - w / 2 + k, cy0 + s) for k in (0, -50, 50, -100, 100, -150, 150) for s in ss]
        for rad in (60, 100, 140, 180, 230):                      # si no cabe cerca, alrededor y mas lejos
            for ang in range(0, 360, 30):
                cands.append((cx0 + rad * math.cos(math.radians(ang)) - w / 2, cy0 + rad * math.sin(math.radians(ang)) - h / 2))
        cabe = lambda bx: seg is None or largo_flecha((bx[0] + 2, bx[1] + 2), (w, h), seg) >= 22
        libre = lambda bx: cx_ok(bx) and cabe(bx) and all(bx[2] <= o[0] or bx[0] >= o[2] or bx[3] <= o[1] or bx[1] >= o[3]
                                                         for o in puestas)
        cx_ok = lambda bx: bx[0] > M0 - 12 and bx[2] < W - 5 and bx[1] > 100 and bx[3] < H - 5
        bx = None
        for cx, cy in cands:
            prueba = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
            if libre(prueba):
                bx = prueba
                break
        if bx is None:
            cx, cy = cands[0]
            bx = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
        puestas.append(bx)
        lab = Image.new("RGBA", (int(tw + 10), int(th + 10)), (255, 255, 255, 0))
        ld = ImageDraw.Draw(lab)
        ld.rounded_rectangle([0, 0, tw + 9, th + 9], radius=5, fill=(255, 255, 255, 235),
                             outline=borde if borde else (255, 255, 255, 0), width=3 if borde else 1)
        ld.text((5, 4), texto, font=f, fill=color)
        if rot:
            lab = lab.rotate(90, expand=True)
        return lab, (int(cx), int(cy))

    def flecha(d, pos, size, seg):
        """Flecha de la etiqueta (pos, size) al punto mas cercano de la dala (seg)."""
        cx, cy = pos[0] + size[0] / 2, pos[1] + size[1] / 2
        (ax_, ay_), (bx_, by_) = seg
        vx, vy = bx_ - ax_, by_ - ay_
        L2 = vx * vx + vy * vy or 1.0
        t = max(0.0, min(1.0, ((cx - ax_) * vx + (cy - ay_) * vy) / L2))
        qx, qy = ax_ + t * vx, ay_ + t * vy
        dx, dy = qx - cx, qy - cy
        dist = math.hypot(dx, dy)
        hw, hh = size[0] / 2 + 3, size[1] / 2 + 3
        k = min(hw / abs(dx) if dx else 9e9, hh / abs(dy) if dy else 9e9)
        if dist < 1 or k >= 1:                       # la etiqueta quedo encima de su dala: no hay flecha
            return False
        ux, uy = dx / dist, dy / dist
        sx, sy = cx + dx * k, cy + dy * k            # sale del borde de la etiqueta
        tx_, ty_ = qx - ux * 8, qy - uy * 8          # punta junto a la dala
        if math.hypot(tx_ - sx, ty_ - sy) < 6:
            return False
        d.line([(sx, sy), (tx_, ty_)], fill=(0, 0, 0), width=3)
        bx2, by2 = tx_ - ux * 16, ty_ - uy * 16
        d.polygon([(tx_, ty_), (bx2 - uy * 7, by2 + ux * 7), (bx2 + uy * 7, by2 - ux * 7)], fill=(0, 0, 0))
        return True

    import openpyxl
    _D = openpyxl.load_workbook(hoja, data_only=True)["DATOS"]
    fr_v, vm, t_tip = _D["C80"].value, _D["C35"].value, _D["C52"].value      # F_R cortante, v'm, t tipico
    # ---------------- hoja 1: planta baja con ratios
    pb = M.order[-2]
    Hm = (M.elev[pb] - M.elev[M.below(pb)]) * 100
    PANEL1 = 780
    img1, d = lienzo("PLANTA BAJA — ratio de cada muro", PANEL1)
    piers = muros(d, pb)
    puestas = []
    especiales = []
    for n, e in piers.items():
        r = res.get(n)
        if r is None:
            continue
        if e["conc"]:
            ratio, no, extra = revisa_concreto(r, e["sec"]["t"], e["sec"]["fc"], Hm)
        else:
            ratio, no, extra = revisa_tabique(r, frfm)
        if extra == "col." or ratio == math.inf:
            # sin ratio de norma (machon o V_R = 0 por tension): cortante como si P = 0, se queda en rojo
            AT = r["CE"] or r["F"] * 100 * t_tip
            f_ = r["CY"] or 1.0
            vr0 = min(fr_v * 0.5 * vm * AT * f_, 1.5 * fr_v * vm * AT * f_) / 1000
            ratio, no = r["BK"] / vr0, no or ["sin V_R"]
            if extra == "col.":
                motivo = f"Machón de {r['F'] * 100:.0f} cm: va como columna"
            else:
                motivo = "Con tensión la norma no cuenta al tabique"
                motivo += {3: "; con alambres no alcanza, va con malla",
                           -1: "; ni con refuerzo horizontal alcanza"}.get(r["ED"], "")
            especiales.append((n, ratio, motivo))
        texto = f"{ratio:.2f}"
        (a0, b0) = ((X(e["co"]), Y(e["h0"])), (X(e["co"]), Y(e["h1"]))) if e["ax"] == "X" else \
                   ((X(e["h0"]), Y(e["co"])), (X(e["h1"]), Y(e["co"])))
        lab, pos = poner(d, puestas, texto, ROJO if no else AZUL, (a0[0] + b0[0]) / 2, (a0[1] + b0[1]) / 2,
                         e["ax"] == "X", F(22, True),
                         vertical=e["ax"] == "X")
        img1.paste(lab, pos, lab)
    # tablita a un lado
    tx, ty = W + 10, 190
    d.text((tx, ty), "Muros sin ratio de norma", font=F(28, True), fill=(0, 0, 0))
    d.text((tx, ty + 38), "(su número es el cortante con P = 0)", font=F(22), fill=(60, 60, 60))
    ty += 88
    c1, c2, c3 = tx, tx + 190, tx + 290
    for c_, t_ in ((c1, "Muro"), (c2, "Ratio"), (c3, "Por qué va en rojo")):
        d.text((c_, ty), t_, font=F(21, True), fill=(80, 80, 80))
    ty += 34
    d.line([(tx, ty), (W + PANEL1 - 20, ty)], fill=(160, 160, 160), width=1)
    ty += 8
    for n, ratio, motivo in sorted(especiales, key=lambda x: -x[1]):
        d.text((c1, ty), n, font=F(21, True), fill=ROJO)
        d.text((c2, ty), f"{ratio:.2f}", font=F(21, True), fill=ROJO)
        lineas, cur = [], ""
        for wd in motivo.split(" "):
            prueba = (cur + " " + wd).strip()
            if d.textlength(prueba, font=F(20)) <= W + PANEL1 - 25 - c3:
                cur = prueba
            else:
                lineas.append(cur); cur = wd
        lineas.append(cur)
        for i, ln in enumerate(lineas):
            d.text((c3, ty + i * 27), ln, font=F(20), fill=(30, 30, 30))
        ty += max(1, len(lineas)) * 27 + 14
        d.line([(tx, ty - 6), (W + PANEL1 - 20, ty - 6)], fill=(225, 225, 225), width=1)

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
        a, b = min(a, e["h0"]), max(b, e["h1"])
        grupos.setdefault((e["ax"], round(e["co"], 3), round(a, 2), round(b, 2)),
                          dict(ax=e["ax"], co=e["co"], a=a, b=b, Lv=res[n]["AM"], sec=pp["seccion"], cambia=True))
    resto = prop.get("_resto", "15×30")              # todas las demas dalas bajo muros de PA
    for n, e in piers.items():
        if n in prop or n not in res:
            continue
        ov = sum(max(0.0, min(e["h1"], p["h1"]) - max(e["h0"], p["h0"]))
                 for (s2, ax2, co2), ps in walls.items() if s2 == abajo and ax2 == e["ax"] and abs(co2 - e["co"]) < 0.02
                 for p in ps)
        Lv = res[n]["AM"]
        if ov / (e["h1"] - e["h0"]) > 0.5 or not isinstance(Lv, (int, float)) or Lv <= 0:
            a, b, Lv = e["h0"], e["h1"], e["h1"] - e["h0"]     # va sobre muro de PB: dala del largo del muro
        else:
            a, b = claro_de(e, Lv, e["ax"], e["co"], walls, abajo, cast, vigas)
            a, b = min(a, e["h0"]), max(b, e["h1"])
        grupos.setdefault((e["ax"], round(e["co"], 3), round(a, 2), round(b, 2)),
                          dict(ax=e["ax"], co=e["co"], a=a, b=b, Lv=Lv, sec=resto, cambia=False))
    juntos = []
    for g in sorted(grupos.values(), key=lambda g: (g["ax"], round(g["co"], 3), g["sec"], g["a"])):
        u = juntos[-1] if juntos else None
        if u and u["ax"] == g["ax"] and abs(u["co"] - g["co"]) < 0.02 and u["sec"] == g["sec"] and g["a"] <= u["b"] + 1.2:
            u["b"] = max(u["b"], g["b"]); u["Lv"] = max(u["Lv"], g["Lv"])
        else:
            juntos.append(dict(g))
    grupos = {i: g for i, g in enumerate(juntos)}
    muros(d, pa)
    puestas = []
    sin_flecha = []
    for g in grupos.values():                        # la viga: linea negra a lo largo de su claro, encima del muro
        if g["ax"] == "X":
            p0, p1 = (X(g["co"]), Y(g["a"])), (X(g["co"]), Y(g["b"]))
        else:
            p0, p1 = (X(g["a"]), Y(g["co"])), (X(g["b"]), Y(g["co"]))
        d.line([p0, p1], fill=(0, 0, 0), width=12)
        g["seg"] = (p0, p1)
        puestas.append((min(p0[0], p1[0]) - 9, min(p0[1], p1[1]) - 9, max(p0[0], p1[0]) + 9, max(p0[1], p1[1]) + 9))
    for g in sorted(grupos.values(), key=lambda g: (not g["cambia"], -g["Lv"])):
        if g["ax"] == "X":
            cx0, cy0 = X(g["co"]), (Y(g["a"]) + Y(g["b"])) / 2
        else:
            cx0, cy0 = (X(g["a"]) + X(g["b"])) / 2, Y(g["co"])
        hmin = math.ceil(g["Lv"] / 14 * 100 / 5 - 1e-9) * 5          # NTC-M 4.1.18: L/h <= 14
        texto = f"{g['sec']} peralte" + (f"\nmín. norma {hmin}" if g["cambia"] and hmin > 30 else "")
        lab, pos = poner(d, puestas, texto, (0, 0, 0), cx0, cy0, False,
                         F(24 if g["cambia"] else 18, True), borde=(0, 0, 0), vertical=g["ax"] == "X", gap=48,
                         seg=g["seg"])
        if not flecha(d, pos, lab.size, g["seg"]):
            sin_flecha.append((g["sec"], g["ax"], round(g["co"], 2), round(g["a"], 2), round(g["b"], 2), pos, lab.size))
        img2.paste(lab, pos, lab)

    img1.save(salida, save_all=True, append_images=[img2], resolution=150)
    print("listo:", salida, "| etiquetas en PA:", len(grupos), "| sin flecha:", sin_flecha or "ninguna")


if __name__ == "__main__":
    main(*sys.argv[1:5])
