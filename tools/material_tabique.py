#!/usr/bin/env python3
"""Cambia solo el material de los muros a tabique rojo recocido.

Cambios en el archivo (nada mas):
  - Agrega el material "Tabique rojo recocido f'm=20kg/cm2", copia del
    material que hoy usan los muros, con peso volumetrico 1300 kg/m3,
    f'm = 20 kg/cm2 y Em = 600 f'm = 12 000 kg/cm2. Lo demas (Poisson,
    coeficiente termico, curva) queda igual al material anterior.
  - La seccion de muro ("Muro 15 cm") apunta al material nuevo. El espesor
    no cambia.

Ademas reporta el peso propio que ETABS deberia dar en "Carga Muerta".

Uso: python3 material_tabique.py ENTRADA.e2k SALIDA.e2k [ORIGINAL.e2k]
"""
import difflib
import math
import re
import sys
from collections import defaultdict

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, NEWLINE  # noqa: E402

NEW_MAT = "Tabique rojo recocido f'm=20kg/cm2"
GRADE = "Mamposteria de Arcilla"
GAMMA_M = 1300.0          # kgf/m3
FM = 20.0e4               # kgf/m2  (20 kg/cm2)
EM = 600 * FM             # kgf/m2  (12 000 kg/cm2)


def shoelace(P):
    s = 0.0
    for i in range(len(P)):
        (x1, y1), (x2, y2) = P[i], P[(i + 1) % len(P)]
        s += x1 * y2 - x2 * y1
    return abs(s) / 2


def in_poly(p, P):
    x, y = p
    inside = False
    for i in range(len(P)):
        (x1, y1), (x2, y2) = P[i], P[(i + 1) % len(P)]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            inside = not inside
    return inside


def props(text):
    dens = {m[1]: float(m[2]) for m in re.finditer(
        r'MATERIAL\s+"([^"]+)"\s+TYPE "[^"]+"\s+GRADE "[^"]*"\s+WEIGHTPERVOLUME (\S+)', text)}
    shell = {m[1]: (m[2], float(m[3])) for m in re.finditer(
        r'SHELLPROP\s+"([^"]+)"\s+PROPTYPE\s+"\w+"\s+MATERIAL "([^"]+)"[^\r\n]*?(?:SLABTHICKNESS|WALLTHICKNESS) (\S+)', text)}
    frame = {m[1]: (m[2], float(m[3]), float(m[4])) for m in re.finditer(
        r'FRAMESECTION\s+"([^"]+)"\s+MATERIAL "([^"]+)"\s+SHAPE "Concrete Rectangular"\s+D (\S+) B (\S+)', text)}
    wmod = {m[1]: float(m[2]) for m in re.finditer(r'FRAMESECTION\s+"([^"]+)"[^\r\n]*?WMOD (\S+)', text)}
    return dens, shell, frame, wmod


def self_weight(text):
    """Peso propio (kgf) que ETABS calcula con SELFWEIGHT 1: muros + losas + barras."""
    M = Model(text)
    dens, shell, frame, wmod = props(text)
    W = defaultdict(float)
    walls, _ = M.wall_pieces()
    for pcs in walls.values():
        for p in pcs:
            mat, t = shell[p["section"]]
            W["muros"] += (p["h1"] - p["h0"]) * (p["z1"] - p["z0"]) * t * dens[mat]
    fl, op = [], []
    for idx, name, story, sec, rest in M.area_assign:
        P = [M.points[p][:2] for p, f in M.areas[name][1]]
        if sec is not None and not M.is_wall_section(sec):
            fl.append((story, sec, P))
        elif 'OPENING "Yes"' in rest:
            op.append((story, P))
    for story, sec, P in fl:
        mat, t = shell[sec]
        W["losas"] += shoelace(P) * t * dens[mat]
    for story, P in op:
        c = (sum(p[0] for p in P) / len(P), sum(p[1] for p in P) / len(P))
        for s2, sec, Q in fl:
            if s2 == story and in_poly(c, Q):
                mat, t = shell[sec]
                W["losas"] -= shoelace(P) * t * dens[mat]
                break
    for ln, p in M.columns.items():
        for story, sec in M.line_assign.get(ln, []):
            if sec in frame:
                mat, D, B = frame[sec]
                W["castillos y dalas"] += (M.elev[story] - M.elev[M.below(story)]) * D * B * dens[mat] * wmod.get(sec, 1.0)
    for ln, (p1, p2) in M.beams.items():
        a, b = M.points[p1], M.points[p2]          # (x, y, dz): la viga puede ir abajo del nivel
        for story, sec in M.line_assign.get(ln, []):
            if sec in frame:
                mat, D, B = frame[sec]
                W["castillos y dalas"] += math.dist(a, b) * D * B * dens[mat] * wmod.get(sec, 1.0)
    W["TOTAL"] = sum(W.values())
    return W


def main(src, dst, original=None):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    L = list(M.lines)
    dens, shell, frame, wmod = props(text)

    wall_props = {sec for idx, n, s, sec, r in M.area_assign if M.is_wall_section(sec)}
    old_mats = {shell[p][0] for p in wall_props}
    if len(old_mats) != 1:
        raise SystemExit(f"Se esperaba un solo material de muro, hay: {old_mats}")
    old = old_mats.pop()
    if NEW_MAT in dens:
        raise SystemExit("El material nuevo ya existe en el archivo")

    pat = re.compile(r'^\s*MATERIAL\s+"' + re.escape(old) + '"')
    blk = [i for i, ln in enumerate(L) if pat.match(ln)]
    new_mat_lines = []
    for i in blk:
        ln = L[i].replace(f'"{old}"', f'"{NEW_MAT}"', 1)
        if "WEIGHTPERVOLUME" in ln:
            ln = re.sub(r'GRADE "[^"]*"', f'GRADE "{GRADE}"', ln)
            ln = re.sub(r"WEIGHTPERVOLUME \S+", f"WEIGHTPERVOLUME {GAMMA_M:g}", ln)
        if "SYMTYPE" in ln:
            ln = re.sub(r"\bE \S+", f"E {EM:.0f}", ln)
        ln = re.sub(r"\bFM \S+", f"FM {FM:.0f}", ln)
        new_mat_lines.append(ln)

    out = []
    for i, ln in enumerate(L):
        m = re.match(r'\s*SHELLPROP\s+"([^"]+)"', ln)
        if m and m[1] in wall_props:
            ln = ln.replace(f'MATERIAL "{old}"', f'MATERIAL "{NEW_MAT}"')
        out.append(ln)
        if i == blk[-1]:
            out.extend(new_mat_lines)
    new_text = NEWLINE.join(out)
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(new_text)

    # Verificacion: solo cambia la linea de la seccion de muro y se agregan las del material
    d = [x for x in difflib.unified_diff(L, out, lineterm="", n=0) if x[:1] in "+-" and x[:3] not in ("+++", "---")]
    removed = [x for x in d if x.startswith("-")]
    added = [x for x in d if x.startswith("+")]
    assert len(removed) == len(wall_props) and all("SHELLPROP" in x for x in removed), removed
    assert len(added) == len(new_mat_lines) + len(wall_props), added
    print(f"cambios: {len(removed)} linea modificada, {len(added) - len(removed)} lineas nuevas")
    for x in removed + added:
        print("  ", x[:170])

    rows = []
    if original:
        rows.append(("original R3.2", self_weight(open(original, encoding="utf-8", newline="").read())))
    rows.append(("muros redibujados", self_weight(text)))
    rows.append(("muros redibujados + tabique", self_weight(new_text)))
    keys = ["muros", "losas", "castillos y dalas", "TOTAL"]
    print("\nPeso propio que ETABS deberia dar en 'Carga Muerta' (t):")
    print(f"{'':30s}" + "".join(f"{k:>19s}" for k in keys))
    for name, W in rows:
        print(f"{name:30s}" + "".join(f"{W.get(k, 0) / 1000:19.1f}" for k in keys))


if __name__ == "__main__":
    main(*sys.argv[1:4])
