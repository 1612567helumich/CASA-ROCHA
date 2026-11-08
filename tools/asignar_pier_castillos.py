#!/usr/bin/env python3
"""Asigna a cada castillo el pier label del muro que confina, para que las
fuerzas de pier de ETABS sumen tabique + castillos.

Regla (por castillo y por piso):
  - confina un solo pier (orilla de hueco, extremo o intermedio)   -> ese pier
  - esquina con un muro que no tiene pier                           -> el pier del otro muro
  - compartido por dos muros con pier (esquina o cruce)             -> sin pier (los dos muros
                                                                       quedan del lado seguro en cortante)
  - solo en muretes, antepechos o dinteles (sin pier)                -> sin pier
Solo cambia lineas LINEASSIGN de castillos (agrega PIER "..."). Nada mas.

Uso: python3 asignar_pier_castillos.py ENTRADA.e2k SALIDA.e2k LISTA.csv
"""
import csv
import difflib
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL, NEWLINE  # noqa: E402


def clasificar(M):
    walls, _ = M.wall_pieces()
    out = []
    for ln, p in M.columns.items():
        x, y, _ = M.points[p]
        for story, sec in M.line_assign.get(ln, []):
            piers, lines, wall_lines = set(), set(), set()
            for (s, ax, co), pcs in walls.items():
                if s != story:
                    continue
                off, h = (x, y) if ax == "X" else (y, x)
                if abs(off - co) > TOL:
                    continue
                for q in pcs:
                    if q["h0"] - TOL <= h <= q["h1"] + TOL:
                        wall_lines.add(ax)
                        if q["pier"]:
                            piers.add(q["pier"])
                            lines.add(ax)
            if not wall_lines:
                caso, pier = "sin muro", None
            elif not piers:
                caso, pier = "solo en murete, antepecho o dintel (sin pier)", None
            elif len(lines) == 2:
                caso, pier = "compartido por dos muros con pier: " + " / ".join(sorted(piers)), None
            elif len(piers) > 1:
                caso, pier = "entre dos piers del mismo eje: " + " / ".join(sorted(piers)), None
            elif len(wall_lines) == 2:
                caso, pier = "esquina con muro sin pier", next(iter(piers))
            else:
                caso, pier = "confina un solo pier", next(iter(piers))
            out.append(dict(castillo=ln, piso=story, x=round(x, 3), y=round(y, 3), pier=pier, caso=caso))
    return out


def main(src, dst, lista):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    known = set(re.findall(r'PIERNAME\s+"([^"]+)"', text))
    asign = clasificar(M)
    want = {(a["castillo"], a["piso"]): a["pier"] for a in asign if a["pier"]}
    assert all(p in known for p in want.values())
    L = M.lines
    out = []
    n = 0
    for ln in L:
        m = re.match(r'(\s*LINEASSIGN\s+"([^"]+)"\s+"([^"]+)"\s+SECTION "[^"]*")(.*)$', ln)
        if m and (m[2], m[3]) in want:
            if "PIER" in m[4]:
                raise SystemExit(f"{m[2]} {m[3]} ya tiene pier")
            ln = f'{m[1]}  PIER "{want[(m[2], m[3])]}"{m[4]}'
            n += 1
        out.append(ln)
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(NEWLINE.join(out))
    d = [x for x in difflib.unified_diff(L, out, lineterm="", n=0) if x[:1] in "+-" and x[:3] not in ("+++", "---")]
    assert all(re.match(r'[+-]\s*LINEASSIGN\s+"C', x) for x in d) and len(d) == 2 * n == 2 * len(want)
    asign.sort(key=lambda a: (a["piso"], a["pier"] or "~", a["castillo"]))
    with open(lista, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["castillo", "piso", "x", "y", "pier", "caso"])
        w.writeheader()
        for a in asign:
            w.writerow({**a, "pier": a["pier"] or "(sin pier)"})
    from collections import Counter
    print(f"castillos con pier asignado: {n} de {len(asign)}")
    for k, v in Counter(a["caso"].split(":")[0] for a in asign).most_common():
        print(f"  {v:4d}  {k}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
