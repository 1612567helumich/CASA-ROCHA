#!/usr/bin/env python3
"""Quita del .e2k los tableros de muro que caen completos dentro de un tramo.

Cada tramo se escribe  PISO:EJE=COORD:H0-H1   (H = coordenada a lo largo del eje)
  ej.  "N+3.00:X=0:0-1.919"   muros del entrepiso N+3.00 sobre X=0, de Y=0 a Y=1.919

Solo borra las lineas AREA/AREAASSIGN de esos tableros y los puntos que
queden sin usar. No toca castillos, dalas, losas, cargas ni nada mas.

Uso: python3 quitar_muros.py ENTRADA.e2k SALIDA.e2k TRAMO [TRAMO ...]
"""
import difflib
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL, NEWLINE  # noqa: E402


def parse_tramo(s):
    story, eje, rng = s.split(":")
    axis, coord = eje.split("=")
    h0, h1 = (float(v) for v in rng.split("-"))
    return story, axis, float(coord), h0, h1


def main(src, dst, tramos):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    walls, _ = M.wall_pieces()
    quitar = []
    for t in map(parse_tramo, tramos):
        story, axis, coord, h0, h1 = t
        found = [p for (s, ax, co), pcs in walls.items() if s == story and ax == axis and abs(co - coord) <= TOL
                 for p in pcs if p["h0"] >= h0 - TOL and p["h1"] <= h1 + TOL]
        if not found:
            raise SystemExit(f"No hay tableros en el tramo {t}")
        quitar += [(p["name"], story, p) for p in found]

    names = {n for n, s, p in quitar}
    drop_assign = {idx for idx, n, s, sec, r in M.area_assign if (n, s) in {(n2, s2) for n2, s2, _ in quitar}}
    still_assigned = {n for idx, n, s, sec, r in M.area_assign if idx not in drop_assign}
    drop_area = {M.areas[n][2] for n in names if n not in still_assigned}
    L = M.lines
    out = [ln for i, ln in enumerate(L) if i not in drop_assign and i not in drop_area]
    new_text = NEWLINE.join(out)

    # puntos que usaban esos tableros y ya no usa nadie
    cand = {p for n in names for p, f in M.areas[n][1]}
    M2 = Model(new_text)
    used = {p for _, conn, _ in M2.areas.values() for p, _ in conn}
    used |= set(M2.columns.values())
    used |= {p for pair in M2.beams.values() for p in pair}
    used |= set(re.findall(r'POINTLOAD\s+"([^"]+)"', new_text))
    orphans = cand - used
    if orphans:
        new_text = NEWLINE.join(ln for ln in new_text.split(NEWLINE)
                                if not (re.match(r'\s*(?:POINT|POINTASSIGN)\s+"([^"]+)"', ln)
                                        and re.match(r'\s*(?:POINT|POINTASSIGN)\s+"([^"]+)"', ln)[1] in orphans))
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(new_text)

    # verificacion: solo se borraron lineas de esos tableros y de puntos huerfanos
    d = [x for x in difflib.unified_diff(L, new_text.split(NEWLINE), lineterm="", n=0)
         if x[:1] in "+-" and x[:3] not in ("+++", "---")]
    assert not [x for x in d if x.startswith("+")], "se agrego algo"
    for x in d:
        m = re.match(r'-\s*(AREA|AREAASSIGN|POINT|POINTASSIGN)\s+"([^"]+)"', x)
        assert m and (m[2] in names or m[2] in orphans), x

    area = sum((p["h1"] - p["h0"]) * (p["z1"] - p["z0"]) for n, s, p in quitar)
    print(f"tableros quitados: {len(quitar)}  ({area:.2f} m2 de muro)")
    for n, s, p in quitar:
        print(f"   {n:6s} {s} {p['axis']}={p['coord']:g}  de {p['h0']:.3f} a {p['h1']:.3f}  z {p['z0']:.2f}-{p['z1']:.2f}  pier {p['pier']}")
    print(f"puntos que quedaron sin usar y se quitaron: {sorted(orphans)}")
    print(f"lineas borradas: {len(d)}, lineas agregadas: 0")
    return area


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
