#!/usr/bin/env python3
"""Quita barras (castillos o vigas) por nombre: sus lineas LINE, LINEASSIGN y
LINELOAD, y los puntos que queden sin usar (con sus POINTASSIGN).
No toca nada mas.
Uso: python3 quitar_barras.py ENTRADA.e2k SALIDA.e2k NOMBRE [NOMBRE ...]"""
import difflib
import re
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, NEWLINE  # noqa: E402


def main(src, dst, names):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    names = set(names)
    missing = names - set(M.columns) - set(M.beams)
    if missing:
        raise SystemExit(f"No existen: {missing}")
    cand = set()
    for n in names:
        if n in M.columns:
            cand.add(M.columns[n])
        else:
            cand.update(M.beams[n])
    L = M.lines
    pat = re.compile(r'\s*(?:LINE|LINEASSIGN|LINELOAD)\s+"([^"]+)"')
    out = [ln for ln in L if not (pat.match(ln) and pat.match(ln)[1] in names)]
    t2 = NEWLINE.join(out)
    M2 = Model(t2)
    used = {p for _, conn, _ in M2.areas.values() for p, _ in conn}
    used |= set(M2.columns.values()) | {p for pair in M2.beams.values() for p in pair}
    used |= set(re.findall(r'POINTLOAD\s+"([^"]+)"', t2))
    orphans = cand - used
    pp = re.compile(r'\s*(?:POINT|POINTASSIGN)\s+"([^"]+)"')
    out = [ln for ln in out if not (pp.match(ln) and pp.match(ln)[1] in orphans)]
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(NEWLINE.join(out))
    d = [x for x in difflib.unified_diff(L, out, lineterm="", n=0) if x[:1] in "+-" and x[:3] not in ("+++", "---")]
    assert not [x for x in d if x.startswith("+")]
    for x in d:
        m = re.match(r'-\s*(?:LINE|LINEASSIGN|LINELOAD|POINT|POINTASSIGN)\s+"([^"]+)"', x)
        assert m and (m[1] in names or m[1] in orphans), x
    print(f"barras quitadas: {sorted(names)}")
    print(f"puntos sin usar quitados (con sus asignaciones): {sorted(orphans)}")
    print(f"lineas borradas: {len(d)}")
    for x in d:
        print("  ", x[:150])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
