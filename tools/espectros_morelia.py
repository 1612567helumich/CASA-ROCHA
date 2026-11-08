#!/usr/bin/env python3
"""Mete los espectros de PRODISIS (CFE 2015, Morelia) al .e2k y los asigna:
  - SXSV, SYSV, SXcimentacion, SYcimentacion -> espectro de diseno, SF = 9.81
  - SXLD, SYLD                             -> espectro transparente, SF = 9.81 / 5.5 (Fser, MDOC 3.1.8)
    (servicio para edificios del Grupo B)
No toca nada mas: la funcion anterior se queda definida pero sin usar.
Uso: python3 espectros_morelia.py ENTRADA.e2k SALIDA.e2k DISENO.txt TRANSPARENTE.txt
"""
import difflib
import re
import sys

NEWLINE = "\r\n"
G = 9.81
F_DIS = "Morelia CFE Diseño (Q2 R2 rho1 a0.8)"
F_TRA = "Morelia CFE Transparente"


def leer(path):
    pts = []
    for ln in open(path, encoding="utf-8"):
        p = ln.split()
        if len(p) == 2:
            pts.append((float(p[0]), float(p[1])))
    assert pts and pts == sorted(pts), path
    return pts


def bloque(nombre, pts):
    out = [f'  FUNCTION "{nombre}"  FUNCTYPE "SPECTRUM"  DAMPRATIO 0.05  SPECTYPE "USER"  ']
    for i in range(0, len(pts), 8):
        par = "  ".join(f"{t:g}  {v:g}" for t, v in pts[i:i + 8])
        out.append(f'  FUNCTION "{nombre}"  TIMEVAL "{par}"  ')
    return out


def main(src, dst, dis, tra):
    L = open(src, encoding="utf-8", newline="").read().split(NEWLINE)
    pd, pt = leer(dis), leer(tra)
    ifun = next(i for i, ln in enumerate(L) if ln.startswith("$ FUNCTIONS"))
    inext = next(i for i in range(ifun + 1, len(L)) if L[i].startswith("$ "))
    last = max(i for i in range(ifun, inext) if L[i].strip())
    out = L[: last + 1] + bloque(F_DIS, pd) + bloque(F_TRA, pt) + L[last + 1:]
    n = 0
    for i, ln in enumerate(out):
        m = re.match(r'(\s*LOADCASE\s+"([^"]+)"\s+ACCEL\s+"U[12]"\s+FUNC\s+)"[^"]+"(\s+SF\s+)\S+(.*)$', ln)
        if not m:
            continue
        caso = m[2]
        if caso.startswith(("SXSV", "SYSV", "SXcimentacion", "SYcimentacion")):
            out[i] = f'{m[1]}"{F_DIS}"{m[3]}{G:g}{m[4]}'
        elif caso.startswith(("SXLD", "SYLD")):
            out[i] = f'{m[1]}"{F_TRA}"{m[3]}{round(G / 5.5, 4):g}{m[4]}'
        else:
            continue
        n += 1
    assert n == 6, n
    open(dst, "w", encoding="utf-8", newline="").write(NEWLINE.join(out))
    d = [x for x in difflib.unified_diff(L, out, lineterm="", n=0) if x[:1] in "+-" and x[:3] not in ("+++", "---")]
    quitadas = [x for x in d if x.startswith("-")]
    assert len(quitadas) == 6 and all("LOADCASE" in x for x in quitadas)
    print(f"funciones nuevas: {len(pd)} y {len(pt)} puntos (T de {pd[0][0]} a {pd[-1][0]} s)")
    print(f"casos cambiados: {n}")
    for x in d:
        if "LOADCASE" in x and x.startswith("+"):
            print("  ", re.sub(r"\s+", " ", x)[:170])


if __name__ == "__main__":
    main(*sys.argv[1:5])
