#!/usr/bin/env python3
"""Quita la ultima columna de MUROS (p. ej. NOTA) sin tocar nada mas:
celdas, ancho, titulos combinados, dimension y area de impresion.
Uso: python3 quitar_columna_muros.py ENTRADA.xlsx SALIDA.xlsx COLUMNA"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from optimizar_hoja_muros import Package, col2num, num2col, q, split_ref  # noqa: E402


def main(src, dst, col):
    pkg = Package(src)
    S = pkg.sheets["MUROS"]
    k = col2num(col)
    prev = num2col(k - 1)
    ultima = max(col2num(split_ref(c.get("r"))[0]) for _, c in S.cells())
    if k != ultima:
        raise SystemExit(f"{col} no es la ultima columna de MUROS")
    n = 0
    for _, c in list(S.cells()):
        if split_ref(c.get("r"))[0] == col:
            if c.find(q("f")) is not None:
                raise SystemExit(f"{c.get('r')} tiene formula")
            c.getparent().remove(c)
            n += 1
    for m in S.root.iter(q("mergeCell")):
        a, b = m.get("ref").split(":")
        if split_ref(b)[0] == col:
            m.set("ref", f"{a}:{prev}{split_ref(b)[1]}")
    cols = S.root.find(q("cols"))
    for c in list(cols):
        a, b = int(c.get("min")), int(c.get("max"))
        if a <= k <= b:
            if a == b:
                cols.remove(c)
            else:
                c.set("max", str(k - 1)) if b == k else None
    dim = S.root.find(q("dimension"))
    dim.set("ref", re.sub(rf"\b{col}(\d+)$", rf"{prev}\1", dim.get("ref")))
    wbx = pkg.data["xl/workbook.xml"].decode("utf-8")
    wbx = wbx.replace(f"MUROS!$A$1:${col}$", f"MUROS!$A$1:${prev}$")
    pkg.data["xl/workbook.xml"] = wbx.encode("utf-8")
    pkg.save(dst)
    print(f"columna {col} quitada: {n} celdas")


if __name__ == "__main__":
    main(*sys.argv[1:4])
