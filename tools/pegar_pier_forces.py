#!/usr/bin/env python3
"""Pega en la hoja de muros solo las combinaciones de diseno de una tabla
Pier Forces de ETABS (Comb4.1-4.8 y Comb2.2), borra lo demas de PIER_FORCES,
recalcula con LibreOffice y guarda los resultados. Tambien escribe la tabla
filtrada como archivo aparte.

Uso: python3 pegar_pier_forces.py HOJA.xlsx PIER_FORCES.xlsx SALIDA.xlsx TABLA_FILTRADA.xlsx
"""
import os
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from optimizar_hoja_muros import (COMBOS, Package, actualizar_caches, fase_croquis, leer_valores, q,  # noqa: E402
                                  recalcular, set_value)


def leer_tabla(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["Pier Forces"].iter_rows(values_only=True))
    hdr = list(rows[1])
    if "Step Number" in hdr:
        k = hdr.index("Step Number")
        rows = [r[:k] + r[k + 1:] for r in rows]
        hdr = list(rows[1])
    assert hdr[:12] == ["Story", "Pier", "Output Case", "Case Type", "Step Type", "Location", "P", "V2", "V3", "T", "M2", "M3"], hdr
    pc = list(wb["Program Control"].iter_rows(values_only=True)) if "Program Control" in wb.sheetnames else []
    return rows[:3], [list(r[:12]) for r in rows[3:] if r[2] in COMBOS], pc


def main(hoja, tabla, salida, filtrada):
    head, data, pc = leer_tabla(tabla)
    pkg = Package(hoja)
    S = pkg.sheets["PIER_FORCES"]
    last = max(S.rows)
    tmpl = {col: S.cell(f"{col}9", create=False) for col in "BCDEFGHIJKLM"}
    for r in range(9, last + 1):
        for col in "BCDEFGHIJKLM":
            c = S.cell(f"{col}{r}", create=False)
            if c is not None and c.find(q("f")) is None:
                set_value(c, None)
    for i, row in enumerate(data):
        for j, col in enumerate("BCDEFGHIJKLM"):
            set_value(S.cell(f"{col}{9 + i}", style_from=tmpl[col]), row[j])
    work = tempfile.mkdtemp(prefix="pegar_pf_")
    paso = os.path.join(work, "paso.xlsx")
    pkg.save(paso)
    vals = leer_valores(recalcular(paso, work))
    log = {}
    actualizar_caches(pkg, vals, log)
    fase_croquis(pkg, vals, log)
    pkg.save(salida)
    # tabla filtrada aparte
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pier Forces"
    for r in head + data:
        ws.append(list(r))
    if pc:
        w2 = wb.create_sheet("Program Control", 0)
        for r in pc:
            w2.append(list(r))
    wb.save(filtrada)
    print(f"renglones pegados: {len(data)} (de la tabla completa solo Comb4.1-4.8 y Comb2.2)")
    return vals


if __name__ == "__main__":
    main(*sys.argv[1:5])
