#!/usr/bin/env python3
"""Pega en DISTORSIONES (desde el renglon 16) la tabla Story Drifts de ETABS,
solo los casos de espectro de seguridad de vida y limitacion de dano
(SXSV, SYSV, SXLD, SYLD), recalcula con LibreOffice y guarda resultados.

Columnas de la hoja: B Story, C Output Case, D Case Type, E Step Type,
F Direction, G Drift, H Label, I X, J Y, K Z.

Uso: python3 pegar_story_drifts.py HOJA.xlsx STORY_DRIFTS.xlsx SALIDA.xlsx
"""
import os
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from optimizar_hoja_muros import Package, actualizar_caches, fase_croquis, leer_valores, q, recalcular, set_value  # noqa: E402

CASOS = ("SXSV", "SYSV", "SXLD", "SYLD")
COLS = "BCDEFGHIJK"
PRIMERO, ULTIMO = 16, 315


def leer(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["Story Drifts"].iter_rows(values_only=True))
    h = {k: i for i, k in enumerate(rows[1])}
    out = []
    for r in rows[3:]:
        caso = r[h["Output Case"]] or ""
        if not caso.startswith(CASOS):
            continue
        out.append([r[h[k]] for k in ("Story", "Output Case", "Case Type", "Step Type", "Direction", "Drift",
                                      "Label", "X", "Y", "Z")])
    vistos = {d[1][:4] for d in out}
    assert vistos == set(CASOS), f"faltan casos: {set(CASOS) - vistos}"
    assert len(out) <= ULTIMO - PRIMERO + 1, len(out)
    return out


def main(hoja, tabla, salida):
    data = leer(tabla)
    pkg = Package(hoja)
    S = pkg.sheets["DISTORSIONES"]
    tmpl = {col: S.cell(f"{col}{PRIMERO}", create=False) for col in COLS}
    for r in range(PRIMERO, ULTIMO + 1):
        for col in COLS:
            c = S.cell(f"{col}{r}", create=False)
            if c is not None and c.find(q("f")) is None:
                set_value(c, None)
    for i, row in enumerate(data):
        for j, col in enumerate(COLS):
            set_value(S.cell(f"{col}{PRIMERO + i}", style_from=tmpl[col]), row[j])
    work = tempfile.mkdtemp(prefix="drifts_")
    paso = os.path.join(work, "paso.xlsx")
    pkg.save(paso)
    vals = leer_valores(recalcular(paso, work))
    log = {}
    actualizar_caches(pkg, vals, log)
    fase_croquis(pkg, vals, log)
    pkg.save(salida)
    print(f"renglones pegados: {len(data)} ({', '.join(CASOS)})")
    return vals


if __name__ == "__main__":
    main(*sys.argv[1:4])
