#!/usr/bin/env python3
"""Llena la tabla de datos de ENTREPISO (W, V_u X, V_u Y por story) con la
tabla Story Forces de ETABS, recalcula con LibreOffice y guarda resultados.

  W     = peso total arriba del entrepiso, sin factor (NTC-M 4.1.12.1):
          P en Bottom de los casos de la fuente de masa (CM, SC, CV inst, ...)
  V_u   = cortante de diseno factorizado (NTC-M 4.1.11.1):
          max |VX| o |VY| en Bottom de Comb4.1-4.8

Uso: python3 llenar_entrepiso.py HOJA.xlsx STORY_FORCES.xlsx SALIDA.xlsx
"""
import os
import sys
import tempfile

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from optimizar_hoja_muros import Package, actualizar_caches, fase_croquis, leer_valores, recalcular, set_value  # noqa: E402

PESO = ["Carga Muerta", "Sobrecarga", "Carga Viva Instantanea", "EquiposAzotea", "Pretil1.1", "RsmpasEscalera",
        "DescansosEscalera", "MurosBlockDivisorio", "Barandal", "CANCELERIA"]
SISMO = [f"Comb4.{i}" for i in range(1, 9)]


def leer(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb["Story Forces"].iter_rows(values_only=True))
    h = {k: i for i, k in enumerate(rows[1])}
    out = {}
    for r in rows[3:]:
        if r[h["Location"]] != "Bottom":
            continue
        e = out.setdefault(r[h["Story"]], dict(W=0.0, VX=0.0, VY=0.0, casos=set()))
        caso = r[h["Output Case"]]
        if caso in PESO:
            e["W"] += r[h["P"]]
            e["casos"].add(caso)
        if caso in SISMO:
            e["VX"] = max(e["VX"], abs(r[h["VX"]]))
            e["VY"] = max(e["VY"], abs(r[h["VY"]]))
            e["casos"].add(caso)
    for s, e in out.items():
        falta = {"Carga Muerta", "Sobrecarga", "Carga Viva Instantanea", *SISMO} - e["casos"]
        assert not falta, f"{s}: faltan {falta}"
    return out


def main(hoja, tabla, salida):
    datos = leer(tabla)
    pkg = Package(hoja)
    S = pkg.sheets["ENTREPISO"]
    puestos = []
    for r in range(6, 10):
        c = S.cell(f"B{r}", create=False)
        story = pkg.text(c) if c is not None else None
        if story in datos:
            e = datos[story]
            for col, v in (("D", e["W"]), ("E", e["VX"]), ("F", e["VY"])):
                set_value(S.cell(f"{col}{r}", create=False), round(v, 2))
            puestos.append((story, round(e["W"], 2), round(e["VX"], 2), round(e["VY"], 2)))
    assert len(puestos) == len(datos), (puestos, list(datos))
    work = tempfile.mkdtemp(prefix="entrepiso_")
    paso = os.path.join(work, "paso.xlsx")
    pkg.save(paso)
    vals = leer_valores(recalcular(paso, work))
    log = {}
    actualizar_caches(pkg, vals, log)
    fase_croquis(pkg, vals, log)
    pkg.save(salida)
    for p in puestos:
        print("story %s: W = %.2f t, V_u X = %.2f t, V_u Y = %.2f t" % p)
    return vals


if __name__ == "__main__":
    main(*sys.argv[1:4])
