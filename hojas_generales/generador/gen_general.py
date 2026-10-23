"""Arma Losas_General_NTC2023.xlsx (hoja general de losas nervadas y macizas, NTC 2023 CDMX).

Uso:
    python3 gen_general.py                 # arma, recalcula con LibreOffice y deja el libro con valores en caché
    python3 gen_general.py --sin-recalculo # sólo arma (fórmulas sin valores)

Ver README.md.
"""
import os
import sys
import glob
import shutil
import subprocess
import tempfile

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import sheet_instrucciones, sheet_datos, sheet_tablas   # noqa: E402
from losas import sheet_nervadas, sheet_macizas                     # noqa: E402
import detalle                                                      # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), 'Losas_General_NTC2023.xlsx')


def build(path):
    wb = Workbook()
    sheet_instrucciones(wb)
    sheet_datos(wb)
    sheet_nervadas(wb)
    sheet_macizas(wb)
    wsD = detalle.sheet_detalle(wb)
    wsK = detalle.sheet_croquis(wb)
    sheet_tablas(wb)
    detalle.build_croquis_datos(wb, wsK)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return path


def recalc_and_inject(path):
    rc = sorted(glob.glob('/root/.claude/skills/synced/*/xlsx/scripts/recalc.py'))
    if not rc:
        print('No encontré recalc.py: el libro queda sin valores en caché (Excel los calcula al abrir).')
        return None
    tmp = tempfile.mkdtemp()
    calc = os.path.join(tmp, 'rc.xlsx')
    shutil.copy(path, calc)
    res = subprocess.run([sys.executable, rc[0], calc, '300'], capture_output=True, text=True)
    print(res.stdout[-3000:])
    if res.returncode != 0:
        print(res.stderr[-3000:])
    final = os.path.join(tmp, 'final.xlsx')
    subprocess.run([sys.executable, os.path.join(HERE, 'inject_values.py'), path, calc, final], check=True)
    shutil.copy(final, path)
    return res.stdout


if __name__ == '__main__':
    build(OUT)
    print('armado:', OUT)
    if '--sin-recalculo' not in sys.argv:
        recalc_and_inject(OUT)
        print('listo:', OUT)
