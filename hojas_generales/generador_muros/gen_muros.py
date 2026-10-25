"""Arma Muros_Mamposteria_NTC2023.xlsx (revisión de muros de mampostería confinada, NTC 2023 CDMX) con el ejemplo de Casa Rocha.

Uso:
    python3 gen_muros.py                 # arma, recalcula con LibreOffice y deja el libro con valores en caché
    python3 gen_muros.py --sin-recalculo # sólo arma (Excel calcula al abrir)

Ver README.md.
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile

from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import muros_comun as C          # noqa: E402
import muros_tabla as T          # noqa: E402
import muros_detalle as D        # noqa: E402
import muros_otros as O          # noqa: E402

OUT = os.path.join(os.path.dirname(HERE), 'Muros_Mamposteria_NTC2023.xlsx')
EJ = os.path.join(HERE, 'ejemplo_casa_rocha.json')

# Story drifts de la hoja Revision_muros_NTC2023.xlsx (ETABS R3.2): Story, Output Case, Case Type, Step Type, Direction, Drift, Label, X, Y, Z
DRIFTS = [
    ('N+6.00', 'SXLD', 'LinRespSpec', 'Max', 'X', 0.000018, 223, 4.33, 1.92, None), ('N+6.00', 'SYLD', 'LinRespSpec', 'Max', 'X', 0.000001, 238, 0, 9.95, None),
    ('N+6.00', 'SYLD', 'LinRespSpec', 'Max', 'Y', 0.000005, 77, 6.42, 6.95, None), ('N+6.00', 'SXSV', 'LinRespSpec', 'Max', 'X', 0.000042, 223, 4.33, 1.92, None),
    ('N+6.00', 'SYSV', 'LinRespSpec', 'Max', 'X', 0.000003, 238, 0, 9.95, None), ('N+6.00', 'SYSV', 'LinRespSpec', 'Max', 'Y', 0.000011, 77, 6.42, 6.95, None),
    ('N+3.00', 'SXLD', 'LinRespSpec', 'Max', 'X', 0.000101, 92, 8.85, 0, None), ('N+3.00', 'SYLD', 'LinRespSpec', 'Max', 'X', 0.000001, 9, 0, 0, None),
    ('N+3.00', 'SYLD', 'LinRespSpec', 'Max', 'Y', 0.000005, 300, 8.43, 6.95, None), ('N+3.00', 'SXSV', 'LinRespSpec', 'Max', 'X', 0.000241, 92, 8.85, 0, None),
    ('N+3.00', 'SYSV', 'LinRespSpec', 'Max', 'X', 0.000004, 9, 0, 0, None), ('N+3.00', 'SYSV', 'LinRespSpec', 'Max', 'Y', 0.000013, 300, 8.43, 6.95, None)]
STORIES = [dict(st='N+3.00', nom='PB', W=253, Vx=22.1, Vy=17.3), dict(st='N+6.00', nom='PA', W=139, Vx=12.1, Vy=10.0)]


def ejemplo():
    out = []
    for w in json.load(open(EJ, encoding='utf-8')):
        e = dict(key=w['key'], st=w['story'], dir=w['dir'], pos='EXTREMO', L=w['L'], rht='NINGUNO',
                 cVu=w['Vu'], cMu=w['Mu'], cP=w['P'], cPmx=w['Pu'], cPg=w['Pu'], cVg=w['Vg'])
        for j, o in enumerate(w['openings'][:C.NV], start=1):
            e.update({'v%dt' % j: o['tipo'], 'v%dx' % j: o['x'], 'v%dw' % j: o['w'], 'v%da' % j: o['ant'], 'v%dh' % j: o['alto']})
        nota = []
        if w['openings']:
            nota.append('Vanos del plano arquitectónico; alturas típicas por confirmar.')
        if w.get('user_tramos'):
            nota.append('Tramos de tu hoja anterior: %s m.' % w['user_tramos'])
        b = w.get('beam')
        if b:
            e.update(sv='SÍ', vnom=b['name'], bv=b['b'], hv=b['h'], Lv=b['Lv'], wv=b['wu'], cv=b['c'])
            nota.append('Viga %s %s×%s propuesta en PRINCIPALES (el modelo R3.2 tiene dala 15×20).' % (b['name'], b['b'], b['h']))
        else:
            e['sv'] = 'NO'
        e['nota'] = ' '.join(nota)
        out.append(e)
    return out


def build(path, sel=None, only=None):
    if sel:
        D.DEFAULT_SEL = sel
    C.N.clear(); C.TB.clear()
    wb = Workbook()
    C.sheet_instrucciones(wb)
    C.sheet_tablas(wb)
    C.sheet_datos(wb)
    T.define_pf_names(wb)
    T.sheet_muros(wb, ejemplo())
    T.sheet_pier_forces(wb)
    D.sheet_detalle(wb)
    croquis = None
    try:
        import muros_croquis as K
        croquis = K
    except ImportError:
        pass
    if croquis:
        wsK = croquis.sheet_croquis(wb)
    O.sheet_entrepiso(wb, STORIES)
    O.sheet_distorsiones(wb, DRIFTS)
    O.sheet_resumen(wb)
    C.sheet_referencias(wb)
    if croquis:
        croquis.build_croquis_datos(wb, wsK)
    order = ['INSTRUCCIONES', 'DATOS', 'MUROS', 'DETALLE', 'CROQUIS', 'RESUMEN', 'ENTREPISO', 'DISTORSIONES', 'PIER_FORCES', 'REFERENCIAS', 'TABLAS', 'CROQUIS_DATOS']
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [s for s in wb._sheets if s.title not in order]
    wb.active = 0
    if only:
        for w in wb.worksheets:
            if w.title != only:
                w.sheet_state = 'hidden'
        wb.active = wb.sheetnames.index(only)
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
    res = subprocess.run([sys.executable, rc[0], calc, '600'], capture_output=True, text=True)
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
