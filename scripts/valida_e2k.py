"""Validación interna del .e2k (referencias cruzadas) y cuantificación de pesos."""
import re
import sys
from collections import defaultdict

Q = re.compile(r'"([^"]*)"')


def main(path):
    lines = open(path, encoding='utf-8').read().splitlines()
    sec = None
    S = defaultdict(list)
    for ln in lines:
        if ln.startswith('$ '):
            sec = ln[2:].strip()
        elif ln.strip():
            S[sec].append(ln)
    err = []
    stories = [Q.findall(l)[0] for l in S['STORIES - IN SEQUENCE FROM TOP']]
    mats = {Q.findall(l)[0] for l in S['MATERIAL PROPERTIES']}
    frames = {Q.findall(l)[0] for l in S['FRAME SECTIONS']}
    shells = {Q.findall(l)[0] for k in ('SLAB PROPERTIES', 'WALL PROPERTIES', 'DECK PROPERTIES') for l in S[k]}
    for k in ('SLAB PROPERTIES', 'WALL PROPERTIES'):
        for l in S[k]:
            m = re.search(r'MATERIAL "([^"]*)"', l)
            if m and m.group(1) not in mats:
                err.append('material inexistente ' + m.group(1))
    for l in S['FRAME SECTIONS']:
        m = re.search(r'MATERIAL "([^"]*)"', l)
        if m and m.group(1) not in mats:
            err.append('material inexistente ' + m.group(1))
    pts = {Q.findall(l)[0]: tuple(map(float, l.split('"')[2].split()[:2])) for l in S['POINT COORDINATES']}
    piers = {Q.findall(l)[0] for l in S['PIER/SPANDREL NAMES'] if 'PIERNAME' in l}
    diaphs = {Q.findall(l)[0] for l in S['DIAPHRAGM NAMES']}
    lines_ = {}
    for l in S['LINE CONNECTIVITIES']:
        q = Q.findall(l)
        lines_[q[0]] = q[1:]
        for p in q[1:]:
            if p not in pts:
                err.append('punto inexistente %s en %s' % (p, q[0]))
    areas = {}
    for l in S['AREA CONNECTIVITIES']:
        q = Q.findall(l)
        typ = l.split()[2]
        areas[q[0]] = (typ, q[1:])
        for p in q[1:]:
            if p not in pts:
                err.append('punto inexistente %s en %s' % (p, q[0]))
    assigned = set()
    for l in S['AREA ASSIGNS']:
        q = Q.findall(l)
        name, st, sect = q[0], q[1], q[2]
        assigned.add((name, st))
        if name not in areas:
            err.append('área no definida ' + name)
        if st not in stories:
            err.append('nivel inexistente ' + st)
        if sect not in shells:
            err.append('sección de área inexistente ' + sect)
        m = re.search(r'PIER "([^"]*)"', l)
        if m and m.group(1) not in piers:
            err.append('pier no declarado ' + m.group(1))
        m = re.search(r'DIAPH  "([^"]*)"', l)
        if m and m.group(1) not in diaphs:
            err.append('diafragma no declarado ' + m.group(1))
    lassigned = set()
    for l in S['LINE ASSIGNS']:
        q = Q.findall(l)
        lassigned.add((q[0], q[1]))
        if q[0] not in lines_:
            err.append('línea no definida ' + q[0])
        if q[2] not in frames:
            err.append('sección de barra inexistente ' + q[2])
    pats = {Q.findall(l)[0] for l in S['LOAD PATTERNS'] if l.strip().startswith('LOADPATTERN')}
    for k in ('FRAME OBJECT LOADS', 'SHELL OBJECT LOADS'):
        for l in S[k]:
            q = Q.findall(l)
            lc = re.search(r'LC "([^"]*)"', l).group(1)
            if lc not in pats:
                err.append('patrón inexistente ' + lc)
            if (q[0], q[1]) not in (lassigned if k.startswith('FRAME') else assigned):
                err.append('carga sobre objeto no asignado %s %s' % (q[0], q[1]))
    for l in S['GROUPS']:
        q = Q.findall(l)
        if len(q) == 3:
            if (q[1], q[2]) not in assigned and (q[1], q[2]) not in lassigned:
                err.append('grupo con objeto inexistente %s' % q[1])
    for l in S['POINT ASSIGNS']:
        q = Q.findall(l)
        if q[0] not in pts:
            err.append('apoyo en punto inexistente ' + q[0])
    cases = {Q.findall(l)[0] for l in S['LOAD CASES']}
    funcs = {Q.findall(l)[0] for l in S['FUNCTIONS']}
    for l in S['LOAD CASES']:
        m = re.search(r'FUNC  "([^"]*)"', l)
        if m and m.group(1) not in funcs:
            err.append('función inexistente ' + m.group(1))
        m = re.search(r'LOADPAT  "([^"]*)"', l)
        if m and m.group(1) not in pats:
            err.append('patrón inexistente en caso ' + m.group(1))
    combos = {Q.findall(l)[0] for l in S['LOAD COMBINATIONS']}
    for l in S['LOAD COMBINATIONS']:
        m = re.search(r'LOADCASE "([^"]*)"', l)
        if m and m.group(1) not in cases:
            err.append('caso inexistente en combinación ' + m.group(1))
        m = re.search(r'LOADCOMBO "([^"]*)"', l)
        if m and m.group(1) not in combos:
            err.append('combinación inexistente ' + m.group(1))
    for l in S['MASS SOURCE']:
        if 'MASSSOURCELOAD' in l and Q.findall(l)[1] not in pats:
            err.append('patrón inexistente en masa ' + Q.findall(l)[1])
    # geometría: muros de longitud cero, losas degeneradas
    for name, (typ, ps) in areas.items():
        xy = [pts[p] for p in ps]
        if typ == 'PANEL':
            L = ((xy[0][0] - xy[1][0]) ** 2 + (xy[0][1] - xy[1][1]) ** 2) ** 0.5
            if L < 0.3:
                err.append('muro muy corto %s L=%.2f' % (name, L))
        else:
            a = 0.5 * abs(sum(xy[i][0] * xy[(i + 1) % len(xy)][1] - xy[(i + 1) % len(xy)][0] * xy[i][1]
                              for i in range(len(xy))))
            if a < 0.05:
                err.append('losa degenerada %s' % name)
    print('Niveles:', stories)
    print('Puntos %d, líneas %d, áreas %d (muros %d, losas %d)' % (
        len(pts), len(lines_), len(areas), sum(1 for v in areas.values() if v[0] == 'PANEL'),
        sum(1 for v in areas.values() if v[0] == 'FLOOR')))
    print('Patrones %d, casos %d, combinaciones %d' % (len(pats), len(cases), len(combos)))
    print('ERRORES:', len(err))
    for e in err[:50]:
        print('  ', e)
    return err


if __name__ == '__main__':
    sys.exit(1 if main(sys.argv[1]) else 0)
