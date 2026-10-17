"""Arma data.json con todo lo que la hoja necesita: tableros, bordes, vacios, castillos propuestos,
tramos de nervadura principal, areas tributarias por tramo y transferencias de reaccion entre tramos.
Fuentes: 001_EST_CROCHA_R3.2.e2k, AREAS_TRIBUTARIAS_R3.2.dxf, 001_PLANO_CROCHA_R3.0.dxf."""
import ezdxf, json, re, sys
from shapely.geometry import Polygon, LineString, Point, box
from shapely.ops import unary_union

import os
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + '/'
m = json.load(open(W + 'work/model.json'))
geo = json.load(open(W + 'work/geo.json'))
T = 0.03
TC = 0.05
LV = ['N+3.00', 'N+6.00']
DX = {'N+3.00': 0.0, 'N+6.00': 14.0}
ABOVE = {'N+3.00': 'N+6.00', 'N+6.00': None}
GX = [('A', 0.0), ('B', 2.925), ('C', 3.427), ('D', 4.426), ('E', 4.925), ('F', 5.422), ('G', 5.923), ('H', 7.25), ('J', 8.85)]
GY = [('1', 0.0), ('2', 1.919), ('3', 2.921), ('4', 5.346), ('5', 5.946), ('6', 6.948), ('7', 7.871), ('8', 9.95),
      ('9', 12.021), ('10', 13.846), ('11', 15.447), ('12', 17.446), ('13', 19.447), ('14', 21.446), ('15', 24.371)]


def lab(ori, c):
    """Nombre del eje de un tramo: 'H' corre en X a y=c, 'V' corre en Y a x=c."""
    if ori == 'H':
        for n, v in GY:
            if abs(v - c) < 0.02:
                return 'Eje %s (y = %.3f)' % (n, c)
        return 'y = %.3f' % c
    for n, v in GX:
        if abs(v - c) < 0.02:
            return 'Eje %s (x = %.3f)' % (n, c)
    return 'x = %.3f' % c


def glab_pt(x, y):
    gx = next((n for n, v in GX if abs(v - x) < 0.02), None)
    gy = next((n for n, v in GY if abs(v - y) < 0.02), None)
    sx = gx if gx else 'x=%.3f' % x
    sy = gy if gy else 'y=%.3f' % y
    return '%s-%s' % (sx, sy)


# ---------------- castillos propuestos (bajar a cimentacion) ----------------
PROP = [
    ('C1', 'N+3.00', 4.327, 1.919, 'Esquina libre al frente de la cochera: ahí rematan las nervaduras de y = 1.919 y de x = 4.327. Ya hay castillo en el primer nivel.'),
    ('C2', 'N+3.00', 4.327, 2.921, 'Cochera: cruce de x = 4.327 con y = 2.921, bajo esquina de muro del primer nivel. Ya hay castillo en el primer nivel.'),
    ('C3', 'N+3.00', 5.923, 2.921, 'Esquina libre de la cochera (lado pickup): ahí rematan y = 2.921 y el eje G. Ya hay castillo en el primer nivel.'),
    ('C4', 'N+3.00', 4.327, 7.871, 'Esquina sur del vacío de doble altura: recibe la nervadura del eje 7. Ya hay castillo en el primer nivel.'),
    ('C5', 'N+3.00', 4.327, 14.021, 'Esquina norte del vacío de doble altura: sin él x = 4.327 queda en volado de 8 m. Ya hay castillo en el primer nivel.'),
    ('C6', 'N+3.00', 5.422, 9.95, 'Esquina escalera/cocina (F-8): recibe la nervadura del eje F que carga muro de planta alta. Ya hay castillo en el primer nivel.'),
    ('C7', 'N+3.00', 5.441, 8.91, 'Llegada de la escalera: amarra el anillo x = 5.44 / y = 8.91. NO hay castillo arriba: es propuesta nueva, revisar con arquitectura.'),
    ('C8', 'N+3.00', 3.427, 15.846, 'Esquina C-10 (arq.): ahí remata y = 15.846 que carga muro de planta alta. Ya hay castillo en el primer nivel.'),
]
EDGE_OVR = {
    ('F80003', 'N+3.00'): {'superior': ('MURO', 'CONTINUO')},
    ('F80008', 'N+3.00'): {'inferior': ('NERV. PRINCIPAL NP-10 / sigue el pasillo', 'DISCONTINUO'), 'superior': ('NERV. PRINCIPAL NP-15', 'CONTINUO'),
                           'izquierdo': ('NERV. PRINCIPAL NP-18, NP-19', 'CONTINUO'), 'derecho': ('MURO / NERV. PRINCIPAL NP-26', 'CONTINUO')},
    ('F80032', 'N+6.00'): {'inferior': ('MURO / vacío del pergolado', 'DISCONTINUO'), 'superior': ('MURO', 'CONTINUO'),
                           'izquierdo': ('MURO', 'CONTINUO'), 'derecho': ('MURO', 'CONTINUO')},
}
SYNTH = {'N+3.00': [((5.441, 14.247), (6.247, 14.247))]}   # prolonga la dala bajo muro de y=14.247 hasta x=5.44
VOLADOS = {('F80012', 'N+3.00'): dict(raiz='derecho', Lv=2.0), ('F80019', 'N+3.00'): dict(raiz='inferior', Lv=2.0)}
HUECOS_T = {('F80028', 'N+6.00'), ('F80029', 'N+6.00')}
# rectangulo representativo para tableros irregulares (claro corto que manda)
REP = {('F80008', 'N+3.00'): (1.995, 3.425, 'Parte x 3.427–5.422, y 14.021–17.446: es la de mayor claro corto. El pasillo de 1.10 m y la llegada de escalera tienen claros menores.'),
       ('F80032', 'N+6.00'): (1.995, 3.425, 'Parte x 3.427–5.422, y 14.021–17.446: la de mayor claro corto. El resto es pasillo de 1.10–1.60 m.'),
       ('F80003', 'N+3.00'): (4.327, 3.025, 'Forma de L: se toma el rectángulo envolvente 4.327 × 3.025 (conservador; el claro corto real es 2.425 en la parte izquierda).'),
       ('F80013', 'N+3.00'): (0.897, 1.999, 'Forma de L: rectángulo envolvente.')}


def merge(segs, tc=TC):
    H = {}
    V = {}
    for (p, q) in segs:
        if abs(p[1] - q[1]) < 1e-3:
            H.setdefault(p[1], []).append((min(p[0], q[0]), max(p[0], q[0])))
        elif abs(p[0] - q[0]) < 1e-3:
            V.setdefault(p[0], []).append((min(p[1], q[1]), max(p[1], q[1])))
    out = []
    for ori, D in (('H', H), ('V', V)):
        groups = []
        for c in sorted(D):
            if groups and abs(c - groups[-1][0]) < tc:
                groups[-1][1].extend(D[c]); groups[-1][2].append(c)
            else:
                groups.append([c, list(D[c]), [c]])
        for c, iv, clist in groups:
            cc = round(sum(clist) / len(clist), 3)
            iv.sort(); cur = list(iv[0])
            for a, b in iv[1:]:
                if a <= cur[1] + T:
                    cur[1] = max(cur[1], b)
                else:
                    out.append([ori, cc, round(cur[0], 3), round(cur[1], 3)]); cur = [a, b]
            out.append([ori, cc, round(cur[0], 3), round(cur[1], 3)])
    return out


def wall_segs(st):
    s = []
    for a in m['areas']:
        if a['story'] == st and a['kind'] == 'PANEL':
            xs = [c[0] for c in a['coords']]; ys = [c[1] for c in a['coords']]
            s.append(((min(xs), min(ys)), (max(xs), max(ys))))
    return s


def cols(st):
    return sorted(set((round(b['i'][0], 3), round(b['i'][1], 3)) for b in m['beams'] if b['story'] == st and b['kind'] == 'COLUMN'))


def beam_segs(st):
    s = [(tuple(b['i']), tuple(b['j'])) for b in m['beams'] if b['story'] == st and b['kind'] == 'BEAM' and b['section'] == '002_D-15x20']
    return s + SYNTH.get(st, [])


def pt(ori, c, s):
    return (s, c) if ori == 'H' else (c, s)


out = dict(tableros=[], tramos=[], trib=[], transfer=[], castillos_prop=[], castillos_sin_cont=[], vacios=geo['vacios'],
           lineloads=[], plantas={})
# ---------------- cargas lineales del E2K ----------------
beams = {(b['name'], b['story']): b for b in m['beams']}
LL = []
for t in m['lineloads']:
    b = beams[(t[1], t[2])]; lc = t[t.index('LC') + 1]
    p, q = b['i'], b['j']; Lb = ((q[0] - p[0]) ** 2 + (q[1] - p[1]) ** 2) ** 0.5
    if 'FVAL' in t:
        f0 = f1 = float(t[t.index('FVAL') + 1]); r0, r1 = 0, 1
    else:
        f0 = float(t[t.index('FSTART') + 1]); f1 = float(t[t.index('FEND') + 1])
        r0 = float(t[t.index('RDSTART') + 1]); r1 = float(t[t.index('RDEND') + 1])
    a = (p[0] + (q[0] - p[0]) * r0, p[1] + (q[1] - p[1]) * r0); bb = (p[0] + (q[0] - p[0]) * r1, p[1] + (q[1] - p[1]) * r1)
    LL.append(dict(beam=t[1], level=t[2], lc=lc, p=a, q=bb, f=(f0 + f1) / 2, L=round(Lb * (r1 - r0), 3)))
    out['lineloads'].append(dict(beam=t[1], level=t[2], lc=lc, x0=round(a[0], 3), y0=round(a[1], 3), x1=round(bb[0], 3), y1=round(bb[1], 3), w=(f0 + f1) / 2, L=round(Lb * (r1 - r0), 3)))
# ---------------- tramos ----------------
tid = 0
for lvl in LV:
    WL = merge(wall_segs(lvl))
    CL0 = cols(lvl)
    PR = [(c, x, y) for c, lv, x, y, _ in PROP if lv == lvl]
    CL = sorted(set(CL0) | set((x, y) for _, x, y in PR))
    DL = merge(beam_segs(lvl))
    WA = merge(wall_segs(ABOVE[lvl])) if ABOVE[lvl] else []
    lines = []
    for (ori, c, a, b) in DL:
        cov = [(max(a, a2), min(b, b2)) for (o2, c2, a2, b2) in WL if o2 == ori and abs(c2 - c) < TC and min(b, b2) > max(a, a2) + T]
        firm = {}
        for (x, y) in CL:
            s, cc = (x, y) if ori == 'H' else (y, x)
            if abs(cc - c) < TC and a - T <= s <= b + T:
                pid = next((cid for cid, px, py in PR if abs(px - x) < 0.01 and abs(py - y) < 0.01), None)
                firm[round(s, 3)] = pid if (x, y) not in CL0 else firm.get(round(s, 3))
        for (o2, c2, a2, b2) in WL:
            if o2 != ori and a2 - TC <= c <= b2 + TC and a - T <= c2 <= b + T:
                firm[round(c2, 3)] = None
        for lo, hi in cov:
            firm[round(lo, 3)] = None; firm[round(hi, 3)] = None
        cuts = sorted(set([a, b] + [f for f in firm if a - T <= f <= b + T]))
        c2l = []
        for x in cuts:
            if not c2l or x - c2l[-1] > 0.05:
                c2l.append(x)
        lines.append(dict(ori=ori, c=c, a=a, b=b, cov=cov, firm=firm, cuts=c2l))
    tr = []
    for L_ in lines:
        for s0, s1 in zip(L_['cuts'][:-1], L_['cuts'][1:]):
            if s1 - s0 < 0.10:
                continue
            mid = (s0 + s1) / 2
            onwall = any(lo - T <= mid <= hi + T for lo, hi in L_['cov'])
            def fk(s):
                for f, pid in L_['firm'].items():
                    if abs(s - f) < 0.05:
                        return True, pid
                return False, None
            f0, p0 = fk(s0); f1, p1 = fk(s1)
            tid += 1
            tr.append(dict(id=tid, level=lvl, ori=L_['ori'], c=L_['c'], s0=s0, s1=s1, L=round(s1 - s0, 3), tipo='MURO' if onwall else 'LIBRE',
                           f=[f0, f1], prop=[p0, p1]))

    def at(t, k):
        return pt(t['ori'], t['c'], t['s0'] if k == 0 else t['s1'])

    def interior_of(P, excl):
        res = []
        for t in tr:
            if t is excl:
                continue
            s, cc = (P[0], P[1]) if t['ori'] == 'H' else (P[1], P[0])
            if abs(cc - t['c']) < TC and t['s0'] + 0.05 < s < t['s1'] - 0.05:
                res.append((t, s - t['s0']))
        return res

    def ends_at(P, excl):
        res = []
        for t in tr:
            if t is excl:
                continue
            for k in (0, 1):
                Q = at(t, k)
                if abs(Q[0] - P[0]) < TC and abs(Q[1] - P[1]) < TC:
                    res.append((t, k))
        return res

    for t in tr:
        t['ext'] = [None, None]
        if t['tipo'] != 'LIBRE':
            continue
        for k in (0, 1):
            if t['f'][k]:
                t['ext'][k] = dict(tipo='FIRME', prop=t['prop'][k]); continue
            inn = interior_of(at(t, k), t)
            if inn:
                rt, a = inn[0]
                t['ext'][k] = dict(tipo='FIRME' if rt['tipo'] == 'MURO' else 'NERVADURA', recibe=None if rt['tipo'] == 'MURO' else rt['id'], a=round(a, 3)); continue
            t['ext'][k] = dict(tipo='NUDO')

    def classify(t):
        sup = [e['tipo'] in ('FIRME', 'NERVADURA', 'PUNTA') for e in t['ext']]
        return 'APOYADA' if all(sup) else ('VOLADO' if any(sup) else '?')

    for it in range(3):
        for t in tr:
            if t['tipo'] == 'LIBRE':
                t['cond'] = classify(t)
        for t in tr:
            if t['tipo'] != 'LIBRE':
                continue
            for k in (0, 1):
                if t['ext'][k]['tipo'] != 'NUDO':
                    continue
                for (o, ko) in ends_at(at(t, k), t):
                    if o['tipo'] == 'LIBRE' and o.get('cond') == 'VOLADO' and o['ext'][ko]['tipo'] == 'NUDO' and o is not t:
                        # punta de volado: vale como apoyo si el volado es corto o si este tramo no tiene otro apoyo
                        if o['L'] <= 0.6 or t['cond'] == '?':
                            t['ext'][k] = dict(tipo='PUNTA', recibe=o['id'], a=round(o['L'] if ko == 1 else 0.0, 3))
                            break
    for t in tr:
        if t['tipo'] == 'LIBRE':
            t['cond'] = classify(t)
            if t['cond'] == '?':
                t['cond'] = 'SIN APOYO'
            for e in t['ext']:
                if e['tipo'] == 'NUDO':
                    e['tipo'] = 'LIBRE'
    # muro de arriba y cargas lineales E2K
    for t in tr:
        ov = 0
        for (o2, c2, a2, b2) in WA:
            if o2 == t['ori'] and abs(c2 - t['c']) < TC:
                ov += max(0, min(t['s1'], b2) - max(t['s0'], a2))
        t['muro_arriba'] = round(ov, 3)
        acc = {}
        for l in LL:
            if l['level'] != lvl:
                continue
            p, q = l['p'], l['q']
            if t['ori'] == 'H' and abs(p[1] - q[1]) < 1e-3 and abs(p[1] - t['c']) < TC:
                lo, hi = min(p[0], q[0]), max(p[0], q[0])
            elif t['ori'] == 'V' and abs(p[0] - q[0]) < 1e-3 and abs(p[0] - t['c']) < TC:
                lo, hi = min(p[1], q[1]), max(p[1], q[1])
            else:
                continue
            ov = max(0, min(t['s1'], hi) - max(t['s0'], lo))
            if ov > 0:
                acc[l['lc']] = acc.get(l['lc'], 0) + l['f'] * ov
        t['e2k_scm'] = round(acc.get('Sobrecarga', 0) / t['L'], 1)
        t['e2k_cv'] = round(acc.get('Carga Viva Máxima', 0) / t['L'], 1)
        t['eje'] = lab(t['ori'], t['c'])
        t['x0'], t['y0'] = at(t, 0); t['x1'], t['y1'] = at(t, 1)
    out['tramos'].extend(tr)
    # castillos de arriba sin continuidad
    if ABOVE[lvl]:
        up = cols(ABOVE[lvl]); dn = set(CL0)
        WLs = [LineString([pt(o, c, a), pt(o, c, b)]) for (o, c, a, b) in WL]
        for P in up:
            if P in dn or any(l.distance(Point(P)) < TC for l in WLs):
                continue
            out['castillos_sin_cont'].append(dict(level=lvl, x=P[0], y=P[1], ref=glab_pt(*P)))
    # geometria para las plantas del croquis
    out['plantas'][lvl] = dict(muros=[[pt(o, c, a), pt(o, c, b)] for (o, c, a, b) in WL],
                               castillos=CL0)
for c, lv, x, y, why in PROP:
    out['castillos_prop'].append(dict(id=c, level=lv, x=x, y=y, ref=glab_pt(x, y), razon=why))

# tramos que son orilla libre de un tablero en volado -> nervadura de borde (no trabaja como trabe)
VOLBORDES = {'N+3.00': [('H', 14.021, -2.0, 0.0), ('H', 17.446, -2.0, 0.0), ('V', -2.0, 14.021, 17.446), ('H', 23.446, 0.0, 7.25)]}
for t in out['tramos']:
    for (o, c, a, b) in VOLBORDES.get(t['level'], []):
        if t['tipo'] == 'LIBRE' and t['ori'] == o and abs(t['c'] - c) < TC and t['s0'] >= a - T and t['s1'] <= b + T:
            t['cond'] = 'BORDE DE VOLADO'
            for e in t['ext']:
                if e and e['tipo'] == 'PUNTA':
                    e['tipo'] = 'LIBRE'; e.pop('recibe', None); e.pop('a', None)
# renumera tramos libres como NP-xx; los de muro como DM-xx
k = 0; km = 0
for t in out['tramos']:
    if t['tipo'] == 'LIBRE':
        k += 1; t['np'] = 'NP-%02d' % k
    else:
        km += 1; t['np'] = 'DM-%02d' % km
idmap = {t['id']: t['np'] for t in out['tramos']}
for t in out['tramos']:
    for e in t.get('ext') or []:
        if e and e.get('recibe'):
            e['recibe'] = idmap[e['recibe']]
# ---------------- tributarias ----------------
dt = ezdxf.readfile(W + 'in/c566697e-AREAS_TRIBUTARIAS_R3.2.dxf'); mt = dt.modelspace()
tabpolys = {(t['name'], t['level']): Polygon(t['pts']) for t in geo['tableros']}
tabgeo = {(t['name'], t['level']): t for t in geo['tableros']}
for p in mt.query('LWPOLYLINE[layer=="TRIB_AREA"]'):
    pts = [(float(x), float(y)) for x, y, *_ in p.get_points()]
    lvl = 'N+6.00' if max(x for x, y in pts) > 13 else 'N+3.00'
    pts = [(x - DX[lvl], y) for x, y in pts]
    P = Polygon(pts).buffer(0)
    if P.area < 1e-4:
        continue
    owner = None
    for (nm, lv), TP in tabpolys.items():
        if lv == lvl and TP.buffer(0.02).contains(P):
            owner = nm; break
    TP = tabpolys[(owner, lvl)]
    if abs(P.area - TP.area) < 0.01:
        # el poligono es todo el tablero (volado): la base es el borde con muro
        tg = tabgeo[(owner, lvl)]; x0, y0, x1, y1 = tg['bbox']
        side = max(tg['bordes'], key=lambda s: tg['bordes'][s]['frac_muro'])
        best = {'inferior': LineString([(x0, y0), (x1, y0)]), 'superior': LineString([(x0, y1), (x1, y1)]),
                'izquierdo': LineString([(x0, y0), (x0, y1)]), 'derecho': LineString([(x1, y0), (x1, y1)])}[side]
    else:
        best = None
        cs = list(P.exterior.coords)
        for (a, b) in zip(cs[:-1], cs[1:]):
            seg = LineString([a, b])
            if seg.length < 0.02:
                continue
            if TP.exterior.buffer(0.01).contains(seg) and (best is None or seg.length > best.length):
                best = seg
    (ax, ay), (bx, by) = best.coords
    ori = 'H' if abs(ay - by) < 1e-3 else 'V'
    c = ay if ori == 'H' else ax
    lo, hi = (min(ax, bx), max(ax, bx)) if ori == 'H' else (min(ay, by), max(ay, by))
    got = 0
    for t in out['tramos']:
        if t['level'] != lvl or t['ori'] != ori or abs(t['c'] - c) > TC:
            continue
        lo2, hi2 = max(lo, t['s0']), min(hi, t['s1'])
        if hi2 <= lo2 + 1e-3:
            continue
        band = box(lo2, -100, hi2, 100) if ori == 'H' else box(-100, lo2, 100, hi2)
        A = P.intersection(band).area
        if A > 1e-4:
            out['trib'].append(dict(tramo=t['np'], tablero=owner, level=lvl, area=round(A, 4))); got += A
    if P.area - got > 0.01:
        out['trib'].append(dict(tramo='SIN DALA', tablero=owner, level=lvl, area=round(P.area - got, 4)))
# ---------------- transferencias ----------------
for t in out['tramos']:
    if t['tipo'] != 'LIBRE' or t['cond'] == 'BORDE DE VOLADO':
        continue
    for k in (0, 1):
        e = t['ext'][k]
        if e['tipo'] in ('NERVADURA', 'PUNTA'):
            out['transfer'].append(dict(de=t['np'], extremo=k, a=e['recibe'], pos=e['a']))
# ---------------- tableros ----------------
trib_by_tab = {}
for r in out['trib']:
    trib_by_tab.setdefault((r['tablero'], r['level']), set()).add(r['tramo'])
npset = {t['np'] for t in out['tramos'] if t['tipo'] == 'LIBRE'}
for tb in geo['tableros']:
    key = (tb['name'], tb['level'])
    x0, y0, x1, y1 = tb['bbox']
    d = dict(name=tb['name'], level=tb['level'], x0=x0, y0=y0, x1=x1, y1=y1, Lx=tb['Lx'], Ly=tb['Ly'], area=tb['area_dxf'],
             area_poly=tb['area_poly'], rect=tb['rect'], pts=tb['pts'], seccion=tb['etabs_section'], loads=tb['loads'])
    rep = REP.get(key)
    d['ax'], d['ay'] = (tb['Lx'], tb['Ly'])
    d['rep_nota'] = ''
    if rep:
        d['ax'], d['ay'] = (rep[0], rep[1]) if tb['Lx'] >= tb['Ly'] or key != ('F80003', 'N+3.00') else (rep[0], rep[1])
        d['rep_nota'] = rep[2]
    d['modo'] = 'HUECO' if key in HUECOS_T else ('VOLADO' if key in VOLADOS else 'PERIMETRAL')
    if key in VOLADOS:
        d['volado'] = VOLADOS[key]
    bordes = {}
    for side, e in tb['bordes'].items():
        vec = [n for n, _ in e['vecinos'] if (n, tb['level']) not in HUECOS_T]
        fvec = sum(ov for n, ov in e['vecinos'] if (n, tb['level']) not in HUECOS_T) / e['L']
        hu = [n for n, _ in e['vecinos'] if (n, tb['level']) in HUECOS_T]
        vac = list(dict.fromkeys(e['vacios'] + (['Hueco/tragaluz (%s)' % ', '.join(hu)] if hu else [])))
        # tramos libres sobre este borde
        side_line = {'inferior': ('H', y0, x0, x1), 'superior': ('H', y1, x0, x1), 'izquierdo': ('V', x0, y0, y1), 'derecho': ('V', x1, y0, y1)}[side]
        o, c, a, b = side_line
        nps = [t['np'] for t in out['tramos'] if t['level'] == tb['level'] and t['tipo'] == 'LIBRE' and t['ori'] == o and abs(t['c'] - c) < TC and min(b, t['s1']) - max(a, t['s0']) > 0.05]
        if e['frac_muro'] >= 0.9:
            apoyo = 'MURO'
        elif nps:
            apoyo = 'NERV. PRINCIPAL ' + ', '.join(nps)
        else:
            apoyo = 'LIBRE'
        cont = 'CONTINUO' if fvec >= 0.5 else 'DISCONTINUO'
        bordes[side] = dict(L=e['L'], frac_muro=e['frac_muro'], apoyo=apoyo, cont=cont, vecinos=', '.join(vec), vacio=', '.join(vac), nps=nps)
    for side, (ap, co) in EDGE_OVR.get(key, {}).items():
        bordes[side]['apoyo'] = ap; bordes[side]['cont'] = co; bordes[side]['ovr'] = True
    d['bordes'] = bordes
    # carga lineal en borde libre de volados (E2K)
    if d['modo'] == 'VOLADO':
        raiz = VOLADOS[key]['raiz']
        opp = {'derecho': 'izquierdo', 'izquierdo': 'derecho', 'inferior': 'superior', 'superior': 'inferior'}[raiz]
        o, c = {'inferior': ('H', y0), 'superior': ('H', y1), 'izquierdo': ('V', x0), 'derecho': ('V', x1)}[opp]
        wsum = 0
        for l in LL:
            if l['level'] != tb['level'] or l['lc'] != 'Sobrecarga':
                continue
            p, q = l['p'], l['q']
            if o == 'H' and abs(p[1] - c) < TC and abs(q[1] - c) < TC:
                wsum = max(wsum, l['f'])
            if o == 'V' and abs(p[0] - c) < TC and abs(q[0] - c) < TC:
                wsum = max(wsum, l['f'])
        d['volado']['borde_libre'] = opp
        d['volado']['carga_borde'] = wsum
    out['tableros'].append(d)
json.dump(out, open(W + 'work/data.json', 'w'), indent=1, ensure_ascii=False)
if '-v' in sys.argv:
    for t in out['tramos']:
        if t['tipo'] == 'LIBRE':
            A = sum(r['area'] for r in out['trib'] if r['tramo'] == t['np'])
            print('%s %s %-20s %6.3f→%6.3f L=%5.2f %-9s %-45s %-45s A=%.2f muro=%.2f scm=%s cv=%s' % (
                t['np'], t['level'], t['eje'], t['s0'], t['s1'], t['L'], t['cond'], t['ext'][0], t['ext'][1], A, t['muro_arriba'], t['e2k_scm'], t['e2k_cv']))
    print(out['transfer'])
    for d in out['tableros']:
        print(d['name'], d['level'], d['modo'], d['ax'], d['ay'], {k: (v['apoyo'], v['cont']) for k, v in d['bordes'].items()})
    print([r for r in out['trib'] if r['tramo'] == 'SIN DALA'])
