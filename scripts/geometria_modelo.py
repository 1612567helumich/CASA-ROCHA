"""Geometría del modelo estructural (coordenadas ETABS: origen en eje A / eje 1, metros)."""
import math
import os
import sys

import ezdxf
from shapely.geometry import Polygon, box
from shapely.ops import unary_union

sys.path.insert(0, os.path.dirname(__file__))
import muros  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DXF = os.path.join(ROOT, '001_PLANO_CROCHA_R1.6.dxf')

X0, Y0 = 31.9354, -0.125          # eje A, eje 1 (coordenadas del DXF, planta baja)

# Ejes (coordenadas DXF de planta baja)
EJES_X = {'A': 31.9354, 'B': 34.8604, 'C': 35.3624, 'D': 36.3614, 'E': 36.8604, 'F': 37.3574,
          'G': 37.8584, 'H': 39.185, 'J': 40.7854}
EJES_Y = {'1': -0.125, '2': 1.794, '3': 2.796, '4': 5.221, '5': 5.821, '6': 6.823, '7': 7.746,
          '8': 9.825, '9': 11.896, '10': 13.721, '11': 15.322, '12': 17.321, '13': 19.322,
          '14': 21.321, '15': 24.246}

L_MIN = 0.50          # muros más cortos que esto no se modelan (remates / jambas)
SNAP = 0.012          # tolerancia para llevar un muro al eje


def rx(x):
    return round(x - X0, 4)


def ry(y):
    return round(y - Y0, 4)


def _snap(v, ejes):
    for e in ejes.values():
        if abs(v - e) <= SNAP:
            return e
    return v


def _merge_lin(segs, gap=0.17):
    """Une segmentos colineales separados por menos de un espesor de muro (encuentros en T)."""
    segs = sorted(segs)
    out = []
    for a, b in segs:
        if out and a <= out[-1][1] + gap:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return out


def muros_planta(doc, region, off):
    H, V = muros.caras(doc, *region, off=off)
    mh, mv = muros.pares(H), muros.pares(V)
    res = []
    for orient, lst, ejes_c, ejes_l in (('H', mh, EJES_Y, EJES_X), ('V', mv, EJES_X, EJES_Y)):
        byc = {}
        for c, a, b in lst:
            cc = _snap(c, ejes_c)
            byc.setdefault(round(cc, 4), []).append((a, b))
        for c, iv in byc.items():
            for a, b in _merge_lin(iv):
                res.append([orient, c, a, b])
    return res


def ajustar_extremos(ws):
    """Lleva los extremos al eje del muro perpendicular (si lo toca) o al eje cercano."""
    for w in ws:
        o, c, a, b = w
        perp = [p for p in ws if p[0] != o]
        for idx in (2, 3):
            v = w[idx]
            best = None
            for p in perp:
                pc, pa, pb = p[1], p[2], p[3]
                if pa - 0.08 <= c <= pb + 0.08 and abs(pc - v) <= 0.16:
                    if best is None or abs(pc - v) < abs(best - v):
                        best = pc
            if best is not None:
                w[idx] = best
            else:
                ejes = EJES_Y if o == 'V' else EJES_X
                for e in ejes.values():
                    if abs(e - v) <= 0.08:
                        w[idx] = e
                        break
    return [w for w in ws if w[3] - w[2] >= L_MIN]


def leer():
    doc = ezdxf.readfile(DXF)
    pb = ajustar_extremos(muros_planta(doc, (29, -1, 44, 26), 0.0))
    pa = ajustar_extremos(muros_planta(doc, (59, -1, 74, 26), 30.0))
    return pb, pa


def ventanas_pa(doc=None):
    """Tramos de cancelería (capa 012_Ventanas) en planta alta, coord. de planta baja."""
    doc = doc or ezdxf.readfile(DXF)
    out = []
    for e in doc.modelspace().query('LWPOLYLINE LINE[layer=="012_Ventanas"]'):
        pts = [(e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)] if e.dxftype() == 'LINE' \
            else [(p[0], p[1]) for p in e.get_points()]
        xs = [p[0] - 30 for p in pts]
        ys = [p[1] for p in pts]
        if not (59 <= pts[0][0] <= 74 and -1 <= pts[0][1] <= 26):
            continue
        if max(xs) - min(xs) > 0.5 and max(ys) - min(ys) < 0.06:
            out.append(('H', round(sum(ys) / len(ys), 3), min(xs), max(xs)))
        elif max(ys) - min(ys) > 0.5 and max(xs) - min(xs) < 0.06:
            out.append(('V', round(sum(xs) / len(xs), 3), min(ys), max(ys)))
    # una línea por ventana (se dibujan 3 paralelas)
    uniq = []
    for v in sorted(out):
        if not any(u[0] == v[0] and abs(u[1] - v[1]) < 0.2 and abs(u[2] - v[2]) < 0.05 and abs(u[3] - v[3]) < 0.05
                   for u in uniq):
            uniq.append(v)
    return uniq


# ----------------------------------------------------------------- losas
def losa_entrepiso():
    """Losa N+3.00 (coord. de planta baja) menos huecos (vacío y escalera)."""
    A, J = EJES_X['A'], EJES_X['J']
    partes = [
        box(A, 1.794, 36.362, 2.796),                 # terraza recámara principal
        box(A, 2.796, 37.859, 6.823),                 # recámara principal y baño
        box(A, 6.823, J, 21.321),                     # cuerpo principal + terraza eje 13-14
        box(39.185, 21.321, J, 24.246),               # cubierta del baño de PB (eje 14-15)
        box(29.936, 11.896, A, 15.322),               # área de tendido en volado
    ]
    s = unary_union(partes)
    # huecos referidos a ejes de muro para no dejar tiras de losa de 7 cm
    vacio = box(EJES_X['A'], EJES_Y['7'], 36.182, EJES_Y['9'])
    escalera = Polygon([(37.41, EJES_Y['6']), (EJES_X['J'], EJES_Y['6']), (EJES_X['J'], EJES_Y['8']),
                        (39.285, EJES_Y['8']), (39.285, 8.75), (37.41, 8.75)])
    return s.difference(vacio).difference(escalera)


def losa_azotea():
    A, J = EJES_X['A'], EJES_X['J']
    s = Polygon([(A, 1.794), (36.361, 1.794), (36.361, 2.796), (37.858, 2.796), (37.858, 6.823),
                 (J, 6.823), (J, 19.322), (A, 19.322)])
    pergola = box(EJES_X['A'], EJES_Y['7'], 36.362, EJES_Y['9'])   # pergolado de vigas 15x40 + cristal
    domo = box(EJES_X['G'], EJES_Y['6'], EJES_X['J'], EJES_Y['8'])  # vacío sobre la escalera
    return s.difference(pergola).difference(domo)


def rectangulos(poly, cortes_x=(), cortes_y=()):
    """Descompone un polígono ortogonal en rectángulos (celdas fusionadas por renglón)."""
    xs, ys = set(cortes_x), set(cortes_y)
    geoms = [poly] if poly.geom_type == 'Polygon' else list(poly.geoms)
    for g in geoms:
        for ring in [g.exterior] + list(g.interiors):
            for x, y in ring.coords:
                xs.add(round(x, 4))
                ys.add(round(y, 4))
    xs, ys = sorted(xs), sorted(ys)
    cells = []
    for j in range(len(ys) - 1):
        row = []
        for i in range(len(xs) - 1):
            cbox = box(xs[i], ys[j], xs[i + 1], ys[j + 1])
            if poly.intersection(cbox).area >= 0.98 * cbox.area:
                row.append(i)
        # fusionar celdas contiguas del renglón
        k = 0
        while k < len(row):
            k2 = k
            while k2 + 1 < len(row) and row[k2 + 1] == row[k2] + 1:
                k2 += 1
            cells.append([xs[row[k]], ys[j], xs[row[k2] + 1], ys[j + 1]])
            k = k2 + 1
    # fusionar verticalmente rectángulos con el mismo ancho
    merged = True
    while merged:
        merged = False
        for a in cells:
            for b in cells:
                if a is not b and abs(a[0] - b[0]) < 1e-6 and abs(a[2] - b[2]) < 1e-6 and abs(a[3] - b[1]) < 1e-6:
                    a[3] = b[3]
                    cells.remove(b)
                    merged = True
                    break
            if merged:
                break
    return cells
