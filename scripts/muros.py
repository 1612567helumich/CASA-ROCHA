"""Extracción de ejes de muro (centro) desde el DXF para el modelo estructural."""
import math
import ezdxf
from collections import defaultdict

T = 0.15          # espesor de muro dibujado
TOL = 0.006


def caras(doc, x0, y0, x1, y1, off=0.0):
    """Segmentos H/V de caras de muro (líneas continuas, sin proyecciones) en una región."""
    msp = doc.modelspace()
    H, V = [], []
    for e in msp.query('LINE LWPOLYLINE[layer=="001_Muros"]'):
        lt = e.dxf.get('linetype', 'BYLAYER').upper()
        if lt not in ('BYLAYER', 'CONTINUOUS'):
            continue
        pts = [(e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)] if e.dxftype() == 'LINE' \
            else [(p[0], p[1]) for p in e.get_points()]
        if e.dxftype() == 'LWPOLYLINE' and e.closed:
            pts.append(pts[0])
        for a, b in zip(pts, pts[1:]):
            if not (x0 <= a[0] <= x1 and y0 <= a[1] <= y1):
                continue
            ax, ay, bx, by = a[0] - off, a[1], b[0] - off, b[1]
            if abs(ay - by) < 1e-3 and abs(ax - bx) > 1e-3:
                H.append((round(ay, 4), min(ax, bx), max(ax, bx)))
            elif abs(ax - bx) < 1e-3 and abs(ay - by) > 1e-3:
                V.append((round(ax, 4), min(ay, by), max(ay, by)))
    return H, V


def _merge(iv):
    iv = sorted(iv)
    out = []
    for a, b in iv:
        if out and a <= out[-1][1] + 0.002:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return out


def pares(L, minlen=0.10):
    """Muros: pares de caras paralelas a 0.15 que se traslapan. Devuelve (c, a, b) centro y extremos."""
    by = defaultdict(list)
    for c, a, b in L:
        by[c].append((a, b))
    keys = sorted(by)
    muros = []
    for c1 in keys:
        for c2 in keys:
            if not (T - TOL <= c2 - c1 <= T + TOL):
                continue
            I1, I2 = _merge(by[c1]), _merge(by[c2])
            for a1, b1 in I1:
                for a2, b2 in I2:
                    a, b = max(a1, a2), min(b1, b2)
                    if b - a > minlen:
                        muros.append(((c1 + c2) / 2, a, b))
    # fusionar colineales
    byc = defaultdict(list)
    for c, a, b in muros:
        byc[round(c, 3)].append((a, b))
    out = []
    for c, iv in byc.items():
        for a, b in _merge(iv):
            out.append((c, a, b))
    return out
