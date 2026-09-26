import ezdxf
from ezdxf import bbox
from ezdxf.math import Vec3

D = 0.10          # alargue: baño de entrada 1.35 -> 1.45 libre
CUT_Y = 5.45      # entre eje 4 (5.221) y eje 5 (5.721)
CUT_X = CUT_Y + 42.5   # misma posicion en Corte A-A' (eje 1 y=-0.125 -> x=42.375)

d = ezdxf.readfile('orig.dxf'); m = d.modelspace()
cache = bbox.Cache()

def region(b):
    x0, y0, x1, y1 = b.extmin.x, b.extmin.y, b.extmax.x, b.extmax.y
    if x1 < 105 and y0 > -3.5: return 'plan'
    if 41 < x0 and x1 < 68.5 and -16 < y0 and y1 < -3.5: return 'secA'
    return None

def S(p, reg):
    p = Vec3(p)
    if reg == 'plan' and p.y > CUT_Y: return p + (0, D, 0)
    if reg == 'secA' and p.x > CUT_X: return p + (D, 0, 0)
    return p

def ref_above(p, reg):
    return S(p, reg) != Vec3(p)

# --- Baño de atras: muro de planta alta alineado con el baño de PB (x 39.11/39.185/39.26 en PB) ---
MAP = {9.285: 9.11, 9.36: 9.185, 9.435: 9.26, 9.437: 9.26}
X0 = 60.0
nfix = 0
def fixp(x, y):
    global nfix
    if 21.1 < y < 24.25:
        for a, bb in MAP.items():
            if abs(x - (X0 + a)) < 0.004:
                nfix += 1
                return X0 + bb
    return x
for e in m.query('LINE LWPOLYLINE[layer=="001_Muros"]'):
    if e.dxftype() == 'LINE':
        s, t_ = e.dxf.start, e.dxf.end
        e.dxf.start = (fixp(s.x, s.y), s.y, s.z); e.dxf.end = (fixp(t_.x, t_.y), t_.y, t_.z)
    else:
        e.set_points([(fixp(x, y), y, s, ee, bu) for x, y, s, ee, bu in e.get_points('xyseb')], 'xyseb')


stats = {}
redim = []
for e in list(m):
    b = bbox.extents([e], cache=cache)
    if not b.has_data: continue
    reg = region(b)
    if not reg: continue
    lo = b.extmin; hi = b.extmax
    k = (reg == 'plan' and lo.y > CUT_Y) or (reg == 'secA' and lo.x > CUT_X)
    cross = (reg == 'plan' and lo.y <= CUT_Y < hi.y) or (reg == 'secA' and lo.x <= CUT_X < hi.x)
    if not (k or cross): continue
    t = e.dxftype()
    vec = (0, D, 0) if reg == 'plan' else (D, 0, 0)
    if t == 'DIMENSION':
        for a in ('defpoint', 'defpoint2', 'defpoint3', 'defpoint4', 'defpoint5', 'text_midpoint'):
            if e.dxf.hasattr(a): e.dxf.set(a, S(e.dxf.get(a), reg))
        blk = d.blocks[e.dxf.geometry]
        for x in blk:
            xt = x.dxftype()
            if xt == 'LINE':
                x.dxf.start = S(x.dxf.start, reg); x.dxf.end = S(x.dxf.end, reg)
            elif xt in ('INSERT', 'POINT', 'MTEXT'):
                a_ = 'location' if xt == 'POINT' else 'insert'
                p0 = Vec3(x.dxf.get(a_))
                if k: x.translate(*vec)
                elif xt == 'MTEXT': x.dxf.insert = p0 + Vec3(vec) * 0.5
                else: x.dxf.set(a_, S(p0, reg))
        if not k:
            old = e.dxf.get('actual_measurement', -1)
            if old > 0: e.dxf.actual_measurement = old + D
            redim.append(e)
    elif k and t != 'HATCH':
        e.translate(*vec)
    elif t == 'LINE':
        e.dxf.start = S(e.dxf.start, reg); e.dxf.end = S(e.dxf.end, reg)
    elif t == 'LWPOLYLINE' and e.dxf.layer != '021_Muebles_Corte':
        pts = [(S((x, y), reg).x, S((x, y), reg).y, s, ee, bu) for x, y, s, ee, bu in e.get_points('xyseb')]
        e.set_points(pts, 'xyseb')
    elif t == 'HATCH':
        if not k: e.dxf.associative = 0
        for path in e.paths:
            if hasattr(path, 'vertices'):
                path.vertices = [(S((v[0], v[1]), reg).x, S((v[0], v[1]), reg).y, v[2]) for v in path.vertices]
            else:
                for ed in path.edges:
                    if ed.type.name == 'LINE':
                        ed.start = S(ed.start, reg).vec2; ed.end = S(ed.end, reg).vec2
                    else:
                        raise RuntimeError('edge %s' % ed.type)
    else:
        # bloques, textos, arcos, muebles de corte: se mueven completos segun su punto base
        base = {'INSERT': 'insert', 'MTEXT': 'insert', 'TEXT': 'insert', 'ARC': 'center', 'CIRCLE': 'center'}.get(t)
        p = e.dxf.get(base) if base else (lo + hi) / 2
        if ref_above(p, reg): e.translate(*vec)
    stats[(reg, t, 'move' if k else 'stretch')] = stats.get((reg, t, 'move' if k else 'stretch'), 0) + 1

import re
for e in redim:
    for x in d.blocks[e.dxf.geometry].query('MTEXT'):
        before = x.text
        x.text = re.sub(r'\d+\.\d+', lambda mm: '%.2f' % (float(mm.group()) + D), x.text, count=1)
        print('dim', e.dxf.handle, repr(before), '->', repr(x.text))

for k_, v in sorted(stats.items()): print(k_, v)
print('dims re-rendered', len(redim), 'points fixed back bath', nfix)
d.saveas('/home/user/CASA-ROCHA/001_PLANO_CROCHA_R1_4.dxf')
