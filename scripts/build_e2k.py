"""Genera 001_CASA_ROCHA_R1.0.e2k (ETABS) a partir del machote 002_HUAYACAN_R2.2.e2k.

Se conservan del machote: materiales, secciones, patrones de carga, casos, combinaciones,
fuente de masa y preferencias de diseño. Se borra todo el proyecto Huayacán (geometría,
asignaciones, cargas, grupos) y se construye la casa: ejes, niveles 0/3/6, muros de
mampostería confinada de tabique rojo, losas reticulares (espesor equivalente), cargas y
espectro sísmico para Morelia.

Uso: python3 scripts/build_e2k.py <ruta_machote.e2k>
"""
import datetime
import math
import os
import sys
from collections import OrderedDict

from shapely.geometry import LineString, Point, Polygon, box
from shapely.ops import unary_union

sys.path.insert(0, os.path.dirname(__file__))
import geometria_modelo as G  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, '001_CASA_ROCHA_R1.0.e2k')
NOMBRE = '001_CASA_ROCHA_R1.0'

# ---------------------------------------------------------------- niveles
ST0, ST1, ST2 = 'N 0.00', 'N+3.00', 'N+6.00'

# --------------------------------------------------------------- materiales
MAT_TAB = "Tabique rojo recocido f'm=15kg/cm2 (NTC-M 2023)"
MAT_BLK = "Block concreto tipo A f'm=20kg/cm2 (NTC-M 2023)"
MAT_CONC = "Concreto f'c=250kg/cm2. Clase 1A"
MAT_LOSA = "Concreto f'c=250 LOSA RETICULAR (peso equivalente)"
EC = 2.2135944E+09                       # 14000 raiz(250) kg/cm2, igual que el machote

# Losa reticular 60x60x20 + 5 (nervios de 10): espesor equivalente por rigidez y peso real
MOD, NERV, CAP, HT = 0.70, 0.10, 0.05, 0.25
_A1, _A2 = MOD * CAP, NERV * (HT - CAP)
_YC = (_A1 * CAP / 2 + _A2 * (CAP + (HT - CAP) / 2)) / (_A1 + _A2)
_I = (MOD * CAP ** 3 / 12 + _A1 * (_YC - CAP / 2) ** 2 + NERV * (HT - CAP) ** 3 / 12
      + _A2 * (CAP + (HT - CAP) / 2 - _YC) ** 2)
H_EQ = round((12 * _I / MOD) ** (1 / 3), 3)                    # m
VOL_M2 = (MOD * MOD * HT - (MOD - NERV) ** 2 * (HT - CAP)) / (MOD * MOD)   # m3 de concreto por m2
GAMMA_LOSA = round(2400 * VOL_M2 / H_EQ, 1)                    # kg/m3 equivalente

SEC_LOSA = 'LOSA RETICULAR h=25 (eq. %.1f cm)' % (H_EQ * 100)
SEC_MURO = 'MURO TABIQUE ROJO 14 cm'
SEC_MURO_BLK = 'MURO BLOCK 15 cm'
SEC_NULA = 'CARGA LINEAL (NULA)'

# ------------------------------------------------------------------ cargas
# NTC Criterios y Acciones 2023 (casa habitación): Wm / Wa / W
CV_ENT = (190, 100, 80)
CV_AZO = (100, 70, 15)
SCM_ENT = 190      # piso + firme + plafón + instalaciones + 40 kg/m2 adicional NTC
SCM_AZO = 250      # relleno para pendientes + enladrillado + impermeabilizante + plafón + 40
W_MURO = 270       # tabique 14 cm (210) + aplanados (60) kg/m2
PRETIL_AZO = round(0.60 * W_MURO)          # 162 kg/m
PRETIL_TERR = round(1.00 * W_MURO)         # 270 kg/m
BARANDAL = 100                             # cristal templado h 0.90
CANCEL = 300                               # cancelería + antepecho/dintel de mampostería
DOMO = 110                                 # domo sobre la escalera (cristal + guarnición)
PERGOLA = 516                              # 6 vigas 15x40 de 4.17 m + cristal -> mitad por apoyo
TRABE_PERG = 144                           # trabe 15x40 que recibe el pergolado (borde oriente)
TINACOS_KG = 2600                          # 2 tinacos de 1,100 L llenos + bases
ESC_CM = 540                               # medio peso de escalera en el borde del descanso
ESC_CV = (420, 220, 176)                   # medio peso de carga viva de escalera (Wm, Wa, W)

# ----------------------------------------------------------- sismo Morelia
# MDOC CFE 2015 (forma del espectro de diseño), datos supuestos a confirmar con PRODISIS:
SIS = OrderedDict(
    a0r=0.17,      # aceleración máxima en roca (g), Morelia (rango publicado 0.11-0.18 g, Tr~500 años)
    Fsit=1.30,     # factor de sitio, terreno tipo II (intermedio, sin estudio de mecánica de suelos)
    Fres=2.60,     # factor de respuesta
    Ts=0.50,       # periodo dominante supuesto del terreno (s)
    Ta=0.10, Tb=0.60, Tc=2.00, r=0.50, k=1.50,
    Q=2.0,         # mampostería confinada de piezas macizas
    R0=2.0,        # sobrerresistencia
    rho=1.0, alfa=1.0,
)
SIS['a0'] = round(SIS['a0r'] * SIS['Fsit'], 4)
SIS['c'] = round(SIS['a0'] * SIS['Fres'], 4)
G_ACC = 9.81


def sa_elastico(T):
    s = SIS
    a0, c, Ta, Tb, Tc, r, k = s['a0'], s['c'], s['Ta'], s['Tb'], s['Tc'], s['r'], s['k']
    if T < Ta:
        return a0 + (c - a0) * T / Ta
    if T < Tb:
        return c
    if T < Tc:
        return c * (Tb / T) ** r
    return c * (Tb / Tc) ** r * (k + (1 - k) * (Tc / T) ** 2) * (Tc / T) ** 2


def q_prima(T):
    s = SIS
    Q, k, Tb = s['Q'], s['k'], s['Tb']
    if T <= Tb:
        return 1 + (Q - 1) * math.sqrt(1.0 / k) * T / Tb
    p = k + (1 - k) * (Tb / T) ** 2
    return 1 + (Q - 1) * math.sqrt(p / k)


def r_sobre(T):
    Ta = SIS['Ta']
    if T <= Ta:
        return 10.0 / (4.0 + math.sqrt(T / Ta)) * SIS['R0'] / 2.0
    return SIS['R0']


def sa_diseno(T):
    return sa_elastico(T) / (q_prima(T) * r_sobre(T) * SIS['rho']) * SIS['alfa']


def periodos():
    Ts = [round(0.02 * i, 2) for i in range(0, 51)] + [round(1.0 + 0.1 * i, 2) for i in range(1, 31)]
    return Ts


# ---------------------------------------------------------------- utilidades
def f(v, nd=6):
    s = ('%.' + str(nd) + 'f') % v
    s = s.rstrip('0').rstrip('.') if '.' in s else s
    return '0' if s in ('-0', '') else s


class Puntos:
    def __init__(self):
        self.p = OrderedDict()

    def get(self, x, y):
        key = (round(x, 4), round(y, 4))
        for k, name in self.p.items():
            if abs(k[0] - key[0]) < 0.002 and abs(k[1] - key[1]) < 0.002:
                return name
        name = str(len(self.p) + 1)
        self.p[key] = name
        return name


def leer_machote(path):
    txt = open(path, encoding='utf-8').read().replace('\r\n', '\n')
    secs = OrderedDict()
    cur = None
    for line in txt.split('\n'):
        if line.startswith('$ '):
            cur = line[2:].strip()
            n = 1
            while cur in secs:
                n += 1
                cur = line[2:].strip() + ' #%d' % n
            secs[cur] = []
        elif cur is not None:
            secs[cur].append(line)
    return txt.split('\n')[0], secs


# ------------------------------------------------------------- geometría
def partir(ws):
    """Divide muros en los puntos donde llegan/cruzan muros perpendiculares."""
    out = []
    for w in ws:
        o, c, a, b = w
        cuts = {a, b}
        for p in ws:
            if p[0] == o:
                continue
            pc, pa, pb = p[1], p[2], p[3]
            if a + 1e-3 < pc < b - 1e-3 and pa - 1e-3 <= c <= pb + 1e-3:
                cuts.add(pc)
        cuts = sorted(cuts)
        keep = [cuts[0]]
        for v in cuts[1:-1]:
            if v - keep[-1] >= 0.30 and cuts[-1] - v >= 0.30:
                keep.append(v)
        keep.append(cuts[-1])
        for s0, s1 in zip(keep, keep[1:]):
            out.append((o, c, s0, s1))
    return out


def plan_xy(o, c, s):
    return (s, c) if o == 'H' else (c, s)


def main(machote):
    head, S = leer_machote(machote)
    pb, pa = G.leer()
    pb_s, pa_s = partir(pb), partir(pa)
    P = Puntos()

    def pt(x, y):
        return P.get(G.rx(x), G.ry(y))

    lines_conn, areas_conn = [], []
    line_assign, area_assign = [], []
    frame_loads, shell_loads, point_assign = [], [], []
    groups = OrderedDict()
    piers = []

    # ---------------- muros
    def muro_continuo(w):
        o, c, a, b = w
        L = 0.0
        for q in pb_s:
            if q[0] == o and abs(q[1] - c) < 0.02:
                L += max(0.0, min(b, q[3]) - max(a, q[2]))
        return L >= 0.9 * (b - a)

    azotea = G.losa_azotea().buffer(0.10)
    nw = 0
    base_pts = set()
    for nivel, lst in ((ST1, pb_s), (ST2, pa_s)):
        for w in lst:
            nw += 1
            name = 'W%d' % nw
            o, c, a, b = w
            p1 = pt(*plan_xy(o, c, a))
            p2 = pt(*plan_xy(o, c, b))
            areas_conn.append('  AREA "%s"  PANEL  4 "%s"  "%s"  "%s"  "%s"  1  1  0  0  ' % (name, p1, p2, p2, p1))
            lado = 'X' if o == 'H' else 'Y'
            pier = '%s-%s%d' % ('PB' if nivel == ST1 else 'PA', lado, nw)
            piers.append(pier)
            area_assign.append('  AREAASSIGN  "%s"  "%s"  SECTION "%s"  PIER "%s"  OBJMESHTYPE "DEFAULT"  '
                               'ADDRESTRAINT "%s"  '
                               % (name, nivel, SEC_MURO, pier, 'Yes' if nivel == ST1 else 'No'))
            if nivel == ST1:
                grp = 'MUROS PB'
                base_pts.update([p1, p2])
            else:
                seg = LineString([plan_xy(o, c, a), plan_xy(o, c, b)])
                if not muro_continuo(w):
                    grp = 'MUROS PA SOBRE LOSA (SIN MURO ABAJO)'
                elif not seg.within(azotea):
                    grp = 'MUROS PA SIN LOSA ARRIBA (BARDAS)'
                else:
                    grp = 'MUROS PA'
                if not seg.within(azotea) and grp != 'MUROS PA SIN LOSA ARRIBA (BARDAS)':
                    grp = 'MUROS PA SOBRE LOSA (SIN MURO ABAJO)'
            groups.setdefault(grp, []).append(('AREA', name, nivel))
            groups.setdefault('MUROS', []).append(('AREA', name, nivel))
    for p in sorted(base_pts, key=int):
        point_assign.append('  POINTASSIGN  "%s"  "%s"  RESTRAINT "UX UY UZ RX RY RZ"  ' % (p, ST0))

    # ---------------- losas
    nf = 0
    for nivel, poly, diaf, scm, cv, extra in (
            (ST1, G.losa_entrepiso(), 'D1', SCM_ENT, CV_ENT, ([], [])),
            (ST2, G.losa_azotea(), 'Azotea', SCM_AZO, CV_AZO, ([], []))):
        cx = [v for v in G.EJES_X.values()] + extra[0]
        cy = [v for v in G.EJES_Y.values()] + extra[1]
        for x0, y0, x1, y1 in G.rectangulos(poly, cx, cy):
            nf += 1
            name = 'F%d' % nf
            ps = [pt(x0, y0), pt(x1, y0), pt(x1, y1), pt(x0, y1)]
            areas_conn.append('  AREA "%s"  FLOOR  4 "%s"  "%s"  "%s"  "%s"  0  0  0  0  ' % (name, *ps))
            area_assign.append('  AREAASSIGN  "%s"  "%s"  SECTION "%s"  DIAPH  "%s"  OBJMESHTYPE "DEFAULT"  '
                               'ADDRESTRAINT "No"  CARDINALPOINT "TOP"  TRANSFORMSTIFFNESSFOROFFSETS "No"  '
                               % (name, nivel, SEC_LOSA, diaf))
            for lc, val in (('Sobrecarga', scm), ('Carga Viva Máxima', cv[0]),
                            ('Carga Viva Instantanea', cv[1]), ('Carga Viva Media', cv[2])):
                shell_loads.append('  AREALOAD  "%s"  "%s"  TYPE "UNIFF"  DIR "GRAV"  LC "%s"  FVAL %s'
                                   % (name, nivel, lc, f(val)))
            groups.setdefault('LOSAS N+3.00' if nivel == ST1 else 'LOSAS N+6.00', []).append(('AREA', name, nivel))

    # ---------------- vigas nulas con cargas lineales
    nb = [0]

    def carga_lineal(nivel, p_a, p_b, cargas, grp):
        nb[0] += 1
        name = 'B%d' % nb[0]
        a, b = pt(*p_a), pt(*p_b)
        lines_conn.append('  LINE  "%s"  BEAM  "%s"  "%s"  0' % (name, a, b))
        line_assign.append('  LINEASSIGN  "%s"  "%s"  SECTION "%s"  CARDINALPT 8  MAXSTASPC 0.5 AUTOMESH "YES"  '
                           'MESHATINTERSECTIONS "YES"  ' % (name, nivel, SEC_NULA))
        for lc, val in cargas:
            frame_loads.append('  LINELOAD  "%s"  "%s"  TYPE "UNIFF"  DIR "GRAV"  LC "%s"  FVAL %s'
                               % (name, nivel, lc, f(val)))
        groups.setdefault(grp, []).append(('LINE', name, nivel))

    A, J = G.EJES_X['A'], G.EJES_X['J']
    # N+3.00: barandales del vacío y de la escalera
    for pa_, pb_ in (((36.182, G.EJES_Y['7']), (36.182, G.EJES_Y['9'])), ((37.41, G.EJES_Y['6']), (37.41, 8.75)),
                     ((37.41, 8.75), (39.285, 8.75))):
        carga_lineal(ST1, pa_, pb_, [('Barandal', BARANDAL)], 'CARGAS LINEALES N+3.00')
    # terraza de la recámara principal (pretil al frente y barandal al oriente)
    carga_lineal(ST1, (A, 1.794), (36.362, 1.794), [('Pretil1.1', PRETIL_TERR)], 'CARGAS LINEALES N+3.00')
    carga_lineal(ST1, (36.362, 1.794), (36.362, 2.796), [('Barandal', BARANDAL)], 'CARGAS LINEALES N+3.00')
    # terraza eje 13-14: pretil sobre el eje 14
    carga_lineal(ST1, (A, 21.321), (G.EJES_X['H'], 21.321), [('Pretil1.1', PRETIL_TERR)], 'CARGAS LINEALES N+3.00')
    # cubierta del baño de PB (eje 14-15, H-J)
    carga_lineal(ST1, (G.EJES_X['H'], 21.321), (G.EJES_X['H'], 24.246), [('Pretil1.1', PRETIL_TERR)],
                 'CARGAS LINEALES N+3.00')
    # área de tendido (volado): barandal en el borde poniente
    carga_lineal(ST1, (29.936, 11.896), (29.936, 15.322), [('Barandal', BARANDAL)], 'CARGAS LINEALES N+3.00')
    # escalera: mitad de su peso en el borde del descanso
    carga_lineal(ST1, (39.285, 8.75), (39.285, 9.75),
                 [('RsmpasEscalera', ESC_CM), ('Carga Viva Máxima', ESC_CV[0]),
                  ('Carga Viva Instantanea', ESC_CV[1]), ('Carga Viva Media', ESC_CV[2])], 'CARGAS LINEALES N+3.00')
    # cancelería de planta alta (ventanales) sobre la losa N+3.00
    for o, c, a, b in G.ventanas_pa():
        pa_, pb_ = plan_xy(o, c, a), plan_xy(o, c, b)
        carga_lineal(ST1, pa_, pb_, [('CANCELERIA', CANCEL)], 'CANCELERIA (VENTANALES)')
    # N+6.00: pretil perimetral de azotea
    az = G.losa_azotea()
    ext = Polygon(az.exterior.coords) if az.geom_type == 'Polygon' else unary_union(az).convex_hull
    coords = list(ext.exterior.coords)
    for p1, p2 in zip(coords, coords[1:]):
        if math.dist(p1, p2) > 0.05:
            carga_lineal(ST2, p1, p2, [('Pretil1.1', PRETIL_AZO)], 'CARGAS LINEALES N+6.00')
    # pergolado sobre el vacío: apoyo en muro A y en trabe del borde oriente
    y7, y9 = G.EJES_Y['7'], G.EJES_Y['9']
    carga_lineal(ST2, (A, y7), (A, y9), [('Pergola', PERGOLA)], 'CARGAS LINEALES N+6.00')
    carga_lineal(ST2, (36.362, y7), (36.362, y9), [('Pergola', PERGOLA + TRABE_PERG)], 'CARGAS LINEALES N+6.00')
    # tinacos: dos apoyos (base) de 1.24 m bajo los 2 tinacos de 1,100 L
    for xb in (32.60, 34.57):
        carga_lineal(ST2, (xb, 13.58), (xb, 14.82), [('EquiposAzotea', round(TINACOS_KG / 2 / 1.24))],
                     'CARGAS LINEALES N+6.00')
    # domo sobre la escalera
    # tinacos: dos apoyos (base) de 1.24 m bajo los 2 tinacos de 1,100 L
    for xb in (32.60, 34.57):
        carga_lineal(ST2, (xb, 13.58), (xb, 14.82), [('EquiposAzotea', round(TINACOS_KG / 2 / 1.24))],
                     'CARGAS LINEALES N+6.00')
    gx, y6, y8 = G.EJES_X['G'], G.EJES_Y['6'], G.EJES_Y['8']
    d = [(gx, y6), (J, y6), (J, y8), (gx, y8), (gx, y6)]
    for p1, p2 in zip(d, d[1:]):
        carga_lineal(ST2, p1, p2, [('CANCELERIA', DOMO)], 'CARGAS LINEALES N+6.00')

    # ================================================================ ESCRITURA
    out = []
    w = out.append
    hoy = datetime.datetime.now().strftime('%d/%m/%Y %I:%M:%S %p').replace('AM', 'a. m.').replace('PM', 'p. m.')
    w('$ File %s.e2k generado %s (a partir del machote 002_HUAYACAN_R2.2)' % (NOMBRE, hoy))
    w(' ')
    for sec, body in S.items():
        base = sec.split(' #')[0]
        if base.startswith('File'):
            continue
        if base == 'CONTROLS':
            w('$ CONTROLS')
            for ln in body:
                if ln.strip().startswith('TITLE1'):
                    ln = '  TITLE1  "%s"  ' % NOMBRE
                elif ln.strip().startswith('TITLE2'):
                    ln = '  TITLE2  "Casa habitación Rocha - Morelia, Mich. - Mampostería confinada"  '
                w(ln)
        elif base == 'STORIES - IN SEQUENCE FROM TOP':
            w('$ ' + sec)
            w('  STORY "%s"  HEIGHT 3 ' % ST2)
            w('  STORY "%s"  HEIGHT 3 ' % ST1)
            w('  STORY "%s"  ELEV 0 ' % ST0)
            w('')
        elif base == 'GRIDS':
            w('$ GRIDS')
            w('  GRIDSYSTEM "Estructural"  TYPE "CARTESIAN"  BUBBLESIZE 1 ')
            for lab, x in G.EJES_X.items():
                w('  GRID "Estructural"  LABEL "%s"  DIR "X"  COORD %s VISIBLE "Yes"  BUBBLELOC "End"  '
                  % (lab, f(G.rx(x), 4)))
            for lab, y in G.EJES_Y.items():
                w('  GRID "Estructural"  LABEL "%s"  DIR "Y"  COORD %s VISIBLE "Yes"  BUBBLELOC "Start"  '
                  % (lab, f(G.ry(y), 4)))
            w('')
        elif base == 'DIAPHRAGM NAMES':
            w('$ DIAPHRAGM NAMES')
            w('  DIAPHRAGM "D1"    TYPE RIGID')
            w('  DIAPHRAGM "Azotea"    TYPE RIGID')
            w('')
        elif base == 'MATERIAL PROPERTIES':
            w('$ MATERIAL PROPERTIES')
            for ln in body:
                if ln.strip():
                    w(ln)
            tab_E = 600 * 15 * 1e4             # Em = 600 f'm (corta duración), kgf/m2
            blk_E = 800 * 20 * 1e4             # Em = 800 f'm (piezas de concreto)
            for name, grade, wv, E, fm in ((MAT_TAB, 'Mamposteria de Arcilla', 1500, tab_E, 150000),
                                           (MAT_BLK, 'Mamposteria de Concreto', 1650, blk_E, 200000)):
                w('  MATERIAL  "%s"    TYPE "Masonry"    GRADE "%s"    WEIGHTPERVOLUME %s' % (name, grade, f(wv)))
                w('  MATERIAL  "%s"    SYMTYPE "Isotropic"  E %s  U 0.25  A 4.500000159168849E-06' % (name, f(E)))
                w('  MATERIAL  "%s"  FM %s' % (name, f(fm)))
                w('  MATERIAL  "%s"  HYSTYPE "Concrete"  ENERGYDEGFACT 0 SSTYPE "Simple"  STRAINATFM %s '
                  'STRAINATULTIMATE 0.005 FINALSLOPE -0.1 IOTENSION 0.01 LSTENSION 0.02 CPTENSION 0.05 '
                  'IOCOMPRESSION 0.003 LSCOMPRESSION 0.006 CPCOMPRESSION 0.015 ' % (name, f(fm / E, 8)))
                # mismas líneas de curva esfuerzo-deformación que traen las mamposterías del machote
                w('  MATERIAL  "%s"  STRAIN -1 STRESS -1 ' % name)
                w('  MATERIAL  "%s"  STRAIN 0 STRESS 0 POINTID "A"  ' % name)
                w('  MATERIAL  "%s"  STRAIN 1 STRESS 1 ' % name)
            n = MAT_LOSA
            w('  MATERIAL  "%s"    TYPE "Concrete"    GRADE "Clase 1A"    WEIGHTPERVOLUME %s' % (n, f(GAMMA_LOSA)))
            w('  MATERIAL  "%s"    SYMTYPE "Isotropic"  E %s  U 0.2  A 5.500000042957254E-06' % (n, f(EC)))
            w('  MATERIAL  "%s"    FC 2500000  FR 316227.78' % n)
            w('  MATERIAL  "%s"  TIMEDEPCONCCODE "CEBFIP90"  ' % n)
            w('  MATERIAL  "%s"  HYSTYPE "Concrete"  ENERGYDEGFACT 0 SSTYPE "Mander"  STRAINATFC 0.0018386021 '
              'STRAINATULTIMATE 0.005 FINALSLOPE -0.1 IOTENSION 0.01 LSTENSION 0.02 CPTENSION 0.05 '
              'IOCOMPRESSION 0.003 LSCOMPRESSION 0.006 CPCOMPRESSION 0.015 ' % n)
            w('')
        elif base == 'FRAME SECTIONS':
            w('$ FRAME SECTIONS')
            body2 = [ln for ln in body if ln.strip()]
            mods = [ln for ln in body2 if 'MMOD' in ln and 'MATERIAL' not in ln]
            defs = [ln for ln in body2 if ln not in mods]
            for ln in defs:
                w(ln)
            w('  FRAMESECTION  "%s"  MATERIAL "%s"  SHAPE "Concrete Rectangular"  D 0.05 B 0.05 '
              % (SEC_NULA, MAT_CONC))
            for ln in mods:
                w(ln)
            w('  FRAMESECTION  "%s"  MMOD 0 WMOD 0 ' % SEC_NULA)
            w('')
        elif base == 'CONCRETE SECTIONS':
            w('$ CONCRETE SECTIONS')
            for ln in body:
                if ln.strip():
                    w(ln)
            w('  CONCRETESECTION  "%s"  LONGBARMATERIAL "Fy=4200"  CONFINEBARMATERIAL "Fy=4200"  TYPE "Beam"  '
              'COVERTOP 0.02 COVERBOTTOM 0.02 ATI 0 ABI 0 ATJ 0 ABJ 0 ' % SEC_NULA)
            w('')
        elif base == 'SLAB PROPERTIES':
            w('$ SLAB PROPERTIES')
            for ln in body:
                if ln.strip():
                    w(ln)
            w('  SHELLPROP  "%s"  PROPTYPE  "Slab"  MATERIAL "%s"  MODELINGTYPE "ShellThin"  SLABTYPE "Slab"  '
              'SLABTHICKNESS %s ' % (SEC_LOSA, MAT_LOSA, f(H_EQ)))
            w('')
        elif base == 'WALL PROPERTIES':
            w('$ WALL PROPERTIES')
            for ln in body:
                if ln.strip():
                    w(ln)
            w('  SHELLPROP  "%s"  PROPTYPE  "Wall"  MATERIAL "%s"  MODELINGTYPE "ShellThick"  WALLTHICKNESS 0.14 '
              % (SEC_MURO, MAT_TAB))
            w('  SHELLPROP  "%s"  PROPTYPE  "Wall"  MATERIAL "%s"  MODELINGTYPE "ShellThick"  WALLTHICKNESS 0.15 '
              % (SEC_MURO_BLK, MAT_BLK))
            w('')
        elif base == 'PIER/SPANDREL NAMES':
            w('$ PIER/SPANDREL NAMES ')
            for p in piers:
                w('  PIERNAME  "%s"  ' % p)
            w('  SPANDRELNAME  "S1"  ')
            w('')
        elif base == 'POINT COORDINATES':
            w('$ POINT COORDINATES')
            for (x, y), name in P.p.items():
                w('  POINT "%s"  %s %s ' % (name, f(x, 4), f(y, 4)))
            w('')
        elif base == 'LINE CONNECTIVITIES':
            w('$ LINE CONNECTIVITIES')
            for ln in lines_conn:
                w(ln)
            w('')
        elif base == 'AREA CONNECTIVITIES':
            w('$ AREA CONNECTIVITIES')
            for ln in areas_conn:
                w(ln)
            w('')
        elif base == 'GROUPS':
            w('$ GROUPS')
            for g, items in groups.items():
                w('  GROUP "%s"  ' % g)
                for typ, name, nivel in items:
                    w('  GROUP "%s"  %s "%s"  "%s"  ' % (g, typ, name, nivel))
            w('')
        elif base == 'POINT ASSIGNS':
            w('$ POINT ASSIGNS')
            for ln in point_assign:
                w(ln)
            w('')
        elif base == 'LINE ASSIGNS':
            w('$ LINE ASSIGNS')
            for ln in line_assign:
                w(ln)
            w('')
        elif base == 'AREA ASSIGNS':
            w('$ AREA ASSIGNS')
            for ln in area_assign:
                w(ln)
            w('')
        elif base == 'LOAD PATTERNS':
            w('$ LOAD PATTERNS')
            for ln in body:
                s = ln.strip()
                if not s:
                    continue
                if s.startswith('SEISMIC'):
                    ln = ('  SEISMIC "Sismo"  "User Coefficient"    DIR "X Y X+ECC Y+ECC X-ECC Y-ECC"  ECC 0.05  '
                          'TOPSTORY "%s"    BOTTOMSTORY "%s"    SHEARCOEFF %s  HEIGHTEXPONENT 1'
                          % (ST2, ST0, f(round(sa_diseno(0.12), 3))))
                elif s.startswith('WIND'):
                    ln = ln.replace('"N+4.08"', '"%s"' % ST2).replace('"N 0.00"', '"%s"' % ST0)
                if s.startswith('LOADPATTERN "CANCELERIA"'):
                    w(ln)
                    w('  LOADPATTERN "Pergola"  TYPE  "Super Dead"  SELFWEIGHT  0')
                    continue
                w(ln)
            w('')
        elif base == 'POINT OBJECT LOADS':
            w('$ POINT OBJECT LOADS')
            w('')
        elif base == 'FRAME OBJECT LOADS':
            w('$ FRAME OBJECT LOADS')
            for ln in frame_loads:
                w(ln)
            w('')
        elif base == 'SHELL OBJECT LOADS':
            w('$ SHELL OBJECT LOADS')
            for ln in shell_loads:
                w(ln)
            w('')
        elif base == 'MASS SOURCE':
            w('$ MASS SOURCE')
            for ln in body:
                if ln.strip():
                    w(ln)
            w('  MASSSOURCELOAD  "MassSOURCE 1"  "Pergola"  1 ')
            w('')
        elif base == 'FUNCTIONS':
            w('$ FUNCTIONS')
            for ln in body:
                if ln.strip() and 'Espectro Base de Diseño' not in ln:
                    w(ln)
            el = '  '.join('%s  %s' % (f(T, 2), f(sa_elastico(T), 5)) for T in periodos())
            di = '  '.join('%s  %s' % (f(T, 2), f(sa_diseno(T), 5)) for T in periodos())
            w('  FUNCTION "Espectro Base de Diseño"  FUNCTYPE "SPECTRUM"  DAMPRATIO 0.05  SPECTYPE "USER"  ')
            w('  FUNCTION "Espectro Base de Diseño"  TIMEVAL "%s"  ' % el)
            w('  FUNCTION "Espectro Diseño SV Morelia (Q=2)"  FUNCTYPE "SPECTRUM"  DAMPRATIO 0.05  SPECTYPE "USER"  ')
            w('  FUNCTION "Espectro Diseño SV Morelia (Q=2)"  TIMEVAL "%s"  ' % di)
            w('')
        elif base == 'LOAD CASES':
            w('$ LOAD CASES')
            for ln in body:
                s = ln.strip()
                if not s:
                    continue
                if s.startswith('LOADCASE "Modal"  MAXMODES'):
                    ln = ln.replace('MAXMODES  12', 'MAXMODES  30')
                if 'FUNC  "Espectro Base de Diseño"' in ln:
                    if 'Limitación de Daño' in ln:
                        pass     # espectro elástico x g/6 (criterio del machote)
                    elif 'Seguridad de Vida' in ln:
                        ln = ln.split('FUNC')[0] + 'FUNC  "Espectro Diseño SV Morelia (Q=2)"  SF  %s ' % f(G_ACC)
                    elif 'cimentacion' in ln:
                        ln = ln.split('FUNC')[0] + 'FUNC  "Espectro Diseño SV Morelia (Q=2)"  SF  %s ' % f(
                            round(G_ACC / 0.65, 4))
                w(ln)
                if s.startswith('LOADCASE "CANCELERIA"  LOADPAT'):
                    w('  LOADCASE "Pergola"  TYPE  "Linear Static"  INITCOND  "PRESET"  ')
                    w('  LOADCASE "Pergola"  LOADPAT  "Pergola"  SF  1 ')
            w('')
        elif base == 'LOAD COMBINATIONS':
            w('$ LOAD COMBINATIONS')
            for ln in body:
                if not ln.strip():
                    continue
                w(ln)
                if ln.strip().startswith('COMBO "SUPERDEAD_TOTAL"  LOADCASE "CANCELERIA"'):
                    w('  COMBO "SUPERDEAD_TOTAL"  LOADCASE "Pergola"  SF 1 ')
            w('')
        elif base == 'STEEL DESIGN OVERWRITES':
            w('$ STEEL DESIGN OVERWRITES')
            w('')
        elif base == 'PROJECT INFORMATION':
            w('$ PROJECT INFORMATION')
            w('  PROJECTINFO    COMPANYNAME "Casa Rocha"    MODELNAME "%s"  ' % NOMBRE)
            w('')
        elif base == 'LOG':
            w('$ LOG')
            w('  STARTCOMMENTS  ')
            w('%s generado a partir del machote 002_HUAYACAN_R2.2 (se borró la geometría de Huayacán).' % NOMBRE)
            w('Sismo: espectro MDOC CFE 2015 para Morelia, a0r=%.2f g, terreno tipo II supuesto; confirmar con PRODISIS.'
              % SIS['a0r'])
            w('  ENDCOMMENTS  ')
            w('')
            w('  END')
        elif base == 'END OF MODEL FILE':
            w('$ END OF MODEL FILE')
        else:
            w('$ ' + sec)
            for ln in body:
                w(ln)
    data = '\r\n'.join(out) + '\r\n'
    open(OUT, 'w', encoding='utf-8', newline='').write(data)
    resumen = dict(muros_pb=len(pb_s), muros_pa=len(pa_s), losas=nf, vigas_nulas=nb[0], puntos=len(P.p),
                   h_eq=H_EQ, gamma=GAMMA_LOSA, vol=VOL_M2, a0=SIS['a0'], c=SIS['c'],
                   coef_estatico=round(sa_diseno(0.12), 3), grupos={g: len(v) for g, v in groups.items()})
    return resumen, pb_s, pa_s


if __name__ == '__main__':
    r, _, _ = main(sys.argv[1])
    for k, v in r.items():
        print(k, v)
    print('OK ->', OUT)
