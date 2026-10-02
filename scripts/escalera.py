"""Geometría paramétrica de la escalera curva (sin el primer escalón original).

Arranque recorrido 0.29 m: 17 peraltes de 3.00/17 = 0.1765 m y 16 huellas compensadas
(balanceadas) para cumplir huella >= 0.25 m medida a 0.40 m del barandal interior.
Todas las coordenadas en planta están referidas a la PLANTA BAJA (m).
"""
import math
import numpy as np
from shapely.geometry import LineString, Point, Polygon

# ----------------------------------------------------------------- datos
H_TOTAL = 3.00
N_PER = 17                      # peraltes
N_HU = 16                       # huellas
P = H_TOTAL / N_PER             # peralte
E_HUELLA = 0.06                 # espesor de huella (placa sobre zanca)
X0_ORIG = 37.551                # arranque original (se elimina el primer escalón)
X0 = 37.841                     # nuevo arranque
XC, YC = 39.285, 8.325          # centro del medio giro
RI, RE = 0.425, 1.425           # radio interior / exterior (ancho 1.00)
RW = 0.925                      # línea de huella (centro del ancho)
R40 = RI + 0.40                 # huella a 0.40 del barandal interior
R_BAR = 0.46                    # eje del barandal interior
LS = XC - X0                    # tramo recto (1.444 m)

# ley de compensación (obtenida por búsqueda numérica, ver PDF)
FIRST, T0, PEXP, QEXP = 0.265, 0.0, 1.2, -0.2


def path_pt(r, s):
    """Punto a longitud s sobre la trayectoria concéntrica de radio r (recto + medio giro)."""
    if s <= LS:
        return (X0 + s, YC - r)
    th = -math.pi / 2 + (s - LS) / r
    return (XC + r * math.cos(th), YC + r * math.sin(th))


def path_len(r):
    return LS + math.pi * r


def path_line(r, n=720):
    pts = [(X0 - 1e-4, YC - r), (XC, YC - r)]
    pts += [(XC + r * math.cos(-math.pi / 2 + math.pi * i / n), YC + r * math.sin(-math.pi / 2 + math.pi * i / n))
            for i in range(1, n + 1)]
    return LineString(pts)


P_IN, P_40, P_W, P_OUT, P_BAR = (path_line(r) for r in (RI, R40, RW, RE, R_BAR))


def _ley():
    La, Lo = path_len(R40), path_len(RE)
    a = [0.0, FIRST] + [FIRST + (La - FIRST) * j / (N_HU - 1) for j in range(1, N_HU)]
    o = []
    for i, ai in enumerate(a):
        t = i / N_HU
        if t <= T0:
            f = 0.0
        else:
            x = (t - T0) / (1 - T0)
            f = min(max(x ** PEXP * (1 + QEXP * (1 - x)), 0.0), 1.0)
        o.append(ai + (Lo - La) * f)
    return a, o


def _segmento(a, o):
    A = np.array(path_pt(R40, a))
    O = np.array(path_pt(RE, o))
    u = (O - A) / np.linalg.norm(O - A)
    ip = LineString([tuple(A), tuple(A - u * 1.2)]).intersection(P_IN)
    pts = [ip] if ip.geom_type == 'Point' else list(ip.geoms)
    Pp = min(pts, key=lambda p: p.distance(Point(A)))
    return (Pp.x, Pp.y), (float(O[0]), float(O[1]))


def _ley_u():
    """Forma de U simple: 5 huellas rectas perpendiculares + 11 radiales en el medio giro."""
    a, o = [], []
    for i in range(6):
        a.append(LS * i / 5); o.append(LS * i / 5)
    for j in range(1, 12):
        th = math.pi * j / 11
        a.append(LS + th * R40); o.append(LS + th * RE)
    return a, o


A_LEY, O_LEY = _ley_u()
NARIZ = [_segmento(a, o) for a, o in zip(A_LEY, O_LEY)]   # 17 líneas: arranque ... llegada


def _s_on(path, seg):
    from shapely.ops import nearest_points
    ln = LineString(seg)
    ip = ln.intersection(path)
    if ip.geom_type == 'Point':
        return path.project(ip)
    if ip.geom_type == 'MultiPoint':
        return path.project(list(ip.geoms)[0])
    return path.project(nearest_points(path, ln)[0])


S_IN = [_s_on(P_IN, s) for s in NARIZ]
S_40 = [_s_on(P_40, s) for s in NARIZ]
S_W = [_s_on(P_W, s) for s in NARIZ]
S_OUT = [_s_on(P_OUT, s) for s in NARIZ]
S_BAR = [_s_on(P_BAR, s) for s in NARIZ]


def sub_path(path, s0, s1, n=24):
    """Puntos de la trayectoria entre s0 y s1 (incluye vértices intermedios)."""
    pts = [path.interpolate(s0 + (s1 - s0) * k / n) for k in range(n + 1)]
    return [(p.x, p.y) for p in pts]


def huella_poly(i):
    """Polígono en planta de la huella i (1..16)."""
    a, b = NARIZ[i - 1], NARIZ[i]
    inner = sub_path(P_IN, S_IN[i - 1], S_IN[i], 12)
    outer = sub_path(P_OUT, S_OUT[i], S_OUT[i - 1], 16)
    return inner + outer


def z_huella(i):
    return i * P


def tabla():
    """Filas: (n, nivel, huella_w, huella_40, huella_in, huella_out)."""
    rows = []
    for i in range(1, N_HU + 1):
        rows.append((i, z_huella(i), S_W[i] - S_W[i - 1], S_40[i] - S_40[i - 1],
                     S_IN[i] - S_IN[i - 1], S_OUT[i] - S_OUT[i - 1]))
    return rows


def t_on(path_s, t):
    """Interpola la posición sobre una trayectoria para parámetro continuo t (0..16)."""
    t = min(max(t, 0.0), N_HU)
    i = min(int(t), N_HU - 1)
    f = t - i
    return path_s[i] + (path_s[i + 1] - path_s[i]) * f


def zanca(n=160):
    """Zanca exterior: lista de (x, y, z_sup, z_inf) a lo largo del borde exterior."""
    out = []
    for k in range(n + 1):
        t = N_HU * k / n
        p = P_OUT.interpolate(t_on(S_OUT, t))
        zs = P - E_HUELLA + t * P
        out.append((p.x, p.y, zs, max(0.0, zs - 0.30)))
    return out


def barandal(n=160):
    """Pasamanos (barandal de cristal) sobre el borde interior, h = 0.90 sobre nariz."""
    out = []
    for k in range(n + 1):
        t = N_HU * k / n
        p = P_BAR.interpolate(t_on(S_BAR, t))
        out.append((p.x, p.y, P + 0.90 + t * P))
    return out


def resumen():
    r = tabla()
    hw = [x[2] for x in r]
    h4 = [x[3] for x in r]
    hi = [x[4] for x in r]
    ho = [x[5] for x in r]
    return {
        'peralte': P, 'n_peraltes': N_PER, 'n_huellas': N_HU,
        'hw_min': min(hw), 'hw_max': max(hw), 'h40_min': min(h4), 'h40_max': max(h4),
        'hi_min': min(hi), 'ho_max': max(ho),
        '2p_h_min': 2 * P + min(hw), '2p_h_max': 2 * P + max(hw),
        'tramo_recto': LS, 'x0': X0,
    }


if __name__ == '__main__':
    for row in tabla():
        print('%2d  z=%.3f  hw=%.3f  h40=%.3f  hin=%.3f  hout=%.3f' % row)
    print(resumen())
