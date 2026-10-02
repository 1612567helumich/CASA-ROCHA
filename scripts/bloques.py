"""Bloques de mobiliario y vehículos (estilo del plano: capa 0, color POR BLOQUE)."""
import math

B45 = math.tan(math.radians(22.5))   # bulge de un arco de 90°
ATR = {'layer': '0', 'color': 0}


def _new_block(doc, name):
    if name in doc.blocks:
        blk = doc.blocks.get(name)
        for e in list(blk):
            blk.delete_entity(e)
        return blk
    return doc.blocks.new(name=name)


def pl(blk, pts, closed=False, **kw):
    a = dict(ATR)
    a.update(kw)
    return blk.add_lwpolyline([tuple(p) for p in pts], format='xyb' if pts and len(pts[0]) == 3 else 'xy',
                              close=closed, dxfattribs=a)


def ln(blk, p1, p2, **kw):
    a = dict(ATR)
    a.update(kw)
    return blk.add_line(p1, p2, dxfattribs=a)


def ci(blk, c, r, **kw):
    a = dict(ATR)
    a.update(kw)
    return blk.add_circle(c, r, dxfattribs=a)


def ar(blk, c, r, a0, a1, **kw):
    a = dict(ATR)
    a.update(kw)
    return blk.add_arc(c, r, a0, a1, dxfattribs=a)


def rrect(blk, x0, y0, x1, y1, r, **kw):
    """Rectángulo con esquinas redondeadas (radio r)."""
    r = min(r, (x1 - x0) / 2, (y1 - y0) / 2)
    if r <= 1e-6:
        return pl(blk, [(x0, y0), (x1, y0), (x1, y1), (x0, y1)], closed=True, **kw)
    p = [(x0 + r, y0, 0), (x1 - r, y0, B45), (x1, y0 + r, 0), (x1, y1 - r, B45),
         (x1 - r, y1, 0), (x0 + r, y1, B45), (x0, y1 - r, 0), (x0, y0 + r, B45)]
    return pl(blk, p, closed=True, **kw)


def superellipse(a, b, n, m=96, taper_front=0.0, taper_rear=0.0):
    """Contorno tipo 'superelipse' (para carrocerías). Frente hacia -y."""
    pts = []
    for i in range(m):
        t = 2 * math.pi * i / m
        c, s = math.cos(t), math.sin(t)
        x = a * math.copysign(abs(c) ** (2 / n), c)
        y = b * math.copysign(abs(s) ** (2 / n), s)
        k = abs(y) / b
        if y < 0:
            x *= 1 - taper_front * k ** 4
        else:
            x *= 1 - taper_rear * k ** 4
        pts.append((x, y))
    return pts


# ------------------------------------------------------------------ COMEDOR
def comedor15(doc):
    """Mesa 1.10 x 4.20 con 15 sillas (7 por lado + 1 en cabecera norte)."""
    name = 'MOB_COMEDOR15_110x420'
    blk = _new_block(doc, name)
    hw, hl = 0.55, 2.10
    rrect(blk, -hw, -hl, hw, hl, 0.03)
    rrect(blk, -hw + 0.04, -hl + 0.04, hw - 0.04, hl - 0.04, 0.02)   # canto de la cubierta

    def silla(cx, cy, ang):
        # asiento 0.45 x 0.45 con respaldo curvo; ang = dirección hacia donde mira la silla
        ca, sa = math.cos(math.radians(ang)), math.sin(math.radians(ang))

        def T(u, v):   # u: hacia la mesa, v: lateral
            return (cx + u * ca - v * sa, cy + u * sa + v * ca)
        seat = [T(-0.225, -0.225), T(0.225, -0.225), T(0.225, 0.225), T(-0.225, 0.225)]
        pl(blk, seat, closed=True)
        back = [T(-0.20, -0.20), T(-0.235, -0.10), T(-0.245, 0.0), T(-0.235, 0.10), T(-0.20, 0.20)]
        pl(blk, back)
    ys = [-1.80, -1.20, -0.60, 0.0, 0.60, 1.20, 1.80]
    for y in ys:
        silla(-hw - 0.225 + 0.175, y, 0)       # lado poniente, mira al oriente
        silla(hw + 0.225 - 0.175, y, 180)      # lado oriente, mira al poniente
    silla(0.0, hl + 0.225 - 0.10, 270)          # cabecera norte
    return name


# --------------------------------------------------------------------- SALA
def sala(doc):
    names = {}
    # Sofá seccional en L: brazo largo a lo largo de y (respaldo al oriente) y brazo corto al norte
    n = 'MOB_SALA_SECCIONAL_L_305x300'
    blk = _new_block(doc, n)
    # origen: esquina nororiente exterior del sofá
    # brazo largo: x -0.95..0, y -3.00..0 ; brazo corto: x -2.05..-0.95, y -0.95..0
    outline = [(0, 0), (-2.05, 0), (-2.05, -0.95), (-0.95, -0.95), (-0.95, -3.00), (0, -3.00)]
    pl(blk, outline, closed=True)
    # respaldos
    rrect(blk, -0.20, -2.80, 0.0, 0.0, 0.06)
    rrect(blk, -2.05, -0.20, -0.20, 0.0, 0.06)
    # brazos (descansabrazos)
    rrect(blk, -0.95, -3.00, 0.0, -2.80, 0.06)
    rrect(blk, -2.05, -0.95, -1.85, -0.20, 0.06)
    # cojines de asiento
    for y0, y1 in ((-2.78, -2.00), (-1.98, -1.20)):
        rrect(blk, -0.93, y0, -0.22, y1, 0.08)
    rrect(blk, -0.93, -1.18, -0.22, -0.22, 0.10)          # módulo esquina
    rrect(blk, -1.83, -0.93, -0.95, -0.22, 0.08)           # módulo brazo corto
    # cojines decorativos
    for (cx, cy, a) in ((-0.30, -2.55, 15), (-0.30, -0.35, 45), (-1.65, -0.30, 80)):
        ca, sa = math.cos(math.radians(a)), math.sin(math.radians(a))
        q = [(cx + u * ca - v * sa, cy + u * sa + v * ca) for u, v in
             ((-0.06, -0.20), (0.06, -0.20), (0.06, 0.20), (-0.06, 0.20))]
        pl(blk, q, closed=True)
    names['seccional'] = n

    # Sillón lounge con respaldo envolvente (mira hacia +x)
    n = 'MOB_SILLON_LOUNGE_80x80'
    blk = _new_block(doc, n)
    rrect(blk, -0.40, -0.40, 0.40, 0.40, 0.18)               # contorno
    # respaldo envolvente (herradura)
    pl(blk, [(0.25, 0.40, 0), (-0.22, 0.40, B45), (-0.40, 0.22, 0), (-0.40, -0.22, B45),
             (-0.22, -0.40, 0), (0.25, -0.40, 0), (0.25, -0.28, 0), (-0.16, -0.28, -B45),
             (-0.28, -0.16, 0), (-0.28, 0.16, -B45), (-0.16, 0.28, 0), (0.25, 0.28, 0)], closed=True)
    rrect(blk, -0.26, -0.26, 0.37, 0.26, 0.12)               # cojín
    names['lounge'] = n

    # Mesas de centro redondas anidadas
    n = 'MOB_MESAS_CENTRO_D90_D50'
    blk = _new_block(doc, n)
    ci(blk, (0, 0), 0.45)
    ci(blk, (0, 0), 0.42)
    ci(blk, (0.48, 0.52), 0.25)
    ci(blk, (0.48, 0.52), 0.22)
    names['mesas'] = n

    # Mesa lateral
    n = 'MOB_MESA_LATERAL_D40'
    blk = _new_block(doc, n)
    ci(blk, (0, 0), 0.20)
    ci(blk, (0, 0), 0.05)
    names['lateral'] = n

    # Tapete
    n = 'MOB_TAPETE_230x240'
    blk = _new_block(doc, n)
    pl(blk, [(-1.15, -1.20), (1.15, -1.20), (1.15, 1.20), (-1.15, 1.20)], closed=True)
    pl(blk, [(-1.08, -1.13), (1.08, -1.13), (1.08, 1.13), (-1.08, 1.13)], closed=True)
    names['tapete'] = n

    # Lámpara de arco (base + arco proyectado)
    n = 'MOB_LAMPARA_ARCO'
    blk = _new_block(doc, n)
    ci(blk, (0, 0), 0.17)
    ci(blk, (0, 0), 0.03)
    ar(blk, (-0.95, 0.0), 0.95, -40, 0)
    ci(blk, (-0.95 + 0.95 * math.cos(math.radians(-40)), 0.95 * math.sin(math.radians(-40))), 0.18)
    names['lampara'] = n
    return names


# -------------------------------------------------------- VEHÍCULOS (PLANTA)
def auto_planta(doc, name, L=4.60, W=1.85):
    """SUV/sedán moderno visto en planta. Frente hacia -y, origen al centro."""
    blk = _new_block(doc, name)
    a, b = W / 2, L / 2
    body = superellipse(a, b, 5.0, 120, taper_front=0.10, taper_rear=0.06)
    pl(blk, body, closed=True)
    # cofre: líneas de carácter
    for s in (-1, 1):
        pl(blk, [(s * 0.46, -b + 1.32), (s * 0.43, -b + 0.70), (s * 0.36, -b + 0.18)])
    # parabrisas
    yb, yt = -b + 1.35, -b + 1.95
    pl(blk, [(-0.78, yb, -0.06), (0.78, yb, 0), (0.67, yt, 0.05), (-0.67, yt, 0)], closed=True)
    # techo y techo panorámico
    rrect(blk, -0.69, yt, 0.69, b - 1.20, 0.14)
    rrect(blk, -0.50, yt + 0.15, 0.50, b - 1.42, 0.08)
    # medallón
    yr0, yr1 = b - 1.20, b - 0.70
    pl(blk, [(-0.69, yr0, 0), (0.69, yr0, 0), (0.62, yr1, 0.04), (-0.62, yr1, 0)], closed=True)
    # línea de cintura (ventanas laterales vistas en planta)
    for s in (-1, 1):
        pl(blk, [(s * 0.80, yb + 0.05), (s * 0.82, yt), (s * 0.82, yr0), (s * 0.76, yr1 + 0.05)])
    # espejos
    for s in (-1, 1):
        pl(blk, [(s * (a - 0.02), yb + 0.10), (s * (a + 0.13), yb + 0.16), (s * (a + 0.13), yb + 0.30),
                 (s * (a - 0.02), yb + 0.33)], closed=True)
    # faros LED delanteros y barra trasera
    for s in (-1, 1):
        pl(blk, [(s * 0.40, -b + 0.12), (s * 0.66, -b + 0.22), (s * 0.66, -b + 0.29), (s * 0.40, -b + 0.19)],
           closed=True)
    pl(blk, [(-0.55, b - 0.21), (0.55, b - 0.21), (0.55, b - 0.15), (-0.55, b - 0.15)], closed=True)
    # llantas (asomando bajo la carrocería)
    for yy in (-b + 0.95, b - 0.95):
        for s in (-1, 1):
            pl(blk, [(s * (a - 0.24), yy - 0.34), (s * (a - 0.02), yy - 0.34), (s * (a - 0.02), yy + 0.34),
                     (s * (a - 0.24), yy + 0.34)], closed=True, linetype='DASHED', ltscale=0.3)
    return name


def pickup_planta(doc, name, L=5.89, W=2.03):
    """Pick-up doble cabina moderna en planta. Frente hacia -y."""
    blk = _new_block(doc, name)
    a, b = W / 2, L / 2
    body = superellipse(a, b, 7.0, 140, taper_front=0.05, taper_rear=0.02)
    pl(blk, body, closed=True)
    # cofre con nervaduras
    for s in (-1, 1):
        pl(blk, [(s * 0.40, -b + 1.45), (s * 0.40, -b + 0.15)])
    ln(blk, (-0.86, -b + 0.10), (0.86, -b + 0.10))      # parrilla / defensa
    # parabrisas
    yb, yt = -b + 1.50, -b + 2.12
    pl(blk, [(-0.86, yb, -0.05), (0.86, yb, 0), (0.78, yt, 0.04), (-0.78, yt, 0)], closed=True)
    # techo de cabina
    yc = 0.62
    rrect(blk, -0.80, yt, 0.80, yc, 0.10)
    ln(blk, (-0.80, yc - 0.12), (0.80, yc - 0.12))       # medallón (vidrio trasero)
    # caja
    y0, y1 = yc + 0.12, b - 0.10
    rrect(blk, -a + 0.06, y0, a - 0.06, y1, 0.04)
    rrect(blk, -a + 0.13, y0 + 0.07, a - 0.13, y1 - 0.12, 0.03)
    for xx in (-0.45, 0.0, 0.45):
        ln(blk, (xx, y0 + 0.15), (xx, y1 - 0.20))
    ln(blk, (-a + 0.06, y1 - 0.12), (a - 0.06, y1 - 0.12))   # compuerta
    # espejos
    for s in (-1, 1):
        pl(blk, [(s * (a - 0.02), yb + 0.05), (s * (a + 0.17), yb + 0.12), (s * (a + 0.17), yb + 0.32),
                 (s * (a - 0.02), yb + 0.35)], closed=True)
    # faros LED y calaveras verticales
    for s in (-1, 1):
        pl(blk, [(s * 0.55, -b + 0.06), (s * 0.92, -b + 0.12), (s * 0.92, -b + 0.20), (s * 0.55, -b + 0.14)],
           closed=True)
        pl(blk, [(s * (a - 0.06), b - 0.40), (s * (a - 0.01), b - 0.40), (s * (a - 0.01), b - 0.08),
                 (s * (a - 0.06), b - 0.08)], closed=True)
    # llantas
    for yy in (-b + 1.05, b - 1.25):
        for s in (-1, 1):
            pl(blk, [(s * (a - 0.27), yy - 0.39), (s * (a - 0.02), yy - 0.39), (s * (a - 0.02), yy + 0.39),
                     (s * (a - 0.27), yy + 0.39)], closed=True, linetype='DASHED', ltscale=0.3)
    return name


# -------------------------------------------------------- VEHÍCULOS (ALZADO)
def _rueda(r_tire, r_rim, cx, cz, n_rayos=10):
    geo = [('circle', (cx, cz), r_tire), ('circle', (cx, cz), r_rim), ('circle', (cx, cz), r_rim * 0.28)]
    for i in range(n_rayos):
        t = 2 * math.pi * i / n_rayos
        geo.append(('line', (cx + r_rim * 0.3 * math.cos(t), cz + r_rim * 0.3 * math.sin(t)),
                    (cx + r_rim * 0.95 * math.cos(t + 0.18), cz + r_rim * 0.95 * math.sin(t + 0.18))))
    return geo


def _arco_rueda(cx, r, z0, n=24):
    """Puntos del arco de salpicadera (de atrás hacia adelante, sentido -x)."""
    pts = []
    for i in range(n + 1):
        t = math.radians(0 + 180 * i / n)
        pts.append((cx + r * math.cos(t), z0 + r * math.sin(t)))
    return pts


def suv_alzado():
    """Silueta y detalles de un SUV moderno de perfil (frente a la izquierda, x=0 en defensa)."""
    L = 4.60
    fx, rx, rw = 0.92, 3.68, 0.36
    sil = [(0.00, 0.47), (0.01, 0.70), (0.08, 0.84), (0.30, 0.90), (0.95, 0.99), (1.55, 1.07),
           (2.08, 1.58), (2.30, 1.64), (3.60, 1.65), (3.95, 1.60), (4.30, 1.42), (4.52, 1.20),
           (4.60, 0.95), (4.59, 0.47)]
    bottom = [(4.59, 0.47), (4.12, 0.44)] + _arco_rueda(rx, 0.44, 0.40) + [(3.24, 0.40), (1.36, 0.40)] \
        + _arco_rueda(fx, 0.44, 0.40) + [(0.48, 0.44), (0.00, 0.47)]
    outline = sil + bottom[1:]
    geo = [('pline', outline, True)]
    # ventanas
    geo.append(('pline', [(1.70, 1.13), (2.13, 1.55), (3.58, 1.57), (3.95, 1.50), (4.18, 1.34), (4.20, 1.15)], True))
    geo.append(('line', (2.90, 1.13), (2.90, 1.565)))       # poste B
    geo.append(('line', (3.92, 1.15), (3.98, 1.49)))       # poste C
    # puertas y manijas
    for x0, x1 in ((1.62, 1.60), (2.92, 2.92), (3.95, 3.92)):
        geo.append(('line', (x0, 0.48), (x1, 1.12)))
    for xh in (2.55, 3.65):
        geo.append(('line', (xh - 0.12, 0.97), (xh + 0.04, 0.97)))
    # línea de carácter y estribo
    geo.append(('pline', [(0.35, 0.86), (1.60, 0.93), (4.40, 1.02)], False))
    geo.append(('line', (1.36, 0.47), (3.24, 0.47)))
    # espejo y faros
    geo.append(('pline', [(1.66, 1.12), (1.60, 1.22), (1.78, 1.22), (1.80, 1.13)], True))
    geo.append(('pline', [(0.03, 0.80), (0.30, 0.86), (0.30, 0.82), (0.05, 0.77)], True))
    geo.append(('pline', [(4.40, 1.25), (4.56, 1.12), (4.56, 1.06), (4.40, 1.18)], True))
    geo += _rueda(rw, 0.25, fx, rw)
    geo += _rueda(rw, 0.25, rx, rw)
    return geo, outline, L


def pickup_alzado():
    """Pick-up doble cabina moderna de perfil (frente a la izquierda)."""
    L = 5.89
    fx, rx, rw = 1.02, 4.62, 0.40
    sil = [(0.00, 0.55), (0.02, 1.00), (0.10, 1.12), (0.95, 1.18), (1.72, 1.24), (2.28, 1.90),
           (2.45, 1.95), (3.62, 1.95), (3.70, 1.88), (3.72, 1.32), (5.84, 1.32), (5.89, 1.25),
           (5.89, 0.58)]
    bottom = [(5.89, 0.58), (5.10, 0.52)] + _arco_rueda(rx, 0.50, 0.46) + [(4.12, 0.46), (1.52, 0.46)] \
        + _arco_rueda(fx, 0.50, 0.46) + [(0.52, 0.52), (0.00, 0.55)]
    outline = sil + bottom[1:]
    geo = [('pline', outline, True)]
    geo.append(('pline', [(1.88, 1.27), (2.35, 1.84), (3.58, 1.86), (3.62, 1.28)], True))
    geo.append(('line', (2.86, 1.27), (2.88, 1.85)))
    for x in (1.80, 2.88, 3.72):
        geo.append(('line', (x, 0.50), (x, 1.26)))
    for xh in (2.55, 3.45):
        geo.append(('line', (xh - 0.12, 1.12), (xh + 0.04, 1.12)))
    geo.append(('line', (3.72, 1.20), (5.89, 1.20)))       # borde de caja
    geo.append(('line', (5.80, 0.62), (5.80, 1.30)))       # compuerta
    geo.append(('pline', [(0.40, 1.02), (5.85, 1.08)], False))
    geo.append(('pline', [(1.78, 1.24), (1.70, 1.38), (1.92, 1.38), (1.94, 1.25)], True))
    geo.append(('pline', [(0.02, 0.98), (0.40, 1.06), (0.40, 1.00), (0.03, 0.93)], True))
    geo += _rueda(rw, 0.27, fx, rw, 6)
    geo += _rueda(rw, 0.27, rx, rw, 6)
    return geo, outline, L
