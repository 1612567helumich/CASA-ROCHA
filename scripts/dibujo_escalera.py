"""Dibujo de la escalera recalculada en plantas, cortes y lámina de detalle."""
import math
from ezdxf import bbox
from shapely.geometry import Polygon, LineString
from shapely.ops import unary_union

import escalera as E

L_ESC = '005_Escalera'
L_BAR = '020_Barandales'
L_TXT = '007_Textos'
VACIO = {'linetype': 'VACIO'}


# ------------------------------------------------------------------ borrar
def borrar_original(c):
    msp = c.msp
    # plantas: todo lo de la capa de escalera dentro del cubo de escalera
    for off in (0.0, 30.0):
        for e in list(msp.query(f'*[layer=="{L_ESC}"]')):
            b = bbox.extents([e])
            if b.has_data and b.extmin.x >= 37.3 + off and b.extmax.x <= 40.8 + off \
                    and b.extmin.y >= 6.85 and b.extmax.y <= 9.8:
                msp.delete_entity(e)
    # cortes: huellas, zanca, pasamanos y postes del barandal de planta alta
    c.kill('554C', '554D', '554E', '554F', '5550', '5551', '5552', '5553', '5554', '5555', '5556',
           '5557', '5558', '5559', '555A', '555B', '555C', '555D', '555E', '555F',
           '5560', '5561', '5562', '5563',
           '5535', '5536', '5537', '5538', '5539', '553A', '553B', '553C', '553D', '553E', '553F',
           '5540', '5541', '5542', '5543', '5544', '5545', '5546', '5547', '5548',
           '5549', '554A', '554B')
    # lámina de detalle completa (se redibuja)
    for e in list(msp):
        b = bbox.extents([e])
        if b.has_data and b.extmin.x >= 108.0 and b.extmax.x <= 126.6 and b.extmin.y >= 6.2 \
                and b.extmax.y <= 13.6:
            msp.delete_entity(e)


# ------------------------------------------------------------------ plantas
def _T(pts, dx, dy, k=1.0, cx=0.0, cy=0.0):
    return [(cx + (x - cx) * k + dx, cy + (y - cy) * k + dy) for x, y in pts]


def _flecha(c, tip, prev, layer, size=0.12, dx=0, dy=0):
    ang = math.atan2(tip[1] - prev[1], tip[0] - prev[0])
    for s in (1, -1):
        a = ang + math.pi - s * math.radians(25)
        c.line((tip[0] + dx, tip[1] + dy), (tip[0] + dx + size * math.cos(a), tip[1] + dy + size * math.sin(a)), layer)


def planta(c, dx=0.0, dy=0.0, corte=True, flecha='SUBE'):
    """Planta de la escalera. corte=True: planta baja (corte a ~1.10 m, resto punteado)."""
    nb = 6   # la línea de corte va entre la nariz 6 y la 7 (huella 6 a 1.06 m)
    sb = lambda S: (S[nb] + S[nb + 1]) / 2
    for path, S in ((E.P_IN, E.S_IN), (E.P_OUT, E.S_OUT)):
        L = path.length
        if corte:
            c.pline(_T(E.sub_path(path, 0, sb(S), 40), dx, dy), L_ESC)
            c.pline(_T(E.sub_path(path, sb(S), L, 80), dx, dy), L_ESC, **VACIO)
        else:
            c.pline(_T(E.sub_path(path, 0, L, 120), dx, dy), L_ESC)
    for i, (p1, p2) in enumerate(E.NARIZ):
        kw = VACIO if (corte and i > nb) else {}
        c.line((p1[0] + dx, p1[1] + dy), (p2[0] + dx, p2[1] + dy), L_ESC, **kw)
    if corte:
        a = E.P_IN.interpolate(sb(E.S_IN))
        b = E.P_OUT.interpolate(sb(E.S_OUT))
        ux, uy = b.x - a.x, b.y - a.y
        L = math.hypot(ux, uy)
        ux, uy = ux / L, uy / L
        nx, ny = -uy, ux
        m = 0.5
        pts = [(a.x, a.y), (a.x + ux * (m - 0.12), a.y + uy * (m - 0.12)),
               (a.x + ux * (m - 0.04) + nx * 0.07, a.y + uy * (m - 0.04) + ny * 0.07),
               (a.x + ux * (m + 0.04) - nx * 0.07, a.y + uy * (m + 0.04) - ny * 0.07),
               (a.x + ux * (m + 0.12), a.y + uy * (m + 0.12)), (b.x, b.y)]
        c.pline(_T(pts, dx, dy), L_ESC)
    # línea de huella con flecha
    if flecha in ('SUBE', 'SUBE_DET'):
        s0 = 0.12
        s1 = (E.S_W[nb] + E.S_W[nb + 1]) / 2 - 0.08 if flecha == 'SUBE' else E.P_W.length - 0.10
        pts = E.sub_path(E.P_W, s0, s1, 40)
        c.circle((pts[0][0] + dx, pts[0][1] + dy), 0.035, L_ESC)
        c.pline(_T(pts, dx, dy), L_ESC)
        _flecha(c, pts[-1], pts[-3], L_ESC, dx=dx, dy=dy)
    else:   # BAJA: de la llegada al arranque
        s0, s1 = E.P_W.length - 0.12, 0.22
        pts = E.sub_path(E.P_W, s0, s1, 120)
        c.circle((pts[0][0] + dx, pts[0][1] + dy), 0.035, L_ESC)
        c.pline(_T(pts, dx, dy), L_ESC)
        _flecha(c, pts[-1], pts[-3], L_ESC, dx=dx, dy=dy)


# ------------------------------------------------------------------- cortes
def corte(c, eje, h0, z0, guard):
    """Escalera vista en corte. eje='y' (corte A-A', horizontal = y+h0) o 'x' (corte B-B')."""
    k = 1 if eje == 'y' else 0
    for i in range(1, E.N_HU + 1):
        poly = E.huella_poly(i)
        u0 = min(p[k] for p in poly) + h0
        u1 = max(p[k] for p in poly) + h0
        zt = z0 + E.z_huella(i)
        c.rect(u0, zt - E.E_HUELLA, u1, zt, L_ESC)
    zan = E.zanca()
    top = [(p[k] + h0, z0 + p[2]) for p in zan]
    bot = [(p[k] + h0, z0 + p[3]) for p in zan]
    c.pline(top, L_ESC)
    c.pline(bot, L_ESC)
    c.line(top[0], bot[0], L_ESC)
    c.line(top[-1], bot[-1], L_ESC)
    bar = [(p[k] + h0, z0 + p[2]) for p in E.barandal()]
    c.pline(bar, L_BAR)
    # barandal del hueco en planta alta (N.P.T. +3.00, h = 0.90)
    zp, zr = z0 + 3.00, z0 + 3.90
    if len(guard) == 2:
        c.line((guard[0], zp), (guard[0], zr), L_BAR)
        c.line((guard[0], zr), (guard[1], zr), L_BAR)
        c.line((guard[1], zp), (guard[1], zr), L_BAR)
    else:
        c.line((guard[0], zp), (guard[0], zr), L_BAR)
        c.line((guard[0], zr), (guard[2], zr), L_BAR)


# --------------------------------------------------------------- isométrico
def _iso_proj():
    cam = (1.0, -1.0, 1.0)
    n = math.sqrt(3)
    cvec = (cam[0] / n, cam[1] / n, cam[2] / n)
    r = (1 / math.sqrt(2), 1 / math.sqrt(2), 0.0)
    up = (-1 / math.sqrt(6), 1 / math.sqrt(6), 2 / math.sqrt(6))

    def pr(p):
        x, y, z = p
        return (x * r[0] + y * r[1] + z * r[2], x * up[0] + y * up[1] + z * up[2],
                x * cvec[0] + y * cvec[1] + z * cvec[2])
    return pr


def isometrico(c, ox, oy, esc=1.0):
    """Isométrico con eliminación de líneas ocultas (huellas, zanca, losa de llegada y barandal)."""
    pr = _iso_proj()
    prims = []   # (profundidad, polígono 3d de oclusión, [polilíneas 3d a dibujar])

    def add(poly3, edges3):
        d = sum(pr(p)[2] for p in poly3) / len(poly3)
        prims.append((d, poly3, edges3))

    def quad(a, b, za0, za1, zb0, zb1):
        return [(a[0], a[1], za0), (b[0], b[1], zb0), (b[0], b[1], zb1), (a[0], a[1], za1)]
    for i in range(1, E.N_HU + 1):
        zt = E.z_huella(i)
        zb = zt - E.E_HUELLA
        a_in, a_out = E.NARIZ[i - 1]
        b_in, b_out = E.NARIZ[i]
        inner = E.sub_path(E.P_IN, E.S_IN[i - 1], E.S_IN[i], 12)
        outer = E.sub_path(E.P_OUT, E.S_OUT[i - 1], E.S_OUT[i], 16)
        top = inner + outer[::-1]
        add([(x, y, zt) for x, y in top], [[(x, y, zt) for x, y in top + [top[0]]]])
        add([(x, y, zb) for x, y in top], [[(x, y, zb) for x, y in top + [top[0]]]])
        for p, q in ((a_in, a_out), (b_in, b_out)):
            qd = quad(p, q, zb, zt, zb, zt)
            add(qd, [qd + [qd[0]]])
        for path in (inner, outer):
            for p, q in zip(path, path[1:]):
                add(quad(p, q, zb, zt, zb, zt), [])
    zan = E.zanca(96)
    for k, (p, q) in enumerate(zip(zan, zan[1:])):
        poly = [(p[0], p[1], p[2]), (q[0], q[1], q[2]), (q[0], q[1], q[3]), (p[0], p[1], p[3])]
        ed = [[poly[0], poly[1]], [poly[3], poly[2]]]
        if k == 0:
            ed.append([poly[0], poly[3]])
        if k == len(zan) - 2:
            ed.append([poly[1], poly[2]])
        add(poly, ed)
    # losa de llegada (descanso N.P.T. +3.00)
    lx0, lx1, ly0, ly1 = 37.43, E.XC, 8.75, 9.75
    for z in (3.0, 2.75):
        f = [(lx0, ly0, z), (lx1, ly0, z), (lx1, ly1, z), (lx0, ly1, z)]
        add(f, [f + [f[0]]])
    for f in ([(lx1, ly0, 3.0), (lx1, ly1, 3.0), (lx1, ly1, 2.75), (lx1, ly0, 2.75)],
              [(lx0, ly0, 3.0), (lx1, ly0, 3.0), (lx1, ly0, 2.75), (lx0, ly0, 2.75)]):
        add(f, [f + [f[0]]])
    prims.sort(key=lambda f: -f[0])     # la más cercana primero
    occ = None
    lines = []
    for d, poly3, edges3 in prims:
        p2 = [(pr(p)[0], pr(p)[1]) for p in poly3]
        for e3 in edges3:
            e2 = LineString([(pr(p)[0], pr(p)[1]) for p in e3])
            lines.append(e2 if occ is None else e2.difference(occ))
        poly = Polygon(p2)
        if not poly.is_valid:
            poly = poly.buffer(0)
        if poly.is_empty or poly.area < 1e-8:
            continue
        occ = poly if occ is None else unary_union([occ, poly])
    # barandal de cristal: pasamanos + juntas de vidrio (se dibujan completos)
    bar = [pr(p) for p in E.barandal(96)]
    lines.append(LineString([(p[0], p[1]) for p in bar]))
    for i in range(1, E.N_HU + 1, 3):
        q = E.P_BAR.interpolate(E.S_BAR[i - 1])
        zt = E.z_huella(i)
        a, b = pr((q.x, q.y, zt)), pr((q.x, q.y, zt + 0.90))
        lines.append(LineString([(a[0], a[1]), (b[0], b[1])]))
    allx = []
    for g in lines:
        for part in ([g] if g.geom_type == 'LineString' else list(getattr(g, 'geoms', []))):
            if part.is_empty or part.length < 0.004:
                continue
            allx.append(part)
    minx = min(p.bounds[0] for p in allx)
    miny = min(p.bounds[1] for p in allx)
    for part in allx:
        pts = [(ox + (x - minx) * esc, oy + (y - miny) * esc) for x, y in part.coords]
        c.pline(pts, L_ESC)
    maxx = max(p.bounds[2] for p in allx)
    maxy = max(p.bounds[3] for p in allx)
    return (ox, oy, ox + (maxx - minx) * esc, oy + (maxy - miny) * esc)


# ------------------------------------------------------------ lámina detalle
DIM_DET = {'dimtxt': 0.08, 'dimasz': 0.06, 'dimexe': 0.06, 'dimexo': 0.04, 'dimgap': 0.03}


def _nivel(c, x, y, texto, h=0.08, lado=1):
    """Símbolo de nivel (triángulo relleno + línea + texto)."""
    c.line((x - 0.10, y), (x + 1.05 * lado, y), '016_Niveles')
    tri = [(x, y), (x - 0.06, y + 0.10), (x + 0.06, y + 0.10)]
    c.pline(tri, '016_Niveles', closed=True)
    hat = c.msp.add_hatch(color=256, dxfattribs={'layer': '016_Niveles'})
    hat.paths.add_polyline_path(tri, is_closed=True)
    c.mtext(texto, (x + 0.12, y + 0.02), h, layer='016_Niveles', attach=7)


def _tabla(c, x0, y_top, cols, header, rows, h_txt=0.065, rh=0.205, hh=0.42):
    """Tabla simple con líneas en capa 010_Notas."""
    lay = '010_Notas'
    W = sum(cols)
    n = len(rows)
    y_bot = y_top - hh - n * rh
    c.rect(x0, y_bot, x0 + W, y_top, lay)
    c.line((x0, y_top - hh), (x0 + W, y_top - hh), lay)
    for k in range(1, n):
        y = y_top - hh - k * rh
        c.line((x0, y), (x0 + W, y), lay)
    x = x0
    for w in cols[:-1]:
        x += w
        c.line((x, y_top), (x, y_bot), lay)
    x = x0
    for w, t in zip(cols, header):
        c.mtext(t, (x + w / 2, y_top - hh / 2), h_txt * 0.92, layer=lay, attach=5, width=w - 0.04)
        x += w
    for k, r in enumerate(rows):
        y = y_top - hh - (k + 0.5) * rh
        x = x0
        for w, t in zip(cols, r):
            c.mtext(t, (x + w / 2, y), h_txt, layer=lay, attach=5)
            x += w
    return y_bot


def lamina_detalle(c):
    R = E.resumen()
    # --- título general
    c.mtext('ESCALERA CURVA — DETALLE (RECALCULADA: SE ELIMINA EL PRIMER ESCALÓN)', (115.2, 12.95), 0.20,
            style='OpenSans-SemiBold', attach=5)
    # ================= PLANTA =================
    dx, dy = 70.12, 1.70
    # muros (proyección) y descanso
    c.line((37.43 + dx, 6.9 + dy), (40.71 + dx, 6.9 + dy), '015_Proyeccion')
    c.line((37.43 + dx, 9.75 + dy), (40.71 + dx, 9.75 + dy), '015_Proyeccion')
    c.line((40.71 + dx, 6.9 + dy), (40.71 + dx, 9.75 + dy), '015_Proyeccion')
    c.rect(37.43 + dx, 8.75 + dy, E.XC + dx, 9.75 + dy, L_ESC)
    c.mtext('DESCANSO\\PN.P.T. +3.00', (38.35 + dx, 9.25 + dy), 0.075, attach=5)
    planta(c, dx, dy, corte=False, flecha='SUBE_DET')
    # escalón eliminado (referencia)
    c.line((E.X0_ORIG + dx, 6.9 + dy), (E.X0_ORIG + dx, 7.9 + dy), L_ESC, linetype='VACIO')
    c.mtext('ESCALÓN\\PELIMINADO', (E.X0_ORIG + dx - 0.05, 8.25 + dy), 0.055, attach=6)
    c.line((E.X0_ORIG + dx - 0.04, 8.2 + dy), (E.X0_ORIG + dx + 0.10, 7.75 + dy), L_TXT)
    # numeración de huellas (en la línea de huella)
    p_num = E.path_line(1.20)
    s_num = [E._s_on(p_num, sg) for sg in E.NARIZ]
    for i in range(1, E.N_HU + 1):
        p = p_num.interpolate((s_num[i - 1] + s_num[i]) / 2)
        c.mtext(str(i), (p.x + dx, p.y + dy), 0.065, attach=5)
    # centro y radios
    c.circle((E.XC + dx, E.YC + dy), 0.02, L_ESC)
    c.mtext('R ext. 1.425\\PR int. 0.425\\Pr huella 0.925', (E.XC - 0.08 + dx, E.YC + dy), 0.05, attach=6)
    # cotas
    c.hdim(37.43 + dx, 40.71 + dx, 9.75 + dy, 10.05 + dy, ovr=DIM_DET)
    c.hdim(E.X0 + dx, E.XC + dx, 6.9 + dy, 6.55 + dy, ovr=DIM_DET)
    c.hdim(E.X0_ORIG + dx, E.X0 + dx, 6.9 + dy, 6.30 + dy, ovr=DIM_DET)
    c.vdim(6.9 + dy, 7.9 + dy, E.X0 + dx, E.X0 - 0.45 + dx, ovr=DIM_DET)
    c.vdim(6.9 + dy, 9.75 + dy, 40.71 + dx, 41.0 + dx, ovr=DIM_DET)
    c.mtext('PLANTA', (39.07 + dx, 7.62), 0.14, attach=5)
    c.mtext('ESC. 1:20', (39.07 + dx, 7.42), 0.07, attach=5)
    # ================= ALZADO =================
    dxe, z0e = 75.17, 7.60
    c.line((37.20 + dxe, z0e), (40.95 + dxe, z0e), '015_Proyeccion')
    c.rect(36.90 + dxe, z0e + 2.75, E.XC + dxe, z0e + 3.00, '013_Cortes')
    c.line((40.71 + dxe, z0e), (40.71 + dxe, z0e + 4.10), '015_Proyeccion')
    corte(c, 'x', dxe, z0e, (37.43 + dxe, None, E.XC + dxe))
    _nivel(c, 41.25 + dxe, z0e, 'N.P.T. ±0.00')
    _nivel(c, 41.25 + dxe, z0e + 3.00, 'N.P.T. +3.00')
    c.vdim(z0e, z0e + 3.0, 40.71 + dxe, 41.10 + dxe, ovr=DIM_DET)
    c.mtext('17 PERALTES DE 0.1765', (41.40 + dxe, z0e + 1.5), 0.06, attach=4, rot=90)
    c.mtext('PASAMANOS h = 0.90\\PSOBRE NARIZ', (38.50 + dxe, z0e + 3.55), 0.06, attach=7)
    c.mtext('ALZADO (VISTA HACIA EJE 15)', (39.07 + dxe, 7.25), 0.14, attach=5)
    c.mtext('ESC. 1:20', (39.07 + dxe, 7.05), 0.07, attach=5)
    # ================= ISOMÉTRICO =================
    bx = isometrico(c, 117.55, 7.45, 1.0)
    cx = (bx[0] + bx[2]) / 2
    c.mtext('ISOMÉTRICO (SIN ESCALA)', (cx, 7.25), 0.14, attach=5)
    # ================= TABLAS =================
    rows = []
    for (n, z, hw, h4, hi, ho) in E.tabla():
        rows.append([str(n), '%.3f' % z, '%.3f' % hw, '%.3f' % h4, '%.3f' % hi, '%.3f' % ho])
    rows.append(['17', '3.000', 'LLEGADA', 'N.P.T.', '+3.00', '—'])
    c.mtext('TABLA 1. GEOMETRÍA POR PELDAÑO (m)', (106.95, 6.62), 0.085, style='OpenSans-SemiBold', attach=7)
    _tabla(c, 106.95, 6.55, [0.45, 0.80, 1.10, 1.10, 0.90, 0.90],
           ['PE-\\PRALTE', 'NIVEL\\PHUELLA', 'HUELLA EN\\PLÍNEA DE HUELLA', 'HUELLA A 0.40\\PDEL BARANDAL INT.',
            'HUELLA\\PBORDE INT.', 'HUELLA\\PBORDE EXT.'], rows)
    rows2 = []
    for i, (p1, p2) in enumerate(E.NARIZ):
        rows2.append([str(i), '%.3f' % (p1[0] - E.XC), '%.3f' % (p1[1] - E.YC),
                      '%.3f' % (p2[0] - E.XC), '%.3f' % (p2[1] - E.YC)])
    c.mtext('TABLA 2. TRAZO DE NARICES (m, ORIGEN EN EL CENTRO DEL GIRO)', (112.45, 6.62), 0.085,
            style='OpenSans-SemiBold', attach=7)
    _tabla(c, 112.45, 6.55, [0.55, 0.80, 0.80, 0.80, 0.80],
           ['NARIZ', 'INT. X', 'INT. Y', 'EXT. X', 'EXT. Y'], rows2)
    # ================= NOTAS =================
    notas = (
        'NOTAS DE LA ESCALERA\\P'
        '1. SE ELIMINA EL PRIMER ESCALÓN: EL ARRANQUE SE RECORRE 0.29 m (DE 37.551 A 37.841 EN EJES DE PLANTA).\\P'
        '2. DESNIVEL 3.00 m = 17 PERALTES DE %.4f m (MÁX. 0.18 ✓).\\P'
        '3. 16 HUELLAS COMPENSADAS (BALANCEADAS) EN TODO EL TRAMO: EN LÍNEA DE HUELLA (r = 0.925) %.3f A %.3f m; '
        'A 0.40 m DEL BARANDAL INTERIOR %.3f A %.3f m (MÍN. 0.25 ✓).\\P'
        '4. 2P + H = %.3f A %.3f m (ENTRE 0.61 Y 0.65 ✓).\\P'
        '5. CON EL MISMO HUECO, 11 HUELLAS RADIALES DARÍAN 0.236 m A 0.40 DEL BARANDAL (NO CUMPLE); '
        'CON 16 PERALTES EL PERALTE SERÍA 0.1875 m (NO CUMPLE). POR ESO SE COMPENSA.\\P'
        '6. TRAZO: CADA NARIZ UNE EL PUNTO DE DIVISIÓN DEL BORDE EXTERIOR CON EL DE LA LÍNEA A 0.40 m (TABLA 2).\\P'
        '7. ANCHO 1.00 m; R EXT. 1.425; R INT. 0.425. HUELLAS e = 0.06 SOBRE ZANCA CURVA EXTERIOR; '
        'BARANDAL DE CRISTAL TEMPLADO h = 0.90.\\P'
        '8. VERIFICAR CON EL D.R.O.: EL REGLAMENTO DE MORELIA PIDE HUELLA 0.28 Y MÁX. 15 PERALTES ENTRE DESCANSOS '
        '(LA ESCALERA ORIGINAL YA TENÍA 17).'
    ) % (R['peralte'], R['hw_min'], R['hw_max'], R['h40_min'], R['h40_max'], R['2p_h_min'], R['2p_h_max'])
    c.mtext(notas, (117.30, 6.62), 0.062, layer='010_Notas', attach=1, width=6.45)


def lamina_vp(c):
    """Viewport de la lámina 03_ESCALERA a 1:20 sobre toda la hoja."""
    lay = c.doc.layouts.get('03_ESCALERA')
    for e in lay.query('VIEWPORT'):
        if e.dxf.id != 1:
            e.dxf.center = (450.0, 300.0)
            e.dxf.width = 870.0
            e.dxf.height = 570.0
            e.dxf.view_center_point = (115.2, 7.6)
            e.dxf.view_height = 11.4
    for m in lay.query('MTEXT'):
        if m.plain_text().startswith('ESCALA:'):
            m.text = 'ESCALA: 1:20'
