"""Genera 001_PLANO_CROCHA_R1.6.dxf a partir de la R1.5 con todas las correcciones.

Uso:  python3 scripts/build_dxf.py
Cada paso está en su propia función para poder revisarlo por separado.
"""
import math
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from cad import Cad, pts_of  # noqa: E402
import bloques  # noqa: E402
import dibujo_escalera as DE  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'originales', '001_PLANO_CROCHA_R1.5.dxf')
OUT = os.path.join(ROOT, '001_PLANO_CROCHA_R1.6.dxf')

# Desplazamientos de cada dibujo respecto a la planta baja
OFF_PA = 30.0          # planta alta: x + 30
OFF_AZ = 60.0          # azotea: x + 60
OFF_AA = 42.5          # corte A-A': horizontal = y_planta + 42.5
OFF_BB = 49.565        # corte B-B': horizontal = x_planta + 49.565

# Ejes corregidos a centro de muro
X_H_OLD, X_H_NEW = 39.3604, 39.185      # eje H (muros reales en 39.185)
Y_10_OLD, Y_10_NEW = 13.821, 13.721     # eje 10 (muros reales en 13.721)


def mover_hatch(h, dx, dy):
    for path in h.paths:
        if hasattr(path, 'vertices'):
            path.vertices = [(v[0] + dx, v[1] + dy, v[2] if len(v) > 2 else 0) for v in path.vertices]
        else:
            for ed in path.edges:
                for attr in ('start', 'end', 'center'):
                    if hasattr(ed, attr):
                        p = getattr(ed, attr)
                        setattr(ed, attr, (p[0] + dx, p[1] + dy))


def near(a, b, t=2e-3):
    return abs(a - b) <= t


# ---------------------------------------------------------------- 1. EJES
def paso_ejes(c: Cad):
    msp = c.msp
    # líneas de eje
    for ln in msp.query('LINE[layer=="002_EjesVerticales" | layer=="003_EjesHorizontales"]'):
        s, e = ln.dxf.start, ln.dxf.end
        for xo, xn in ((X_H_OLD, X_H_NEW), (X_H_OLD + OFF_PA, X_H_NEW + OFF_PA),
                       (X_H_OLD + OFF_BB, X_H_NEW + OFF_BB)):
            if near(s.x, xo) and near(e.x, xo):
                ln.dxf.start = (xn, s.y, 0)
                ln.dxf.end = (xn, e.y, 0)
        if near(s.y, Y_10_OLD) and near(e.y, Y_10_OLD) and s.x > 55:
            ln.dxf.start = (s.x, Y_10_NEW, 0)
            ln.dxf.end = (e.x, Y_10_NEW, 0)
    # globos de eje
    for ins in msp.query('INSERT[name=="EJE"]'):
        lab = {a.dxf.tag: a.dxf.text for a in ins.attribs}.get('EJE')
        p = ins.dxf.insert
        if lab == 'H':
            for xo, xn in ((X_H_OLD, X_H_NEW), (X_H_OLD + OFF_PA, X_H_NEW + OFF_PA),
                           (X_H_OLD + OFF_BB, X_H_NEW + OFF_BB)):
                if near(p.x, xo, 3e-3):
                    ins.translate(xn - p.x, 0, 0)
        if lab == '10' and near(p.y, Y_10_OLD):
            ins.translate(0, Y_10_NEW - p.y, 0)
    # polilínea suelta en capa de ejes (duplicaba el eje 14 en planta alta)
    for pl in list(msp.query('LWPOLYLINE[layer=="002_EjesVerticales"]')):
        msp.delete_entity(pl)
    # cotas que tocan el eje H (incluye las dos que estaban volteadas en planta alta)
    est = {0.0: c.ovr_de('1378'), OFF_PA: c.ovr_de('144A'), OFF_BB: c.ovr_de('494F')}
    c.kill('1369', '1378', '144A', '40B7', '494F', '495E')
    for off in (0.0, OFF_PA):
        c.hdim(37.858 + off, X_H_NEW + off, 24.321, 25.321, base=est[off])
        c.hdim(X_H_NEW + off, 40.785 + off, 24.321, 25.321, base=est[off])
    c.hdim(37.858 + OFF_BB, X_H_NEW + OFF_BB, -12.82, -13.12, base=est[OFF_BB])
    c.hdim(X_H_NEW + OFF_BB, 40.785 + OFF_BB, -12.82, -13.12, base=est[OFF_BB])


# ------------------------------------------- 2. RECÁMARA PRINCIPAL (PA)
def paso_recamara(c: Cad):
    """Muro baño/recámara: línea de centro, vano de puerta 0.75 (hoja 0.70) y TV reubicada."""
    msp = c.msp
    y_old, y_new = 4.896, 4.996           # jamba inferior del vano del baño
    # prolongar las caras del muro (x 66.187 / 66.337) y su remate
    for h in ('FEA', 'FE7'):
        pl = c.e(h)
        pts = [(p[0], p[1]) for p in pl.get_points()]
        pts = [(x, y_new if near(y, y_old) else y) for x, y in pts]
        pl.set_points(pts)
    rem = c.e('FF5')
    rem.dxf.start = (rem.dxf.start.x, y_new, 0)
    rem.dxf.end = (rem.dxf.end.x, y_new, 0)
    # línea de centro faltante (la "rayita en medio")
    c.line((66.262, 2.871), (66.262, y_new), '001_Muros')
    # segundo muro sin línea de centro detectado en la revisión: eje D, planta alta
    c.line((66.36, 15.322), (66.36, 19.247), '001_Muros')
    # mueble de TV fuera del vano: queda sobre el muro, de y 3.15 a 4.75
    tv = [i for i in msp.query('INSERT[name=="MOB_TV_160x45"]')][0]
    tv.dxf.insert = (tv.dxf.insert.x, 3.95, 0)


# ------------------------------------------------------------- 3. COMEDOR
def paso_comedor(c: Cad):
    msp = c.msp
    for ins in list(msp.query('INSERT[name=="MOB_COMEDOR8_200x330"]')):
        msp.delete_entity(ins)
    name = bloques.comedor15(c.doc)
    # mesa 1.10 x 4.20 (y 12.90 a 17.10) + cabecera al norte; 1.3 m libres a muro y a bancos de isla
    msp.add_blockref(name, (34.15, 15.00), dxfattribs={'layer': '006_Mobiliario'})
    # etiqueta COMEDOR fuera de los bancos de la isla
    lab = [m for m in msp.query('MTEXT') if m.plain_text().strip() == 'COMEDOR' and m.dxf.insert.y > 0][0]
    lab.set_location((34.15, 12.48, 0), attachment_point=5)
    # cota 4.78 se baja de la mesa: va entre la cabecera y la alberca
    est = c.ovr_de('4AB6')
    c.kill('4AB6')
    c.hdim(32.010, 36.785, 18.05, 18.05, base=est)


# ---------------------------------------------------------------- 4. SALA
def paso_sala(c: Cad):
    msp = c.msp
    for nm in ('MOB_SOFA_230x90', 'MOB_MESA_C_120x60', 'MOB_SILLON_80x80'):
        for ins in list(msp.query(f'INSERT[name=="{nm}"]')):
            if ins.dxf.insert.x < 45:          # sólo planta baja
                msp.delete_entity(ins)
    n = bloques.sala(c.doc)
    lay = {'layer': '006_Mobiliario'}
    msp.add_blockref(n['tapete'], (34.60, 9.50), dxfattribs=lay)
    msp.add_blockref(n['seccional'], (36.25, 11.10), dxfattribs=lay)
    msp.add_blockref(n['mesas'], (34.35, 9.30), dxfattribs=lay)
    msp.add_blockref(n['lounge'], (33.15, 8.60), dxfattribs=lay)
    msp.add_blockref(n['lounge'], (33.15, 9.90), dxfattribs=lay)
    msp.add_blockref(n['lateral'], (33.08, 9.25), dxfattribs=lay)
    msp.add_blockref(n['lampara'], (36.05, 7.85), dxfattribs=dict(lay, rotation=270))
    # rótulo VACÍO de planta baja: se recorre para no quedar sobre los sillones
    lab = [m for m in msp.query('MTEXT') if m.plain_text().strip() == 'VACÍO' and m.dxf.insert.x < 40][0]
    lab.set_location((33.75, 7.92, 0), attachment_point=7)


# --------------------------------------------------------------- 5. AUTOS
def paso_autos(c: Cad):
    msp, doc = c.msp, c.doc
    bloques.auto_planta(doc, 'MOB_AUTO_185x460', 4.60, 1.85)
    bloques.auto_planta(doc, 'MOB_AUTO_185x455', 4.55, 1.85)
    bloques.pickup_planta(doc, 'MOB_PICKUP_203x589', 5.89, 2.03)
    # Corte A-A': se borran los autos viejos (entourage en la cochera) y se dibujan perfiles nuevos
    from ezdxf import bbox
    for e in list(msp.query('*[layer=="019_Entourage"]')):
        b = bbox.extents([e])
        if b.has_data and b.extmin.x >= 42.79 and b.extmax.x <= 47.70 and b.extmin.y >= -12.56 \
                and b.extmax.y <= -10.50:
            msp.delete_entity(e)
    z0 = -12.50                       # N.P.T. +-0.00 en el corte
    x_suv = 42.5 + 0.35               # defensa delantera del auto 1 (y planta 0.35)
    x_pk = 42.5 + 0.355               # defensa delantera de la pick-up
    x_muro = 47.646                   # muro cortado del eje 4 (oculta lo que queda detrás)
    geo_s, out_s, _ = bloques.suv_alzado()
    geo_p, out_p, _ = bloques.pickup_alzado()
    from shapely.geometry import Polygon, LineString, box
    from shapely.ops import unary_union
    from shapely.geometry import Point
    suv_poly = Polygon([(x_suv + x, z0 + z) for x, z in out_s]).buffer(0.005)
    llantas = [Point(x_suv + g[1][0], z0 + g[1][1]).buffer(g[2] + 0.005) for g in geo_s
               if g[0] == 'circle' and g[2] > 0.3]
    tapa = unary_union([suv_poly, box(x_muro, z0 - 1, x_muro + 10, z0 + 5)] + llantas)
    lay = {'layer': '019_Entourage'}

    def draw(geo, x0, clip=None):
        for g in geo:
            if g[0] == 'circle':
                (cx, cz), r = g[1], g[2]
                pts = [(x0 + cx + r * math.cos(2 * math.pi * i / 48), z0 + cz + r * math.sin(2 * math.pi * i / 48))
                       for i in range(49)]
                geom = LineString(pts)
            elif g[0] == 'line':
                geom = LineString([(x0 + g[1][0], z0 + g[1][1]), (x0 + g[2][0], z0 + g[2][1])])
            else:
                pts = [(x0 + x, z0 + z) for x, z in g[1]]
                if g[2]:
                    pts.append(pts[0])
                geom = LineString(pts)
            if clip is not None:
                geom = geom.difference(clip)
            parts = [geom] if geom.geom_type == 'LineString' else list(getattr(geom, 'geoms', []))
            for p in parts:
                if p.is_empty or p.length < 0.01:
                    continue
                msp.add_lwpolyline(list(p.coords), dxfattribs=lay)
    draw(geo_p, x_pk, clip=tapa)     # pick-up atrás (oculta por el SUV y por el muro cortado)
    draw(geo_s, x_suv)               # SUV al frente


# ------------------------------------------------------------- 6. ESCALERA
def paso_escalera(c: Cad):
    DE.borrar_original(c)
    DE.planta(c, 0.0, 0.0, corte=True, flecha='SUBE')
    DE.planta(c, OFF_PA, 0.0, corte=False, flecha='BAJA')
    DE.corte(c, 'y', OFF_AA, -12.50, (49.38, 51.27))
    DE.corte(c, 'x', OFF_BB, -12.50, (86.975, None, 88.85))
    DE.lamina_detalle(c)
    DE.lamina_vp(c)


# ------------------------------------------------- 7. TEXTOS QUE SE ENCIMAN
def paso_textos(c: Cad):
    def mt(h):
        return c.e(h)
    # PB: "MÁQUINAS Y SERVICIO" en una línea, fuera de la cota 3.77 y de la cota vertical 1.85
    # etiquetas de varias palabras como TEXT de un renglón (algunos visores parten el MTEXT)
    from ezdxf.enums import TextEntityAlignment

    def texto_1r(h_old, txt, xy):
        old = mt(h_old)
        a = {'layer': old.dxf.layer, 'style': old.dxf.style, 'height': old.dxf.char_height}
        c.msp.add_text(txt, dxfattribs=a).set_placement(xy, align=TextEntityAlignment.MIDDLE_CENTER)
        c.msp.delete_entity(old)
    texto_1r('89F', 'MÁQUINAS Y SERVICIO', (39.20, 16.95))
    # PB: RECÁMARA fuera de la cota 3.15
    mt('956').set_location((36.98, 20.95, 0), attachment_point=1)
    # PA: RECÁMARA (eje 9-10) y su cota 3.05 recorrida al hueco entre escritorio y cama
    mt('109D').set_location((67.50, 12.75, 0), attachment_point=1)
    est = c.ovr_de('4BD3')
    c.kill('4BD3')
    c.vdim(10.600, 13.647, 68.60, 68.60, base=est)
    # PA: ÁREA DE LAVADO en una línea, arriba de la cota 3.28
    texto_1r('10D4', 'ÁREA DE LAVADO', (63.65, 13.25))
    # PA: la recámara del eje D-J (11-13) tenía su rótulo repetido; se deja el del centro del cuarto
    c.kill('113A')
    mt('62B1').set_location((67.35, 17.40, 0), attachment_point=1)
    # PA: baño eje 10-11 (cotas 2.88 y 1.45 encimadas)
    est = c.ovr_de('4C1E')
    c.kill('4C1E', '4C2D')
    c.hdim(67.832, 70.710, 14.92, 14.92, base=est)
    c.vdim(13.797, 15.247, 69.60, 69.60, base=est)
    mt('10F8').set_location((68.70, 14.60, 0), attachment_point=7)
    # cadenas de ejes con claros de 0.50 m: el texto de los claros chicos va abajo de la línea
    est = c.ovr_de('1486')
    c.kill('1486', '1459')
    c.hdim(64.8604, 65.3624, 24.321, 25.321, ovr={'dimtad': 4}, base=est)
    c.hdim(67.3574, 67.8584, 24.321, 25.321, ovr={'dimtad': 4}, base=est)
    est = c.ovr_de('4904')
    c.kill('4904', '4922', '4931', '4940')
    c.hdim(84.4254, 84.9274, -12.82, -13.12, ovr={'dimtad': 4}, base=est)
    c.hdim(85.9264, 86.4254, -12.82, -13.12, ovr={'dimtad': 4}, base=est)
    c.hdim(86.4254, 86.9224, -12.82, -13.12, ovr={'dimtad': 1}, base=est)
    c.hdim(86.9224, 87.4234, -12.82, -13.12, ovr={'dimtad': 4}, base=est)


# ------------------------------------------------------------- 8. NOTAS
NOTAS = [
    '1. COTAS Y NIVELES EN METROS; NIVELES REFERIDOS AL N.P.T. DE PLANTA BAJA (±0.00).',
    '2. ENTREPISO: LOSA ALIGERADA DE CASETONES 60×60×20 CON NERVADURAS DE 10 cm',
    '    Y CAPA DE 5 cm (h = 0.25); MACIZADA EN APOYOS.',
    '3. AZOTEA: LOSA ALIGERADA RETICULAR, MISMO SISTEMA QUE EL ENTREPISO (PERALTE',
    '    FINAL SEGÚN CÁLCULO ESTRUCTURAL). SOBRE EL VACÍO: VIGAS DE CONCRETO 15×40',
    '    @ 0.75 CON CRISTAL TEMPLADO.',
    '4. ALTURAS PROPUESTAS: ENTREPISO 3.00; PUERTAS 2.10; VENTANAS 0.90 A 2.10;',
    '    PRETIL AZOTEA 0.60; PRETIL TERRAZAS 1.00.',
    '5. PUERTAS: EL VANO INCLUYE EL ESPESOR DE LA HOJA (EJ. PUERTA DE 0.75 = HOJA 0.70 + 0.05).',
    '6. ESCALERA CURVA SIN EL PRIMER ESCALÓN ORIGINAL: 17 PERALTES DE 0.1765 Y 16 HUELLAS',
    '    EN U: 5 RECTAS DE 0.29 + 11 RADIALES DE 0.264 EN LÍNEA DE HUELLA (MÍN. 0.25).',
    '    VER LÁMINA A-03.',
    '7. FIGURAS HUMANAS, VEGETACIÓN, VEHÍCULOS Y MOBILIARIO: REFERENCIA DE ESCALA.',
]
REFS = [
    '• COLCHONES (MEDIDAS MÉXICO): INDIVIDUAL 1.00×1.90, MATRIMONIAL 1.35×1.90, QUEEN',
    '   1.50×1.90, KING 2.00×1.90 — MERCADO LIBRE MÉXICO, «MEDIDAS DE COLCHONES».',
    '• ESCALERAS (PERALTE MÁX. 0.18, HUELLA MÍN. 0.25 MEDIDA A 0.40 DEL BARANDAL INTERIOR',
    '   EN COMPENSADAS): NTC PARA EL PROYECTO ARQUITECTÓNICO, RCDF. REGLAMENTO DE',
    '   CONSTRUCCIÓN DE MORELIA: HUELLA MÍN. 0.28, 2P + H ENTRE 0.61 Y 0.65, MÁX. 15',
    '   PERALTES ENTRE DESCANSOS — VERIFICAR CON EL D.R.O.',
    '• MOBILIARIO, BAÑOS Y COCINA: NEUFERT, «ARTE DE PROYECTAR EN ARQUITECTURA».',
    '• COMEDOR: 0.60 m POR COMENSAL Y 0.90 m LIBRES DETRÁS DE LAS SILLAS (NEUFERT).',
    '• PICKUP DOBLE CABINA DE TAMAÑO COMPLETO ≈ 5.89 × 2.03 × 1.96 m — FICHA DEL FABRICANTE.',
]


def paso_notas(c: Cad):
    y = -5.20
    c.e('4E19').set_location((95.6, y, 0), attachment_point=1)
    body = c.e('4E1A')
    body.text = '\\P'.join(NOTAS)
    body.dxf.width = 11.2
    body.set_location((95.6, y - 0.40, 0), attachment_point=1)
    y = y - 0.40 - len(NOTAS) * 0.14 * 1.70 - 0.30
    c.e('4E1B').set_location((95.6, y, 0), attachment_point=1)
    refs = c.e('4E1C')
    refs.text = '\\P'.join(REFS)
    refs.dxf.width = 11.2
    refs.set_location((95.6, y - 0.40, 0), attachment_point=1)
    y = y - 0.40 - len(REFS) * 0.12 * 1.70 - 0.30
    tit = c.e('4E1D')
    dy = y - tit.dxf.insert.y
    tit.translate(0, dy, 0)
    for h in ('4E1E', '4E1F', '4E20', '4E21', '4E22', '4E23', '4E24', '4E25', '4E26', '4E27', '4E28',
              '4E29', '4E2A', '4E2B', '4E2C'):
        ent = c.e(h)
        if ent.dxftype() == 'HATCH':
            mover_hatch(ent, 0, dy)      # translate() de ezdxf reescala el patrón: se mueve sólo el contorno
        else:
            ent.translate(0, dy, 0)


# ------------------------------------------------------------ 9. LAYOUTS
REV_TXT = ('REVISIONES\\P'
           'R1.5  SEP-2026  EMISIÓN ANTERIOR.\\P'
           'R1.6  OCT-2026  COMEDOR PARA 15 PERSONAS; ESCALERA SIN EL PRIMER\\P'
           '                 ESCALÓN (17 PERALTES, HUELLAS EN U); EJES H Y 10\\P'
           '                 A CENTRO DE MURO; MURO, PUERTA Y MUEBLE EN RECÁMARA\\P'
           '                 PRINCIPAL; BLOQUES DE AUTOS Y SALA; COTAS Y NOTAS\\P'
           '                 SIN ENCIMARSE; DIBUJOS AMPLIADOS EN LÁMINAS.')


def _vp(lay, center, size, vcenter, vheight):
    v = lay.add_viewport(center=center, size=size, view_center_point=vcenter, view_height=vheight)
    v.dxf.layer = '024_Ventanas'
    v.dxf.status = 1 + sum(1 for e in lay.query('VIEWPORT') if e.dxf.id != 1)
    return v


def _membrete(lay, escala, rev_xy):
    for m in lay.query('MTEXT'):
        t = m.plain_text()
        x = m.dxf.insert.x
        if t.startswith('ESCALA'):
            m.text = escala
        elif t.startswith('FECHA'):
            m.text = 'FECHA: OCTUBRE 2026'
            m.set_location((x, 25.0, 0), attachment_point=4)
        elif t.startswith('DIBUJÓ'):
            m.set_location((x, 19.5, 0), attachment_point=4)
        elif t.startswith('ARCHIVO'):
            m.text = 'ARCHIVO: 001_PLANO_CROCHA_R1.6.dxf'
            m.set_location((x, 14.5, 0), attachment_point=4)
        elif t.startswith('CLAVE'):
            m.set_location((x, 25.0, 0), attachment_point=4)
        elif t.startswith('REVISÓ'):
            m.set_location((x, 19.5, 0), attachment_point=4)
    a = {'layer': '023_Membrete', 'style': 'OpenSans', 'char_height': 2.4}
    m = lay.add_mtext(REV_TXT, dxfattribs=a)
    m.set_location((rev_xy[0], rev_xy[1], 0), attachment_point=1)
    lay.add_lwpolyline([(rev_xy[0] - 4, rev_xy[1] + 4), (886, rev_xy[1] + 4), (886, rev_xy[1] - 34),
                        (rev_xy[0] - 4, rev_xy[1] - 34)], close=True, dxfattribs={'layer': '023_Membrete'})


def paso_layouts(c: Cad):
    doc = c.doc
    for name in ('01_PLANTAS', '02_CORTES'):
        lay = doc.layouts.get(name)
        for v in list(lay.query('VIEWPORT')):
            if v.dxf.id != 1:
                lay.delete_entity(v)
    # A-01: tres plantas a 1:75 + notas a 1:50
    lay = doc.layouts.get('01_PLANTAS')
    k = 1000 / 75
    w, h = 15.10, 31.90
    for i, xm in enumerate((36.36, 66.36, 96.36)):
        _vp(lay, (118.7 + i * (w * k + 8), 300.0), (w * k, h * k), (xm, 11.43), h)
    _vp(lay, (767.5, 490.0), (11.4 * 20, 9.2 * 20), (101.2, -9.70), 9.2)
    _membrete(lay, 'ESCALA: 1:75 / NOTAS 1:50', (660, 380))
    # A-02: cortes a 1:40 + notas a 1:50
    lay = doc.layouts.get('02_CORTES')
    k = 1000 / 40
    _vp(lay, (450.0, 447.5), (29.05 * k, 11.0 * k), (54.775, -10.05), 11.0)
    _vp(lay, (184.5, 157.0), (13.55 * k, 10.95 * k), (86.175, -10.075), 10.95)
    _vp(lay, (482.0, 207.0), (11.4 * 20, 9.2 * 20), (101.2, -9.70), 9.2)
    _membrete(lay, 'ESCALA: 1:40 / NOTAS 1:50', (660, 290))
    # A-03: escalera a 1:20 (viewport ya ajustado en paso_escalera)
    lay = doc.layouts.get('03_ESCALERA')
    _membrete(lay, 'ESCALA: 1:20', (660, 160))


# ---------------------------------------- 10. MOBILIARIO EN CORTE A-A' (PB)
def paso_corte_muebles(c: Cad):
    """Redibuja sala, comedor de 15, bancos, isla y barra del corte A-A' con líneas ocultas."""
    from ezdxf import bbox
    from shapely.geometry import Polygon, LineString, box
    from shapely.ops import unary_union
    msp = c.msp
    for e in list(msp.query('*[layer=="021_Muebles_Corte"]')):
        b = bbox.extents([e])
        if b.has_data and b.extmin.x >= 50.0 and b.extmax.x <= 60.3 and b.extmin.y >= -12.52 \
                and b.extmax.y <= -11.2:
            msp.delete_entity(e)
    z0 = -12.50
    X = lambda y: y + OFF_AA
    objs = []      # (polígono de oclusión, [polilíneas]) del más cercano al más lejano

    def obj(outline, extra=()):
        pts = [(X(y), z0 + z) for y, z in outline]
        lines = [pts + [pts[0]]] + [[(X(y), z0 + z) for y, z in ln] for ln in extra]
        objs.append((Polygon(pts).buffer(0), lines))
    # sillones lounge (respaldo envolvente visto por detrás)
    for yc in (8.60, 9.90):
        obj([(yc - 0.40, 0.0), (yc - 0.40, 0.62), (yc - 0.30, 0.74), (yc + 0.30, 0.74), (yc + 0.40, 0.62),
             (yc + 0.40, 0.0)], [[(yc - 0.40, 0.10), (yc + 0.40, 0.10)]])
    # mesa lateral
    obj([(9.07, 0.47), (9.43, 0.47), (9.43, 0.50), (9.07, 0.50)], [])
    obj([(9.235, 0.0), (9.265, 0.0), (9.265, 0.47), (9.235, 0.47)], [])
    # sillas poniente del comedor (se ve el respaldo)
    for yc in (13.20, 13.80, 14.40, 15.00, 15.60, 16.20, 16.80):
        obj([(yc - 0.225, 0.0), (yc - 0.225, 0.92), (yc + 0.225, 0.92), (yc + 0.225, 0.0),
             (yc + 0.19, 0.0), (yc + 0.19, 0.45), (yc - 0.19, 0.45), (yc - 0.19, 0.0)],
            [[(yc - 0.225, 0.45), (yc + 0.225, 0.45)], [(yc - 0.18, 0.60), (yc + 0.18, 0.60)]])
    # silla de cabecera (de perfil)
    obj([(17.10, 0.0), (17.10, 0.45), (17.48, 0.45), (17.48, 0.92), (17.55, 0.92), (17.55, 0.0),
         (17.52, 0.0), (17.52, 0.42), (17.13, 0.42), (17.13, 0.0)], [])
    # mesas de centro
    obj([(8.85, 0.36), (9.75, 0.36), (9.75, 0.40), (8.85, 0.40)], [])
    obj([(9.05, 0.0), (9.55, 0.0), (9.48, 0.36), (9.12, 0.36)], [])
    obj([(9.57, 0.41), (10.07, 0.41), (10.07, 0.45), (9.57, 0.45)], [])
    obj([(9.80, 0.0), (9.84, 0.0), (9.84, 0.41), (9.80, 0.41)], [])
    # mesa del comedor 1.10 x 4.20 (h = 0.76) con patas en extremos
    obj([(12.90, 0.72), (17.10, 0.72), (17.10, 0.76), (12.90, 0.76)], [])
    for y0 in (12.98, 16.96):
        obj([(y0, 0.0), (y0 + 0.06, 0.0), (y0 + 0.06, 0.72), (y0, 0.72)], [])
    # sofá seccional (brazo corto al frente, módulo largo atrás)
    obj([(10.15, 0.0), (10.15, 0.60), (11.10, 0.60), (11.10, 0.0)], [[(10.15, 0.12), (11.10, 0.12)]])
    obj([(8.10, 0.0), (8.10, 0.60), (8.30, 0.60), (8.30, 0.80), (11.10, 0.80), (11.10, 0.0)],
        [[(8.30, 0.42), (10.15, 0.42)], [(8.30, 0.12), (10.15, 0.12)], [(9.08, 0.12), (9.08, 0.42)],
         [(9.86, 0.12), (9.86, 0.42)]])
    # bancos de la isla (h asiento 0.65)
    for yc in (11.80, 12.57, 13.34):
        obj([(yc - 0.20, 0.62), (yc + 0.20, 0.62), (yc + 0.20, 0.66), (yc - 0.20, 0.66)], [])
        obj([(yc - 0.02, 0.05), (yc + 0.02, 0.05), (yc + 0.02, 0.62), (yc - 0.02, 0.62)],
            [[(yc - 0.14, 0.28), (yc + 0.14, 0.28)]])
        obj([(yc - 0.18, 0.0), (yc + 0.18, 0.0), (yc + 0.18, 0.05), (yc - 0.18, 0.05)], [])
    # isla (cubierta 0.90) y barra de cocina al fondo
    obj([(11.32, 0.0), (13.82, 0.0), (13.82, 0.90), (11.32, 0.90)], [[(11.32, 0.86), (13.82, 0.86)],
                                                                      [(11.32, 0.10), (13.82, 0.10)]])
    obj([(9.90, 0.0), (15.25, 0.0), (15.25, 0.90), (9.90, 0.90)], [[(9.90, 0.86), (15.25, 0.86)],
                                                                    [(9.90, 0.10), (15.25, 0.10)]])
    occ = None
    for poly, lines in objs:
        for ln in lines:
            g = LineString(ln)
            if occ is not None:
                g = g.difference(occ)
            for part in ([g] if g.geom_type == 'LineString' else list(getattr(g, 'geoms', []))):
                if not part.is_empty and part.length > 0.005:
                    msp.add_lwpolyline(list(part.coords), dxfattribs={'layer': '021_Muebles_Corte'})
        occ = poly if occ is None else unary_union([occ, poly])
    # rótulo de la losa del corte en dos renglones (antes cruzaba el muro del baño)
    for m in msp.query('MTEXT'):
        t = m.plain_text()
        if t.startswith('LOSA ALIGERADA DE CASETONES'):
            m.text = 'LOSA ALIGERADA DE CASETONES 60×60×20\\PCAPA DE COMPRESIÓN 5 cm (h = 0.25)'
            m.dxf.width = 4.0
            m.set_location((42.85, -10.02, 0), attachment_point=1)
            c.e('48C3').set_points([(46.0, -9.65), (42.8, -9.98)])
        if t.startswith('2 TINACOS'):
            m.set_location((93.585, 13.45, 0), attachment_point=2)


# ---------------------------------------- 11. MOBILIARIO EN CORTE B-B' (PB)
def paso_corte_bb_muebles(c: Cad):
    """Sala y comedor nuevos vistos desde el corte B-B' (vista hacia el eje 15)."""
    from ezdxf import bbox
    from shapely.geometry import Polygon, LineString
    from shapely.ops import unary_union
    msp = c.msp
    for e in list(msp.query('*[layer=="021_Muebles_Corte"]')):
        b = bbox.extents([e])
        if b.has_data and b.extmin.x >= 82.95 and b.extmax.x <= 85.35 and b.extmin.y >= -12.52 \
                and b.extmax.y <= -11.45:
            msp.delete_entity(e)
    z0 = -12.50
    X = lambda x: x + OFF_BB
    objs = []

    def obj(outline, extra=()):
        pts = [(X(x), z0 + z) for x, z in outline]
        lines = [pts + [pts[0]]] + [[(X(x), z0 + z) for x, z in ln] for ln in extra]
        objs.append((Polygon(pts).buffer(0), lines))
    # sillón lounge (cortado), de perfil con respaldo al poniente
    obj([(32.75, 0.0), (32.75, 0.62), (32.85, 0.74), (33.02, 0.74), (33.02, 0.42), (33.55, 0.42), (33.55, 0.0)],
        [[(33.02, 0.40), (33.55, 0.40)], [(32.75, 0.10), (33.55, 0.10)]])
    # sofá seccional: brazo sur del módulo largo (cortado) con respaldo al oriente
    obj([(35.30, 0.0), (35.30, 0.60), (36.05, 0.60), (36.05, 0.80), (36.25, 0.80), (36.25, 0.0)],
        [[(35.30, 0.12), (36.25, 0.12)]])
    # mesas de centro
    obj([(33.90, 0.36), (34.80, 0.36), (34.80, 0.40), (33.90, 0.40)])
    obj([(34.10, 0.0), (34.60, 0.0), (34.55, 0.36), (34.15, 0.36)])
    obj([(34.58, 0.41), (35.08, 0.41), (35.08, 0.45), (34.58, 0.45)])
    obj([(34.81, 0.0), (34.85, 0.0), (34.85, 0.41), (34.81, 0.41)])
    # brazo corto del seccional (al fondo)
    obj([(34.20, 0.0), (34.20, 0.80), (35.30, 0.80), (35.30, 0.0)],
        [[(34.40, 0.42), (35.30, 0.42)], [(34.40, 0.42), (34.40, 0.60)], [(34.20, 0.60), (34.40, 0.60)],
         [(34.20, 0.12), (35.30, 0.12)]])
    # comedor: cabecera sur de la mesa y respaldos de las sillas laterales
    obj([(33.60, 0.72), (34.70, 0.72), (34.70, 0.76), (33.60, 0.76)])
    obj([(33.65, 0.0), (34.65, 0.0), (34.65, 0.72), (33.65, 0.72)])
    for xa, xb, xr in ((33.325, 33.775, 33.325), (34.525, 34.975, 34.94)):
        obj([(xa, 0.0), (xa, 0.45), (xb, 0.45), (xb, 0.0)])
        obj([(xr, 0.45), (xr, 0.92), (xr + 0.035, 0.92), (xr + 0.035, 0.45)])
    occ = None
    for poly, lines in objs:
        for ln in lines:
            g = LineString(ln)
            if occ is not None:
                g = g.difference(occ)
            for part in ([g] if g.geom_type == 'LineString' else list(getattr(g, 'geoms', []))):
                if not part.is_empty and part.length > 0.005:
                    msp.add_lwpolyline(list(part.coords), dxfattribs={'layer': '021_Muebles_Corte'})
        occ = poly if occ is None else unary_union([occ, poly])


def build():
    c = Cad(SRC)
    paso_ejes(c)
    paso_recamara(c)
    paso_comedor(c)
    paso_sala(c)
    paso_autos(c)
    paso_escalera(c)
    paso_textos(c)
    paso_notas(c)
    paso_corte_muebles(c)
    paso_corte_bb_muebles(c)
    paso_layouts(c)
    c.save(OUT)
    print('OK ->', OUT)


if __name__ == '__main__':
    build()
