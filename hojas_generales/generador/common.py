"""Parte común del libro general de losas: constantes, ayudantes de fórmulas, TABLAS, DATOS e INSTRUCCIONES.

Las ecuaciones, los coeficientes y el estilo de citas son los de la hoja de Casa Rocha
(losas_nervadas/generador/scripts/gen_a.py y gen_b.py en HEAD), ya auditada contra la NTC 2023.
"""
from openpyxl.styles import Alignment
from openpyxl.workbook.defined_name import DefinedName
from xlstyle import *          # noqa: F401,F403  (font, fill, Narr, cell, hdr_cell, unit_cell, dv_list, add_cf ...)

NTCC = 'NTC-Concreto 2023 (CDMX)'
NTCCA = 'NTC-Criterios y Acciones 2023 (CDMX)'
N = {}                          # nombre -> referencia absoluta (DATOS!$C$12 ...)
NROWS = 40                      # renglones de captura en NERVADAS y MACIZAS
R0 = 9                          # primer renglón de captura
R1 = R0 + NROWS - 1

# ---------------------------------------------------------------- TABLAS (posiciones fijas)
TB_M = 5
CASOS = ['INTERIOR', 'BORDE-CORTO', 'BORDE-LARGO', 'ESQUINA', 'EXTREMO-LARGO', 'EXTREMO-CORTO', 'AISLADO']
MOMS = ['NIC', 'NIL', 'NDC', 'NDL', 'PC', 'PL']
TB_C0 = 6
TB_C1 = TB_C0 + len(CASOS) * len(MOMS) - 1
BARS = [(2, 0.64, 0.32, 0.248), (2.5, 0.79, 0.49, 0.384), (3, 0.95, 0.71, 0.559), (4, 1.27, 1.27, 0.994), (5, 1.59, 1.98, 1.552),
        (6, 1.91, 2.85, 2.235), (8, 2.54, 5.07, 3.973)]
NBARS = len(BARS)
TB_V0 = TB_C1 + 5
TB_V1 = TB_V0 + NBARS - 1
MALLAS = [('6x6-10/10', 3.43), ('6x6-8/8', 4.11), ('6x6-6/6', 4.88), ('6x6-4/4', 5.72)]
TB_W0 = TB_V1 + 5
TB_W1 = TB_W0 + len(MALLAS) - 1
TB_L0 = TB_W1 + 5
LISTS = [('LST_CASO', 'Forzar caso', ['AUTO'] + CASOS),
         ('LST_TIPO', 'Tipo de tablero', ['PERIMETRAL', 'VOLADO']),
         ('LST_BORDE', 'Borde', ['CONTINUO', 'DISCONTINUO']),
         ('LST_COLADO', 'Colado', ['MONOLÍTICO (caso I)', 'NO MONOLÍTICO (caso II)']),
         ('LST_SINO', 'Sí / no', ['SÍ', 'NO']),
         ('LST_AGR', 'Agregado grueso', ['CALIZO', 'BASALTO'])]
TB_T0 = TB_L0 + 14             # lista compacta de tableros para DETALLE/CROQUIS
NTAB = 2 * NROWS
BARCOL = "TABLAS!$B$%d:$B$%d" % (TB_V0, TB_V1)
DBCOL = "TABLAS!$C$%d:$C$%d" % (TB_V0, TB_V1)
ABCOL = "TABLAS!$D$%d:$D$%d" % (TB_V0, TB_V1)
KGCOL = "TABLAS!$E$%d:$E$%d" % (TB_V0, TB_V1)
LDCOL = "TABLAS!$F$%d:$F$%d" % (TB_V0, TB_V1)
LDHCOL = "TABLAS!$G$%d:$G$%d" % (TB_V0, TB_V1)
LTCOL = "TABLAS!$H$%d:$H$%d" % (TB_V0, TB_V1)
LDOCOL = "TABLAS!$I$%d:$I$%d" % (TB_V0, TB_V1)
N.update(BARCOL=BARCOL, DBCOL=DBCOL, ABCOL=ABCOL, KGCOL=KGCOL, LDCOL=LDCOL, LDHCOL=LDHCOL, LTCOL=LTCOL, LDOCOL=LDOCOL)
N['MROW'] = "TABLAS!$D$%d:$J$%d" % (TB_M, TB_M)
N['CKEY'] = "TABLAS!$B$%d:$B$%d" % (TB_C0, TB_C1)
N['CVAL'] = "TABLAS!$D$%d:$J$%d" % (TB_C0, TB_C1)
N['MALLAS'] = "TABLAS!$B$%d:$B$%d" % (TB_W0, TB_W1)
N['MALLA_D'] = "TABLAS!$C$%d:$C$%d" % (TB_W0, TB_W1)
N['MALLA_A'] = "TABLAS!$D$%d:$D$%d" % (TB_W0, TB_W1)
N['LST_TAB'] = "TABLAS!$C$%d:$C$%d" % (TB_T0 + 1, TB_T0 + NTAB)


def db(v):
    return 'INDEX(%s,MATCH(%s,%s,0))' % (DBCOL, v, BARCOL)


def ab(v):
    return 'INDEX(%s,MATCH(%s,%s,0))' % (ABCOL, v, BARCOL)


def ld(v):
    return 'INDEX(%s,MATCH(%s,%s,0))' % (LDCOL, v, BARCOL)


def ldo(v):
    return 'INDEX(%s,MATCH(%s,%s,0))' % (LDOCOL, v, BARCOL)


def ldh(v):
    return 'INDEX(%s,MATCH(%s,%s,0))' % (LDHCOL, v, BARCOL)


def asflex(M, b, d):
    """As por flexión (cm²) para M en t·m; -1 si la sección no alcanza o rebasa p_max (ε_t ≥ 0.005). Igual que gen_b.asflex."""
    rad = '(1-2*%s*1E5/({FRf}*%s*%s^2*{fpc}))' % (M, b, d)
    asf = '(1-SQRT(%s))*{fpc}*%s*%s/{fy}' % (rad, b, d)
    return 'IF(%s<=0,0,IF(%s<0,-1,IF(%s>{pmax}*%s*%s,-1,%s)))' % (M, rad, asf, b, d, asf)


def red391(d):
    """3.9.1 d) y 3.9.2: el peralte del lecho superior se reduce 20 mm si d ≤ 200 mm y no se anota la precaución en planos."""
    return '(%s-IF(AND(%s<=20,{ctl_bw}<>"SÍ"),2,0))' % (d, d)


COEF = {   # NTC-Concreto 2023, Tabla 3.4.2.1 (§3.4.2, p. 65), caso I — copiada sin cambios de gen_a.py (HEAD)
    'INTERIOR': {'NIC': [998, 1160, 907, 736, 607, 503, 418], 'NIL': [516, 550, 517, 492, 466, 436, 418], 'NDC': [0] * 7, 'NDL': [0] * 7, 'PC': [630, 312, 331, 330, 298, 256, 209], 'PL': [175, 139, 174, 200, 213, 216, 209]},
    'BORDE-CORTO': {'NIC': [998, 1163, 965, 785, 652, 546, 460], 'NIL': [516, 561, 525, 498, 470, 440, 421], 'NDC': [0] * 7, 'NDL': [326, 258, 248, 236, 222, 206, 190], 'PC': [630, 329, 292, 331, 311, 276, 238], 'PL': [179, 232, 263, 266, 266, 263, 286]},
    'BORDE-LARGO': {'NIC': [1060, 1292, 981, 777, 628, 512, 421], 'NIL': [587, 566, 540, 513, 482, 483, 461], 'NDC': [651, 362, 321, 283, 250, 219, 190], 'NDL': [0] * 7, 'PC': [751, 458, 433, 383, 323, 293, 286], 'PL': [185, 217, 237, 248, 250, 243, 238]},
    'ESQUINA': {'NIC': [1060, 1382, 1046, 830, 675, 557, 465], 'NIL': [600, 580, 545, 516, 484, 486, 465], 'NDC': [651, 362, 321, 277, 250, 219, 190], 'NDL': [326, 333, 320, 298, 270, 241, 190], 'PC': [751, 452, 437, 397, 338, 306, 296], 'PL': [191, 298, 294, 292, 288, 274, 289]},
    'EXTREMO-LARGO': {'NIC': [1060, 1470, 1108, 882, 730, 650, 570], 'NIL': [0] * 7, 'NDC': [651, 370, 340, 310, 280, 250, 220], 'NDL': [220, 361, 337, 310, 280, 220, 220], 'PC': [751, 730, 670, 610, 550, 490, 430], 'PL': [185, 430, 430, 430, 430, 430, 430]},
    'EXTREMO-CORTO': {'NIC': [0] * 7, 'NIL': [570, 570, 570, 570, 570, 570, 570], 'NDC': [570, 480, 420, 370, 310, 270, 220], 'NDL': [330, 448, 402, 220, 313, 220, 220], 'PC': [1100, 960, 840, 730, 620, 540, 430], 'PL': [200, 430, 430, 430, 430, 430, 430]},
    'AISLADO': {'NIC': [0] * 7, 'NIL': [0] * 7, 'NDC': [570, 550, 530, 470, 430, 380, 330], 'NDL': [330, 446, 330, 330, 330, 330, 330], 'PC': [1100, 830, 800, 720, 640, 570, 500], 'PL': [200, 500, 500, 500, 500, 500, 500]},
}
CASODESC = {'INTERIOR': 'Interior — todos los bordes continuos', 'BORDE-CORTO': 'De borde — un lado corto discontinuo', 'BORDE-LARGO': 'De borde — un lado largo discontinuo',
            'ESQUINA': 'De esquina — dos lados adyacentes discontinuos', 'EXTREMO-LARGO': 'Extremo — tres bordes discontinuos, un lado largo continuo',
            'EXTREMO-CORTO': 'Extremo — tres bordes discontinuos, un lado corto continuo', 'AISLADO': 'Aislado — cuatro lados discontinuos'}
MOMDESC = {'NIC': 'negativo en bordes continuos, claro corto', 'NIL': 'negativo en bordes continuos, claro largo', 'NDC': 'negativo en bordes discontinuos, claro corto',
           'NDL': 'negativo en bordes discontinuos, claro largo', 'PC': 'positivo, claro corto', 'PL': 'positivo, claro largo'}

# cargas vivas: (destino, W_m, ¿azotea?, "W_a / W", referencia)
CVROWS = [
    ('HABITACIÓN', 190, 'NO', '100 / 80', NTCCA + ' Tabla 6.1.2.2 a), p. 11'),
    ('OFICINAS Y DESPACHOS', 250, 'NO', '180 / 100', NTCCA + ' Tabla 6.1.2.2 b), p. 11'),
    ('AULAS', 250, 'NO', '180 / 100', NTCCA + ' Tabla 6.1.2.2 c), p. 11'),
    ('PASILLOS, ESCALERAS Y VESTÍBULOS', 350, 'NO', '150 / 40', NTCCA + ' Tabla 6.1.2.2 d), p. 11'),
    ('LUGARES DE REUNIÓN SIN ASIENTOS INDIVIDUALES', 450, 'NO', '350 / 40', NTCCA + ' Tabla 6.1.2.2 e), p. 11'),
    ('OTROS LUGARES DE REUNIÓN (bibliotecas, templos, restaurantes…)', 350, 'NO', '250 / 40', NTCCA + ' Tabla 6.1.2.2 f), p. 11'),
    ('AZOTEA (pendiente ≤ 5 %)', 100, 'SÍ', '70 / 15', NTCCA + ' Tabla 6.1.2.2 h), p. 11'),
    ('AZOTEA (pendiente de 6 a 10 %)', 60, 'SÍ', '30 / 10', NTCCA + ' Tabla 6.1.2.2 h), p. 11'),
    ('AZOTEA (pendiente de 11 a 20 %)', 40, 'SÍ', '20 / 5', NTCCA + ' Tabla 6.1.2.2 h), p. 11'),
    ('AZOTEA (pendiente > 20 %)', 30, 'SÍ', '20 / 5', NTCCA + ' Tabla 6.1.2.2 h), p. 11'),
    ('VOLADOS EN VÍA PÚBLICA (balcones)', 300, 'NO', '70 / 15', NTCCA + ' Tabla 6.1.2.2 j), p. 11'),
    ('ESTACIONAMIENTO (sólo autos)', 250, 'NO', '100 / 40', NTCCA + ' Tabla 6.1.2.2 k), p. 11'),
    ('OTRO (captura del usuario)', 190, 'NO', '— / —', 'CAPTURA del usuario: cambia el nombre y los valores'),
]


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def cel(ref):
    """'DATOS!$C$5' -> 'C5' (para validación y formato condicional)."""
    return ref.split('!')[1].replace('$', '')


# =====================================================================
def sheet_datos(wb):
    ws = wb.create_sheet('DATOS'); S = Narr(ws, N)
    S.title('DATOS GENERALES — LOSAS NERVADAS Y MACIZAS, NTC 2023 (CDMX)')
    S.subtitle('Datos que valen para todo el libro. Lo de cada tablero va en NERVADAS y MACIZAS. Cada número capturado dice de dónde sale (columna «Referencia»); '
               'lo que es criterio de esta hoja y no de la norma se dice así.', 30)
    S.gap()
    S.notice('CÓMO SE USA: sólo se tocan las celdas AMARILLAS con número AZUL. Lo demás está calculado y se rompe si lo escribes encima. '
             'Los valores que trae la hoja son un punto de partida (los de la hoja auditada de Casa Rocha): cámbialos por los de tu proyecto.')
    S.gap()
    S.section('0. PROYECTO')
    S.header()
    S.row('Proyecto', 'NOMBRE DEL PROYECTO', '—', 'CAPTURA. Sale en DETALLE y en CROQUIS.', '—', 'input', '@', 'proj')
    S.row('Elaboró / revisó', '—', '—', 'CAPTURA.', '—', 'input', '@', 'elab')
    S.gap()
    S.section('1. MATERIALES')
    S.header()
    S.row("f'c — resistencia del concreto", 250, 'kg/cm²', "CAPTURA. Clase 1A (250 ≤ f'c < 400). Mínimo para exposición A1, A2 y B1.",
          NTCC + ' 2.2.2 a), p. 30 · Tabla 2.2.6.3, p. 33 · 4.4.1.1 y 4.5.2 a), p. 78', 'input', '0', 'fc')
    S.row('fy — fluencia de la varilla corrugada', 4200, 'kg/cm²', 'CAPTURA. Grado 42, NMX-B-506.', NTCC + ' Tabla 2.4.2.1.2.a, p. 43', 'input', '0', 'fy')
    S.row('fy de la malla electrosoldada', 5000, 'kg/cm²', 'CAPTURA. Alambre de malla NMX-B-290. Sólo se usa en la capa de compresión.', NTCC + ' 2.4.2.1.6 y 2.4.2.1.7, p. 44', 'input', '0', 'fym')
    S.row('γc — peso volumétrico del concreto reforzado', 2.4, 't/m³', 'CAPTURA. Valor máximo probable. Si tu modelo usa otro (Casa Rocha: 2.44), pon el del modelo.', NTCCA + ' 5.1.1.2, p. 9', 'input', '0.00', 'gc')
    S.row('γ del casetón (poliestireno)', 0.015, 't/m³', 'CAPTURA. Poliestireno expandido de 15 kg/m³. Si el casetón es recuperable (de plástico), pon 0.', 'Criterio de esta hoja (dato del fabricante)', 'input', '0.000', 'gpol')
    S.row('Agregado grueso del concreto', 'CALIZO', '—', "SELECTOR. Cambia Ec: calizo 14 000·√f'c, basáltico 11 000·√f'c. Especificarlo en planos.", NTCC + ' Tabla 2.2.1, p. 31 · 2.2.8.4, p. 34', 'input', '@', 'agr')
    S.row('Ec — módulo de elasticidad', '=IF({agr}="BASALTO",11000,14000)*SQRT({fc})', 'kg/cm²', "CALCULADA = 14 000·√f'c (calizo) u 11 000·√f'c (basalto). Informativa (flechas).", NTCC + ' Tabla 2.2.1, p. 31', 'calc', '#,##0', 'Ec')
    S.row('Es — módulo de elasticidad del acero', 2000000, 'kg/cm²', 'CAPTURA.', NTCC + ' 2.4.2.2.2, p. 45', 'input', '#,##0', 'Es')
    S.row("f''c — esfuerzo del bloque equivalente", '=0.85*{fc}', 'kg/cm²', "CALCULADA = 0.85·f'c.", NTCC + ' 3.6.1 e), p. 68', 'calc', '0.0', 'fpc')
    S.row('β1 — profundidad del bloque', '=IF({fc}<=300,0.85,MAX(0.65,1.05-{fc}/1400))', '—', "CALCULADA. 0.85 hasta f'c = 300 kg/cm².", NTCC + ' ec. 3.6.1, p. 68', 'calc', '0.00', 'beta1')
    S.gap()
    S.section('2. FACTORES DE CARGA Y DE RESISTENCIA')
    S.header()
    S.row('FC_CM — factor de carga muerta', 1.3, '—', 'CAPTURA. Grupo B (Grupo A: 1.5).', NTCCA + ' 3.4.1 a), p. 7 (combinación 2.3.1 a), p. 5)', 'input', '0.00', 'FCcm')
    S.row('FC_CV — factor de carga viva', 1.5, '—', 'CAPTURA. Grupo B (Grupo A: 1.7).', NTCCA + ' 3.4.1 a), p. 7', 'input', '0.00', 'FCcv')
    S.row('FC_acc — factor en combinaciones con acciones accidentales', 1.1, '—', 'CAPTURA. Combinación de granizo en azotea: 1.1·(CM + W_g), sin carga viva.', NTCCA + ' 3.4.1 b), p. 7 · 7.3.1.3, p. 17', 'input', '0.00', 'FCacc')
    S.row('F_R — flexión', 0.9, '—', 'CAPTURA. Vale si la sección es controlada por tensión (ε_t ≥ 0.005); la hoja limita el acero para garantizarlo (p_max).',
          NTCC + ' Tabla 3.8.2.1 a) y Tabla 3.8.2.2, p. 70 · 3.8.2.4, p. 71 · 6.7.3.3.1, p. 169', 'input', '0.00', 'FRf')
    S.row('F_R — cortante', 0.75, '—', 'CAPTURA.', NTCC + ' Tabla 3.8.2.1 b), p. 70', 'input', '0.00', 'FRv')
    S.row('Fracción de la cuantía balanceada que se permite', 0.90, '—', 'CAPTURA. Tope de ductilidad baja; en la práctica manda antes ε_t ≥ 0.005.', NTCC + ' 6.3.5.2.1, p. 122', 'input', '0.00', 'pbf')
    S.row('Incremento de V_cR en losas encasetonadas', 1.1, '—', 'CAPTURA. Sólo en tableros que cumplen 6.7.8.1.4 (claro corto ≥ 6·a_c); en los demás vale 1.0. No aplica a losas macizas.',
          NTCC + ' 6.7.8.1.5, p. 187 · 6.7.8.1.9, p. 188', 'input', '0.00', 'VcRinc')
    S.row('Incremento de cortante con bordes continuos y discontinuos', 1.15, '—', 'CAPTURA. De la fórmula de cortante de losas apoyadas en su perímetro de la NTC 2017, que la 2023 ya no trae; se conserva como «método reconocido».',
          NTCC + ' 3.4.1.1, p. 64 (métodos reconocidos) · NTC-Concreto 2017', 'input', '0.00', 'Vinc')
    S.gap()
    S.section('3. CARGA MUERTA ADICIONAL Y GRANIZO')
    S.header()
    S.row('Incremento del peso propio por losa colada en el lugar', 20, 'kg/m²', 'CAPTURA. Se suma al peso propio de todos los tableros.', NTCCA + ' 5.1.2.1, p. 9', 'input', '0', 'inc_losa')
    S.row('Incremento por capa de mortero sobre la losa (pegapiso, firme o entortado)', 20, 'kg/m²',
          'CAPTURA. Si tu sobrecarga ya trae estos 20 kg/m², pon 0 para no contarlos dos veces. No se aplica donde la carga muerta es favorable.', NTCCA + ' 5.1.2.1 y 5.1.2.2, p. 9', 'input', '0', 'inc_mort')
    S.row('H_s — altura de granizo', 0.125, 'm', 'CAPTURA. Azotea con pendiente < 5°, Grupo B.', NTCCA + ' Tabla 7.3.2.2, p. 18', 'input', '0.000', 'Hs')
    S.row('C_s — coeficiente de forma', 1.0, '—', 'CAPTURA.', NTCCA + ' Tabla 7.3.2.3, p. 18', 'input', '0.00', 'Cs')
    S.row('W_g — granizo en azotea = 900·H_s·C_s ≥ 25', '=MAX(25,900*{Hs}*{Cs})', 'kg/m²', 'CALCULADA, no menos que 25 kg/m² (7.3.2.1). Va con FC = 1.1 y sin carga viva; la hoja toma el mayor de esa combinación y la gravitacional en los destinos marcados como azotea.',
          NTCCA + ' ec. 7.3.2.1, p. 17 · 3.4.1 b), p. 7', 'calc', '0.0', 'Wg')
    S.gap()
    S.section('4. CARGAS VIVAS UNITARIAS (lista de destinos)')
    S.secnote('Los tableros escogen su destino de esta lista. W_m (máxima) es la de diseño por gravedad; W_a y W son informativas. «¿Azotea?» = SÍ activa la combinación con granizo. '
              'Los valores están cotejados contra la Tabla 6.1.2.2 de la NTC-Criterios 2023 (p. 11). Puedes cambiar nombres y valores.', 42)
    S.header(('Destino', 'W_m', '¿Azotea? (granizo)', 'W_a / W (informativos)', 'Referencia'))
    c0 = S.r
    for (nm, wm, az, wa, ref) in CVROWS:
        r = S.r
        cell(ws, r, 2, nm, 'input', '@', al='left', sz=10)
        cell(ws, r, 3, wm, 'input', '0', sz=10)
        cell(ws, r, 4, az, 'input', '@', sz=10)
        cell(ws, r, 5, wa, 'input', '@', al='left', sz=9)
        cell(ws, r, 6, ref, 'ref', al='left', wrap=True)
        ws.row_dimensions[r].height = 24
        S.r += 1
    c1 = S.r - 1
    N['CV_NAMES'] = 'DATOS!$B$%d:$B$%d' % (c0, c1)
    N['CV_WM'] = 'DATOS!$C$%d:$C$%d' % (c0, c1)
    N['CV_AZ'] = 'DATOS!$D$%d:$D$%d' % (c0, c1)
    dv_list(ws, 'D%d:D%d' % (c0, c1), '=LST_SINO')
    S.gap()
    S.section('5. RECUBRIMIENTOS, EXPOSICIÓN Y DETALLADO')
    S.header()
    S.row('r — recubrimiento libre en losas y nervaduras (interior, A1)', 2.0, 'cm',
          'CAPTURA. El mayor de 15 mm (colocación), 0.75 × 25 mm = 19 mm (A1) y 20 mm (losas y viguetas no expuestas). Lo usan las losas macizas y las nervaduras que cumplen 6.7.8.1.4.',
          NTCC + ' 4.13.1.1 y 4.13.2.2 b), p. 83 · Tabla 4.14.16.a, p. 85', 'input', '0.0', 'rec')
    S.row('r de viga — nervaduras de tableros que no cumplen 6.7.8.1.4', 4.0, 'cm',
          'CAPTURA. Si a_c > claro corto/6 la losa se diseña como losa apoyada en vigas (6.7.8.1.9): sus nervaduras ya no son «viguetas» y llevan el recubrimiento de viga (40 mm). También es el que se descuenta del ancho del apoyo al revisar anclajes.',
          NTCC + ' Tabla 4.14.16.a, p. 85 · 6.7.8.1.9, p. 188', 'input', '0.0', 'rec_np')
    S.row('r_ext — recubrimiento de caras a la intemperie (barras ≤ No. 5)', 4.0, 'cm',
          'CAPTURA. Lecho inferior de todo volado y lecho superior de volados y azoteas si así lo indican los selectores de abajo. Con No. 6 o mayor son 5 cm.',
          NTCC + ' Tabla 4.14.16.a, p. 85 · Tabla 4.3.1, p. 76', 'input', '0.0', 'rec_ext')
    S.row('¿Se anotan en planos las precauciones de 3.9.2? (casetón de medida fija y silletas que fijan el lecho superior)', 'SÍ', '—',
          'SELECTOR. Si dices NO: b_w − 2 cm en flexión negativa y cortante cuando b_w < 20 cm, y d − 2 cm en el lecho superior donde d ≤ 20 cm (nervadas y macizas).',
          NTCC + ' 3.9.1 d) y e) y 3.9.2, p. 74', 'input', '@', 'ctl_bw')
    S.row('¿La cara superior de las AZOTEAS queda expuesta directo a la intemperie?', 'NO', '—',
          'SELECTOR. Con relleno, entortado e impermeabilizante se toma como protegida (criterio del Responsable). Si dices SÍ, el lecho superior de azotea usa r_ext.',
          NTCC + ' 4.3.2 y Tabla 4.3.1, p. 76', 'input', '@', 'exp_az')
    S.row('¿La cara superior de los VOLADOS queda a la intemperie?', 'NO', '—',
          'SELECTOR. NO = terraza o balcón con impermeabilizante y piso (decisión D12 de Casa Rocha). Si el volado queda sin protección recibe la lluvia (A2): pon SÍ y su lecho superior lleva r_ext '
          '(la malla al centro de una capa de 6 cm no alcanza ese recubrimiento: la sección 7 lo marca). Su lecho inferior siempre lleva r_ext.',
          NTCC + ' Tabla 4.3.1 c), p. 76 · Tabla 4.14.16.a, p. 85', 'input', '@', 'exp_vol')
    S.row('Tamaño máximo del agregado grueso (TMA)', 1.27, 'cm', 'CAPTURA. No más de 2/3 del recubrimiento ni de la separación libre entre barras.', NTCC + ' 2.2.5.2.1, p. 32', 'input', '0.00', 'tmag')
    S.row('¿El agregado cabe? (≤ 2/3 de r y de la separación libre)', '=IF({tmag}<=2/3*MIN({rec},MAX(2.5,1.5*{tmag})),"CUMPLE","NO CUMPLE — usa agregado más chico")', '—', 'CALCULADA.', NTCC + ' 2.2.5.2.1, p. 32', 'calc', '@')
    S.row('Separación libre mínima entre varillas', '=MAX(2.5,1.5*{tmag})', 'cm', 'CALCULADA = el mayor de 25 mm, d_b y 1.5·TMA (d_b se revisa barra por barra).', NTCC + ' 14.2.1, p. 347', 'calc', '0.00', 'slib')
    S.row('No. de varilla preferida — lecho INFERIOR de nervaduras', 4, '—', 'CAPTURA. Mínimo No. 4: toda nervadura lleva al menos 1 No. 4 abajo y 1 No. 4 arriba, corridas (si capturas menos, NERVADAS usa No. 4). Si no alcanza, la hoja sube sola de número.',
          NTCC + ' 6.7.6.1.3, p. 175', 'input', 'General', 'var_inf')
    S.row('No. de varilla preferida — lecho SUPERIOR de nervaduras', 4, '—', 'CAPTURA. Mínimo No. 4 (si capturas menos, NERVADAS usa No. 4). Una varilla superior es corrida y las demás son bastones.', NTCC + ' 6.7.6.1.3, p. 175', 'input', 'General', 'var_sup')
    S.row('Máximo de varillas por lecho en una nervadura', 2, 'pzas', 'CAPTURA. Además se revisa que quepan con la separación libre mínima.', NTCC + ' 14.2.1, p. 347', 'input', '0', 'nmax_sec')
    S.row('Mínimo de varillas por lecho y por nervadura', 1, 'pzas', 'CAPTURA. 1 corrida. Si As colocada < 4/3 de la requerida, la hoja pide 2 (As mín de viga).', NTCC + ' 6.7.6.1.3, p. 175 · 6.3.5.1.1 y 6.3.5.1.3, p. 122', 'input', '0', 'nmin_sec')
    S.row('Ancho del apoyo en bordes DISCONTINUOS (para anclar)', 40, 'cm',
          'CAPTURA. Dala de borde armada (o macizo de borde) donde rematan con gancho de 180° la varilla inferior (1.25 fy en el paño interior, ≥ 15 cm dentro) y el bastón superior (fy), como en la decisión D9 de Casa Rocha. '
          'Con un muro de 15 cm sin dala ancha, pon 15: la hoja pedirá una dala más ancha.',
          NTCC + ' 6.7.8.1.7, p. 188 · 6.7.7.4.1.1 a) y b), p. 179 · 14.4.3.1, p. 355', 'input', '0', 'b_borde')
    S.row('Peralte de la dala de borde', 25, 'cm',
          'CAPTURA. Donde se dobla el gancho de 180° (alto 8·d_b) entre los recubrimientos de viga arriba y abajo. En nervadas suele ser h; bajo una losa maciza, la dala suele bajar más que la losa.',
          NTCC + ' Tabla 14.3.1, p. 349 · 14.4.3.1, p. 355', 'input', '0', 'h_borde')
    S.gap()
    S.section('6. LOSA NERVADA — SECCIÓN TIPO')
    S.secnote('Casetón a_c × a_c × h_c, nervadura b_w y capa t. Cada tablero de NERVADAS puede usar otro b_w (p. ej. 20 cm en volados): entonces conserva la separación s y el casetón se angosta a s − b_w.')
    S.header()
    S.row('a_c — lado del casetón (separación libre entre nervaduras)', 60, 'cm', 'CAPTURA. ≤ 80 cm y ≤ claro corto/6 para que cuente como encasetonada (se revisa por tablero).', NTCC + ' 6.7.8.1.4, p. 187', 'input', '0', 'ac')
    S.row('h_c — altura del casetón', 19, 'cm', 'CAPTURA. Con t = 6 cm da h = 25 cm (LR-25, casetón 60×60×19 de la hoja de Casa Rocha, decisión D3).', 'Dato del proyecto', 'input', '0', 'hc')
    S.row('b_w — ancho de la nervadura', 10, 'cm', 'CAPTURA. Mínimo 10 cm en cualquier punto del peralte.', NTCC + ' 6.7.8.1.2, p. 187', 'input', '0', 'bw')
    S.row('t — espesor de la capa de compresión', 6, 'cm', 'CAPTURA. Mínimo el mayor de a_c/13 y 5 cm; 6 cm si algún claro es ≥ 6 m (diafragma).', NTCC + ' 6.7.8.3.1, p. 188 · 6.8.3.1.2, p. 189', 'input', '0', 't')
    S.row('h — peralte total', '={hc}+{t}', 'cm', 'CALCULADA = h_c + t. Máximo 3.5·b_w.', NTCC + ' 6.7.8.1.3, p. 187', 'key', '0', 'h')
    S.row('s — separación entre ejes de nervaduras', '={ac}+{bw}', 'cm', 'CALCULADA = a_c + b_w. Ancho tributario de cada nervadura.', 'Geometría', 'key', '0', 's')
    S.row('b_p — ancho del apoyo (muro, dala o trabe) en los bordes', 15, 'cm', 'CAPTURA. Claro libre L_n = claro a ejes − b_p (en volados − b_p/2). Lo usan nervadas y macizas.', NTCC + ' 6.3.2.1, p. 117', 'input', '0', 'bp')
    S.row('¿La sección cumple las limitaciones de losa encasetonada? (b_w, h, t)', '=IF(AND({bw}>=10,{h}<=3.5*{bw},{t}>=MAX({ac}/13,5),{ac}<=80),"CUMPLE","NO CUMPLE")', '—',
          'CALCULADA: b_w ≥ 10 cm, h ≤ 3.5·b_w, t ≥ máx(a_c/13, 5 cm), a_c ≤ 80 cm. El claro ≥ 6·a_c se revisa por tablero.', NTCC + ' 6.7.8.1.2 a 6.7.8.1.4, p. 187 · 6.7.8.3.1, p. 188', 'key', '@')
    S.row('Volumen de concreto por m²', '={h}/100-({ac}/100)^2*({hc}/100)/({s}/100)^2', 'm³/m²', 'CALCULADA = h − a_c²·h_c/s².', 'Geometría', 'calc', '0.0000', 'Vm2')
    S.row('Peso propio geométrico', '=({Vm2}*{gc}+({ac}/100)^2*({hc}/100)/({s}/100)^2*{gpol})*1000', 'kg/m²', 'CALCULADA = V·γc + casetón·γ_casetón.', 'Geometría', 'key', '0.0', 'PP')
    S.row('Peso propio de diseño (+ incrementos de la sección 3)', '={PP}+{inc_losa}+{inc_mort}', 'kg/m²', 'CALCULADA. Con b_w propio en un tablero, NERVADAS lo recalcula en ese renglón.', NTCCA + ' 5.1.2.1, p. 9', 'key', '0.0', 'PPd')
    S.row('A_T — área de la sección T de una nervadura', '={s}*{t}+{bw}*{hc}', 'cm²', 'CALCULADA = s·t + b_w·h_c.', 'Geometría', 'calc', '0.0', 'AT')
    S.row('ȳ — centroide medido desde arriba', '=({s}*{t}*{t}/2+{bw}*{hc}*({t}+{hc}/2))/{AT}', 'cm', 'CALCULADA.', 'Geometría', 'calc', '0.00', 'ybar')
    S.row('I_T — inercia bruta de una nervadura con su patín', '={s}*{t}^3/12+{s}*{t}*({ybar}-{t}/2)^2+{bw}*{hc}^3/12+{bw}*{hc}*({t}+{hc}/2-{ybar})^2', 'cm⁴', 'CALCULADA (Steiner).', 'Geometría', 'calc', '#,##0', 'IT')
    S.row('h_eq — losa maciza de la misma inercia', '=(12*{IT}/{s})^(1/3)', 'cm', 'CALCULADA = ∛(12·I_T/s). La hoja compara h_eq (no h) contra el espesor mínimo, del lado seguro.', 'Criterio de esta hoja · ' + NTCC + ' 6.7.3.1, p. 167–168', 'key', '0.00', 'heq')
    S.row('Malla electrosoldada de la capa', '6x6-4/4', '—', 'SELECTOR (lista en TABLAS). Sus revisiones están en la sección 7.', NTCC + ' 13.8.2.1.1, p. 345 · 6.7.8.1.8, p. 188 · 6.8.6.4, p. 196', 'input', '@', 'malla_sel')
    S.row('Ø del alambre de la malla', '=INDEX({MALLA_D},MATCH({malla_sel},{MALLAS},0))/10', 'cm', 'CALCULADA. La malla va al centro de la capa y las varillas superiores pasan DEBAJO de ella: su d pierde 2·Ø.', 'TABLAS sección C', 'calc', '0.000', 'fim')
    S.row('d inferior típico (varilla preferida)', '={h}-{rec}-' + db('{var_inf}') + '/2', 'cm', 'CALCULADA = h − r − d_b/2. NERVADAS lo recalcula con la varilla real, la dirección y el r de cada tablero.', 'Geometría', 'calc', '0.00', 'd_inf')
    S.row('d superior típico (varilla preferida)', '={h}-MAX({rec},({t}-2*{fim})/2)-2*{fim}-' + db('{var_sup}') + '/2', 'cm', 'CALCULADA = h − máx(r, (t − 2Ø)/2) − 2·Ø malla − d_b/2: la malla va al centro de la capa y la varilla pasa debajo; si r es mayor, manda r.', 'Geometría · ' + NTCC + ' 3.9.1 d), p. 74', 'calc', '0.00', 'd_sup')
    S.row('p_b — cuantía balanceada', '={fpc}/{fy}*6000*{beta1}/({fy}+6000)', '—', "CALCULADA = (f''c/fy)·6000β1/(fy + 6000).", NTCC + ' ec. 5.2.1.3.1, p. 88', 'calc', '0.0000', 'pb')
    S.row('p para ε_t = 0.005 (sección controlada por tensión)', '=0.375*{beta1}*{fpc}/{fy}', '—', "CALCULADA: c ≤ 0.375·d ⇒ p ≤ 0.375·β1·f''c/fy sobre el ancho a compresión.", NTCC + ' Tabla 3.8.2.2, p. 70 · 3.8.2.4, p. 71 · 6.3.3.3.1, p. 121', 'calc', '0.0000', 'pten')
    S.row('p_max — cuantía máxima', '=MIN({pbf}*{pb},{pten})', '—', 'CALCULADA = el menor de 0.90·p_b y el de ε_t = 0.005. Vale para nervadas y macizas.', NTCC + ' 6.3.5.2.1, p. 122 · 6.7.3.3.1, p. 169', 'key', '0.0000', 'pmax')
    S.row('p_min — cuantía mínima por flexión de nervaduras', '=MAX(0.8*SQRT({fc}),14)/{fy}', '—', "CALCULADA = el mayor de 0.80·√f'c/fy y 14/fy, sobre b_w·d. Se dispensa si As colocada ≥ 4/3 de la requerida.", NTCC + ' 6.3.5.1.2 y 6.3.5.1.3, p. 122', 'calc', '0.00000', 'pmin')
    S.row('α_fm por omisión (rigidez de los apoyos del tablero)', 99, '—',
          'CAPTURA. 99 = muros o trabes muy rígidas en los cuatro bordes (α_f → ∞). Con trabes de concreto calcula α_f = I_b/I_s por borde, promedia y ponlo en el tablero.',
          NTCC + ' Tabla 6.7.3.1.2 nota [1], p. 168', 'input', '0.00', 'afm_def')
    S.row('¿Los bordes discontinuos tienen muro o viga con α_f ≥ 0.8?', 'SÍ', '—', 'SELECTOR. Si dices NO, el espesor mínimo de la Tabla 6.7.3.1.2 se aumenta 10 % en los tableros con algún borde discontinuo.', NTCC + ' 6.7.3.1.2.1, p. 168', 'input', '@', 'borde_viga')
    S.gap()
    S.section('7. CAPA DE COMPRESIÓN Y MALLA (IGUAL PARA TODOS LOS TABLEROS NERVADOS)')
    S.header()
    S.row('As de la malla', '=INDEX({MALLA_A},MATCH({malla_sel},{MALLAS},0))', 'cm²/m', 'TABLAS sección C.', 'TABLAS', 'calc', '0.000', 'asm')
    S.row('Temperatura: As mín = 0.0018·t·100', '=0.0018*{t}*100', 'cm²/m', 'Ya no depende de fy.', NTCC + ' 13.8.2.1.1, p. 345 · 6.7.8.1.8, p. 188', 'calc', '0.000', 'as1')
    S.row('Estado (temperatura)', '=IF({asm}>={as1},"CUMPLE","NO CUMPLE")', '—', '—', '—', 'key', '@', 'ck_t')
    S.row('Diafragma: cuantía mínima por cortante = 0.0025·t·100', '=0.0025*{t}*100', 'cm²/m', 'Sólo cuenta el acero de la capa.', NTCC + ' 6.8.6.4 y 6.8.6.3, p. 196', 'calc', '0.000', 'asd')
    S.row('Estado (diafragma)', '=IF({asm}>={asd}-0.0005,"CUMPLE","NO CUMPLE: "&TEXT({asm}/{asd},"0%")&" — sube la malla o la capa")', '—', '—', '—', 'key', '@', 'ck_d')
    S.row('Recubrimiento de la malla centrada en la capa (¿cabe?)', '=({t}-2*{fim})/2', 'cm', '(t − 2Ø)/2 ≥ r (≥ r_ext donde la cara superior queda expuesta).', NTCC + ' Tabla 4.14.16.a, p. 85 · 4.3.2, p. 76', 'calc', '0.00', 'rm')
    S.row('Estado (cabe con su recubrimiento)', '=IF({rm}>=IF(OR({exp_az}="SÍ",{exp_vol}="SÍ"),{rec_ext},{rec})-0.005,"CUMPLE",IF({rm}>={rec}-0.005,"NO CUMPLE en zonas expuestas: protegerlas o engrosar ahí la capa","NO CUMPLE: no cabe"))', '—', '—', '—', 'key', '@', 'ck_r')
    S.row('Separación de la malla (15.24 cm) ≤ máx. permitida', '=IF(15.24<=MIN(5*{t},45,30),"CUMPLE","NO CUMPLE")', '—', 'El menor de 5t, 45 cm y 30 cm entre alambres soldados.', NTCC + ' 6.8.7.3, p. 196 · 2.4.2.1.6.3, p. 44', 'calc', '@', 'ck_s')
    S.row('Espesor de piso y rellenos que reparten la carga concentrada', 0, 'cm', 'CAPTURA. 0 = la carga actúa directo sobre la capa (lo más estricto).', 'Criterio de esta hoja', 'input', '0.0', 'tpz')
    S.row('Carga concentrada en la capa entre nervaduras (10 × 10 cm)', 1000, 'kg', 'CAPTURA. Mínimo de la norma en la posición más desfavorable.', NTCC + ' 6.7.8.1.6, p. 187', 'input', '0', 'Pcap')
    S.row('P_u = FC_CV × P', '={FCcv}*{Pcap}', 'kg', 'CALCULADA.', NTCCA + ' 3.4.1 a), p. 7', 'calc', '0', 'Pu')
    S.row('d de la capa (malla al centro)', '={t}/2', 'cm', '—', 'Geometría', 'calc', '0.00', 'dc')
    S.row('b₀ — perímetro crítico a d/2', '=4*(10+2*{tpz}+{dc})', 'cm', '4·(c + d).', NTCC + ' 5.6, p. 99–100', 'calc', '0.0', 'bo')
    S.row('V_cR punzonamiento de la capa', '={FRv}*MIN(1,0.27*(2+40*{dc}/{bo}))*MIN(1,SQRT(2/(1+0.04*{dc})))*SQRT({fc})*{bo}*{dc}', 'kg',
          "F_R·(el menor de 1.0, 0.5(1 + 2/β) y 0.27(2 + α_s·d/b₀))·λs·√f'c·b₀·d, β = 1, α_s = 40.", NTCC + ' Tabla 5.6.2.1.1, p. 100', 'calc', '0', 'vcp')
    S.row('Estado (punzonamiento con 1 000 kg)', '=IF({Pu}<={vcp},"CUMPLE","NO CUMPLE: P_u/V_cR = "&TEXT({Pu}/{vcp},"0.00"))', '—', '—', '—', 'key', '@', 'ck_p')
    S.row('v_uv — esfuerzo cortante en la sección crítica', '={Pu}/({bo}*{dc})', 'kg/cm²', 'P_u/(b₀·d).', NTCC + ' 6.7.6.1.2, p. 175', 'calc', '0.00', 'vuv')
    S.row("Límite 0.5·F_R·λs·√f'c", '=0.5*{FRv}*MIN(1,SQRT(2/(1+0.04*{dc})))*SQRT({fc})', 'kg/cm²', 'Si v_uv lo rebasa se pide acero mínimo alrededor de la carga.', NTCC + ' 6.7.6.1.2, p. 175', 'calc', '0.00', 'vlim')
    S.row('b_losa — ancho donde va ese acero', '=10+2*{tpz}+2*1.5*{t}', 'cm', 'Ancho cargado + 1.5h a cada lado.', NTCC + ' Tabla 6.7.4.2.2.3, p. 171', 'calc', '0.0', 'bls')
    S.row('As,mín alrededor de la carga = 5·v_uv·b_losa·b₀/(F_R·α_s·fy)', '=IF({vuv}>{vlim},5*{vuv}*{bls}*{bo}/({FRv}*40*{fym}),0)', 'cm²', 'α_s = 40; fy de la malla.', NTCC + ' ec. 6.7.6.1.2, p. 175', 'calc', '0.000', 'asc')
    S.row('Estado (acero alrededor de la carga)', '=IF({asm}/100*{bls}>={asc}-0.0005,"CUMPLE","NO CUMPLE: la malla da "&TEXT({asm}/100*{bls},"0.00")&" de "&TEXT({asc},"0.00")&" cm²")', '—', '—', '—', 'key', '@', 'ck_c')
    S.row('m_R de la capa por unidad de ancho', '={FRf}*{asm}/100*{fym}*({dc}-{asm}/100*{fym}/{fpc}/2)', 'kg·cm/cm', 'F_R·As·fy·(d − a/2) por cm.', NTCC + ' ec. 5.2.2.1.1.1.b, p. 89', 'calc', '0.0', 'mR')
    S.row('m requerido = P_u/(4π) (abanico de líneas de fluencia)', '={Pu}/(4*PI())', 'kg·cm/cm', 'Malla al centro: m = m′. Ignorar el área cargada queda del lado seguro.', 'Criterio de esta hoja · ' + NTCC + ' 3.4.1.1, p. 64', 'calc', '0.0', 'mq')
    S.row('Estado (flexión de la capa)', '=IF({mq}<={mR},"CUMPLE","NO CUMPLE")', '—', '—', '—', 'key', '@', 'ck_f')
    S.row('RESUMEN DE LA CAPA', '=IF(AND(LEFT({ck_t},6)="CUMPLE",LEFT({ck_d},6)="CUMPLE",LEFT({ck_r},6)="CUMPLE",LEFT({ck_s},6)="CUMPLE",LEFT({ck_p},6)="CUMPLE",LEFT({ck_c},6)="CUMPLE",LEFT({ck_f},6)="CUMPLE"),"CUMPLE","REVISAR: capa o malla (ver renglones de arriba)")',
          '—', 'Si no cumple: capa de 6 cm o malla más gruesa (y que quepa).', '—', 'key', '@', 'capa')
    S.gap()
    S.section('8. LOSA MACIZA — VALORES POR OMISIÓN')
    S.secnote('MACIZAS usa estos valores cuando el renglón deja vacía la celda correspondiente. La separación máxima es CRITERIO DE ESTA HOJA y está para que la ajustes: '
              'la hoja no pudo cotejar el numeral de separación máxima de refuerzo en losas de la NTC 2023; los valores son los del ACI 318-19, del que deriva.', 42)
    S.header()
    S.row('h — espesor de la losa maciza por omisión', 12, 'cm', 'CAPTURA. Cada tablero puede llevar su propio h.', 'Dato del proyecto', 'input', '0', 'h_mac')
    S.row('No. de varilla preferida en losa maciza', 3, '—', 'CAPTURA. Si con la separación mínima no alcanza, la hoja sube de número.', 'Dato del proyecto', 'input', 'General', 'var_mac')
    S.row('Cuantía de As mínima por lecho (temperatura y losa)', 0.0018, '—', 'CAPTURA. As mín = 0.0018·b·h en cada lecho y dirección (del lado seguro: la norma la pide sobre el área total).', NTCC + ' 6.7.6.1.1, p. 175 · 13.8.2.1.1, p. 345', 'input', '0.0000', 'pmac')
    S.row('Separación máxima — tableros perimetrales: múltiplo de h', 2, '× h', 'CAPTURA. Secciones críticas de losas en dos direcciones: 2h.', 'Criterio de esta hoja = ACI 318-19 8.7.2.2 · confirmar con la NTC-Concreto 2023', 'input', '0.0', 'kcrit')
    S.row('Separación máxima — volados: múltiplo de h', 3, '× h', 'CAPTURA. Losas en una dirección: 3h.', 'Criterio de esta hoja = ACI 318-19 7.7.2.3 · confirmar con la NTC-Concreto 2023', 'input', '0.0', 'kotr')
    S.row('Separación máxima absoluta', 45, 'cm', 'CAPTURA.', 'Criterio de esta hoja = ACI 318-19 7.7.2.3 y 8.7.2.2 (450 mm)', 'input', '0', 'stope')
    S.row('Separación mínima práctica', 10, 'cm', 'CAPTURA. Debajo de esto la hoja sube de número de varilla.', 'Criterio de esta hoja (práctica de obra)', 'input', '0', 'smin')
    S.row('Redondeo de la separación', 2.5, 'cm', 'CAPTURA. La separación se redondea hacia abajo a múltiplos de este valor.', 'Criterio de esta hoja (práctica de obra)', 'input', '0.0', 'spaso')
    S.row('h mínimo de volados macizos = L / …', 10, '—', 'CAPTURA. Losa en voladizo sin elementos frágiles: L/10 (L = longitud del volado a ejes). Si carga pretil o muro frágil, calcula la flecha.', 'Criterio de esta hoja = ACI 318-19 Tabla 7.3.1.1 · confirmar con la NTC-Concreto 2023', 'input', '0', 'kvol')
    S.gap()
    for nm in ('agr',):
        dv_list(ws, cel(N[nm]), '=LST_AGR')
    for nm in ('ctl_bw', 'exp_az', 'exp_vol', 'borde_viga'):
        dv_list(ws, cel(N[nm]), '=LST_SINO')
    dv_list(ws, cel(N['malla_sel']), '=LST_MALLA')
    add_cf(ws, 'C1:F%d' % S.r)
    ws.sheet_view.zoomScale = 90
    ws.freeze_panes = 'A4'
    print_setup(ws, 'B1:F%d' % S.r, landscape=False)
    return ws


def print_setup(ws, area, landscape=True):
    ws.print_area = area
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = '&A — página &P de &N'
    ws.oddFooter.center.size = 8


# =====================================================================
def sheet_tablas(wb):
    ws = wb.create_sheet('TABLAS')
    S = Narr(ws, N, {'A': 1.78, 'B': 22, 'C': 46, 'D': 8, 'E': 8, 'F': 8, 'G': 8, 'H': 8, 'I': 8, 'J': 8, 'K': 40}, last='K')
    S.title('TABLAS — COEFICIENTES, VARILLAS, MALLAS Y LISTAS')
    S.subtitle('Todas las hojas leen de aquí con INDEX/MATCH. Los coeficientes son los de la NTC; L_d, L_dh y traslapes se recalculan solos con f\'c y fy de DATOS.', 18)
    S.r = TB_M - 1
    S.section('A. COEFICIENTES DE MOMENTO PARA TABLEROS RECTANGULARES, FRANJAS CENTRALES — CASO I (LOSA COLADA MONOLÍTICAMENTE CON SUS APOYOS) · ×10⁻⁴·w·a1²')
    r = TB_M
    hdr_cell(ws, r, 2, 'Clave'); hdr_cell(ws, r, 3, 'Caso — momento  (m = a1/a2 →)')
    for k, mv in enumerate([0, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]):
        x = hdr_cell(ws, r, 4 + k, mv); x.number_format = '0.0'
    hdr_cell(ws, r, 11, 'Referencia')
    rr = TB_C0
    for cs in CASOS:
        for mm in MOMS:
            cell(ws, rr, 2, '%s|%s' % (cs, mm), al='left', sz=8)
            cell(ws, rr, 3, '%s — %s' % (CASODESC[cs], MOMDESC[mm]), al='left', sz=8)
            for k, v in enumerate(COEF[cs][mm]):
                cell(ws, rr, 4 + k, v, 'calc', '0')
            cell(ws, rr, 11, NTCC + ' · Tabla 3.4.2.1, caso I (§3.4.2, p. 65)', 'ref', al='left')
            rr += 1
    assert rr - 1 == TB_C1
    S.r = TB_C1 + 1
    S.notice('FUENTE: NTC-Concreto 2023, Tabla 3.4.2.1 «Coeficientes de momentos flexionantes para tableros rectangulares, franjas centrales», columna I (losa colada monolíticamente con sus apoyos), p. 65; '
             'copiada sin cambios de la hoja auditada de Casa Rocha (cotejada celda por celda contra la imagen de la tabla). SÓLO ESTÁ EL CASO I: el caso II (losa no colada monolíticamente) no está transcrito '
             'y la hoja marca FUERA DE ALCANCE a esos tableros. Renglones en cero = momentos que ese caso no tiene. Para m < 0.5 se interpola entre 0 y 0.5. '
             'Se conservan tal como vienen impresas dos rarezas de NDL que parecen erratas (Extremo corto en m = 0.7: 220; Aislado en m = 0.5: 446).', 66)
    S.r = TB_V0 - 2
    S.section('B. VARILLA CORRUGADA — ÁREAS, PESOS, DESARROLLO Y TRASLAPES (se recalculan con DATOS)')
    r = TB_V0 - 1
    for k, t in enumerate(['No.', 'd_b (cm)', 'A_s (cm²)', 'kg/m', 'L_d (cm)', 'L_dh gancho (cm)', 'Traslape clase B (cm)', 'L_d otros casos (cm)']):
        hdr_cell(ws, r, 2 + k, t)
    hdr_cell(ws, r, 10, 'Referencia'); ws.merge_cells('J%d:K%d' % (r, r))
    for i, (n_, d_, a_, w_) in enumerate(BARS):
        rr = TB_V0 + i
        cell(ws, rr, 2, n_, fmt='General'); cell(ws, rr, 3, d_, fmt='0.00'); cell(ws, rr, 4, a_, fmt='0.00'); cell(ws, rr, 5, w_, fmt='0.000')
        if n_ == 2:
            for cc in (6, 7, 8, 9):
                cell(ws, rr, cc, 'no aplica', al='center')
        else:
            cell(ws, rr, 6, '=MAX(%s/(IF(B%d<=6,6.6,5.3)*SQRT(%s))*C%d,30)' % (N['fy'], rr, N['fc'], rr), fmt='0')
            cell(ws, rr, 7, '=MAX(0.075*%s/SQRT(%s)*C%d,8*C%d,15)' % (N['fy'], N['fc'], rr, rr), fmt='0')
            cell(ws, rr, 8, '=MAX(1.3*F%d,30)' % rr, fmt='0')
            cell(ws, rr, 9, '=MAX(%s/(IF(B%d<=6,4.4,3.5)*SQRT(%s))*C%d,30)' % (N['fy'], rr, N['fc'], rr), fmt='0')
        ws.merge_cells('J%d:K%d' % (rr, rr))
        cell(ws, rr, 10, ('NMX-B-506 · L_d: NTC-Concreto 14.4.2.1 y Tabla 14.4.2.4, pp. 352–353 · L_dh: 14.4.3.1, p. 355 · traslape clase B: Tabla 14.5.2.1, p. 367' if n_ != 2 else
                               'Alambrón liso 1/4" (No. 2): no se usa como refuerzo longitudinal'), 'ref', al='left')
    S.r = TB_V1 + 1
    S.secnote('L_d (Tabla 14.4.2.4, p. 353) con separación libre ≥ 2d_b y recubrimiento ≥ d_b: fy·d_b/(6.6·√f\'c) hasta No. 6 y fy·d_b/(5.3·√f\'c) de No. 7 en adelante, ≥ 30 cm. '
              '«Otros casos» (separación libre < 2d_b o recubrimiento < d_b): fy·d_b/(4.4·√f\'c) y fy·d_b/(3.5·√f\'c). '
              'L_dh = el mayor de 0.075·fy·d_b/√f\'c, 8d_b y 15 cm; gancho de 90° con extensión de 12d_b (14.4.3.1, p. 355; Tabla 14.3.1, p. 349).', 42)
    S.r = TB_W0 - 2
    S.section('C. MALLA ELECTROSOLDADA (separación 15.24 cm en las dos direcciones)')
    r = TB_W0 - 1
    for k, t in enumerate(['Malla', 'Ø alambre (mm)', 'As (cm²/m)', 'kg/m²']):
        hdr_cell(ws, r, 2 + k, t)
    hdr_cell(ws, r, 6, 'Referencia'); ws.merge_cells('F%d:K%d' % (r, r))
    for i, (nm, dmm) in enumerate(MALLAS):
        rr = TB_W0 + i
        cell(ws, rr, 2, nm); cell(ws, rr, 3, dmm, fmt='0.00')
        cell(ws, rr, 4, '=PI()*(C%d/10)^2/4*100/15.24' % rr, fmt='0.000')
        cell(ws, rr, 5, '=2*PI()*(C%d/1000)^2/4*7850/0.1524' % rr, fmt='0.00')
        ws.merge_cells('F%d:K%d' % (rr, rr))
        cell(ws, rr, 6, 'Calibre de alambre y separación de 6" = 15.24 cm · NMX-B-290 · NTC-Concreto 2.4.2.1.6.3, p. 44 y 6.8.7.3, p. 196 · traslape ≥ un cuadro + 5 cm (14.5.4.1, p. 369)', 'ref', al='left')
    S.r = TB_L0 - 1
    S.section('D. LISTAS DE LOS SELECTORES')
    for j, (nm, title, vals) in enumerate(LISTS):
        hdr_cell(ws, TB_L0, 2 + j, title)
        for i, v in enumerate(vals):
            cell(ws, TB_L0 + 1 + i, 2 + j, v, al='left')
        col = CL(2 + j)
        add_name(wb, nm, "TABLAS!$%s$%d:$%s$%d" % (col, TB_L0 + 1, col, TB_L0 + len(vals)))
    add_name(wb, 'LST_DEST', N['CV_NAMES'])
    add_name(wb, 'LST_MALLA', N['MALLAS'])
    S.r = TB_T0 - 1
    S.section('E. LISTA DE TABLEROS PARA DETALLE Y CROQUIS (la arma el libro: N = nervada, M = maciza)')
    hdr_cell(ws, TB_T0, 2, 'k'); hdr_cell(ws, TB_T0, 3, 'Tablero (para el selector)')
    nn = "NERVADAS!$B$%d:$B$%d" % (R0, R1); nc = "NERVADAS!$A$%d:$A$%d" % (R0, R1)
    mn = "MACIZAS!$B$%d:$B$%d" % (R0, R1); mc = "MACIZAS!$A$%d:$A$%d" % (R0, R1)
    for i in range(NTAB):
        rr = TB_T0 + 1 + i
        cell(ws, rr, 2, i + 1, fmt='0')
        f = ('=IF($B{r}<=MAX({nc}),"N · "&INDEX({nn},MATCH($B{r},{nc},0)),IF($B{r}-MAX({nc})<=MAX({mc}),"M · "&INDEX({mn},MATCH($B{r}-MAX({nc}),{mc},0)),""))'
             .format(r=rr, nc=nc, nn=nn, mc=mc, mn=mn))
        cell(ws, rr, 3, f, al='left')
    add_name(wb, 'LST_TAB', N['LST_TAB'])
    ws.column_dimensions['B'].width = 24; ws.column_dimensions['C'].width = 58; ws.column_dimensions['D'].width = 12; ws.column_dimensions['E'].width = 12
    ws.freeze_panes = 'A4'
    print_setup(ws, 'B1:K%d' % (TB_W1 + 1), landscape=True)
    return ws


# =====================================================================
def sheet_instrucciones(wb):
    ws = wb.active; ws.title = 'INSTRUCCIONES'
    S = Narr(ws, N, {'A': 1.78, 'B': 30, 'C': 16, 'D': 9, 'E': 54, 'F': 40}, last='F')
    S.title('LOSAS NERVADAS Y MACIZAS — NTC 2023 (CDMX) · HOJA GENERAL')
    S.subtitle('Diseño por flexión y cortante de tableros de losa con el método de coeficientes de la NTC-Concreto 2023 (Tabla 3.4.2.1), para cualquier proyecto. '
               'Sale de la hoja auditada de Casa Rocha: mismas ecuaciones, mismas citas («NTC-Concreto 2023 (CDMX) 6.7.8.1.4, p. 187»).', 30)
    S.gap()
    S.section('1. QUÉ HACE')
    for t in [
        '• NERVADAS: hasta 40 tableros de losa nervada (encasetonada). Por tablero: caso de la tabla, coeficientes interpolados en m = a1/a2, momentos por nervadura en las dos direcciones, '
        'As, número y cantidad de varillas (sube de número si no cabe), M_R, ε_t ≥ 0.005, cortante con V_cR (×1.1 sólo si cumple 6.7.8.1.4), hileras de casetones a macizar, espesor mínimo contra h_eq, bastones y estado.',
        '• MACIZAS: hasta 40 tableros de losa maciza, por franja de 1 m: los mismos pasos, con As mínima 0.0018·b·h, varilla y separación («No. 3 @ 20 cm»), cortante de losa y h mínimo.',
        '• VOLADOS (en las dos hojas): momento y cortante en el paño del apoyo con la carga lineal en la orilla y el peso de los macizos (nervadas), y acero superior que entra al tablero de respaldo. '
        'En nervadas, si el volado lleva nervaduras más anchas, cuántas hileras del respaldo siguen con ellas (decisión D10 de Casa Rocha).',
        '• DETALLE: memoria paso a paso del tablero que escojas (fórmula y referencia de cada paso), lista para imprimir.',
        '• CROQUIS: planta, corte y alzado del tablero escogido, a escala, que se redibujan solos al cambiar de tablero.',
    ]:
        S.bullet(t)
    S.gap()
    S.section('2. ALCANCE Y LÍMITES — LÉELO')
    for t in [
        '• Método de coeficientes (NTC-Concreto 2023, 3.4.2 y Tabla 3.4.2.1) para tableros rectangulares APOYADOS EN SU PERÍMETRO sobre muros o trabes. Sólo está transcrito el CASO I (losa colada monolíticamente con sus apoyos); un tablero marcado como caso II sale FUERA DE ALCANCE.',
        '• Condición 3.4.2.1 c): los momentos negativos de tableros vecinos en un borde común no deben diferir más de 50 % del menor. La hoja NO conoce a los vecinos: captura el M⁻ del vecino (columna opcional) y la hoja arma con el mayor y te avisa si difieren más de 50 %. Si es así, distribuye el desequilibrio (3.4.3.1) o confirma con un modelo.',
        '• Condición 3.4.2.1 d) (W_m ≤ 2.5·CM aprox.) y las demás de 3.4.2.1: revísalas tú; la hoja supone tableros rectangulares de cargas uniformes. Para tableros irregulares usa un rectángulo representativo.',
        '• Claros A EJES (a_x, a_y), del lado seguro: la tabla permite claros libres en el caso I. Los claros libres (L_n = a − b_p) se usan en cortante, patín, espesor mínimo y bastones.',
        '• VOLADO: un solo borde empotrado (márcalo CONTINUO; los otros tres DISCONTINUO). Su largo es el claro perpendicular a ese borde, del eje del apoyo a la orilla. El bastón entra al tablero de respaldo hasta pasar el punto donde se anula el momento con 0.9·PP y sin carga viva.',
        '• No incluye: trabes y nervaduras principales, flechas calculadas (sólo el espesor mínimo que permite omitirlas), vibración, sismo, losas postensadas, cargas concentradas grandes (tinacos, equipos) ni huecos dentro del tablero.',
    ]:
        S.bullet(t)
    S.gap()
    S.section('3. COLORES')
    r = S.r
    cell(ws, r, 2, 'Celda de captura', 'input', al='left', sz=10); ws.merge_cells('C%d:F%d' % (r, r)); cell(ws, r, 3, 'Amarillo con número azul: aquí se escribe. Todo lo demás es fórmula.', al='left', sz=9); S.r += 1
    r = S.r
    cell(ws, r, 2, 'Fórmula', 'calc', al='left', sz=10); ws.merge_cells('C%d:F%d' % (r, r)); cell(ws, r, 3, 'Negro: calculado. No se escribe encima.', al='left', sz=9); S.r += 1
    r = S.r
    cell(ws, r, 2, 'Resultado clave', 'key', al='left', sz=10); ws.merge_cells('C%d:F%d' % (r, r)); cell(ws, r, 3, 'Azul pálido: resultado que importa (w_u, Mu/M_R, armado).', al='left', sz=9); S.r += 1
    for txt, desc in [('CUMPLE', 'Verde: pasa.'), ('MACIZAR 1 hilera(s)', 'Ámbar: pasa con una acción (macizar, calcular flecha).'), ('NO CUMPLE', 'Rojo: no pasa, faltan datos o está fuera de alcance.')]:
        r = S.r
        cell(ws, r, 2, txt, 'calc', al='left', sz=10); ws.merge_cells('C%d:F%d' % (r, r)); cell(ws, r, 3, desc, al='left', sz=9)
        add_cf(ws, 'B%d' % r); S.r += 1
    S.gap()
    S.section('4. CÓMO SE USA, PASO A PASO')
    for t in [
        '1) DATOS: captura materiales, factores, recubrimientos, la sección nervada (a_c, h_c, b_w, t), la malla y los valores por omisión de macizas. Revisa la lista de destinos (cargas vivas).',
        '2) NERVADAS o MACIZAS: un renglón por tablero. Borra o sobrescribe los renglones «EJEMPLO». Captura nombre, tipo (PERIMETRAL o VOLADO), destino, a_x y a_y a ejes, '
        'y cada borde como CONTINUO (sigue losa del otro lado) o DISCONTINUO (orilla, hueco o losa que no sigue). El caso de la tabla sale solo (puedes forzarlo).',
        '3) Captura la sobrecarga muerta SCM (acabados, rellenos, plafón, muros divisorios repartidos). La carga viva sale del destino; si quieres otra, escríbela en «CV propia».',
        '4) Volados: carga lineal en la orilla (barandal, pretil) y claro del tablero de respaldo L_b. En nervadas puedes darle al volado otra b_w (20 cm suele hacer falta): '
        'la columna «Respaldo del volado» dice cuántas hileras del tablero de respaldo siguen con esa nervadura; anótalas también en el renglón de ese tablero.',
        '5) Lee ESTADO y NOTA. Si dice NO CUMPLE, la nota y las columnas en rojo dicen qué falla (flexión, no caben, ε_t, cortante, peralte).',
        '6) DETALLE: escoge el tablero en la celda amarilla y obtén la memoria paso a paso. CROQUIS dibuja ese mismo tablero.',
        '7) Imprime DETALLE y CROQUIS (ya tienen área de impresión) y pasa a planos el armado de la columna «ARMADO».',
    ]:
        S.bullet(t)
    S.gap()
    S.section('5. REVISIONES QUE HACE CADA RENGLÓN')
    for t in [
        '• Flexión en los cuatro lechos (inferior y superior, claro corto y largo): Mu ≤ M_R, As ≥ mínima, p ≤ p_max y ε_t ≥ 0.005 (F_R = 0.9) — 3.8.2, 6.3.5, 6.7.6.1.',
        '• Que las varillas quepan con la separación libre mínima (nervadas) o que la separación no pase de la máxima (macizas) — 14.2.1.',
        '• Cortante a d del paño: V_u ≤ V_cR (ec. 5.5.3.2.1); en nervadas, hileras de casetones a macizar si no alcanza; ×1.1 sólo si cumple 6.7.8.1.4.',
        '• Losa encasetonada (6.7.8.1.2 a 6.7.8.1.4): si a_c > claro corto/6 el tablero se diseña como losa apoyada en vigas (6.7.8.1.9): recubrimiento de viga y sin ×1.1.',
        '• Espesor mínimo para omitir flechas (Tabla 6.7.3.1.2 con α_fm; h_eq en nervadas); volados L/8 (nervadas) y L/10 (macizas, criterio de esta hoja).',
        '• Capa de compresión y malla al centro de la capa (DATOS sección 7): temperatura, diafragma, que quepa, punzonamiento y flexión con 1 000 kg. Las varillas superiores pasan debajo de la malla.',
        '• Anclaje en bordes discontinuos en la dala de borde de DATOS: 1.25·L_dh de la varilla inferior, L_dh del bastón y gancho de 180° que quepa en su peralte (6.7.8.1.7, 6.7.7.4.1.1 y Tabla 14.3.1).',
        '• 3.4.2.1 c) contra el M⁻ del vecino, si lo capturas.',
    ]:
        S.bullet(t)
    S.gap()
    S.notice('RESPONSABILIDAD: la hoja es una herramienta de cálculo; el Director Responsable de Obra / Corresponsable revisa datos, hipótesis y resultados. '
             'Los renglones EJEMPLO vienen de la hoja de Casa Rocha (nervadas) y de tableros típicos (macizas) sólo para mostrar cómo se llena.', 42)
    ws.sheet_view.zoomScale = 100
    print_setup(ws, 'B1:F%d' % S.r, landscape=False)
    return ws
