"""CROQUIS (dibujos a escala del muro escogido en DETALLE, se redibujan solos) y CROQUIS_DATOS (sus coordenadas).

Mismo método que la hoja de cimentación del espectacular y las de losas: gráficas de dispersión cuyas coordenadas son fórmulas vivas.
Cada trazo es un par de columnas de CROQUIS_DATOS; las filas que sobran repiten el último punto y un trazo que no aplica se manda a −1000.
Coordenadas en cm del muro: s a lo largo (0 = extremo izquierdo), z hacia arriba (0 = desplante); en la planta, y a través del espesor.
"""
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.series import SeriesLabel, StrRef
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.marker import Marker
from openpyxl.chart.legend import LegendEntry
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.title import Title
from openpyxl.chart.text import Text, RichText
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties, RegularTextRun
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import cm_to_EMU
from openpyxl.utils import column_index_from_string as CI
from openpyxl.styles import Alignment
from openpyxl.worksheet.pagebreak import Break
from muros_comun import *        # noqa: F401,F403
from muros_detalle import V

CD = 'CROQUIS_DATOS'
Y0 = 8            # primera fila de coordenadas
NPT = 1200        # filas de coordenadas
NK = 16           # castillos que se dibujan
NZ = 34           # estribos por castillo que se dibujan
PPC = 2 * NZ + 2  # puntos por castillo en el trazo de estribos
NL = 30           # líneas de refuerzo horizontal por tramo
PPT = 2 * NL + 2
NH = 70           # líneas de achurado por rectángulo
NDS = 140         # estribos de dala que se dibujan
LAB0 = Y0 + NPT + 4
PAR0 = LAB0 + 260   # parámetros del dibujo
ROW = dict(t=4, a1=6, a2=37, pl=68, sec=86, dat=110)   # renglones de la hoja CROQUIS


def croquis_params():
    """Parámetros del dibujo (en CROQUIS_DATOS): todo enlazado a DETALLE, MUROS y DATOS, en cm."""
    P = []

    def kp(label, formula, nm, fmt='0.0', unit='cm', how='', ref=''):
        P.append((label, formula, nm, fmt, unit))
    kp('¿Hay muro válido?', '=IF({d_ok}="SÍ",1,0)', 'k_ok', '0', '—')
    kp('Muro', '=IF({k_ok}=1,{d_sel},"(escoge en DETALLE un muro con datos completos)")', 'k_key', '@', '—')
    kp('Estado', '={d_est}', 'k_est', '@', '—')
    kp('L — longitud', '=IF({k_ok}=1,' + V('L') + '*100,300)', 'k_L')
    kp('H — altura libre', '=IF({k_ok}=1,' + V('H') + '*100,250)', 'k_H')
    kp('t — espesor', '=IF({k_ok}=1,' + V('t') + ',15)', 'k_t')
    kp('Vanos', '=IF({k_ok}=1,' + V('nv') + ',0)', 'k_nv', '0', '—')
    for j in range(1, NV + 1):
        kp('Vano %d: x / ancho / antepecho / alto' % j, '=IF({k_nv}>=%d,%s*100,0)' % (j, V('v%dx' % j)), 'k_v%dx' % j)
        kp('…', '=IF({k_nv}>=%d,%s*100,0)' % (j, V('v%dw' % j)), 'k_v%dw' % j)
        kp('…', '=IF({k_nv}>=%d,N(%s)*100,0)' % (j, V('v%da' % j)), 'k_v%da' % j)
        kp('…', '=IF({k_nv}>=%d,%s*100,0)' % (j, V('v%dh' % j)), 'k_v%dh' % j)
        kp('…tipo', '=IF({k_nv}>=%d,%s,"")' % (j, V('v%dt' % j)), 'k_v%dt' % j, '@', '—')
    for i in range(1, NT + 1):
        kp('Tramo %d: inicio / largo / ¿cuenta? / castillos' % i, '=0' if i == 1 else '=IF({k_nv}>=%d,{k_v%dx}+{k_v%dw},{k_L})' % (i - 1, i - 1, i - 1), 'k_s%d' % i)
        kp('…', '=IF({k_ok}=1,' + V('T%d' % i) + '*100,IF(%d=1,{k_L},0))' % i, 'k_T%d' % i)
        kp('…', '=IF({k_T%d}>={Lmin}*100-0.01,1,0)' % i, 'k_c%d' % i, '0', '—')
        kp('…', '=IF({k_T%d}<=0.01,0,IF({k_T%d}<2*{hc},1,2+INT(({k_T%d}-{hc})/{k_smax}-0.000001)))' % (i, i, i), 'k_n%d' % i, '0', 'pzas')
    kp('Separación máx. de castillos', '=MIN(1.5*{k_H},400)', 'k_smax')
    kp('Castillos acumulados por tramo', '={k_n1}', 'k_C1', '0', 'pzas')
    for i in range(2, NT + 1):
        kp('…', '={k_C%d}+{k_n%d}' % (i - 1, i), 'k_C%d' % i, '0', 'pzas')
    kp('Castillos que se dibujan', '=MIN({k_C%d},%d)' % (NT, NK), 'k_K', '0', 'pzas', 'Hasta %d.' % NK)
    kp('h_c / b_c / h_d / b_d', '={hc}', 'k_hc'); kp('…', '={bc}', 'k_bc'); kp('…', '={hd}', 'k_hd'); kp('…', '={bd}', 'k_bd')
    kp('Recubrimiento / d_b castillo / d_b estribo / d_b dala', '={rc}', 'k_rc')
    kp('…', '=' + db('{vc}'), 'k_db', '0.00'); kp('…', '=' + db('{ve}'), 'k_dbe', '0.00'); kp('…', '=' + db('{vd}'), 'k_dbd', '0.00')
    kp('d_b del castillo extremo', '=IF({k_ok}=1,' + db('IF(' + V('vce') + '<>"",' + V('vce') + ',{vc})') + ',{k_db})', 'k_dbx', '0.00')
    kp('s / s₀ / H₀ / s de dala', '={se}', 'k_se'); kp('…', '={se0}', 'k_se0'); kp('…', '=MAX({k_H}/6,40)', 'k_H0'); kp('…', '={sed}', 'k_sed')
    kp('Estribos: en cada extremo / en medio / total', '=MAX(1,INT({k_H0}/{k_se0}))', 'k_nb', '0', 'pzas')
    kp('…', '=MAX(0,ROUNDUP(({k_H}-2*{k_nb}*{k_se0})/{k_se},0)-1)', 'k_nm', '0', 'pzas')
    kp('…', '=MIN(%d,2*{k_nb}+{k_nm})' % NZ, 'k_nz', '0', 'pzas')
    kp('Nivel de las varillas de la dala (inferior / superior)', '={k_H}+{k_rc}+{k_dbe}+{k_dbd}/2', 'k_zd1')
    kp('…', '={k_H}+{k_hd}-{k_rc}-{k_dbe}-{k_dbd}/2', 'k_zd2')
    kp('¿Desplanta sobre viga?', '=IF(AND({k_ok}=1,' + V('sv') + '="SÍ"),1,0)', 'k_sv', '0', '—')
    kp('Viga: b / h / L_v / c', '=IF({k_sv}=1,N(' + V('bv') + '),0)', 'k_bv'); kp('…', '=IF({k_sv}=1,N(' + V('hv') + '),0)', 'k_hv')
    kp('…', '=IF({k_sv}=1,N(' + V('Lv') + ')*100,0)', 'k_Lv'); kp('…', '=IF({k_sv}=1,N(' + V('cv') + ')*100,0)', 'k_cv')
    kp('Viga: inicio / fin (centrada en L/2 + c)', '={k_L}/2+{k_cv}-{k_Lv}/2', 'k_sb0'); kp('…', '={k_sb0}+{k_Lv}', 'k_sb1')
    kp('Elemento de abajo: peralte', '=IF({k_sv}=1,MAX({k_hv},10),{k_hd})', 'k_hb')
    kp('Varillas del elemento de abajo (inferior / superior)', '=-{k_hb}+{k_rc}+{k_dbe}+{k_dbd}/2', 'k_zb1')
    kp('…', '=-{k_rc}-{k_dbe}-{k_dbd}/2', 'k_zb2')
    kp('¿Refuerzo horizontal?', '=IF(AND({k_ok}=1,OR(' + V('rht') + '="VARILLA EN JUNTAS",' + V('rht') + '="MALLA",' + V('rht') + '="ESCALERILLA")),1,0)', 'k_rh', '0', '—')
    kp('Separación de las líneas de refuerzo horizontal que se dibujan', '=IF(' + V('rht') + '="MALLA",15.24,MAX(N(' + V('sh') + '),5))', 'k_shd')
    kp('Líneas por tramo', '=MIN(%d,MAX(0,INT(({k_H}-1)/{k_shd})))' % NL, 'k_nl', '0', 'pzas')
    kp('Estribos de dala que se dibujan / separación', '=MIN(%d,MAX(1,INT({k_L}/{k_sed})))' % NDS, 'k_nds', '0', 'pzas')
    kp('…', '={k_L}/({k_nds}+1)', 'k_sdd')
    # --- escalas
    kp('Alto del alzado (elemento de abajo + muro + dala)', '={k_H}+{k_hd}+{k_hb}', 'k_Hc')
    kp('Exageración vertical de los alzados', '=MAX(1,MIN(3,ROUNDUP({k_L}/({k_Hc}*2.3),0)))', 'k_E', '0', '×')
    for g, (mL, mR, mB, mT) in (('a', (2, 2, 2.5, 2)), ('b', (9, 7, 10, 4))):
        kp('Alzado %s: escala' % g, '=MIN(%g/{k_L},%g/({k_Hc}*{k_E}))' % (100 - mL - mR, 45 - mB - mT), 'k_sc' + g, '0.00000', 'u/cm')
        kp('Alzado %s: origen s / z' % g, '=%g+(%g-{k_L}*{k_sc%s})/2' % (mL, 100 - mL - mR, g), 'k_ox' + g, '0.00', 'u')
        kp('…', '=%g+(%g-{k_Hc}*{k_E}*{k_sc%s})/2+{k_hb}*{k_E}*{k_sc%s}' % (mB, 45 - mB - mT, g, g), 'k_oz' + g, '0.00', 'u')
    kp('Alzado: paso del achurado', '=1.6/{k_scb}', 'k_ph')
    kp('Planta: escala / exageración del espesor', '=92/{k_L}', 'k_sp', '0.00000', 'u/cm')
    kp('…', '=MAX(1,4.5/({k_t}*{k_sp}))', 'k_Ep', '0.00', '×')
    kp('Planta: origen s / y', '=6', 'k_opx', '0.00', 'u'); kp('…', '=9.5', 'k_opy', '0.00', 'u')
    kp('Secciones: escala', '=MIN(84/({k_hc}+45+{k_bd}+18),40/(MAX({k_bc},{k_hd})+14))', 'k_ss', '0.0000', 'u/cm')
    kp('Secciones: x del castillo / x de la dala / y base', '=8', 'k_sx1', '0.00', 'u'); kp('…', '={k_sx1}+({k_hc}+45)*{k_ss}', 'k_sx2', '0.00', 'u'); kp('…', '=12', 'k_sy0', '0.00', 'u')
    return P


PARAMS = []


def register_params():
    PARAMS[:] = croquis_params()
    for i, (lab, f, nm, fmt, unit) in enumerate(PARAMS):
        N[nm] = "%s!$C$%d" % (CD, PAR0 + i)


def sheet_croquis(wb):
    register_params()
    ws = wb.create_sheet('CROQUIS'); S = Narr(ws, N, {'A': 1.78, 'B': 46, 'C': 17, 'D': 9, 'E': 54, 'F': 44})
    S.title('CROQUIS DEL MURO')
    S.subtitle('Dibuja el muro escogido en DETALLE, a escala. Se redibuja solo.', 16)
    r = S.r
    ws.merge_cells('B%d:F%d' % (r, r))
    c = ws['B%d' % r]; c.value = '=HYPERLINK("#DETALLE!%s","▶ Cambiar de muro (DETALLE)")' % cel(N['d_sel']); c.font = font(10, True, color='0000CC')
    # encabezados de los dibujos
    for rr, txt in ((ROW['t'], '="Muro: "&{k_key}&"   ·   L = "&TEXT({k_L}/100,"0.00")&" m   ·   H = "&TEXT({k_H}/100,"0.00")&" m   ·   t = "&{k_t}&" cm   ·   "&{k_est}'),):
        ws.merge_cells('B%d:F%d' % (rr, rr))
        x = ws['B%d' % rr]; x.value = txt.format(**N); x.font = font(11, True, color='FFFFFF'); x.fill = fill(SEC_BG)
        x.alignment = Alignment(vertical='center', wrap_text=True); ws.row_dimensions[rr].height = 34
    # tabla corta de resultados junto a las secciones
    r0 = ROW['sec'] + 1
    hdr_cell(ws, r0, 5, 'Resultado del muro'); hdr_cell(ws, r0, 6, 'Valor')
    res = [('V_u / V_R (cortante)', '=IF({k_ok}=1,TEXT(' + V('Vu') + ',"0.00")&" / "&TEXT(' + V('VR') + ',"0.00")&" t = "&TEXT(' + V('rV') + ',"0.00"),"—")'),
           ('M_u / M_R (flexocompresión)', '=IF({k_ok}=1,TEXT(' + V('Mu') + ',"0.00")&" / "&TEXT(' + V('MR') + ',"0.00")&" t·m = "&TEXT(' + V('rM') + ',"0.00"),"—")'),
           ('P_u / P_R (compresión)', '=IF({k_ok}=1,TEXT(' + V('Pmx') + ',"0.0")&" / "&TEXT(' + V('PR') + ',"0.0")&" t = "&TEXT(' + V('rPR') + ',"0.00"),"—")'),
           ('Tramos que cuentan', '=IF({k_ok}=1,' + V('ncnt') + '&" · ΣL = "&TEXT(' + V('Lc') + ',"0.00")&" m","—")'),
           ('Castillos', '=IF({k_ok}=1,' + V('nck') + '&" pzas ("&' + V('nci') + '&" intermedios)","—")'),
           ('Refuerzo horizontal', '=IF({k_ok}=1,' + V('rhok') + ',"—")'),
           ('Qué refuerzo horizontal le falta', '=IF({k_ok}=1,' + V('rhsug') + ',"—")'),
           ('Muro sobre viga (4.1.18)', '=IF({k_ok}=1,' + V('m418') + ',"—")'),
           ('ESTADO', '={k_est}')]
    for i, (t_, f_) in enumerate(res):
        rr = r0 + 1 + i
        cell(ws, rr, 5, t_, 'calc', '@', al='left', bold=True)
        c = cell(ws, rr, 6, f_.format(**N), 'key' if t_ == 'ESTADO' else 'calc', '@', al='left', wrap=True)
        ws.row_dimensions[rr].height = 30
    add_cf(ws, 'F%d:F%d' % (r0 + 1, r0 + len(res)))
    ws.row_breaks.append(Break(id=ROW['a2'] - 1)); ws.row_breaks.append(Break(id=ROW['pl'] - 1)); ws.row_breaks.append(Break(id=ROW['dat'] - 2))
    ws.sheet_view.zoomScale = 75
    print_setup(ws, 'B1:F%d' % (ROW['dat'] - 2), landscape=True)
    ws.page_setup.fitToHeight = 0
    return ws


# =====================================================================
def build_croquis_datos(wb, wsK):
    ws = wb.create_sheet(CD); ws.sheet_view.showGridLines = False
    ws['A1'] = 'COORDENADAS DE LOS DIBUJOS DE LA HOJA CROQUIS (las genera el libro; no se tocan)'; ws['A1'].font = font(12, True, color=HDR_FG)
    ws['A2'] = ('Cada par de columnas es un trazo continuo; las filas que sobran repiten el último punto y un trazo que no aplica se manda a −1000 (fuera del dibujo). '
                'Columnas B a H: los castillos del muro (tramo, posición del centro y ancho) y la altura de cada estribo.')
    ws['A2'].font = font(9, i=True, color=GRAY)
    ws.column_dimensions['A'].width = 52
    ws['A%d' % (PAR0 - 1)] = 'PARÁMETROS DEL DIBUJO (cm, enlazados a DETALLE, MUROS y DATOS)'; ws['A%d' % (PAR0 - 1)].font = font(10, True, color=HDR_FG)
    for i, (lab, f, nm, fmt, unit) in enumerate(PARAMS):
        r = PAR0 + i
        ws.cell(r, 1, lab).font = font(8); ws.cell(r, 2, unit).font = font(8, color=GRAY)
        c = ws.cell(r, 3, f.format(**N) if isinstance(f, str) and f.startswith('=') else f); c.number_format = fmt; c.font = font(8)
    for i in range(NPT):
        ws['A%d' % (Y0 + i)] = i + 1
    ws['A4'] = 'k'
    # ---------------- tabla auxiliar de castillos (B:G) y alturas de estribos (H)
    for c, h in zip('BCDEFGH', ('castillo q', 'tramo', 'm en el tramo', 'centro s (cm)', 'ancho (cm)', '¿se dibuja?', 'z del estribo k (cm)')):
        ws['%s4' % c] = h; ws['%s4' % c].font = font(8, True)
    for q in range(1, NK + 1):
        r = Y0 + q - 1
        ws['B%d' % r] = q
        ws['C%d' % r] = ('=1+(B{r}>{k_C1})+(B{r}>{k_C2})+(B{r}>{k_C3})+(B{r}>{k_C4})'.replace('{r}', str(r))).format(**N)
        ws['D%d' % r] = ('=B{r}-CHOOSE(MIN(C{r},5),0,{k_C1},{k_C2},{k_C3},{k_C4})'.replace('{r}', str(r))).format(**N)
        nq = 'CHOOSE(MIN(C{r},5),{k_n1},{k_n2},{k_n3},{k_n4},{k_n5})'.replace('{r}', str(r))
        sq = 'CHOOSE(MIN(C{r},5),{k_s1},{k_s2},{k_s3},{k_s4},{k_s5})'.replace('{r}', str(r))
        tq = 'CHOOSE(MIN(C{r},5),{k_T1},{k_T2},{k_T3},{k_T4},{k_T5})'.replace('{r}', str(r))
        ws['E%d' % r] = ('=IF(G{r}=0,0,%s+IF(%s=1,%s/2,{k_hc}/2+(D{r}-1)*(%s-{k_hc})/MAX(%s-1,1)))'.replace('{r}', str(r)) % (sq, nq, tq, tq, nq)).format(**N)
        ws['F%d' % r] = ('=IF(G{r}=0,0,IF(%s=1,MIN({k_hc},%s),{k_hc}))'.replace('{r}', str(r)) % (nq, tq)).format(**N)
        ws['G%d' % r] = ('=IF(AND({k_ok}=1,B{r}<={k_K}),1,0)'.replace('{r}', str(r))).format(**N)
    for k in range(1, NZ + 1):
        r = Y0 + k - 1
        ws['H%d' % r] = ('=IF({k}<={k_nb},{k}*{k_se0},IF({k}<={k_nb}+{k_nm},{k_nb}*{k_se0}+({k}-{k_nb})*({k_H}-2*{k_nb}*{k_se0})/({k_nm}+1),{k_H}-(2*{k_nb}+{k_nm}+1-{k})*{k_se0}))'
                         .replace('{k}', str(k)).format(**N))
    CQ = lambda col, q: 'INDEX($%s$%d:$%s$%d,%s)' % (col, Y0, col, Y0 + NK - 1, q)
    ZK = lambda k: 'INDEX($H$%d:$H$%d,%s)' % (Y0, Y0 + NZ - 1, k)
    labels = {}; lab_r = [LAB0]; series = {}; col = [10]

    def label(name, formula):
        r = lab_r[0]; ws.cell(r, 1, formula.format(**N) if formula.startswith('=') else formula)
        labels[name] = '%s!$A$%d' % (CD, r); lab_r[0] += 1

    def dyn(name, xf, yf, npt, lab=None, hide='1', lab_hide=None, lab_else='"—"'):
        cx, cy = CL(col[0]), CL(col[0] + 1); col[0] += 2
        ws['%s4' % cx] = name + ' x'; ws['%s4' % cy] = name + ' y'
        for i in range(npt):
            r = Y0 + i; k = '$A%d' % r
            ws['%s%d' % (cx, r)] = ('=IF(AND({k_ok}=1,%s),%s,-1000)' % (hide, xf(k))).format(**N)
            ws['%s%d' % (cy, r)] = ('=IF(AND({k_ok}=1,%s),%s,-1000)' % (hide, yf(k))).format(**N)
        series[name] = (cx, cy, Y0, Y0 + npt - 1)
        if lab:
            lh = lab_hide or hide
            if lh != '1' and lab.startswith('='):
                lab = '=IF(AND({k_ok}=1,%s),%s,%s)' % (lh, lab[1:], lab_else)
            label(name, lab)

    ch_ = lambda k, n, vals: 'CHOOSE(MIN(MAX(%s,1),%d),%s)' % (k, n, ','.join(vals))
    AXa = lambda s: '({k_oxa}+(%s)*{k_sca})' % s
    AZa = lambda z: '({k_oza}+(%s)*{k_E}*{k_sca})' % z
    AXb = lambda s: '({k_oxb}+(%s)*{k_scb})' % s
    AZb = lambda z: '({k_ozb}+(%s)*{k_E}*{k_scb})' % z
    AZEb = lambda ze: '({k_ozb}+(%s)*{k_scb})' % ze      # z ya multiplicada por la exageración
    PX = lambda s: '({k_opx}+(%s)*{k_sp})' % s
    PY = lambda y: '({k_opy}+(%s)*{k_sp}*{k_Ep})' % y
    SX1 = lambda x: '({k_sx1}+(%s)*{k_ss})' % x
    SX2 = lambda x: '({k_sx2}+(%s)*{k_ss})' % x
    SY = lambda y: '({k_sy0}+(%s)*{k_ss})' % y

    def rect(x0, y0, x1, y1, tx, ty):
        return (lambda k: ch_(k, 5, [tx(x0), tx(x1), tx(x1), tx(x0), tx(x0)]),
                lambda k: ch_(k, 5, [ty(y0), ty(y0), ty(y1), ty(y1), ty(y0)]))

    def cota(p0, p1, along_x, at, name, lab, hide, tick=0.9):
        """Cota en unidades del dibujo: de p0 a p1 sobre una línea horizontal (along_x) o vertical en la coordenada at."""
        a = [p0, p0, p0, '(%s+%s)/2' % (p0, p1), p1, p1, p1]
        b = ['%s-%s' % (at, tick), '%s+%s' % (at, tick), at, at, at, '%s+%s' % (at, tick), '%s-%s' % (at, tick)]
        fx = (lambda k: ch_(k, 7, a)) if along_x else (lambda k: ch_(k, 7, b))
        fy = (lambda k: ch_(k, 7, b)) if along_x else (lambda k: ch_(k, 7, a))
        dyn(name, fx, fy, 7, lab, hide)

    def point(name, x, y, lab, hide):
        dyn(name, lambda k: x, lambda k: y, 1, lab, hide)

    def hatch(name, x0, z0, x1, z1, hide):
        """Achurado a 45° (en el dibujo) de un rectángulo de s = x0..x1, z = z0..z1 (cm), en serpentina."""
        Z0 = '(%s)*{k_E}' % z0; Z1 = '(%s)*{k_E}' % z1
        n = 'MAX(1,INT((((%s)-(%s))+(%s-%s))/{k_ph}-0.000001))' % (x1, x0, Z1, Z0)
        j = lambda k: 'MIN(INT((%s-1)/2)+1,%s)' % (k, n)
        c = lambda k: '((%s)-%s+%s*{k_ph})' % (x0, Z1, j(k))
        ya = lambda k: 'MAX(%s,(%s)-%s)' % (Z0, x0, c(k))
        yb = lambda k: 'MIN(%s,(%s)-%s)' % (Z1, x1, c(k))
        # primero o segundo punto de la línea: en las impares A→B, en las pares B→A; pasado el final se repite el último
        second = lambda k: 'IF(INT((%s-1)/2)+1>%s,1,MOD(%s-1,2))' % (k, n, k)
        useB = lambda k: 'IF(MOD(%s,2)=1,%s,1-%s)' % (j(k), second(k), second(k))
        yy = lambda k: 'IF(%s=1,%s,%s)' % (useB(k), yb(k), ya(k))
        dyn(name, lambda k: AXb('%s+%s' % (yy(k), c(k))), lambda k: AZEb(yy(k)), 2 * NH, None,
            '%s,(%s)-(%s)>0.5,(%s)-(%s)>0.5' % (hide, x1, x0, z1, z0))

    # ================================================= ALZADO ARMADO (A)
    L_ = '{k_L}'; H_ = '{k_H}'
    # castillos: contorno (la unión entre castillos corre por z = 0, que tapa el contorno del muro)
    def cq_pt(k, ppb, pts):
        q = 'MIN(INT((%s-1)/%d)+1,{k_K})' % (k, ppb)
        j = 'IF(INT((%s-1)/%d)+1>{k_K},%d,MOD(%s-1,%d)+1)' % (k, ppb, ppb, k, ppb)
        return q, j
    cL = lambda q: '(%s-%s/2)' % (CQ('E', q), CQ('F', q))
    cR = lambda q: '(%s+%s/2)' % (CQ('E', q), CQ('F', q))

    def cast_outline(k):
        q, j = cq_pt(k, 4, None)
        return ch_(j, 4, [cL(q), cL(q), cR(q), cR(q)]), ch_(j, 4, ['0', H_, H_, '0'])
    dyn('Acast', lambda k: AXa(cast_outline(k)[0]), lambda k: AZa(cast_outline(k)[1]), 4 * NK,
        '="Castillos "&{hc}&"×"&{bc}&" ("&' + V('nck') + '&")"&IF({k_C5}>%d," — se dibujan %d","")' % (NK, NK))
    # varillas longitudinales de castillos: bajan del nivel de la varilla superior de la dala y suben; la unión corre por esa varilla
    ins = lambda q, extreme: '({k_rc}+{k_dbe}+IF(%s,{k_dbx},{k_db})/2)' % extreme

    def cast_bars(k):
        q, j = cq_pt(k, 4, None)
        ext = 'OR(%s=1,%s={k_K})' % (q, q)
        xl = '(%s+MIN(%s,%s/2-0.5))' % (cL(q), ins(q, ext), CQ('F', q))
        xr = '(%s-MIN(%s,%s/2-0.5))' % (cR(q), ins(q, ext), CQ('F', q))
        return ch_(j, 4, [xl, xl, xr, xr]), ch_(j, 4, ['{k_zd2}', '{k_zb1}', '{k_zb1}', '{k_zd2}'])
    dyn('Acvar', lambda k: AXa(cast_bars(k)[0]), lambda k: AZa(cast_bars(k)[1]), 4 * NK,
        '="Varillas: "&{nc}&" No. "&{vc}&IF(OR(' + V('nce') + '<>"",' + V('vce') + '<>"")," (extremos "&IF(ISNUMBER(' + V('nce') + '),' + V('nce') + ',{nc})&" No. "&IF(' + V('vce') + '<>"",' + V('vce') + ',{vc})&")","")')

    # estribos de castillos: serpentina por castillo, de una cara a la otra; entra y sale por z = H
    def cast_stir(k):
        q, j = cq_pt(k, PPC, None)
        kk = 'MIN(INT((%s-2)/2)+1,{k_nz})' % j
        first = 'MOD(%s-2,2)' % j
        side_tick = 'IF(MOD(%s,2)=1,IF(%s=0,0,1),IF(%s=0,1,0))' % (kk, first, first)     # 0 = izquierda, 1 = derecha
        last_side = 'IF(MOD({k_nz},2)=1,1,0)'
        is_entry = '%s=1' % j
        is_exit = '%s>2*{k_nz}+1' % j
        side = 'IF(%s,0,IF(%s,%s,%s))' % (is_entry, is_exit, last_side, side_tick)
        x = 'IF(%s=0,%s,%s)' % (side, cL(q), cR(q))
        z = 'IF(OR(%s,%s),%s,%s)' % (is_entry, is_exit, H_, ZK(kk))
        return x, z
    dyn('Acest', lambda k: AXa(cast_stir(k)[0]), lambda k: AZa(cast_stir(k)[1]), PPC * NK,
        '="Estribos No. "&{ve}&" @ "&{se}&" (@ "&{se0}&" en los extremos)"')
    # refuerzo horizontal: líneas en las juntas de cada tramo que cuenta, entre las caras de sus castillos extremos
    def rh_pt(k):
        i = 'MIN(INT((%s-1)/%d)+1,5)' % (k, PPT)
        j = 'MOD(%s-1,%d)+1' % (k, PPT)
        si = ch_(i, 5, ['{k_s%d}' % t for t in range(1, NT + 1)])
        ti = ch_(i, 5, ['{k_T%d}' % t for t in range(1, NT + 1)])
        ci = ch_(i, 5, ['{k_c%d}' % t for t in range(1, NT + 1)])
        xa = '(%s+IF(%s=1,{k_hc},0))' % (si, ci); xb = '(%s+%s-IF(%s=1,{k_hc},0))' % (si, ti, ci)
        nl = 'IF(%s=1,{k_nl},0)' % ci
        kk = 'MIN(INT((%s-2)/2)+1,MAXa(%s,1))' % (j, nl)
        first = 'MOD(%s-2,2)' % j
        side_t = 'IF(MOD(%s,2)=1,IF(%s=0,0,1),IF(%s=0,1,0))' % (kk, first, first)
        last_side = 'IF(MOD(%s,2)=1,1,0)' % nl
        is_entry = 'OR(%s=1,%s=0)' % (j, nl)
        is_exit = '%s>2*%s+1' % (j, nl)
        side = 'IF(%s,0,IF(%s,%s,%s))' % (is_entry, is_exit, last_side, side_t)
        x = 'IF(%s=0,%s,%s)' % (side, xa, xb)
        z = 'IF(OR(%s,%s),%s,%s*{k_shd})' % (is_entry, is_exit, H_, kk)
        return x, z
    dyn('Arh', lambda k: AXa(rh_pt(k)[0]), lambda k: AZa(rh_pt(k)[1]), PPT * NT,
        '="Refuerzo horizontal: "&IF(' + V('rht') + '="MALLA","malla "&IF(' + V('rhv') + '<>"",' + V('rhv') + ',{mallad}),IF(ISNUMBER(' + V('rhn') + '),' + V('rhn') + ',2)&" × "&IF(' + V('rhv') + '<>"",' + V('rhv') + ',{rhvd})&" cada "&IF(ISNUMBER(' + V('rhs') + '),' + V('rhs') + ',2)&" hiladas")',
        '{k_rh}=1', lab_else='"Sin refuerzo horizontal"')
    # dala superior: varillas, estribos y contorno
    dyn('Advar', lambda k: AXa(ch_(k, 4, ['{k_rc}', '(%s-{k_rc})' % L_, '(%s-{k_rc})' % L_, '{k_rc}'])), lambda k: AZa(ch_(k, 4, ['{k_zd1}', '{k_zd1}', '{k_zd2}', '{k_zd2}'])), 4,
        '="Dala "&{bd}&"×"&{hd}&": "&{nd}&" No. "&{vd}&", E No. "&{ve}&" @ "&{sed}')
    jj = lambda k: 'MIN(INT((%s-1)/2)+1,{k_nds})' % k
    up = lambda k: 'IF(MOD(%s,2)=1,MOD(%s-1,2),1-MOD(%s-1,2))' % (jj(k), k, k)
    dyn('Adest', lambda k: AXa('%s*{k_sdd}' % jj(k)), lambda k: AZa('IF(INT((%s-1)/2)+1>{k_nds},IF(MOD({k_nds},2)=1,{k_H}+{k_hd},{k_H}),IF(%s=1,{k_H}+{k_hd},{k_H}))' % (k, up(k))), 2 * NDS)
    fx, fy = rect('0', H_, L_, '%s+{k_hd}' % H_, AXa, AZa)
    dyn('Adala', fx, fy, 5)
    # elemento de abajo: dala de desplante o viga
    b0 = 'IF({k_sv}=1,{k_sb0},0)'; b1 = 'IF({k_sv}=1,{k_sb1},%s)' % L_
    dyn('Abvar', lambda k: AXa(ch_(k, 4, ['(%s+{k_rc})' % b0, '(%s-{k_rc})' % b1, '(%s-{k_rc})' % b1, '(%s+{k_rc})' % b0])), lambda k: AZa(ch_(k, 4, ['{k_zb1}', '{k_zb1}', '{k_zb2}', '{k_zb2}'])), 4)
    fx, fy = rect(b0, '-{k_hb}', b1, '0', AXa, AZa)
    dyn('Abase', fx, fy, 5, '=IF({k_sv}=1,"Viga de apoyo "&' + V('vnom') + '&" "&{k_bv}&"×"&{k_hv},"Dala o losa de desplante")')
    for side in (0, 1):
        sx = '{k_sb0}' if side == 0 else '{k_sb1}'
        dyn('Aapo%d' % side, lambda k, sx=sx: AXa(ch_(k, 4, [sx, '%s-10' % sx, '%s+10' % sx, sx])), lambda k: AZa(ch_(k, 4, ['-{k_hb}', '-{k_hb}-14', '-{k_hb}-14', '-{k_hb}'])), 4, None, '{k_sv}=1')
    # vanos, cerramientos y repisones
    for j in range(1, NV + 1):
        x0 = '{k_v%dx}' % j; x1 = '({k_v%dx}+{k_v%dw})' % (j, j); z0 = '{k_v%da}' % j; z1 = '({k_v%da}+{k_v%dh})' % (j, j)
        hj_ = '{k_nv}>=%d' % j
        fx, fy = rect(x0, z0, x1, z1, AXa, AZa)
        dyn('Av%d' % j, fx, fy, 5, '="Vanos con cerramiento (y repisón en ventanas)"' if j == 1 else None, hj_, lab_hide='{k_nv}>=1', lab_else='"Sin vanos"')
        dyn('Avx%d' % j, lambda k, x0=x0, x1=x1: AXa(ch_(k, 5, [x0, x1, x1, x0, x1])), lambda k, z0=z0, z1=z1: AZa(ch_(k, 5, [z0, z1, z1, z0, z0])), 5, None, hj_)
        fx, fy = rect('(%s-{k_hc})' % x0, z1, '(%s+{k_hc})' % x1, '(%s+{k_hd})' % z1, AXa, AZa)
        dyn('Acer%d' % j, fx, fy, 5, None, '%s,%s+{k_hd}<{k_H}-2' % (hj_, z1))
        fx, fy = rect('(%s-{k_hc})' % x0, '(%s-10)' % z0, '(%s+{k_hc})' % x1, z0, AXa, AZa)
        dyn('Arep%d' % j, fx, fy, 5, None, '%s,%s>=12' % (hj_, z0))
    fx, fy = rect('0', '0', L_, H_, AXa, AZa)
    dyn('Amuro', fx, fy, 5, '="Muro "&{k_key}&" · t = "&{k_t}&" cm"&IF({k_E}>1," · escala vertical ×"&{k_E},"")')

    # ================================================= ALZADO ACOTADO (B)
    fx, fy = rect('0', '0', L_, H_, AXb, AZb)
    dyn('Bmuro', fx, fy, 5, '="Mampostería"')
    fx, fy = rect('0', H_, L_, '%s+{k_hd}' % H_, AXb, AZb)
    dyn('Bdala', fx, fy, 5, '="Dala"')
    fx, fy = rect(b0, '-{k_hb}', b1, '0', AXb, AZb)
    dyn('Bbase', fx, fy, 5, '=IF({k_sv}=1,"Viga de apoyo","Dala o losa de desplante")')
    for i in range(1, NT + 1):
        hatch('Bh%d' % i, '{k_s%d}' % i, '0', '({k_s%d}+{k_T%d})' % (i, i), H_, '{k_T%d}>0' % i)
    for j in range(1, NV + 1):
        x0 = '{k_v%dx}' % j; x1 = '({k_v%dx}+{k_v%dw})' % (j, j); z0 = '{k_v%da}' % j; z1 = '({k_v%da}+{k_v%dh})' % (j, j)
        hatch('Bha%d' % j, x0, '0', x1, z0, '{k_nv}>=%d' % j)
        hatch('Bhd%d' % j, x0, z1, x1, H_, '{k_nv}>=%d' % j)
        fx, fy = rect(x0, z0, x1, z1, AXb, AZb)
        dyn('Bv%d' % j, fx, fy, 5, None, '{k_nv}>=%d' % j)
        point('Bvl%d' % j, AXb('(%s+%s)/2' % (x0, x1)), AZb('(%s+%s)/2' % (z0, z1)),
              '=IF({k_nv}>=%d,{k_v%dt}&" "&TEXT({k_v%dw}/100,"0.00")&" × "&TEXT({k_v%dh}/100,"0.00")&IF({k_v%da}>0," (antepecho "&TEXT({k_v%da}/100,"0.00")&")",""),"")' % (j, j, j, j, j, j),
              '{k_nv}>=%d' % j)
    # cadena de cotas: tramos y vanos alternados, y la total
    zc1 = '(%s-3.5)' % AZb('-{k_hb}'); zc2 = '(%s-7.5)' % AZb('-{k_hb}')
    edges = []
    for i in range(1, NT + 1):
        edges.append(('{k_s%d}' % i, '({k_s%d}+{k_T%d})' % (i, i), '{k_T%d}>0.5,{k_nv}>0' % i, 'TEXT({k_T%d}/100,"0.00")' % i))
        if i <= NV:
            edges.append(('{k_v%dx}' % i, '({k_v%dx}+{k_v%dw})' % (i, i), '{k_nv}>=%d' % i, '"("&TEXT({k_v%dw}/100,"0.00")&")"' % i))
    for n_, (a_, b_, h_, t_) in enumerate(edges):
        cota(AXb(a_), AXb(b_), True, zc1, 'Bc%d' % n_, '=' + t_, h_)
    cota(AXb('0'), AXb(L_), True, zc2, 'BcL', '="L = "&TEXT({k_L}/100,"0.00")&" m"', '1')
    cota(AZb('0'), AZb(H_), False, '(%s-3)' % AXb('0'), 'BcH', '="H = "&TEXT({k_H}/100,"0.00")&" m"', '1')
    cota(AZb(H_), AZb('%s+{k_hd}' % H_), False, '(%s+2)' % AXb(L_), 'Bchd', '="h_d = "&{hd}', '1')
    for j in range(1, NV + 1):
        cota(AZb('0'), AZb('{k_v%da}' % j), False, '(%s+1.6)' % AXb('({k_v%dx}+{k_v%dw})' % (j, j)), 'Bca%d' % j, '="ant. "&TEXT({k_v%da}/100,"0.00")' % j, '{k_nv}>=%d,{k_v%da}>0' % (j, j))
    # posición de los castillos (marcas sobre la dala) y sección crítica
    kx = lambda k: CQ('E', 'MIN(%s,{k_K})' % k)
    dyn('Bkpos', lambda k: AXb(kx(k)), lambda k: '(%s+1.5)' % AZb('%s+{k_hd}' % H_), NK, '="Castillos"')
    zcrit = 'IF({k_nv}>0,MAXb(5,MIN({k_H}-5,({k_v1a}+{k_v1h})/2)),{k_H}/2)'
    dyn('Bcrit', lambda k: AXb(ch_(k, 2, ['-25', '%s+25' % L_])), lambda k: AZb(ch_(k, 2, [zcrit, zcrit])), 2,
        '="Sección crítica: ΣL = "&TEXT(' + V('Lc') + ',"0.00")&" m"')

    # ================================================= PLANTA — SECCIÓN CRÍTICA (P)
    T_ = '{k_t}'
    for i in range(1, NT + 1):
        s0 = '{k_s%d}' % i; s1 = '({k_s%d}+{k_T%d})' % (i, i)
        fx, fy = rect(s0, '0', s1, T_, PX, PY)
        dyn('Pt%d' % i, fx, fy, 5, None, '{k_T%d}>0.5,{k_c%d}=1' % (i, i))
        dyn('Pm%d' % i, fx, fy, 5, '="Machón (no cuenta)"' if i == 1 else None, '{k_T%d}>0.5,{k_c%d}=0' % (i, i),
            lab_hide='OR({k_c1}+{k_c2}+{k_c3}+{k_c4}+{k_c5}<{k_nv}+1)', lab_else='"Tramos de muro"')
        point('Ptl%d' % i, PX('(%s+%s)/2' % (s0, s1)), '(%s+3.2)' % PY(T_),
              '=IF(AND({k_T%d}>0.5,{k_c%d}=1),"T%d: "&TEXT({k_T%d}/100,"0.00")&" m · V_mR "&TEXT(%s,"0.0")&IF(%s>0," + V_sR "&TEXT(%s,"0.0"),"")&" t","")' % (i, i, i, i, V('V%d' % i), V('S%d' % i), V('S%d' % i)),
              '{k_T%d}>0.5,{k_c%d}=1' % (i, i))
    # castillos en planta (cuadros) y sus varillas
    def pc(k):
        q, j = cq_pt(k, 5, None)
        y0 = '(({k_t}-{k_bc})/2)'; y1 = '(({k_t}+{k_bc})/2)'
        return ch_(j, 5, [cL(q), cR(q), cR(q), cL(q), cL(q)]), ch_(j, 5, [y0, y0, y1, y1, y0])
    dyn('Pcast', lambda k: PX(pc(k)[0]), lambda k: PY(pc(k)[1]), 5 * NK, '="Castillos"')

    def pbars(k):
        q = 'INT((%s-1)/4)+1' % k
        m = 'MOD(%s-1,4)' % k
        ext = 'OR(%s=1,%s={k_K})' % (q, q)
        dx = '(%s/2-%s)' % (CQ('F', 'MIN(%s,%d)' % (q, NK)), ins(q, ext))
        dyb = '({k_bc}/2-%s)' % ins(q, ext)
        x = '(%s+IF(OR(%s=0,%s=3),-1,1)*%s)' % (CQ('E', 'MIN(%s,%d)' % (q, NK)), m, m, dx)
        y = '({k_t}/2+IF(%s<2,-1,1)*%s)' % (m, dyb)
        vis = '%s<={k_K}' % q
        return 'IF(%s,%s,-1000)' % (vis, PX(x)), 'IF(%s,%s,-1000)' % (vis, PY(y))
    dyn('Pvar', lambda k: pbars(k)[0], lambda k: pbars(k)[1], 4 * NK, None)
    zp1 = '({k_opy}-3.5)'; zp2 = '({k_opy}-7.5)'
    for n_, (a_, b_, h_, t_) in enumerate(edges):
        cota(PX(a_), PX(b_), True, zp1, 'Pc%d' % n_, '=' + t_, h_)
    cota(PX('0'), PX(L_), True, zp2, 'PcL', '="L = "&TEXT({k_L}/100,"0.00")&" m"', '1')
    cota(PY('0'), PY(T_), False, '(%s-2.5)' % PX('0'), 'Pct', '="t = "&{k_t}&" cm"', '1')
    # ================================================= SECCIONES: CASTILLO Y DALA (S)
    # castillo: h_c a lo largo del muro (horizontal en el dibujo) y b_c a través
    fx, fy = rect('0', '0', '{k_hc}', '{k_bc}', SX1, SY)
    dyn('Sc', fx, fy, 5, '="Castillo "&{hc}&"×"&{bc}&": "&{nc}&" No. "&{vc}&", E No. "&{ve}&" @ "&{se}&" (@ "&{se0}&" en H₀)"')
    e = '({k_rc}-{k_dbe}/2)'
    hook = lambda X, w, h: ([X('%s+3' % e), X(e), X(e), X('%s-%s' % (w, e)), X('%s-%s' % (w, e)), X(e), X(e), X('%s+3.5' % e)],
                            [SY('%s-%s' % (h, e)), SY('%s-%s' % (h, e)), SY(e), SY(e), SY('%s-%s' % (h, e)), SY('%s-%s' % (h, e)), SY('%s-3.5' % '(%s-%s)' % (h, e)), SY('%s-%s-0.5' % (h, e))])
    hx, hy = hook(SX1, '{k_hc}', '{k_bc}')
    dyn('Sce', lambda k: ch_(k, 8, hx), lambda k: ch_(k, 8, hy), 8)
    ic = '({k_rc}+{k_dbe}+{k_db}/2)'

    def bars_rect(X, w, h, n, nb_name):
        # 4, 6 u 8 varillas: esquinas, luego a la mitad de los lados largos y de los cortos
        xs = [X(ic), X('%s-%s' % (w, ic)), X('%s-%s' % (w, ic)), X(ic), X('%s/2' % w), X('%s/2' % w), X(ic), X('%s-%s' % (w, ic))]
        ys = [SY(ic), SY(ic), SY('%s-%s' % (h, ic)), SY('%s-%s' % (h, ic)), SY(ic), SY('%s-%s' % (h, ic)), SY('%s/2' % h), SY('%s/2' % h)]
        dyn(nb_name, lambda k: 'IF(%s<=%s,%s,-1000)' % (k, n, ch_(k, 8, xs)), lambda k: 'IF(%s<=%s,%s,-1000)' % (k, n, ch_(k, 8, ys)), 8)
    bars_rect(SX1, '{k_hc}', '{k_bc}', '{nc}', 'Scv')
    cota(SX1('0'), SX1('{k_hc}'), True, '(%s-3)' % SY('0'), 'Sch', '="h_c = "&{hc}', '1')
    cota(SY('0'), SY('{k_bc}'), False, '(%s-3)' % SX1('0'), 'Scb', '="b_c = "&{bc}', '1', 0.7)
    # muro que llega al castillo (para ver el espesor)
    fx, fy = rect('{k_hc}', '({k_bc}-{k_t})/2', '{k_hc}+30', '({k_bc}+{k_t})/2', SX1, SY)
    dyn('Sm', lambda k: ch_(k, 4, [SX1('{k_hc}+30'), SX1('{k_hc}'), SX1('{k_hc}'), SX1('{k_hc}+30')]),
        lambda k: ch_(k, 4, [SY('({k_bc}-{k_t})/2'), SY('({k_bc}-{k_t})/2'), SY('({k_bc}+{k_t})/2'), SY('({k_bc}+{k_t})/2')]), 4, '="Muro t = "&{k_t}&" cm"')
    # dala: b_d horizontal y h_d vertical
    fx, fy = rect('0', '0', '{k_bd}', '{k_hd}', SX2, SY)
    dyn('Sd', fx, fy, 5, '="Dala "&{bd}&"×"&{hd}&": "&{nd}&" No. "&{vd}&", E No. "&{ve}&" @ "&{sed}')
    hx, hy = hook(SX2, '{k_bd}', '{k_hd}')
    dyn('Sde', lambda k: ch_(k, 8, hx), lambda k: ch_(k, 8, hy), 8)
    ic = '({k_rc}+{k_dbe}+{k_dbd}/2)'
    bars_rect(SX2, '{k_bd}', '{k_hd}', '{nd}', 'Sdv')
    cota(SX2('0'), SX2('{k_bd}'), True, '(%s-3)' % SY('0'), 'Sdb', '="b_d = "&{bd}', '1')
    cota(SY('0'), SY('{k_hd}'), False, '(%s+3)' % SX2('{k_bd}'), 'Sdh', '="h_d = "&{hd}', '1', 0.7)
    point('Src', SX1('{k_hc}/2'), '(%s+5)' % SY('MAX({k_bc},{k_hd})'), '="Recubrimiento "&{rc}&" cm"', '1')
    ws['A%d' % (LAB0 - 1)] = 'TEXTOS DE LEYENDAS Y COTAS'; ws['A%d' % (LAB0 - 1)].font = font(10, True, color=HDR_FG)

    # ================================================= gráficas
    small = lambda sz: RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=sz)), endParaRPr=CharacterProperties(sz=sz))])

    def mk_title(txt, sz=1000):
        cp = CharacterProperties(sz=sz, b=True, solidFill='1F3864')
        return Title(tx=Text(rich=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=txt)])])), overlay=False)

    def mk_chart(title, xmin, xmax, ymin, ymax, w_cm, h_cm, hfrac, ytop=0.09):
        ch = ScatterChart(); ch.style = 13; ch.scatterStyle = 'lineMarker'
        ch.title = mk_title(title)
        ch.x_axis.scaling.min = xmin; ch.x_axis.scaling.max = xmax; ch.y_axis.scaling.min = ymin; ch.y_axis.scaling.max = ymax
        ch.x_axis.delete = True; ch.y_axis.delete = True
        ch.x_axis.majorGridlines = None; ch.y_axis.majorGridlines = None
        ch.width = w_cm; ch.height = h_cm
        ch.legend.position = 'b'
        ch.legend.txPr = small(750)
        ch.display_blanks = 'gap'
        wfrac = (xmax - xmin) / (ymax - ymin) * hfrac * h_cm / w_cm
        if wfrac > 0.97:
            hfrac = hfrac * 0.97 / wfrac; wfrac = 0.97
        ch.layout = Layout(manualLayout=ManualLayout(layoutTarget='inner', xMode='edge', yMode='edge', x=(1 - wfrac) / 2, y=ytop, w=wfrac, h=hfrac))
        return ch

    def add(ch, name, color='1F4E79', w=1.5, dash='solid', marker=None, msize=6, lbl_idx=None, noleg=False, pos='t', lsz=750, bold=False):
        cx, cy, r0, r1 = series[name]
        xr = Reference(ws, min_col=CI(cx), min_row=r0, max_row=r1)
        yr = Reference(ws, min_col=CI(cy), min_row=r0, max_row=r1)
        s = Series(yr, xr)
        if name in labels:
            s.tx = SeriesLabel(strRef=StrRef(labels[name]))
        s.smooth = False
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = int(w * 12700)
        s.graphicalProperties.line.prstDash = dash
        if marker:
            s.marker = Marker(symbol=marker, size=msize)
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
            s.marker.graphicalProperties.line.solidFill = color
            s.graphicalProperties.line = LineProperties(noFill=True)
        else:
            s.marker = Marker(symbol='none')
        if lbl_idx is not None:
            tp = RichText(bodyPr=RichTextProperties(wrap='none'), p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=lsz, b=bold, solidFill=(color if bold else None))), endParaRPr=CharacterProperties(sz=lsz, b=bold))])
            s.dLbls = DataLabelList(dLbl=[DataLabel(idx=lbl_idx, showSerName=True, showVal=False, showCatName=False, showLegendKey=False, showPercent=False, dLblPos=pos, txPr=tp)])
            s.dLbls.showVal = False; s.dLbls.showSerName = False; s.dLbls.showLegendKey = False; s.dLbls.showCatName = False
        if (noleg or name not in labels or lbl_idx is not None) and ch.legend is not None:
            ch.legend.legendEntry.append(LegendEntry(idx=len(ch.series), delete=True))
        ch.series.append(s)
        return s

    def place(ch, row, col=1):
        anchor = OneCellAnchor(_from=AnchorMarker(col=col, colOff=0, row=row - 1, rowOff=0), ext=XDRPositiveSize2D(cm_to_EMU(ch.width), cm_to_EMU(ch.height)))
        wsK.add_chart(ch, anchor)

    RED, GRN, BLU, GRY, PUR, ORG = 'C00000', '00B050', '1F3864', '7F7F7F', '7030A0', 'ED7D31'
    # --- A: alzado armado
    ch = mk_chart('ALZADO — ARMADO', 0, 100, 0, 45, 32.5, 15.5, 0.66)
    add(ch, 'Acest', GRN, 0.75)
    add(ch, 'Arh', ORG, 1.0, 'dash')
    add(ch, 'Acvar', RED, 1.25)
    add(ch, 'Adest', GRN, 0.5, noleg=True)
    add(ch, 'Advar', RED, 1.25)
    add(ch, 'Abvar', RED, 1.25, noleg=True)
    add(ch, 'Acast', BLU, 1.25)
    for j in range(1, NV + 1):
        add(ch, 'Avx%d' % j, 'BFBFBF', 0.5, 'dash', noleg=True)
        add(ch, 'Av%d' % j, GRY, 1.25, noleg=j > 1)
        add(ch, 'Acer%d' % j, PUR, 1.0, noleg=True)
        add(ch, 'Arep%d' % j, PUR, 1.0, noleg=True)
    add(ch, 'Adala', BLU, 1.5, noleg=True)
    add(ch, 'Abase', GRY, 1.75)
    add(ch, 'Aapo0', GRY, 1.25, noleg=True); add(ch, 'Aapo1', GRY, 1.25, noleg=True)
    add(ch, 'Amuro', BLU, 2.25)
    place(ch, ROW['a1'])
    # --- B: alzado acotado
    ch = mk_chart('ALZADO — SILUETA ACOTADA', 0, 100, 0, 45, 32.5, 15.5, 0.66)
    for i in range(1, NT + 1):
        add(ch, 'Bh%d' % i, 'BFBFBF', 0.5, noleg=True)
    for j in range(1, NV + 1):
        add(ch, 'Bha%d' % j, 'BFBFBF', 0.5, noleg=True); add(ch, 'Bhd%d' % j, 'BFBFBF', 0.5, noleg=True)
    add(ch, 'Bdala', '8EA9DB', 1.25)
    add(ch, 'Bbase', GRY, 1.25)
    for j in range(1, NV + 1):
        add(ch, 'Bv%d' % j, GRY, 1.25, noleg=True)
        add(ch, 'Bvl%d' % j, GRY, 0.5, lbl_idx=0, pos='ctr', lsz=700)
    add(ch, 'Bmuro', BLU, 2.0)
    add(ch, 'Bcrit', ORG, 1.25, 'dashDot')
    add(ch, 'Bkpos', BLU, marker='triangle', msize=5)
    for n_ in range(len(edges)):
        add(ch, 'Bc%d' % n_, PUR, 0.75, lbl_idx=3, pos='t', lsz=700)
    add(ch, 'BcL', RED, 0.75, lbl_idx=3, pos='b')
    add(ch, 'BcH', RED, 0.75, lbl_idx=3, pos='l')
    add(ch, 'Bchd', GRY, 0.75, lbl_idx=3, pos='r', lsz=650)
    for j in range(1, NV + 1):
        add(ch, 'Bca%d' % j, PUR, 0.75, lbl_idx=3, pos='r', lsz=650)
    place(ch, ROW['a2'])
    # --- P: planta, sección crítica
    ch = mk_chart('PLANTA — SECCIÓN CRÍTICA (espesor exagerado)', 0, 100, 0, 24, 32.5, 9.0, 0.64, 0.11)
    for i in range(1, NT + 1):
        add(ch, 'Pt%d' % i, BLU, 1.75, noleg=True)
        add(ch, 'Pm%d' % i, ORG, 1.25, 'dash', noleg=i > 1)
        add(ch, 'Ptl%d' % i, BLU, 0.5, lbl_idx=0, pos='t', lsz=650)
    add(ch, 'Pcast', GRY, 1.0)
    add(ch, 'Pvar', RED, marker='circle', msize=3)
    for n_ in range(len(edges)):
        add(ch, 'Pc%d' % n_, PUR, 0.75, lbl_idx=3, pos='t', lsz=650)
    add(ch, 'PcL', RED, 0.75, lbl_idx=3, pos='b', lsz=650)
    add(ch, 'Pct', GRY, 0.75, lbl_idx=3, pos='l', lsz=600)
    place(ch, ROW['pl'])
    # --- S: secciones
    ch = mk_chart('SECCIONES — CASTILLO Y DALA (cm)', 0, 100, 0, 60, 13.4, 11.0, 0.62)
    add(ch, 'Sm', 'BFBFBF', 1.0)
    add(ch, 'Sc', BLU, 2.0)
    add(ch, 'Sce', GRN, 1.25, noleg=True)
    add(ch, 'Scv', RED, marker='circle', msize=7, noleg=True)
    add(ch, 'Sd', '2E75B6', 2.0)
    add(ch, 'Sde', GRN, 1.25, noleg=True)
    add(ch, 'Sdv', RED, marker='circle', msize=7, noleg=True)
    for nm, ps in (('Sch', 'b'), ('Scb', 'l'), ('Sdb', 'b'), ('Sdh', 'r')):
        add(ch, nm, GRY, 0.75, lbl_idx=3, pos=ps, lsz=700)
    add(ch, 'Src', GRY, 0.5, lbl_idx=0, pos='t', lsz=650)
    place(ch, ROW['sec'])
    return ws
