"""Columnas de la hoja MUROS (un renglón por muro) y rangos que comparten todas las hojas."""
from openpyxl.utils import get_column_letter as CL

R0 = 11            # primer renglón de muros
NW = 60            # muros que caben
R1 = R0 + NW - 1
PF0 = 9            # primer renglón de PIER_FORCES
NPF = 3000         # renglones de PIER_FORCES
PF1 = PF0 + NPF - 1
DR0 = 16           # primer renglón de DISTORSIONES
NDR = 300
DR1 = DR0 + NDR - 1
NV = 4             # vanos por muro
NT = NV + 1        # tramos por muro

# (clave, encabezado, ancho, unidad)
SPEC = [
    ('key', 'Muro (etiqueta del pier en ETABS)', 13, ''), ('st', 'Story de ETABS', 9, ''), ('dir', 'Dirección', 6, 'X / Y'),
    ('pos', 'Posición (para F_E)', 10, ''),
    ('L', 'L — longitud total', 7, 'm'), ('Hi', 'H — altura libre (vacío = DATOS)', 7.5, 'm'), ('ti', 't — espesor (vacío = DATOS)', 7.5, 'cm'),
]
for _j in range(1, NV + 1):
    SPEC += [('v%dt' % _j, 'Vano %d — tipo' % _j, 9, ''), ('v%dx' % _j, 'Vano %d — x desde el extremo izquierdo' % _j, 7.5, 'm'),
             ('v%dw' % _j, 'Vano %d — ancho' % _j, 6.5, 'm'), ('v%da' % _j, 'Vano %d — antepecho (puerta = 0)' % _j, 7, 'm'),
             ('v%dh' % _j, 'Vano %d — alto del vano' % _j, 6.5, 'm')]
SPEC += [
    ('nce', 'Varillas por castillo extremo (vacío = DATOS)', 7.5, 'pzas'), ('vce', 'No. de varilla del castillo extremo (vacío = DATOS)', 7.5, 'No.'),
    ('rht', 'Refuerzo horizontal', 13, ''), ('rhn', 'Barras por junta (malla: caras)', 7, 'pzas'), ('rhv', 'Varilla / alambre / malla (vacío = DATOS)', 10, ''),
    ('rhs', 'Cada cuántas hiladas (varilla)', 7, 'hiladas'),
    ('sv', '¿Desplanta sobre viga? (4.1.18)', 8, 'SÍ / NO'), ('vnom', 'Viga', 8, ''), ('bv', 'b de la viga', 6, 'cm'), ('hv', 'h de la viga', 6, 'cm'),
    ('Lv', 'L_v — claro de la viga', 6.5, 'm'), ('wv', 'w_u sobre la viga sin el muro', 7, 't/m'), ('cv', 'c — del eje del muro al centro de la viga', 7.5, 'm'),
    ('PD', 'Carga permanente del muro sin factor (vacío = P_u grav./FC)', 8.5, 't'),
    ('cVu', 'V_u', 7, 't'), ('cMu', 'M_u (en el plano)', 7, 't·m'), ('cP', 'P mínima SIN factor (para V_mR)', 8, 't'),
    ('cPmx', 'P_u máxima (compresión)', 7.5, 't'), ('cPmn', 'P_u mínima factorizada (vacío = P·1.1)', 8, 't'), ('cPg', 'P_u de gravedad (4.1.18; vacío = P_u máx.)', 8, 't'),
    ('cVg', 'V_u de la carga vertical (opcional)', 7.5, 't'),
    ('kst', 'Story|Muro', 15, ''), ('eN', 'Renglones de sismo', 6.5, ''), ('eNg', 'Renglones de gravedad', 6.5, ''), ('eVu', 'V_u máx. (sismo)', 7, 't'),
    ('eMu', 'M_u máx. (sismo)', 7, 't·m'), ('ePmn', 'P_u mín. (sismo)', 7, 't'), ('ePmx', 'P_u máx. (sismo)', 7, 't'), ('ePg', 'P_u máx. (gravedad)', 7, 't'),
    ('eVg', 'V_u máx. (gravedad)', 7, 't'), ('eRw', 'Renglón del V_u máx.', 6.5, ''), ('eCb', 'Combinación y lugar del V_u máx.', 16, ''),
    ('ok', '¿Datos completos?', 7, ''), ('fsr', 'Fuerzas de', 12, ''), ('Vu', 'V_u', 7, 't'), ('Mu', 'M_u', 7, 't·m'), ('P', 'P mín. sin factor', 7, 't'),
    ('Pmn', 'P_u mín.', 7, 't'), ('Pmx', 'P_u máx.', 7, 't'), ('Pg', 'P_u gravedad', 7, 't'), ('Vg', 'V_u por carga vertical', 7, 't'), ('pg', '% del V_u por carga vertical', 7, '%'),
    ('H', 'H usada', 6, 'm'), ('t', 't usado', 6, 'cm'), ('nv', 'Vanos', 5, ''), ('vok', 'Vanos bien capturados', 12, ''),
]
for _i in range(1, NT + 1):
    SPEC.append(('T%d' % _i, 'Tramo %d en la sección crítica' % _i, 6.5, 'm'))
SPEC += [
    ('Lc', 'ΣL de tramos que cuentan (≥ L mín.)', 7.5, 'm'), ('ncnt', 'Tramos que cuentan', 6, ''), ('Ltot', 'ΣL de todos los tramos', 7, 'm'),
    ('AT', 'A_T que cuenta (cortante)', 8, 'cm²'), ('ATt', 'A_T de todos los tramos', 8, 'cm²'), ('Ht', 'H/t', 5.5, ''), ('gok', 'Espesor y esbeltez (6.1.3)', 13, ''),
    ('smax', 'Separación máx. de castillos', 7, 'm'), ('nci', 'Castillos intermedios', 7, 'pzas'), ('nck', 'Castillos en el muro', 7, 'pzas'),
    ('Ase', 'A_s de un castillo extremo', 7, 'cm²'), ('AsT', 'ΣA_s de todos los castillos', 7, 'cm²'), ('H0', 'H₀ (estribos cerrados)', 6.5, 'cm'),
    ('FE', 'F_E', 5.5, ''), ('PR', 'P_R', 7.5, 't'), ('rPR', 'P_u / P_R', 6.5, '≤ 1'),
    ('dp', "d' entre castillos extremos", 7.5, 'cm'), ('dd', 'd', 7, 'cm'), ('M0', 'M₀ = A_s·f_y·d\'', 7.5, 't·m'), ('MRa', 'M_R con P_u mín.', 7.5, 't·m'),
    ('MRb', 'M_R con P_u máx.', 7.5, 't·m'), ('MR', 'M_R', 7.5, 't·m'), ('rM', 'M_u / M_R', 6.5, '≤ 1'),
]
for _i in range(1, NT + 1):
    SPEC.append(('f%d' % _i, 'f tramo %d' % _i, 5.5, ''))
for _i in range(1, NT + 1):
    SPEC.append(('V%d' % _i, 'V_mR tramo %d' % _i, 6.5, 't'))
SPEC += [('VmR', 'V_mR = Σ tramos', 7.5, 't'),
         ('Ash', 'A_sh por junta (malla: por cm)', 7.5, 'cm²'), ('sh', 's_h', 6, 'cm'), ('phf', 'p_h·f_yh', 7, 'kg/cm²'), ('phfe', '(p_h·f_yh)e', 7, 'kg/cm²'),
         ('etas', 'η_s', 5.5, ''), ('k1', 'k₁', 5.5, '')]
for _i in range(1, NT + 1):
    SPEC.append(('k0%d' % _i, 'k₀ tramo %d' % _i, 5.5, ''))
for _i in range(1, NT + 1):
    SPEC.append(('S%d' % _i, 'V_sR tramo %d' % _i, 6.5, 't'))
SPEC += [
    ('VsR', 'V_sR = Σ tramos', 7.5, 't'), ('rhok', 'Refuerzo horizontal (6.5.3 / 6.5.4)', 18, ''),
    ('VR', 'V_R = V_mR + V_sR', 7.5, 't'), ('rV', 'V_u / V_R', 6.5, '≤ 1'), ('phreq', 'p_h·f_yh necesario para V_u', 8, 'kg/cm²'), ('rhsug', 'Qué refuerzo horizontal le falta', 26, ''),
    ('Iv', 'I de la viga', 8, 'cm⁴'), ('av', 'a = L/L_v', 5.5, ''), ('Wp', 'W_p', 6.5, 'kg/cm'), ('fI', 'f_I', 5.5, ''), ('K', 'K', 5.5, ''),
    ('kc', 'k_c', 5.5, ''), ('FCE', 'F_CE', 5.5, ''), ('fCM', 'f_CM', 6.5, 'kg/cm²'), ('Lhr', 'L_v / h_v', 6, '≤ 14'), ('Lhd', 'L_v / h_v máx. por flecha', 7, ''),
    ('m418', 'Muro sobre viga (4.1.18)', 18, ''),
    ('Am', 'Mampostería neta', 7, 'm²'), ('Vcc', 'Concreto de castillos', 7, 'm³'), ('Kgc', 'Acero de castillos', 7, 'kg'), ('Kgh', 'Refuerzo horizontal', 7, 'kg'),
    ('fal', 'Lo que no cumple', 20, ''), ('pen', 'Pendiente o aviso', 20, ''), ('est', 'ESTADO', 32, ''), ('nota', 'NOTA', 40, ''),
]
COL = {}
COLS = {}
for _i, (_k, _t, _w, _u) in enumerate(SPEC):
    COL[_k] = CL(2 + _i)
    COLS[_k] = (_t, _w, _u)

INPUTS = ['key', 'st', 'dir', 'pos', 'L', 'Hi', 'ti'] + ['v%d%s' % (j, c) for j in range(1, NV + 1) for c in 'txwah'] + \
         ['nce', 'vce', 'rht', 'rhn', 'rhv', 'rhs', 'sv', 'vnom', 'bv', 'hv', 'Lv', 'wv', 'cv', 'PD', 'cVu', 'cMu', 'cP', 'cPmx', 'cPmn', 'cPg', 'cVg', 'nota']

GROUPS = [('key', 'ti', 'MURO — CAPTURA'), ('v1t', 'v%dh' % NV, 'VANOS DE IZQUIERDA A DERECHA — CAPTURA'), ('nce', 'rhs', 'CASTILLOS EXTREMOS Y REFUERZO HORIZONTAL — CAPTURA'),
          ('sv', 'PD', 'VIGA DE APOYO — CAPTURA'), ('cVu', 'cVg', 'FUERZAS CAPTURADAS (si DATOS dice CAPTURA)'), ('kst', 'eCb', 'FUERZAS QUE SALEN DE PIER_FORCES (ETABS)'),
          ('ok', 'pg', 'FUERZAS DE DISEÑO QUE SE USAN'), ('H', 'gok', 'SECCIÓN CRÍTICA'), ('smax', 'H0', 'CASTILLOS (6.2)'), ('FE', 'rPR', 'COMPRESIÓN (6.4.1)'),
          ('dp', 'rM', 'FLEXOCOMPRESIÓN EN EL PLANO (6.4.2.2)'), ('f1', 'VmR', 'CORTANTE DE LA MAMPOSTERÍA (6.5.2)'), ('Ash', 'rhok', 'REFUERZO HORIZONTAL (6.5.3 Y 6.5.4)'),
          ('VR', 'rhsug', 'CORTANTE TOTAL'), ('Iv', 'm418', 'MURO SOBRE VIGA (4.1.18)'), ('Am', 'Kgh', 'CANTIDADES'), ('fal', 'nota', 'RESULTADO')]


def mr(c, r0=R0, r1=R1, sheet='MUROS'):
    """Rango absoluto de una columna de MUROS."""
    return "%s!$%s$%d:$%s$%d" % (sheet, COL[c], r0, COL[c], r1)
