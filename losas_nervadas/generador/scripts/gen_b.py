"""Parte B: TABLEROS, NERVADURA, PRINCIPALES, TRIB, TRANSFER, MUROS."""
from gen_a import *
from openpyxl.utils import column_index_from_string as CI

VOLB = {'F80012': 3.029, 'F80019': 4.0}   # claro de respaldo (m) del volado: tablero más corto detrás del muro raíz
for _t in TABS:
    if _t['modo'] == 'VOLADO':
        _t['volado'].setdefault('Lb', VOLB[_t['name']])


# ---------------------------------------------------------------- TABLEROS
TC = {}
BLKS = ('Pc', 'Nc', 'Pl', 'Nl')
_tcols = ['key', 'nivel', 'modo', 'dest', 'Lx', 'Ly', 'ax', 'ay', 'a1', 'a2', 'm', 'dirc', 'ln1', 'ln2', 'bfc', 'bfl', 'rb', 'rt',
          'scm', 'cv', 'pp', 'cm', 'ws', 'wu',
          'nd', 'ndc', 'ndl', 'cauto', 'cforce', 'caso', 'idx', 'fr', 'cNIC', 'cNIL', 'cNDC', 'cNDL', 'cPC', 'cPL', 'pbor', 'MPc', 'MNc', 'MPl', 'MNl']
for _b in BLKS:
    _tcols += ['As' + _b, 'v' + _b, 'n' + _b, 'd' + _b, 'MR' + _b, 'FU' + _b, 'e' + _b]
_tcols += ['g678', 'Vu', 'VcR', 'cort', 'afm', 'hmin', 'heqv', 'peral', 'tdia', 'Lbc', 'Lbl', 'ncx', 'ncy', 'apoy', 'estado', 'nota']
for i, nm in enumerate(_tcols):
    TC[nm] = CL(2 + i)
TC_LAST = CL(2 + len(_tcols) - 1)


def tcol(nm):
    return "TABLEROS!$%s$%d:$%s$%d" % (TC[nm], T0, TC[nm], T1)


def asflex(M, b, d, neg):
    """As por flexión (cm²) para M en t·m; -1 si la sección no alcanza o rebasa p_max (ε_t ≥ 0.005)."""
    rad = '(1-2*%s*1E5/({FRf}*%s*%s^2*{fpc}))' % (M, b, d)
    asf = '(1-SQRT(%s))*{fpc}*%s*%s/{fy}' % (rad, b, d)
    return 'IF(%s<=0,0,IF(%s<0,-1,IF(%s>{pmax}*%s*%s,-1,%s)))' % (M, rad, asf, b, d, asf)


def asreq(M, b, d, neg, bmin='{bw}'):
    """As requerida = el mayor de la de flexión y el menor de (As mín de 6.3.5.1.2, 4/3 de la de flexión) — 6.3.5.1.3."""
    f = asflex(M, b, d, neg)
    return 'IF(%s<0,-1,MAX(%s,MIN({pmin}*%s*%s,4/3*%s)))' % (f, f, bmin, d, f)


def nfit(var, avail):
    return 'INT((%s+{slib})/(%s+{slib}))' % (avail, db(var))


def varsel2(As, pref, nmax, avail, nmin='1'):
    """Varilla más chica (desde la preferida) cuyas piezas necesarias (al menos nmin) caben en el ancho libre y no pasan del máximo."""
    i0 = 'MATCH(%s,%s,0)' % (pref, BARCOL)
    out = 'INDEX(%s,%d)' % (BARCOL, NBARS)
    for k in range(4, -1, -1):
        idx = 'MIN(%s+%d,%d)' % (i0, k, NBARS)
        a_ = 'INDEX(%s,%s)' % (ABCOL, idx); d_ = 'INDEX(%s,%s)' % (DBCOL, idx)
        ok = 'AND(%s>=0,MAX(%s,CEILING(%s/%s,1))<=MIN(%s,INT((%s+MAX({slib},%s))/(%s+MAX({slib},%s)))))' % (As, nmin, As, a_, nmax, avail, d_, d_, d_)
        out = 'IF(%s,INDEX(%s,%s),%s)' % (ok, BARCOL, idx, out)
    return out


def edge_af(t, side, k):
    """αf del borde (fórmula): muro completo = 99; nervadura principal = I_b/I_s; nada = 0."""
    e = t['bordes'][side]
    row = EDGE[(k, side)]
    perp = t['ay'] if side in ('inferior', 'superior') else t['ax']
    Is = '(%.3f*100*IF(APOYOS!$G$%d="CONTINUO",1,0.5)*{heq}^3/12)' % (perp, row)
    if e['apoyo'].startswith('MURO') and e['frac_muro'] >= 0.9:
        return '99'
    nps = e.get('nps') or []
    if nps and nps[0] in PROW:
        pr = PROW[nps[0]]
        return '(PRINCIPALES!$%s$%d*PRINCIPALES!$%s$%d^3/12/%s)' % ('{PCb}', pr, '{PCh}', pr, Is)
    return '0'


def sheet_tableros(wb):
    ws = wb.create_sheet('TABLEROS'); ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 1.78
    last = TC_LAST
    ws.merge_cells('B1:%s1' % last)
    c = ws['B1']; c.value = 'TABLEROS — NERVADURAS SECUNDARIAS DE CADA TABLERO (MÉTODO DE COEFICIENTES, NTC-CONCRETO 2023 TABLA 3.4.2.1)'
    c.font = font(13, True, color='FFFFFF'); c.fill = fill(TITLE_BG); c.alignment = Alignment(vertical='center'); ws.row_dimensions[1].height = 26
    ws.merge_cells('B2:%s2' % last)
    ws['B2'].value = ('Un renglón por tablero. Momentos = coeficiente × 10⁻⁴ × w_u × a1² por metro; por nervadura × s. Abajo trabaja la sección T con el patín efectivo de la Tabla 5.2.1.4.2; arriba sólo la nervadura (b_w de cálculo). '
                      'Toda nervadura lleva al menos 1 No. 4 corrida abajo y 1 No. 4 corrida arriba (6.7.6.1.3); arriba, las demás varillas son bastones. Lo AMARILLO se puede cambiar. El detalle paso a paso está en NERVADURA.')
    ws['B2'].font = font(9, i=True, color=GRAY); ws['B2'].alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[2].height = 30
    ws.merge_cells('B3:%s3' % last)
    ws['B3'].value = ('ESTADO: «CUMPLE» = flexión, ε_t ≥ 0.005, cortante, peralte, capa y apoyos bien. «VER PRINCIPALES» = el tablero pasa pero alguna nervadura principal que lo recibe no. «MACIZAR n hilera(s)» = quitar casetones junto al apoyo por cortante. '
                      '«6.7.8: NO» = el claro corto es menor que 6·a_c: la losa se diseña como losa apoyada en vigas (6.7.8.1.9), sin el ×1.1 de V_cR.')
    ws['B3'].font = font(9, True, color=WARN_FG); ws['B3'].fill = fill(WARN_BG); ws['B3'].alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[3].height = 30
    groups = [('key', 'dest', 'TABLERO'), ('Lx', 'rt', 'GEOMETRÍA'), ('scm', 'wu', 'CARGAS (kg/m²)'), ('nd', 'caso', 'CASO DE LA TABLA'),
              ('idx', 'cPL', 'COEFICIENTES ×10⁻⁴ (Tabla 3.4.2.1, interpolados en m)'), ('pbor', 'MNl', 'MOMENTOS POR NERVADURA (t·m)'),
              ('AsPc', 'ePc', 'INFERIOR — CLARO CORTO'), ('AsNc', 'eNc', 'SUPERIOR — CLARO CORTO'), ('AsPl', 'ePl', 'INFERIOR — CLARO LARGO'),
              ('AsNl', 'eNl', 'SUPERIOR — CLARO LARGO'), ('g678', 'cort', 'CORTANTE'), ('afm', 'tdia', 'PERALTE Y CAPA'), ('Lbc', 'nota', 'DETALLE Y RESUMEN')]
    for a, b, t in groups:
        ws.merge_cells('%s6:%s6' % (TC[a], TC[b]))
        x = ws['%s6' % TC[a]]; x.value = t; x.font = font(9, True, color='FFFFFF'); x.fill = fill(SEC_BG); x.alignment = Alignment(horizontal='center')
    heads = {'key': ('Tablero', 15), 'nivel': ('Nivel', 8), 'modo': ('Modo', 12), 'dest': ('Destino', 12), 'Lx': ('Lx', 6.5), 'Ly': ('Ly', 6.5), 'ax': ('a_x usado', 7.5), 'ay': ('a_y usado', 7.5),
             'a1': ('a1 corto', 7), 'a2': ('a2 largo', 7), 'm': ('m = a1/a2', 7), 'dirc': ('Corto ‖', 6), 'ln1': ('Claro libre corto', 7), 'ln2': ('Claro libre largo', 7),
             'bfc': ('Patín b_f corto', 7), 'bfl': ('Patín b_f largo', 7), 'rb': ('r inferior', 6), 'rt': ('r superior', 6),
             'scm': ('SCM', 7), 'cv': ('CV W_m', 7), 'pp': ('PP diseño', 7), 'cm': ('CM', 7), 'ws': ('w serv.', 7), 'wu': ('w_u', 7),
             'nd': ('Bordes discont.', 7.5), 'ndc': ('… cortos', 6.5), 'ndl': ('… largos', 6.5), 'cauto': ('Caso automático', 15), 'cforce': ('Forzar caso', 12), 'caso': ('Caso usado', 15),
             'idx': ('col. m', 5), 'fr': ('interp.', 6), 'cNIC': ('M⁻ cont. corto', 7.5), 'cNIL': ('M⁻ cont. largo', 7.5), 'cNDC': ('M⁻ disc. corto', 7.5), 'cNDL': ('M⁻ disc. largo', 7.5),
             'cPC': ('M⁺ corto', 7.5), 'cPL': ('M⁺ largo', 7.5), 'pbor': ('Carga en orilla de volado', 8.5), 'MPc': ('M⁺ corto', 7.5), 'MNc': ('M⁻ corto', 7.5), 'MPl': ('M⁺ largo', 7.5), 'MNl': ('M⁻ largo', 7.5),
             'g678': ('6.7.8 cumple', 8), 'Vu': ('V_u por nerv.', 7.5), 'VcR': ('V_cR', 7.5), 'cort': ('Estado', 15),
             'afm': ('α_fm', 6.5), 'hmin': ('h mín. Tabla 6.7.3.1.2 / volado L/8', 9), 'heqv': ('h_eq / h', 7), 'peral': ('Estado', 12), 'tdia': ('Capa (diafragma)', 11),
             'Lbc': ('Bastón ‖ corto desde paño', 8), 'Lbl': ('Bastón ‖ largo desde paño', 8), 'ncx': ('Casetones ‖X', 7.5), 'ncy': ('Casetones ‖Y', 7.5), 'apoy': ('Apoyos', 15), 'estado': ('ESTADO', 26), 'nota': ('Nota', 60)}
    for blk in BLKS:
        heads['As' + blk] = ('As req.', 7); heads['v' + blk] = ('Var. No.', 6); heads['n' + blk] = ('Pzas', 5.5); heads['d' + blk] = ('d', 6); heads['MR' + blk] = ('M_R', 7); heads['FU' + blk] = ('Mu/M_R', 7); heads['e' + blk] = ('ε_t', 7)
    units = {'Lx': 'm', 'Ly': 'm', 'ax': 'm', 'ay': 'm', 'a1': 'm', 'a2': 'm', 'ln1': 'cm', 'ln2': 'cm', 'bfc': 'cm', 'bfl': 'cm', 'rb': 'cm', 'rt': 'cm',
             'scm': 'kg/m²', 'cv': 'kg/m²', 'pp': 'kg/m²', 'cm': 'kg/m²', 'ws': 'kg/m²', 'wu': 'kg/m²', 'pbor': 'kg/m (serv.)',
             'MPc': 't·m', 'MNc': 't·m', 'MPl': 't·m', 'MNl': 't·m', 'Vu': 't', 'VcR': 't', 'hmin': 'cm', 'heqv': 'cm', 'Lbc': 'cm', 'Lbl': 'cm'}
    for blk in BLKS:
        units['As' + blk] = 'cm²'; units['MR' + blk] = 't·m'; units['d' + blk] = 'cm'; units['e' + blk] = '≥ 0.005'
    for nm, (t, w) in heads.items():
        hdr_cell(ws, 7, CI(TC[nm]), t, w)
        unit_cell(ws, 8, CI(TC[nm]), units.get(nm, ''))
    ws.row_dimensions[7].height = 42
    NN = dict(N); NN['PCb'] = PC['bq']; NN['PCh'] = PC['hq']
    F = lambda s, r: s.format(r=r, **{k: v for k, v in TC.items()}, **NN)
    for t in TABS:
        k = key(t); r = TROW[k]
        L = t['loads']
        scm = L.get('Sobrecarga', 270.0 if t['name'] == 'F2' else 0)
        cvm = L.get('Carga Viva Máxima', 100.0)
        dest = 'HABITACIÓN' if abs(cvm - 190) < 1 else 'AZOTEA'
        e = {s: EDGE[(k, s)] for s in SIDES}
        D_ = {s: '(APOYOS!$G$%d="DISCONTINUO")' % e[s] for s in SIDES}
        dbe = dbz('{est_sec}')
        vals = {
            'key': (k, 'calc', '@'), 'nivel': (t['level'], 'calc', '@'), 'modo': (t['modo'], 'input', '@'), 'dest': (dest, 'input', '@'),
            'Lx': (t['Lx'], 'data', '0.000'), 'Ly': (t['Ly'], 'data', '0.000'), 'ax': (t['ax'], 'input', '0.000'), 'ay': (t['ay'], 'input', '0.000'),
            'a1': ('=MIN({ax}{r},{ay}{r})', 'calc', '0.000'), 'a2': ('=MAX({ax}{r},{ay}{r})', 'calc', '0.000'), 'm': ('=IF({a2}{r}>0,{a1}{r}/{a2}{r},0)', 'calc', '0.000'),
            'dirc': ('=IF({ax}{r}<={ay}{r},"X","Y")', 'calc', '@'),
            'ln1': ('=MAX(0,{a1}{r}*100-IF({modo}{r}="VOLADO",{bp}/2,{bp}))', 'calc', '0'), 'ln2': ('=MAX(0,{a2}{r}*100-{bp})', 'calc', '0'),
            'bfc': ('=MIN({s},{bw}+2*MIN(8*{t},{ac}/2,{ln1}{r}/8))', 'calc', '0.0'), 'bfl': ('=MIN({s},{bw}+2*MIN(8*{t},{ac}/2,{ln2}{r}/8))', 'calc', '0.0'),
            'rb': ('=IF({modo}{r}="VOLADO",{rec_ext},{rec})', 'calc', '0.0'), 'rt': ('=IF(AND({nivel}{r}="N+6.00",{exp_az}="SÍ"),{rec_ext},{rec})', 'calc', '0.0'),
            'scm': (scm, 'input', '0'), 'cv': ('=IF({dest}{r}="HABITACIÓN",{cv_hab},{cv_az})', 'calc', '0'),
            'pp': ('={PPd}', 'calc', '0'), 'cm': ('={pp}{r}+{scm}{r}', 'calc', '0'), 'ws': ('={cm}{r}+{cv}{r}', 'calc', '0'),
            'wu': ('=MAX({FCcm}*{cm}{r}+{FCcv}*{cv}{r},IF({dest}{r}="AZOTEA",{FCacc}*({cm}{r}+{Wg}),0))', 'key', '0'),
            'nd': ('=' + '+'.join(D_[s] for s in SIDES), 'calc', '0'),
            'ndc': ('=IF({dirc}{r}="X",%s+%s,%s+%s)' % (D_['inferior'], D_['superior'], D_['izquierdo'], D_['derecho']), 'calc', '0'),
            'ndl': ('={nd}{r}-{ndc}{r}', 'calc', '0'),
            'cauto': ('=IF({nd}{r}=0,"INTERIOR",IF({nd}{r}=1,IF({ndc}{r}=1,"BORDE-CORTO","BORDE-LARGO"),IF({nd}{r}=2,IF({ndc}{r}=1,"ESQUINA",IF({ndc}{r}=2,"EXTREMO-LARGO","EXTREMO-CORTO")),IF({nd}{r}=3,IF({ndc}{r}=2,"EXTREMO-LARGO","EXTREMO-CORTO"),"AISLADO"))))', 'calc', '@'),
            'cforce': ('AUTO', 'input', '@'), 'caso': ('=IF({cforce}{r}="AUTO",{cauto}{r},{cforce}{r})', 'key', '@'),
            'idx': ('=MATCH({m}{r},{MROW},1)', 'calc', '0'),
            'fr': ('=IF({idx}{r}>=7,0,({m}{r}-INDEX({MROW},{idx}{r}))/(INDEX({MROW},{idx}{r}+1)-INDEX({MROW},{idx}{r})))', 'calc', '0.00'),
            'pbor': (t.get('volado', {}).get('carga_borde', 0.0), 'input', '0.0'),
        }
        for mm in MOMS:
            idx = 'MATCH({caso}{r}&"|%s",{CKEY},0)' % mm
            vals['c' + mm] = ('=INDEX({CVAL},%s,{idx}{r})+{fr}{r}*(INDEX({CVAL},%s,MIN({idx}{r}+1,7))-INDEX({CVAL},%s,{idx}{r}))' % (idx, idx, idx), 'calc', '0')
        fac = '*{wu}{r}*{a1}{r}^2*1E-4*{s}/100/1000'
        vals['MPc'] = ('=IF({modo}{r}="PERIMETRAL",{cPC}{r}%s,0)' % fac, 'calc', '0.000')
        vals['MNc'] = ('=IF({modo}{r}="PERIMETRAL",MAX({cNIC}{r},{cNDC}{r})%s,IF({modo}{r}="VOLADO",({wu}{r}*({ln1}{r}/100)^2/2+{FCcm}*{pbor}{r}*{ln1}{r}/100)*{s}/100/1000,0))' % fac, 'calc', '0.000')
        vals['MPl'] = ('=IF({modo}{r}="PERIMETRAL",{cPL}{r}%s,0)' % fac, 'calc', '0.000')
        vals['MNl'] = ('=IF({modo}{r}="PERIMETRAL",MAX({cNIL}{r},{cNDL}{r})%s,0)' % fac, 'calc', '0.000')
        # peraltes efectivos: el lecho del claro corto va por fuera; el del largo, encima (o debajo) de él
        dq = {'Pc': '({h}-{rb}{r}-%s-%s/2)' % (dbe, db('{var_inf}')),
              'Pl': '({h}-{rb}{r}-%s-%s-%s/2)' % (dbe, db('{var_inf}'), db('{var_inf}')),
              'Nc': '({h}-{rt}{r}-2*{fim}-%s-%s/2)' % (dbe, db('{var_sup}')),
              'Nl': '({h}-{rt}{r}-2*{fim}-%s-%s-%s/2)' % (dbe, db('{var_sup}'), db('{var_sup}'))}
        dact = {'Pc': '{h}-{rb}{r}-%s-%s/2' % (dbe, db('{vPc}{r}')),
                'Pl': '{h}-{rb}{r}-%s-%s-%s/2' % (dbe, db('{vPc}{r}'), db('{vPl}{r}')),
                'Nc': '{h}-{rt}{r}-2*{fim}-%s-%s/2' % (dbe, db('{vNc}{r}')),
                'Nl': '{h}-{rt}{r}-2*{fim}-%s-%s-%s/2' % (dbe, db('{vNc}{r}'), db('{vNl}{r}'))}
        for blk, M, neg in (('Pc', 'MPc', False), ('Nc', 'MNc', True), ('Pl', 'MPl', False), ('Nl', 'MNl', True)):
            b = '{bwc}' if neg else '{bf%s}{r}' % blk[1].lower()
            pref = '{var_sup}' if neg else '{var_inf}'
            av = '({bw}-2*{rec}-2*%s)' % dbe if neg else '({bw}-2*{rb}{r}-2*%s)' % dbe
            bmin = 'IF({modo}{r}="VOLADO",2*{bw},{bw})' if blk == 'Nc' else '{bw}'
            Mref = '{%s}{r}' % M
            vals['As' + blk] = ('=IF({modo}{r}="HUECO",0,%s)' % asreq(Mref, b, dq[blk], neg, bmin), 'calc', '0.00')
            vals['v' + blk] = ('=' + varsel2('{As%s}{r}' % blk, pref, '{nmax_sec}', av), 'calc', 'General')
            ab_ = ab('{v%s}{r}' % blk)
            n0 = 'MAX({nmin_sec},CEILING(MAX({As%s}{r},0)/%s,1))' % (blk, ab_)
            vals['n' + blk] = ('=IF({modo}{r}="HUECO",0,IF({As%s}{r}<0,0,IF(%s*%s>=4/3*%s-0.0001,%s,MAX(%s,2))))' % (blk, n0, ab_, asflex(Mref, b, dq[blk], neg), n0, n0), 'calc', '0')
            vals['d' + blk] = ('=' + dact[blk], 'calc', '0.00')
            asp = '{n%s}{r}*%s' % (blk, ab_)
            vals['MR' + blk] = ('=IF({n%s}{r}=0,0,{FRf}*%s*{fy}*({d%s}{r}-%s*{fy}/({fpc}*%s)/2)/1E5)' % (blk, asp, blk, asp, b), 'calc', '0.000')
            vals['FU' + blk] = ('=IF({As%s}{r}<0,9.99,IF({MR%s}{r}>0,{%s}{r}/{MR%s}{r},0))' % (blk, blk, M, blk), 'key', '0.00')
            cexp = '(%s*{fy}/({fpc}*%s)/{beta1})' % (asp, b)
            vals['e' + blk] = ('=IF({n%s}{r}=0,"—",0.003*({d%s}{r}-%s)/%s)' % (blk, blk, cexp, cexp), 'calc', '0.0000')
        vals['g678'] = ('=IF({modo}{r}="HUECO","—",IF(AND({modo}{r}="PERIMETRAL",{a1}{r}*100>=6*{ac},{ac}<=80,{bw}>=10,{h}<=3.5*{bw},{t}>=MAX({ac}/13,5)),"SÍ","NO"))', 'calc', '@')
        lam = 'MIN(1,SQRT(2/(1+0.04*{dNc}{r})))'
        pexp = '({nNc}{r}*%s/({bwc}*{dNc}{r}))' % ab('{vNc}{r}')
        vals['Vu'] = ('=IF({modo}{r}="PERIMETRAL",MAX(0,({ln1}{r}/2-{dNc}{r})/100)*(0.95-0.5*{m}{r})*{wu}{r}*IF(AND({nd}{r}>0,{nd}{r}<4),{Vinc},1)*{s}/100/1000,'
                      'IF({modo}{r}="VOLADO",({wu}{r}*MAX(0,{ln1}{r}-{dNc}{r})/100+{FCcm}*{pbor}{r})*{s}/100/1000,0))', 'calc', '0.000')
        vals['VcR'] = ('=IF({modo}{r}="HUECO",0,{FRv}*2*%s*%s^(1/3)*SQRT({fc})*{bwc}*{dNc}{r}*IF({g678}{r}="SÍ",{VcRinc},1)/1000)' % (lam, pexp), 'calc', '0.000')
        vals['cort'] = ('=IF({modo}{r}="HUECO","—",IF({Vu}{r}<={VcR}{r},"CUMPLE","MACIZAR "&MAX(1,CEILING(IF({modo}{r}="VOLADO",{a1}{r},{a1}{r}/2)*(1-{VcR}{r}/{Vu}{r})*100/{s},1))&" hilera(s)"))', 'calc', '@')
        afs = [edge_af(t, s_, k) for s_ in SIDES]
        vals['afm'] = ('=IF({modo}{r}="PERIMETRAL",(%s)/4,0)' % '+'.join(afs), 'calc', '0.00')
        disc_nobeam = 'OR(' + ','.join('AND(%s,%s<0.8)' % (D_[s_], a_) for s_, a_ in zip(SIDES, afs)) + ')'
        beta = '({ln2}{r}/MAX({ln1}{r},1))'
        k_ = '(0.8+{fy}/14000)'
        hd = 'MAX(9,{ln2}{r}*%s/(36+9*%s))' % (k_, beta)
        hb = 'MAX(13,{ln2}{r}*%s/(36+5*%s*({afm}{r}-0.2)))' % (k_, beta)
        ha = 'MAX(13,{ln2}{r}/IF({nd}{r}>0,30,33))'
        vals['hmin'] = ('=IF({modo}{r}="PERIMETRAL",IF({afm}{r}>2,%s,IF({afm}{r}>0.2,%s,%s))*IF(AND({afm}{r}>0.2,%s),1.1,1),IF({modo}{r}="VOLADO",MAX(25,{a1}{r}*100/8),0))' % (hd, hb, ha, disc_nobeam), 'calc', '0.0')
        vals['heqv'] = ('=IF({modo}{r}="VOLADO",{h},{heq})', 'calc', '0.0')
        vals['peral'] = ('=IF({modo}{r}="HUECO","—",IF({heqv}{r}>={hmin}{r},"CUMPLE",IF({modo}{r}="VOLADO","NO CUMPLE","CALCULAR FLECHA")))', 'calc', '@')
        vals['tdia'] = ('=IF({modo}{r}="HUECO","—",IF({t}>=IF(IF({modo}{r}="VOLADO",{a1}{r},{a2}{r})>=6,6,5),"CUMPLE","CAPA ≥ 6 cm"))', 'calc', '@')
        # bastones: desde el paño, punto de inflexión + max(d, 12db, Ln/16), y no menos que L_d
        xpi_c = '{ln1}{r}*(1-SQRT({MPc}{r}/MAX({MPc}{r}+{MNc}{r},1E-9)))/2'
        xpi_l = '{ln2}{r}*(1-SQRT({MPl}{r}/MAX({MPl}{r}+{MNl}{r},1E-9)))/2'
        ext_c = 'MAX({dNc}{r},12*%s,{ln1}{r}/16)' % db('{vNc}{r}')
        ext_l = 'MAX({dNl}{r},12*%s,{ln2}{r}/16)' % db('{vNl}{r}')
        if t['modo'] == 'VOLADO':
            lb = t['volado']['Lb']
            # claro de respaldo con 0.9·PP geométrico y sin carga viva (efecto favorable): donde se anula el momento negativo
            wD = '(0.9*{PP}*{s}/100/1000)'
            A_ = '(0.45*%s)' % wD; B_ = '(0.45*%s*%.3f+{MNc}{r}/%.3f)' % (wD, lb, lb); C_ = '{MNc}{r}'
            x0 = 'IF(%s^2-4*%s*%s<0,%.3f,MIN(%.3f,(%s-SQRT(%s^2-4*%s*%s))/(2*%s)))*100' % (B_, A_, C_, lb, lb, B_, B_, A_, C_, A_)
            vals['Lbc'] = ('=CEILING(MAX(%s+MAX({dNc}{r},12*%s,%.1f/16),%s),5)' % (x0, db('{vNc}{r}'), lb * 100, ld('{vNc}{r}')), 'key', '0')
            vals['Lbl'] = (0, 'calc', '0')
        else:
            vals['Lbc'] = ('=IF({modo}{r}="HUECO",0,CEILING(MAX(%s+%s,%s),5))' % (xpi_c, ext_c, ld('{vNc}{r}')), 'key', '0')
            vals['Lbl'] = ('=IF({modo}{r}="HUECO",0,CEILING(MAX(%s+%s,%s),5))' % (xpi_l, ext_l, ld('{vNl}{r}')), 'key', '0')
        vals['ncx'] = ('=MAX(0,INT(({Lx}{r}*100-{bp}+{bw})/{s}))', 'calc', '0')
        vals['ncy'] = ('=MAX(0,INT(({Ly}{r}*100-{bp}+{bw})/{s}))', 'calc', '0')
        libre = '+'.join('(APOYOS!$F$%d="LIBRE")' % e[s] for s in SIDES)
        vals['apoy'] = ('=IF(COUNTIFS(TRIB!$C$%d:$C$%d,{key}{r},TRIB!$K$%d:$K$%d,"NO CUMPLE*")+COUNTIFS(TRIB!$C$%d:$C$%d,{key}{r},TRIB!$K$%d:$K$%d,"SIN APOYO*")>0,"VER PRINCIPALES",IF(%s>0,"REVISAR: borde libre","OK"))'
                        % (R0, R1, R0, R1, R0, R1, R0, R1, libre), 'calc', '@')
        fus = 'MAX({FUPc}{r},{FUNc}{r},{FUPl}{r},{FUNl}{r})'
        nn = 'OR(' + ','.join('{n%s}{r}>MIN({nmax_sec},%s)' % (k_b, nfit('{v%s}{r}' % k_b, ('({bw}-2*{rec}-2*%s)' % dbe) if k_b[0] == 'N' else ('({bw}-2*{rb}{r}-2*%s)' % dbe))) for k_b in BLKS) + ')'
        etl = 'OR(' + ','.join('AND(ISNUMBER({e%s}{r}),{e%s}{r}<0.005)' % (k_b, k_b) for k_b in BLKS) + ')'
        vals['estado'] = ('=IF({modo}{r}="HUECO","HUECO — marco de nervaduras principales",IF(AND(%s<=1,NOT(%s),NOT(%s),{peral}{r}="CUMPLE",{tdia}{r}="CUMPLE"),'
                          'IF({apoy}{r}="OK",IF({cort}{r}="CUMPLE","CUMPLE","CUMPLE — "&{cort}{r}),IF({apoy}{r}="VER PRINCIPALES","VER PRINCIPALES","REVISAR APOYOS")),'
                          'IF(AND(%s<=1,NOT(%s),NOT(%s),{peral}{r}="CALCULAR FLECHA"),"REVISAR: calcular flecha (h_eq < h mín.)",'
                          '"NO CUMPLE: "&IF(%s>1,"flexión ","")&IF(%s,"no caben ","")&IF(%s,"ε_t < 0.005 ","")&IF({peral}{r}="NO CUMPLE","peralte ","")&IF({tdia}{r}<>"CUMPLE","capa < 6 cm",""))))'
                          % (fus, nn, etl, fus, nn, etl, fus, nn, etl), 'key', '@')
        notes = []
        if t['rep_nota']:
            notes.append(t['rep_nota'])
        if t['seccion'] == 'Slab1':
            notes.append('En el E2K está como «Slab1» (maciza 20 cm, 4000 psi): corregir a LR-25.')
        if t['name'] == 'F2':
            notes.append('Sin cargas en el E2K: se toma la SCM de F80023 (270 kg/m²). REVISAR.')
        if t['modo'] == 'VOLADO':
            notes.append('Volado de %.2f m desde el borde %s (muro); orilla libre con %.1f kg/m del E2K. Cara inferior a la intemperie (r = r_ext). Si la orilla lleva pretil de block u otro elemento frágil, el L/8 no basta: calcular flecha (13.4). '
                         'Bastón superior: se prolonga hacia el claro de respaldo (%.2f m) hasta pasar el punto donde se anula el momento con 0.9·PP y sin carga viva (NTC-Criterios 3.4.1 c y 5.1.2.2).' % (t['volado']['Lv'], t['volado']['raiz'], t['volado'].get('carga_borde', 0), t['volado']['Lb']))
        if t['modo'] == 'HUECO':
            notes.append('Tragaluz de 0.30 m: casi todo el tablero es hueco (quedan %.2f m² de losa).' % t['area'])
        if min(t['Lx'], t['Ly']) < 0.75:
            notes.append('No cabe un casetón en el claro corto: el tablero se cuela MACIZO.')
        txt = ' '.join(notes) or '—'
        vals['nota'] = ('=IF({g678}{r}="NO","No cumple 6.7.8.1.4 (claro corto < 6·a_c): losa apoyada en vigas, 6.7.8.1.9 — V_cR sin ×1.1. ","")&"%s"' % txt.replace('"', '""'), 'calc', '@')
        for nm in _tcols:
            v, kind, fmt = vals[nm]
            if isinstance(v, str) and v.startswith('='):
                v = F(v, r)
            x = cell(ws, r, CI(TC[nm]), v, kind, fmt, al='left' if nm in ('nota', 'estado', 'cort', 'apoy', 'key') else 'center', sz=9)
            if nm == 'nota':
                x.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[r].height = 30
        dv_list(ws, '%s%d' % (TC['modo'], r), '=' + N['LIST_1'])
        dv_list(ws, '%s%d' % (TC['dest'], r), '=' + N['LIST_2'])
        dv_list(ws, '%s%d' % (TC['cforce'], r), '=' + N['LIST_0'])
    ws.freeze_panes = 'C9'
    add_cf(ws, 'B9:%s%d' % (TC['estado'], T1))
    for c_ in ('g678', 'peral', 'tdia', 'cort'):
        add_cf(ws, '%s9:%s%d' % (TC[c_], TC[c_], T1))
    ws.sheet_view.zoomScale = 80
    return ws


# ---------------------------------------------------------------- NERVADURA (detalle)
def sheet_nervadura(wb):
    ws = wb.create_sheet('NERVADURA'); S = Narr(ws, N)
    S.title('NERVADURA — DISEÑO PASO A PASO DE UN TABLERO')
    S.subtitle('Escoge el tablero en la celda amarilla. La hoja vuelve a hacer, con la fórmula a la vista, lo mismo que hace TABLEROS en un renglón; al final compara los dos resultados. La sección 9 revisa la capa de compresión (es igual para todos los tableros).', 30)
    S.gap()
    S.section('0. TABLERO')
    S.header()
    S.row('Tablero a revisar (y a dibujar en CROQUIS)', 'F80016 N+3.00', '—', 'SELECTOR. Lista de TABLEROS columna B.', '—', 'input', '@', 'sel')
    dv_list(ws, N['sel'].split('!')[1].replace('$', ''), '=' + N['TKEYS'])
    S.row('Renglón en TABLEROS', '=MATCH({sel},{TKEYS},0)', '—', 'Auxiliar.', '—', 'calc', '0', 'nr')

    def tv(c):
        return 'INDEX(%s,{nr})' % tcol(c)
    S.gap()
    S.section('1. GEOMETRÍA Y CARGAS')
    S.header()
    S.row('Modo', '=' + tv('modo'), '—', 'PERIMETRAL = apoyado en sus cuatro bordes; VOLADO; HUECO.', 'APOYOS / TABLEROS', 'key', '@', 'n_modo')
    S.row('Nivel', '=' + tv('nivel'), '—', '—', E2K, 'calc', '@', 'n_niv')
    S.row('a_x — claro en X usado (a ejes)', '=' + tv('ax'), 'm', 'De TABLEROS (para irregulares es el rectángulo representativo). Usar claros a ejes queda del lado seguro: la tabla permite claros libres en el caso I.', DXFT + ' · TRIB_TABLERO · ' + NTCC + ' Tabla 3.4.2.1 nota [2], p. 65', 'calc', '0.000', 'n_ax')
    S.row('a_y — claro en Y usado (a ejes)', '=' + tv('ay'), 'm', '—', DXFT, 'calc', '0.000', 'n_ay')
    S.row('a1 — claro corto', '=MIN({n_ax},{n_ay})', 'm', 'CALCULADA.', '—', 'key', '0.000', 'n_a1')
    S.row('a2 — claro largo', '=MAX({n_ax},{n_ay})', 'm', 'CALCULADA.', '—', 'calc', '0.000', 'n_a2')
    S.row('m = a1/a2', '={n_a1}/{n_a2}', '—', 'CALCULADA.', '—', 'calc', '0.000', 'n_m')
    S.row('Dirección del claro corto (ahí trabajan las nervaduras que más cargan)', '=IF({n_ax}<={n_ay},"‖ X","‖ Y")', '—', 'Su acero va por fuera; el del claro largo se apoya sobre él (d menor).', '—', 'key', '@', 'n_dir')
    S.row('Claros libres L_n corto / largo', '=TEXT(' + tv('ln1') + ',"0")&" / "&TEXT(' + tv('ln2') + ',"0")', 'cm', 'Claro a ejes menos el ancho de la nervadura principal o dala.', NTCC + ' 6.3.2.1, p. 117', 'calc', '@')
    S.row('Patín efectivo b_f claro corto / largo', '=TEXT(' + tv('bfc') + ',"0.0")&" / "&TEXT(' + tv('bfl') + ',"0.0")', 'cm', 'b_w + 2·el menor de 8t, a_c/2 y L_n/8 (no más que s).', NTCC + ' Tabla 5.2.1.4.2, p. 89', 'calc', '@')
    S.row('SCM — sobrecarga muerta', '=' + tv('scm'), 'kg/m²', 'Del E2K, tablero por tablero.', E2K + ' · AREALOAD "Sobrecarga"', 'calc', '0', 'n_scm')
    S.row('W_m — carga viva', '=' + tv('cv'), 'kg/m²', 'Según el destino del tablero.', NTCCA + ' Tabla 6.1.2.2, p. 11', 'calc', '0', 'n_cv')
    S.row('PP de diseño (geométrico + 20 + 20)', '={PPd}', 'kg/m²', 'CARGAS sección 1.', NTCCA + ' 5.1.2.1, p. 9', 'calc', '0.0', 'n_pp')
    S.row('CM = PP + SCM', '={n_pp}+{n_scm}', 'kg/m²', 'CALCULADA. W_m/CM < 2.5: se pueden usar los coeficientes.', NTCC + ' 3.4.2.1 d), p. 64', 'calc', '0', 'n_cm')
    S.row('w — carga de servicio', '={n_cm}+{n_cv}', 'kg/m²', 'CM + W_m.', NTCCA + ' 3.4.1 d), p. 7', 'calc', '0', 'n_ws')
    S.row('w_u — carga última', '=MAX({FCcm}*{n_cm}+{FCcv}*{n_cv},IF(' + tv('dest') + '="AZOTEA",{FCacc}*({n_cm}+{Wg}),0))', 'kg/m²', 'El mayor de 1.3·CM + 1.5·W_m y, en azotea, 1.1·(CM + W_g).', NTCCA + ' 3.4.1 a) y b), p. 7 · 7.3.2.1, p. 17', 'key', '0', 'n_wu')
    S.gap()
    S.section('2. CASO Y COEFICIENTES')
    S.header()
    S.row('Bordes discontinuos (de APOYOS)', '=' + tv('nd'), '—', 'Cuenta de bordes con «DISCONTINUO» en APOYOS sección C.', 'APOYOS', 'calc', '0')
    S.row('Caso de la tabla', '=' + tv('caso'), '—', 'Automático según cuántos bordes discontinuos son cortos o largos; se puede forzar en TABLEROS. Caso I: losa colada monolíticamente con sus apoyos.', NTCC + ' Tabla 3.4.2.1, p. 65', 'key', '@', 'n_caso')
    S.row('Coef. M⁻ en bordes continuos, claro corto', '=' + tv('cNIC'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cNIC')
    S.row('Coef. M⁻ en bordes discontinuos, claro corto', '=' + tv('cNDC'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cNDC')
    S.row('Coef. M⁺ claro corto', '=' + tv('cPC'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cPC')
    S.row('Coef. M⁻ en bordes continuos, claro largo', '=' + tv('cNIL'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cNIL')
    S.row('Coef. M⁻ en bordes discontinuos, claro largo', '=' + tv('cNDL'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cNDL')
    S.row('Coef. M⁺ claro largo', '=' + tv('cPL'), '×10⁻⁴', 'Interpolado en m.', 'TABLAS sección A', 'calc', '0', 'n_cPL')
    S.row('Carga en la orilla del volado (servicio)', '=' + tv('pbor'), 'kg/m', 'Sólo para volados: barandal del E2K.', E2K + ' · FRAME OBJECT LOADS', 'calc', '0.0', 'n_pb')
    S.gap()
    S.section('3. MOMENTOS POR NERVADURA')
    S.secnote('Por metro: M = coef × 10⁻⁴ × w_u × a1². Por nervadura: × s. En volado: M = (w_u·L_n²/2 + 1.3·P·L_n)·s, con L_n medido al paño del muro.')
    S.header()
    fac = '*{n_wu}*{n_a1}^2*1E-4*{s}/100/1000'
    S.row('M⁺ claro corto', '=IF({n_modo}="PERIMETRAL",{n_cPC}%s,0)' % fac, 't·m', 'Coef. M⁺ corto × 10⁻⁴ × w_u × a1² × s.', NTCC + ' 3.4.2.1, p. 64 y Tabla 3.4.2.1, p. 65', 'key', '0.000', 'n_MPc')
    S.row('M⁻ claro corto', '=IF({n_modo}="PERIMETRAL",MAX({n_cNIC},{n_cNDC})%s,IF({n_modo}="VOLADO",({n_wu}*(%s/100)^2/2+{FCcm}*{n_pb}*%s/100)*{s}/100/1000,0))' % (fac, tv('ln1'), tv('ln1')), 't·m', 'El mayor entre borde continuo y discontinuo (envolvente: no se redistribuye el desequilibrio de 3.4.3.1); en volado, el momento en el paño.', NTCC + ' 3.4.3.1, p. 64', 'key', '0.000', 'n_MNc')
    S.row('M⁺ claro largo', '=IF({n_modo}="PERIMETRAL",{n_cPL}%s,0)' % fac, 't·m', 'Mismo procedimiento.', NTCC + ' Tabla 3.4.2.1, p. 65', 'key', '0.000', 'n_MPl')
    S.row('M⁻ claro largo', '=IF({n_modo}="PERIMETRAL",MAX({n_cNIL},{n_cNDL})%s,0)' % fac, 't·m', 'Mismo procedimiento.', NTCC + ' Tabla 3.4.2.1, p. 65', 'key', '0.000', 'n_MNl')
    sec = 4
    for blk, M, neg, lbl in (('Pc', 'n_MPc', False, 'LECHO INFERIOR — CLARO CORTO'), ('Nc', 'n_MNc', True, 'LECHO SUPERIOR — CLARO CORTO'),
                             ('Pl', 'n_MPl', False, 'LECHO INFERIOR — CLARO LARGO'), ('Nl', 'n_MNl', True, 'LECHO SUPERIOR — CLARO LARGO')):
        S.gap()
        S.section('%d. %s' % (sec, lbl)); sec += 1
        S.header()
        b = '{bwc}' if neg else tv('bf' + blk[1].lower())
        S.row('b — ancho que trabaja a compresión', '=' + b, 'cm', 'Arriba en tensión: comprime la nervadura (b_w de cálculo).' if neg else 'Arriba comprime el patín efectivo b_f (sección T).', NTCC + (' 3.9.1 e), p. 74' if neg else ' Tabla 5.2.1.4.2, p. 89 · 5.2.2.1.3.1, p. 90'), 'calc', '0.0', 'n_b' + blk)
        S.row('d — peralte efectivo con la varilla que se usa', '=' + tv('d' + blk), 'cm', 'h − r − d_b/2' + (' − 2·Ø de la malla (la varilla va bajo la malla)' if neg else '') + (' − d_b del lecho del claro corto' if blk[1] == 'l' else '') + '. r = ' + ('r superior' if neg else 'r inferior (r_ext en volados)') + '.', NTCC + ' Tabla 4.14.16.a, p. 85', 'calc', '0.00', 'n_d' + blk)
        S.row("As mín = máx(0.80·√f'c, 14)/fy · b_w · d", '={pmin}*IF(AND(%s,{n_modo}="VOLADO"),2*{bw},{bw})*{n_d%s}' % ('TRUE' if blk == 'Nc' else 'FALSE', blk), 'cm²', 'Sobre b_w (en el volado, patín a tensión: 2b_w). Se dispensa si la As colocada es ≥ 4/3 de la requerida.', NTCC + ' 6.3.5.1.2 y 6.3.5.1.3, p. 122', 'calc', '0.00', 'n_min' + blk)
        S.row('As requerida (TABLEROS)', '=' + tv('As' + blk), 'cm²', "q = 1 − √(1 − 2Mu/(F_R·b·d²·f''c)); As = q·f''c·b·d/fy; luego el mayor de esa y el menor de (As mín, 4/3·As). −1 = no alcanza o rebasa p_max.", NTCC + ' ecs. 5.2.2.1.1.1.a y b, p. 89 · 6.3.5.1.3, p. 122', 'key', '0.00', 'n_As' + blk)
        S.row('Varilla que se usa', '=' + tv('v' + blk), 'No.', 'La preferida (No. 4) o la siguiente que quepa.', NTCC + ' 6.7.6.1.3, p. 175', 'calc', 'General', 'n_v' + blk)
        S.row('Número de varillas', '=' + tv('n' + blk), 'pzas', 'Al menos 1 (la No. 4 corrida). Si As colocada < 4/3 de la de flexión, al menos 2.' + (' Arriba: 1 corrida + el resto bastones.' if neg else ' Todas corridas.'), NTCC + ' 6.7.6.1.3, p. 175 · 6.3.5.1.1, p. 122', 'key', '0', 'n_n' + blk)
        asp = '{n_n%s}*%s' % (blk, ab('{n_v%s}' % blk))
        S.row('a — profundidad del bloque', '=%s*{fy}/({fpc}*{n_b%s})' % (asp, blk), 'cm', "a = As·fy/(f''c·b)." + ('' if neg else ' Debe quedar dentro del patín (a ≤ t).'), NTCC + ' 3.6.1 e), p. 68', 'calc', '0.00', 'n_a' + blk)
        S.row('M_R — momento resistente', '={FRf}*%s*{fy}*({n_d%s}-{n_a%s}/2)/1E5' % (asp, blk, blk), 't·m', "F_R·As·fy·(d − a/2).", NTCC + ' ec. 5.2.2.1.1.1.b, p. 89 · ' + NAIM24 + ' pág. 13', 'calc', '0.000', 'n_MR' + blk)
        S.row('Mu / M_R', '=IF({n_As%s}<0,9.99,IF({n_MR%s}>0,{%s}/{n_MR%s},0))' % (blk, blk, M, blk), '—', 'Debe ser ≤ 1. 9.99 = la sección no alcanza.', '—', 'key', '0.00', 'n_FU' + blk)
        S.row('ε_t — deformación del acero de tensión', '=0.003*({n_d%s}-{n_a%s}/{beta1})/({n_a%s}/{beta1})' % (blk, blk, blk), '—', 'ε_t = 0.003·(d − c)/c con c = a/β1. Debe ser ≥ 0.005 para usar F_R = 0.9 (sección controlada por tensión).', NTCC + ' Tabla 3.8.2.2, p. 70 · 6.7.3.3.1, p. 169', 'calc', '0.0000', 'n_e' + blk)
        S.row('¿Caben en la nervadura?', '=IF({n_n%s}=0,"—",IF({n_n%s}*%s+({n_n%s}-1)*MAX({slib},%s)<={bw}-2*%s-2*%s+0.01,"CUMPLE","NO CABE"))' % (blk, blk, db('{n_v%s}' % blk), blk, db('{n_v%s}' % blk), ('{rec}' if neg else tv('rb')), dbz('{est_sec}')),
              '—', 'n·d_b + (n−1)·s_libre ≤ b_w − 2r − 2·estribo.', NTCC + ' 14.2.1, p. 347', 'calc', '@')
        S.row('Estado', '=IF(AND({n_FU%s}<=1,{n_e%s}>=0.005),"CUMPLE","NO CUMPLE")' % (blk, blk), '—', '—', '—', 'key', '@')
    S.gap()
    S.section('%d. CORTANTE EN LA NERVADURA' % sec); sec += 1
    S.secnote('V por metro a d del paño = (L_n/2 − d)·(0.95 − 0.5·m)·w_u, × 1.15 si hay bordes continuos y discontinuos (fórmula de la NTC 2017 que la 2023 ya no trae; se usa como método reconocido, 3.4.1.1). Por nervadura × s. '
              'V_cR sin estribos con la ec. 5.5.3.2.1 y la cuantía del acero superior.', 42)
    S.header()
    S.row('V_u por nervadura', '=' + tv('Vu'), 't', 'En volado: V = (w_u·(L_n − d) + 1.3·P)·s.', NTCC + ' 3.4.1.1, p. 64 · 6.7.4.3.2, p. 172 (sección crítica a d del paño)', 'key', '0.000', 'n_Vu')
    S.row('p — cuantía del acero superior en el apoyo', '={n_nNc}*%s/({bwc}*{n_dNc})' % ab('{n_vNc}'), '—', 'As_sup / (b_w·d).', NTCC + ' 5.5.3.2.1, p. 95', 'calc', '0.0000', 'n_p')
    S.row('λs — factor de tamaño', '=MIN(1,SQRT(2/(1+0.04*{n_dNc})))', '—', '√(2/(1 + 0.04d)) ≤ 1, d en cm.', NTCC + ' ec. 5.5.3.2.1.b, p. 95', 'calc', '0.000', 'n_ls')
    S.row('¿Cumple 6.7.8 (losa encasetonada)?', '=' + tv('g678'), '—', 'Claro corto ≥ 6·a_c, a_c ≤ 80, b_w ≥ 10, h ≤ 3.5b_w, t ≥ máx(a_c/13, 5). Si NO, se diseña como losa apoyada en vigas y no se usa el ×1.1.', NTCC + ' 6.7.8.1.2 a 6.7.8.1.4, p. 187 · 6.7.8.1.9, p. 188', 'key', '@', 'n_678')
    S.row('V_cR — cortante que toma el concreto', "={FRv}*2*{n_ls}*{n_p}^(1/3)*SQRT({fc})*{bwc}*{n_dNc}*IF({n_678}=\"SÍ\",{VcRinc},1)/1000", 't',
          "F_R·2·λs·p^(1/3)·√f'c·b_w·d (× 1.1 si cumple 6.7.8). Sin estribos: elemento en que no es práctico colocarlos.", NTCC + ' 5.5.3.8.1, p. 97 · ec. 5.5.3.2.1.a, p. 95 · 6.7.8.1.5, p. 187 · Tabla 3.8.2.1, p. 70', 'key', '0.000', 'n_VcR')
    S.row('Estado', '=' + tv('cort'), '—', 'Si no pasa, se quita(n) la(s) primera(s) hilera(s) de casetones junto al apoyo (macizo). Donde V_u > V_cR sin macizar haría falta al menos A_v,mín.', NTCC + ' Tabla 6.3.5.4.1 c) y 6.3.5.4.1, p. 123', 'key', '@')
    S.gap()
    S.section('%d. ESPESOR MÍNIMO (SE PUEDE OMITIR EL CÁLCULO DE FLECHAS)' % sec); sec += 1
    S.secnote('Tabla 6.7.3.1.2 con α_fm = promedio de α_f de los cuatro bordes: un muro bajo todo el borde cuenta como viga infinitamente rígida (α_f = ∞); una nervadura principal, como I_b/I_s. '
              'α_fm > 2: h = L_n·(0.8 + fy/14 000)/(36 + 9β) ≥ 9 cm; 0.2 < α_fm ≤ 2: h = L_n·(0.8 + fy/14 000)/(36 + 5β(α_fm − 0.2)) ≥ 13 cm; α_fm ≤ 0.2: Tabla 6.7.3.1.1 (L_n/30 o L_n/33, ≥ 13 cm). '
              '+10% con borde discontinuo sin viga de α_f ≥ 0.8. L_n = claro libre largo, β = L_n largo / L_n corto. La hoja compara contra h_eq (losa maciza de la misma inercia), no contra los 25 cm.', 66)
    S.header()
    S.row('α_fm', '=' + tv('afm'), '—', 'Promedio de los cuatro bordes (99 = muro).', NTCC + ' Tabla 6.7.3.1.2 nota [1], p. 168 · comentario de 6.7.1, p. 165', 'calc', '0.00')
    S.row('h mínimo (o L/8 si es volado)', '=' + tv('hmin'), 'cm', 'Volado: las nervaduras son vigas, h ≥ L/8 y ≥ 25 cm (si no carga elementos frágiles).', NTCC + ' Tabla 6.7.3.1.2 y 6.7.3.1.2.1, p. 168 · Tabla 6.7.3.1.1, p. 167 · Tabla 6.3.3.1.1, p. 120', 'key', '0.00')
    S.row('h que se compara (h_eq, o h si es volado)', '=' + tv('heqv'), 'cm', 'DATOS sección 4.', 'Criterio de esta hoja', 'calc', '0.00')
    S.row('Estado', '=' + tv('peral'), '—', 'Si no cumple: calcular flechas (13.4) y revisar vibración (13.7).', NTCC + ' 6.7.3.2.1, p. 168 · 13.7.1.1, p. 343', 'key', '@')
    S.gap()
    S.section('%d. ACOMODO DE CASETONES, BASTONES Y DETALLE' % sec); sec += 1
    S.header()
    S.row('Casetones en X', '=MAX(0,INT(({n_ax}*100-{bp}+{bw})/{s}))', 'pzas', 'INT((a_x − b_p + b_w)/s). Lo que sobra se reparte como macizo en las dos orillas.', 'Geometría', 'key', '0', 'n_ncx')
    S.row('Macizo en cada orilla X', '=IF({n_ncx}>0,({n_ax}*100-{bp}-({n_ncx}*{ac}+({n_ncx}-1)*{bw}))/2,{n_ax}*100/2)', 'cm', 'Del paño de la nervadura principal al primer casetón.', 'Geometría', 'calc', '0.0', 'n_mx')
    S.row('Casetones en Y', '=MAX(0,INT(({n_ay}*100-{bp}+{bw})/{s}))', 'pzas', 'Mismo cálculo.', 'Geometría', 'key', '0', 'n_ncy')
    S.row('Macizo en cada orilla Y', '=IF({n_ncy}>0,({n_ay}*100-{bp}-({n_ncy}*{ac}+({n_ncy}-1)*{bw}))/2,{n_ay}*100/2)', 'cm', 'Mismo cálculo.', 'Geometría', 'calc', '0.0', 'n_my')
    S.row('Nervaduras secundarias ‖ X (entre casetones)', '=MAX(0,{n_ncy}-1)', 'pzas', 'Una entre cada par de hileras de casetones.', 'Geometría', 'calc', '0', 'n_nrx')
    S.row('Nervaduras secundarias ‖ Y', '=MAX(0,{n_ncx}-1)', 'pzas', 'Mismo criterio.', 'Geometría', 'calc', '0', 'n_nry')
    S.row('Bastón superior ‖ claro corto: largo desde el paño', '=' + tv('Lbc'), 'cm', 'Punto de inflexión x = L_n·(1 − √(M⁺/(M⁺ + M⁻)))/2 más el mayor de d, 12d_b y L_n/16, y no menos que L_d. En volado: hacia el claro de respaldo.', NTCC + ' 6.3.7.3.4, p. 127 · 6.3.7.3.10.4, p. 129 · 14.4.2.1, p. 352', 'key', '0', 'n_Lbc')
    S.row('Bastón superior ‖ claro largo: largo desde el paño', '=' + tv('Lbl'), 'cm', 'Mismo criterio.', NTCC + ' 6.3.7.3.10.4, p. 129', 'key', '0', 'n_Lbl')
    S.row('Varilla inferior corrida en bordes discontinuos', 'entra ≥ 15 cm en el apoyo', '—', 'Una varilla inferior por nervadura, continua y anclada para 1.25 fy en el paño (L_dh·1.25 ≈ 32 cm para No. 4): en dalas de 15 cm no cabe recta; rematar con gancho dentro de la dala y continuidad con la nervadura vecina.',
          NTCC + ' 6.7.8.1.7, p. 188 · 6.7.7.4.1.1 a), p. 179 · 14.4.3.1, p. 355', 'calc', '@')
    S.row('Bastón superior en bordes discontinuos', 'gancho de 90° en la dala', '—', 'Se dobla o se ancla en el muro para que desarrolle fy en el paño (L_dh ≈ 25 cm No. 4; extensión 12d_b).', NTCC + ' 6.7.7.4.1.1 b), p. 179 · Tabla 14.3.1, p. 349', 'calc', '@')
    S.row('Refuerzo de esquina exterior (tableros de esquina sobre muros)', 'arriba y abajo en L_largo/5', '—', 'Para el M⁺ máximo del tablero, en un cuadro de un quinto del claro largo desde la esquina.', NTCC + ' 6.7.7.3.1, p. 178', 'calc', '@')
    S.gap()
    S.section('%d. CAPA DE COMPRESIÓN — MALLA, DIAFRAGMA Y CARGA CONCENTRADA (IGUAL PARA TODOS LOS TABLEROS)' % sec); sec += 1
    S.header()
    S.row('As de la malla escogida', '=INDEX(TABLAS!$D$%d:$D$%d,MATCH({malla},TABLAS!$B$%d:$B$%d,0))' % (TB_W0, TB_W1, TB_W0, TB_W1), 'cm²/m', 'TABLAS sección C.', 'TABLAS', 'calc', '0.000', 'n_asm')
    S.row('Ø del alambre', '=INDEX(TABLAS!$C$%d:$C$%d,MATCH({malla},TABLAS!$B$%d:$B$%d,0))/10' % (TB_W0, TB_W1, TB_W0, TB_W1), 'cm', 'TABLAS sección C.', 'TABLAS', 'calc', '0.000', 'n_fim')
    S.row('Contracción y temperatura: As mín = 0.0018·t·100', '=0.0018*{t}*100', 'cm²/m', 'Ya no depende de fy ni de la exposición (la fórmula 660x₁/(fy(x₁+100)) era de 2017).', NTCC + ' 13.8.2.1.1, p. 345 · 6.7.8.1.8, p. 188', 'calc', '0.000', 'n_as1')
    S.row('Estado (temperatura)', '=IF({n_asm}>={n_as1},"CUMPLE","NO CUMPLE")', '—', '—', '—', 'key', '@')
    S.row('Diafragma: cuantía mínima por cortante = 0.0025·t·100', '=0.0025*{t}*100', 'cm²/m', 'Sólo cuenta el acero de la capa (el de cambios volumétricos sí contribuye; el de flexión de las nervaduras no).', NTCC + ' 6.8.6.4 y 6.8.6.3, p. 196 · 6.8.1.1 a), p. 188', 'calc', '0.000', 'n_asd')
    S.row('Estado (diafragma)', '=IF({n_asm}>={n_asd}-0.0005,"CUMPLE","NO CUMPLE: "&TEXT({n_asm}/{n_asd},"0%")&" — capa de 6 cm con 6x6-4/4")', '—', 'Con capa de 5 cm ninguna malla de la lista cumple a la vez diafragma y recubrimiento.', '—', 'key', '@')
    S.row('Recubrimiento de la malla centrada en la capa', '=({t}-2*{n_fim})/2', 'cm', '(t − 2Ø)/2. Debe ser ≥ r (2 cm, interior A1).', NTCC + ' Tabla 4.14.16.a, p. 85', 'calc', '0.00', 'n_rm')
    S.row('Estado (recubrimiento de la malla)', '=IF({n_rm}>={rec}-0.005,"CUMPLE","NO CUMPLE")', '—', '—', '—', 'key', '@')
    S.row('Separación de la malla (15.24 cm) ≤ máx. permitida', '=IF(15.24<=MIN(5*{t},45,30),"CUMPLE","NO CUMPLE")', '—', 'El menor de 5t y 45 cm (diafragma) y 30 cm entre alambres soldados.', NTCC + ' 6.8.7.3, p. 196 · 2.4.2.1.6.3, p. 44', 'calc', '@')
    S.row('Espesor de pisos y rellenos que reparten la carga concentrada (45°)', 0, 'cm', 'CAPTURA. 0 = la carga actúa directo sobre la capa (lo más estricto). Si cuentas el piso y el pegapiso (≈ 6 cm), pon aquí su espesor.', 'Criterio de esta hoja', 'input', '0.0', 'tpz')
    S.row('P_u = FC_CV × 1 000 kg en 10 × 10 cm', '={FCcv}*' + N['Pcap'], 'kg', 'Carga concentrada mínima entre nervaduras.', NTCC + ' 6.7.8.1.6, p. 187 · ' + NTCCA + ' 3.4.1 a), p. 7', 'calc', '0', 'n_Pu')
    S.row('d de la capa (malla al centro)', '={t}/2', 'cm', '—', 'Geometría', 'calc', '0.00', 'n_dc')
    S.row('b₀ — perímetro crítico a d/2', '=4*(10+2*{tpz}+{n_dc})', 'cm', '4·(c + d) con c = 10 cm más el reparto.', NTCC + ' 5.6, p. 99–100', 'calc', '0.0', 'n_bo')
    S.row('V_cR punzonamiento de la capa', "={FRv}*MIN(1,0.5*(1+2/1),0.27*(2+40*{n_dc}/{n_bo}))*MIN(1,SQRT(2/(1+0.04*{n_dc})))*SQRT({fc})*{n_bo}*{n_dc}", 'kg', "v_cR = F_R·(el menor de 1.0, 0.5(1 + 2/β) y 0.27(2 + α_s·d/b₀))·λs·√f'c, con β = 1 y α_s = 40.", NTCC + ' Tabla 5.6.2.1.1, p. 100', 'calc', '0', 'n_vcp')
    S.row('Estado (punzonamiento)', '=IF({n_Pu}<={n_vcp},"CUMPLE","NO CUMPLE: P_u/V_cR = "&TEXT({n_Pu}/{n_vcp},"0.00")&" — capa de 6 cm o contar el reparto del piso")', '—', '—', '—', 'key', '@')
    S.row('m_R de la capa por unidad de ancho', "={FRf}*{n_asm}/100*{fym}*({n_dc}-{n_asm}/100*{fym}/({fpc})/2)", 'kg·cm/cm', "F_R·As·fy·(d − a/2) por cm.", NTCC + ' ec. 5.2.2.1.1.1.b, p. 89', 'calc', '0.0', 'n_mR')
    S.row('m requerido = P_u/(4π) (abanico de líneas de fluencia, m = m′)', '={n_Pu}/(4*PI())', 'kg·cm/cm', 'Mecanismo de abanico alrededor de la carga; la malla única al centro da m = m′. Ignorar el área cargada queda del lado seguro.', 'Criterio de esta hoja · ' + NTCC + ' 3.4.1.1, p. 64 (líneas de fluencia) · 6.7.8.1.8, p. 188', 'calc', '0.0', 'n_mq')
    S.row('Estado (flexión de la capa)', '=IF({n_mq}<={n_mR},"CUMPLE","NO CUMPLE")', '—', '—', '—', 'key', '@')
    S.gap()
    S.section('%d. COMPROBACIÓN CONTRA TABLEROS' % sec)
    S.header()
    S.row('Mu/M_R inferior corto (aquí / TABLEROS)', '=TEXT({n_FUPc},"0.000")&" / "&TEXT(%s,"0.000")' % tv('FUPc'), '—', 'Deben coincidir.', '—', 'calc', '@')
    S.row('Mu/M_R superior corto (aquí / TABLEROS)', '=TEXT({n_FUNc},"0.000")&" / "&TEXT(%s,"0.000")' % tv('FUNc'), '—', 'Deben coincidir.', '—', 'calc', '@')
    S.row('¿Coinciden?', '=IF(AND(ABS({n_FUPc}-%s)<0.0005,ABS({n_FUNc}-%s)<0.0005,ABS({n_FUPl}-%s)<0.0005,ABS({n_FUNl}-%s)<0.0005,ABS({n_VcR}-%s)<0.0005),"CUMPLE — coinciden","REVISAR")' % (tv('FUPc'), tv('FUNc'), tv('FUPl'), tv('FUNl'), tv('VcR')),
          '—', 'Control interno de la hoja.', '—', 'key', '@')
    S.row('ESTADO DEL TABLERO (TABLEROS)', '=' + tv('estado'), '—', '—', '—', 'key', '@')
    add_cf(ws, 'C1:F%d' % S.r)
    ws.sheet_view.zoomScale = 85
    return ws


# ---------------------------------------------------------------- PRINCIPALES
PC = {}
_pcols = ['np', 'nivel', 'eje', 'de', 'a', 'L', 'ei', 'ej', 'ci', 'cj', 'cond', 'raiz', 'bq', 'hq', 'cambio',
          'At', 'wt', 'Lm', 'wm', 'waz', 'scm', 'cv', 'wl', 'wpp', 'w', 'wD', 'wL', 'wW',
          'SP', 'SM', 'SV', 'SPm', 'SMs', 'Ri', 'Rj', 'Ris', 'Rjs', 'Mp', 'Mn', 'Vu', 'Ms',
          'd', 'AsP', 'vP', 'nPx', 'nP', 'MRP', 'FUP', 'eP', 'AsN', 'vN', 'nN', 'MRN', 'FUN', 'eN',
          'VcR', 'se', 'sext', 'cort', 'bfq', 'Ig', 'Ie', 'dlt', 'dlim', 'Pu', 'Kf', 'fcm', 'm418', 'htab', 'h418', 'hflex', 'hrec', 'estado', 'nota', 'nmax']
for i, nm in enumerate(_pcols):
    PC[nm] = CL(2 + i)
PC_LAST = CL(2 + len(_pcols) - 1)
XROW = {}
# vigas a lo largo del perímetro del edificio (fachadas y = 1.919, y = 21.446, x = 0, x = 8.85): integridad 6.3.7.7.1 y 2h de 6.3.7.7.7
PERIM = {t['np'] for t in NPS if t['cond'] != 'BORDE DE VOLADO' and ((t['ori'] == 'H' and min(abs(t['c'] - 1.919), abs(t['c'] - 21.446)) < 0.01) or (t['ori'] == 'V' and min(abs(t['c'] - 0.0), abs(t['c'] - 8.85)) < 0.01))}
_ntab = {}
for _tr in D['trib']:
    if _tr['tramo'].startswith('NP') and _tr['area'] > 0.05:
        _ntab.setdefault(_tr['tramo'], set()).add(_tr['tablero'])
ONESIDE = {t['np'] for t in NPS if t['np'] in PERIM or len(_ntab.get(t['np'], ())) < 2}
PROP = {}
try:
    PROP = json.load(open(W + 'work/propuesta_np.json'))
except Exception:
    PROP = {}


def trib_sum(col, tramo):
    return 'SUMIFS(TRIB!$%s$%d:$%s$%d,TRIB!$B$%d:$B$%d,"%s")' % (col, R0, col, R1, R0, R1, tramo)


def sheet_principales(wb):
    ws = wb.create_sheet('PRINCIPALES'); ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 1.78
    last = PC_LAST
    ws.merge_cells('B1:%s1' % last)
    c = ws['B1']; c.value = 'PRINCIPALES — NERVADURAS PRINCIPALES SIN MURO ABAJO (TRABAJAN COMO TRABE)'
    c.font = font(13, True, color='FFFFFF'); c.fill = fill(TITLE_BG); ws.row_dimensions[1].height = 26
    ws.merge_cells('B2:%s2' % last)
    ws['B2'].value = ('Cada renglón es un tramo de dala del modelo que NO tiene muro abajo, entre dos apoyos. Carga = áreas tributarias del DXF × w_u del tablero + muro de planta alta encima + la AZOTEA que ese muro recibe por su dala de N+6.00 + cargas lineales del E2K + peso extra de la trabe + reacciones de otras principales (TRANSFER, directas o por el muro). '
                      'Apoyada: M⁺ = wL²/8 + ΣPab/L, M⁻ = 0.667·M⁺. Volado: M⁻ = wL²/2 + ΣP·brazo. Recubrimiento de viga (4 cm), estribos cerrados No. 3 y A_v,mín en todo el claro; flecha con la inercia efectiva de la Tabla 13.4.3.2; con muro encima, además la NTC-Mampostería 4.1.18.')
    ws['B2'].font = font(9, i=True, color=GRAY); ws['B2'].alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[2].height = 42
    ws.merge_cells('B3:%s3' % last)
    ws['B3'].value = ('b y h (amarillo) son la SECCIÓN PROPUESTA: la hoja la dimensionó para que el tramo cumpla todo; donde dice «PERALTAR» la dala de 15×25 del modelo no alcanza y debe ir como trabe descolgada de esa sección. '
                      'Los extremos con castillo (C1…C8) dependen de lo que decidas en APOYOS sección B. «h requerido» es el mínimo por reglas de peralte y por flexión; la flecha y el muro sobre la viga se revisan con la h capturada.')
    ws['B3'].font = font(9, True, color=WARN_FG); ws['B3'].fill = fill(WARN_BG); ws['B3'].alignment = Alignment(wrap_text=True, vertical='top'); ws.row_dimensions[3].height = 30
    groups = [('np', 'raiz', 'TRAMO Y APOYOS'), ('bq', 'cambio', 'SECCIÓN PROPUESTA'), ('At', 'wW', 'CARGA REPARTIDA (última; servicio D / W_m / W)'), ('SP', 'Rjs', 'CARGAS PUNTUALES Y REACCIONES'),
              ('Mp', 'Ms', 'ELEMENTOS MECÁNICOS'), ('d', 'eN', 'FLEXIÓN'), ('VcR', 'cort', 'CORTANTE'), ('bfq', 'dlim', 'FLECHA (13.4) — SECCIÓN T CON LA CAPA'), ('Pu', 'm418', 'MURO SOBRE VIGA (NTC-M 4.1.18)'), ('htab', 'nmax', 'PERALTE Y ESTADO')]
    for a, b, t in groups:
        ws.merge_cells('%s6:%s6' % (PC[a], PC[b]))
        x = ws['%s6' % PC[a]]; x.value = t; x.font = font(9, True, color='FFFFFF'); x.fill = fill(SEC_BG); x.alignment = Alignment(horizontal='center')
    heads = {'np': ('Tramo', 7), 'nivel': ('Nivel', 7.5), 'eje': ('Eje / línea', 16), 'de': ('De (x, y)', 13), 'a': ('A (x, y)', 13), 'L': ('L', 6.5), 'ei': ('Extremo i', 10.5), 'ej': ('Extremo j', 10.5),
             'ci': ('Castillo en i', 7), 'cj': ('Castillo en j', 7), 'cond': ('Condición', 13), 'raiz': ('Raíz del volado', 6.5), 'bq': ('b', 6), 'hq': ('h', 6), 'cambio': ('Respecto al modelo (15×25)', 16),
             'At': ('Área trib.', 7), 'wt': ('w losa', 7), 'Lm': ('Muro encima', 7), 'wm': ('w muro', 7), 'waz': ('w azotea por el muro', 7.5), 'scm': ('SCM lineal E2K', 7.5), 'cv': ('CV lineal E2K', 7.5),
             'wl': ('w lineal', 7), 'wpp': ('w peso extra', 7), 'w': ('w_u total', 7.5), 'wD': ('w_D serv.', 7), 'wL': ('w_L serv. (W_m)', 7), 'wW': ('w_L media (W)', 7),
             'SP': ('ΣP_u recibidas', 7.5), 'SM': ('ΣM_u por P', 7.5), 'SV': ('ΣV_u por P', 7.5), 'SPm': ('ΣP_u por el muro', 7.5), 'SMs': ('ΣM serv. por P', 7.5),
             'Ri': ('R_i', 7), 'Rj': ('R_j', 7), 'Ris': ('R_i serv.', 7), 'Rjs': ('R_j serv.', 7), 'Mp': ('Mu⁺', 7), 'Mn': ('Mu⁻', 7), 'Vu': ('Vu', 7), 'Ms': ('M serv.', 7), 'd': ('d', 6),
             'AsP': ('As⁺ req.', 7), 'vP': ('Var.', 5.5), 'nPx': ('Extra por flecha', 6), 'nP': ('Pzas', 5), 'MRP': ('M_R⁺', 7), 'FUP': ('Mu⁺/M_R', 7), 'eP': ('ε_t⁺', 6.5),
             'AsN': ('As⁻ req.', 7), 'vN': ('Var.', 5.5), 'nN': ('Pzas', 5), 'MRN': ('M_R⁻', 7), 'FUN': ('Mu⁻/M_R', 7), 'eN': ('ε_t⁻', 6.5),
             'VcR': ('V_cR', 7), 'se': ('Sep. estribo claro', 7), 'sext': ('Sep. en 2h extremos', 7), 'cort': ('Cortante', 20),
             'bfq': ('Patín b_f', 6.5), 'Ig': ('I_g sección T', 8.5), 'Ie': ('I_e', 8), 'dlt': ('Flecha total', 7), 'dlim': ('Límite', 7), 'Pu': ('P_u del muro', 7), 'Kf': ('K', 6), 'fcm': ('f_CM', 7), 'm418': ('Estado muro', 16),
             'htab': ('h mín. (Tabla 6.3.3.1.1 / NTC-M 4.1.18.4)', 9), 'h418': ('h por NTC-M 4.1.18.3', 8), 'hflex': ('h req. flexión', 8), 'hrec': ('h requerido', 8), 'estado': ('ESTADO', 30),
             'nota': ('Nota', 50), 'nmax': ('Ancho libre para varillas', 7)}
    units = {'L': 'm', 'bq': 'cm', 'hq': 'cm', 'At': 'm²', 'wt': 't/m', 'Lm': 'm', 'wm': 't/m', 'waz': 't/m', 'scm': 'kg/m', 'cv': 'kg/m', 'wl': 't/m', 'wpp': 't/m', 'w': 't/m', 'wD': 't/m', 'wL': 't/m', 'wW': 't/m',
             'SP': 't', 'SM': 't·m', 'SV': 't', 'SPm': 't', 'SMs': 't·m', 'Ri': 't', 'Rj': 't', 'Ris': 't', 'Rjs': 't', 'Mp': 't·m', 'Mn': 't·m', 'Vu': 't', 'Ms': 't·m', 'd': 'cm',
             'AsP': 'cm²', 'MRP': 't·m', 'AsN': 'cm²', 'MRN': 't·m', 'eP': '≥ 0.005', 'eN': '≥ 0.005', 'VcR': 't', 'se': 'cm', 'sext': 'cm', 'bfq': 'cm', 'Ig': 'cm⁴', 'Ie': 'cm⁴', 'dlt': 'cm', 'dlim': 'cm',
             'Pu': 't', 'fcm': 'kg/cm²', 'htab': 'cm', 'h418': 'cm', 'hflex': 'cm', 'hrec': 'cm', 'nmax': 'cm'}
    for nm, (t, w) in heads.items():
        hdr_cell(ws, 7, CI(PC[nm]), t, w); unit_cell(ws, 8, CI(PC[nm]), units.get(nm, ''))
    ws.row_dimensions[7].height = 42
    for i, x in enumerate(D['transfer']):
        XROW[i] = X0 + i
    F = lambda s, r: s.format(r=r, **{k: v for k, v in PC.items()}, **N)
    sup = lambda c: 'OR({%s}{r}="FIRME",{%s}{r}="NERVADURA",{%s}{r}="PUNTA")' % (c, c, c)
    for t in NPS:
        r = PROW[t['np']]
        ext = t['ext']

        def etxt(k):
            e = ext[k]
            if e.get('prop'):
                return '=IF(%s="SÍ","FIRME","LIBRE")' % PROPROW[e['prop']]
            return e['tipo']
        ei = etxt(0); ej = etxt(1)
        base = t['cond']
        incoming = [(i, x) for i, x in enumerate(D['transfer']) if x['a'] == t['np']]
        sumcol = lambda col, via=None: ('+'.join('TRANSFER!$%s$%d' % (col, XROW[i]) for i, x in incoming if via is None or x['via'] == via)) or '0'
        bq0, hq0, nx0 = (PROP.get(t['np'], []) + [None, None, 0])[:3]
        waz = WAZ.get(t['np'], [])
        wz = lambda col: ('(' + '+'.join('%s*%s' % (trib_sum(col, dm), f) for dm, f in waz) + ')/{L}{r}') if waz else '0'
        wzD = ('(' + '+'.join('(%s-%s)*%s' % (trib_sum('I', dm), trib_sum('L', dm), f) for dm, f in waz) + ')/{L}{r}') if waz else '0'
        vals = {
            'np': (t['np'], 'key', '@'), 'nivel': (t['level'], 'calc', '@'), 'eje': (t['eje'], 'calc', '@'),
            'de': ('(%.3f, %.3f)' % (t['x0'], t['y0']), 'calc', '@'), 'a': ('(%.3f, %.3f)' % (t['x1'], t['y1']), 'calc', '@'),
            'L': (t['L'], 'input', '0.000'), 'ei': (ei, 'calc', '@'), 'ej': (ej, 'calc', '@'),
            'ci': (ext[0].get('prop') or '—', 'calc', '@'), 'cj': (ext[1].get('prop') or '—', 'calc', '@'),
            'cond': (('BORDE DE VOLADO', 'calc', '@') if base == 'BORDE DE VOLADO' else
                     ('=IF(AND(%s,%s),"APOYADA",IF(OR(%s,%s),"VOLADO","SIN APOYO"))' % (sup('ei'), sup('ej'), sup('ei'), sup('ej')), 'key', '@')),
            'raiz': ('=IF({cond}{r}="VOLADO",IF(%s,"i","j"),"—")' % sup('ei'), 'calc', '@'),
            'bq': (bq0 if bq0 else '={bp}', 'input', '0'), 'hq': (hq0 if hq0 else '={h}', 'input', '0'),
            'cambio': ('=IF({cond}{r}="BORDE DE VOLADO","—",IF(OR({bq}{r}<>{bp},{hq}{r}<>{h}),"PERALTAR: trabe "&TEXT({bq}{r},"0")&"×"&TEXT({hq}{r},"0")&IF({hq}{r}>{h}," (descuelga "&TEXT({hq}{r}-{h},"0")&" cm)",""),"Queda en la losa (15×25)"))', 'calc', '@'),
            'At': ('=' + trib_sum('E', t['np']), 'calc', '0.00'),
            'wt': ('=' + trib_sum('G', t['np']) + '/{L}{r}', 'calc', '0.000'),
            'Lm': (t['muro_arriba'], 'input', '0.00'), 'wm': ('={FCcm}*{Lm}{r}*{wmuro}/{L}{r}', 'calc', '0.000'),
            'waz': ('=' + wz('G'), 'calc', '0.000'),
            'scm': (t['e2k_scm'], 'input', '0.0'), 'cv': (t['e2k_cv'], 'input', '0.0'), 'wl': ('=({FCcm}*{scm}{r}+{FCcv}*{cv}{r})/1000', 'calc', '0.000'),
            'wpp': ('={FCcm}*({bq}{r}/100*{hq}{r}/100*{gc}-{bq}{r}/100*{PP}/1000)', 'calc', '0.000'),
            'w': ('={wt}{r}+{wm}{r}+{waz}{r}+{wl}{r}+{wpp}{r}', 'key', '0.000'),
            'wD': ('=(%s-%s)/{L}{r}+{Lm}{r}*{wmuro}/{L}{r}+%s+{scm}{r}/1000+{wpp}{r}/{FCcm}' % (trib_sum('I', t['np']), trib_sum('L', t['np']), wzD), 'calc', '0.000'),
            'wL': ('=%s/{L}{r}+%s+{cv}{r}/1000' % (trib_sum('L', t['np']), wz('L')), 'calc', '0.000'),
            'wW': ('=%s/{L}{r}+%s+{cv}{r}/1000*{cvw_hab}/{cv_hab}' % (trib_sum('M', t['np']), wz('M')), 'calc', '0.000'),
            'SP': ('=' + sumcol('F'), 'calc', '0.000'), 'SM': ('=' + sumcol('H'), 'calc', '0.000'), 'SV': ('=' + sumcol('I'), 'calc', '0.000'),
            'SPm': ('=' + sumcol('F', 'muro'), 'calc', '0.000'), 'SMs': ('=' + sumcol('N'), 'calc', '0.000'),
            'Ri': ('=IF({cond}{r}="APOYADA",{w}{r}*{L}{r}/2+%s,IF(AND({cond}{r}="VOLADO",{raiz}{r}="i"),{w}{r}*{L}{r}+{SP}{r},0))' % sumcol('J'), 'calc', '0.000'),
            'Rj': ('=IF({cond}{r}="APOYADA",{w}{r}*{L}{r}/2+%s,IF(AND({cond}{r}="VOLADO",{raiz}{r}="j"),{w}{r}*{L}{r}+{SP}{r},0))' % sumcol('K'), 'calc', '0.000'),
            'Ris': ('=IF({cond}{r}="APOYADA",({wD}{r}+{wL}{r})*{L}{r}/2+%s,IF(AND({cond}{r}="VOLADO",{raiz}{r}="i"),({wD}{r}+{wL}{r})*{L}{r}+%s,0))' % (sumcol('O'), sumcol('M')), 'calc', '0.000'),
            'Rjs': ('=IF({cond}{r}="APOYADA",({wD}{r}+{wL}{r})*{L}{r}/2+%s,IF(AND({cond}{r}="VOLADO",{raiz}{r}="j"),({wD}{r}+{wL}{r})*{L}{r}+%s,0))' % (sumcol('P'), sumcol('M')), 'calc', '0.000'),
            'Mp': ('=IF({cond}{r}="APOYADA",{w}{r}*{L}{r}^2/8+{SM}{r},0)', 'key', '0.000'),
            'Mn': ('=IF({cond}{r}="APOYADA",{Mrat}*{Mp}{r},IF({cond}{r}="VOLADO",{w}{r}*{L}{r}^2/2+{SM}{r},0))', 'key', '0.000'),
            'Vu': ('=IF({cond}{r}="APOYADA",{w}{r}*{L}{r}/2+{SV}{r},IF({cond}{r}="VOLADO",{w}{r}*{L}{r}+{SV}{r},0))', 'key', '0.000'),
            'Ms': ('=IF({cond}{r}="APOYADA",({wD}{r}+{wL}{r})*{L}{r}^2/8+{SMs}{r},IF({cond}{r}="VOLADO",({wD}{r}+{wL}{r})*{L}{r}^2/2+{SMs}{r},0))', 'calc', '0.000'),
            'd': ('={hq}{r}-{rec_np}-' + db('{est_np}') + '-' + db('{var_np}') + '/2', 'calc', '0.00'),
            'nmax': ('={bq}{r}-2*{rec_np}-2*' + db('{est_np}'), 'calc', '0.00'),
        }
        vals['nPx'] = (nx0 or 0, 'input', '0')
        for blk, M in (('P', 'Mp'), ('N', 'Mn')):
            asr = asreq('{%s}{r}' % M, '{bq}{r}', '{d}{r}', True, '{bq}{r}')
            vals['As' + blk] = ('=IF(OR({cond}{r}="BORDE DE VOLADO",{cond}{r}="SIN APOYO"),0,%s)' % asr, 'calc', '0.00')
            vals['v' + blk] = ('=IF({As%s}{r}<0,0,' % blk + varsel2('{As%s}{r}' % blk, '{var_np}', '99', '{nmax}{r}', '{nmin_np}') + ')', 'calc', 'General')
            extra = '+{nPx}{r}' if blk == 'P' else ''
            vals['n' + blk] = ('=IF(OR({cond}{r}="BORDE DE VOLADO",{cond}{r}="SIN APOYO"),0,IF({As%s}{r}<0,0,MAX({nmin_np},CEILING({As%s}{r}/%s,1))%s))' % (blk, blk, ab('IF({v%s}{r}=0,%s,{v%s}{r})' % (blk, BARS[-1][0], blk)), extra), 'calc', '0')
            asp = '{n%s}{r}*%s' % (blk, ab('{v%s}{r}' % blk))
            dact = '({hq}{r}-{rec_np}-%s-%s/2)' % (db('{est_np}'), db('{v%s}{r}' % blk))
            vals['MR' + blk] = ('=IF({n%s}{r}=0,0,{FRf}*%s*{fy}*(%s-%s*{fy}/({fpc}*{bq}{r})/2)/1E5)' % (blk, asp, dact, asp), 'calc', '0.000')
            vals['FU' + blk] = ('=IF({As%s}{r}<0,9.99,IF({MR%s}{r}>0,{%s}{r}/{MR%s}{r},0))' % (blk, blk, M, blk), 'key', '0.00')
            cexp = '(%s*{fy}/({fpc}*{bq}{r})/{beta1})' % asp
            vals['e' + blk] = ('=IF({n%s}{r}=0,"—",0.003*(%s-%s)/%s)' % (blk, dact, cexp, cexp), 'calc', '0.0000')
        # cortante con A_v,mín: V_cR = F_R·máx(0.5, 2p^(1/3))·√f'c·b·d ≤ F_R·1.25·√f'c·b·d
        pmn = 'MAX(1E-6,MIN({nP}{r}*%s,{nN}{r}*%s)/({bq}{r}*{d}{r}))' % (ab('MAX({vP}{r},2)'), ab('MAX({vN}{r},2)'))
        vals['VcR'] = ('=IF({nN}{r}=0,0,{FRv}*MIN(1.25,MAX(0.5,2*%s^(1/3)))*SQRT({fc})*{bq}{r}*{d}{r}/1000)' % pmn, 'calc', '0.000')
        av = '2*' + ab('{est_np}')
        vs = 'MAX(0,{Vu}{r}-{VcR}{r})/{FRv}'
        smaxv = 'IF(%s<=1.1*SQRT({fc})*{bq}{r}*{d}{r}/1000,MIN({d}{r}/2,60),MIN({d}{r}/4,30))' % vs
        sreq = 'IF({Vu}{r}>{VcR}{r},{FRv}*%s*{fyt}*{d}{r}/(({Vu}{r}-{VcR}{r})*1000),999)' % av
        savm = '%s*{fyt}/(MAX(0.2*SQRT({fc}),3.5)*{bq}{r})' % av
        s147 = 'MIN(16*%s,48*%s)' % (db('MAX({vP}{r},{vN}{r},2)'), db('{est_np}'))
        vals['se'] = ('=IF({Vu}{r}=0,0,FLOOR(MIN(%s,%s,%s,%s,{smax_np}),0.5))' % (sreq, smaxv, savm, s147), 'calc', '0.0')
        if t['np'] in PERIM:
            vals['sext'] = ('=IF({Vu}{r}=0,0,FLOOR(MIN({se}{r},{d}{r}/4,8*%s,24*%s,30),0.5))' % (db('MIN(MAX({vP}{r},2),MAX({vN}{r},2))'), db('{est_np}')), 'calc', '0.0')
        else:
            vals['sext'] = ('={se}{r}', 'calc', '0.0')
        vals['cort'] = ('=IF({Vu}{r}=0,"—",IF({Vu}{r}>{FRv}*({VcR}{r}/{FRv}+2.2*SQRT({fc})*{bq}{r}*{d}{r}/1000),"SECCIÓN INSUFICIENTE",'
                        '"E No. "&TEXT({est_np},"0")&IF({sext}{r}<{se}{r}," @ "&TEXT({sext}{r},"0.0")&" en 2h y"," @")&" "&TEXT({se}{r},"0.0")&" cm"))', 'calc', '@')
        # flecha: Bischoff (Tabla 13.4.3.2), inmediata + diferida, con la sección T que forma la capa de compresión (Tabla 5.2.1.4.2)
        if t['np'] in ONESIDE:
            vals['bfq'] = ('={bq}{r}+MIN(6*{t},{ac}/2,{L}{r}*100/12)', 'calc', '0')
        else:
            vals['bfq'] = ('={bq}{r}+2*MIN(8*{t},{ac}/2,{L}{r}*100/8)', 'calc', '0')
        AgT = '({bfq}{r}*{t}+{bq}{r}*({hq}{r}-{t}))'
        ybT = '(({bfq}{r}*{t}^2/2+{bq}{r}*({hq}{r}-{t})*({t}+({hq}{r}-{t})/2))/%s)' % AgT
        vals['Ig'] = ('={bfq}{r}*{t}^3/12+{bfq}{r}*{t}*(%s-{t}/2)^2+{bq}{r}*({hq}{r}-{t})^3/12+{bq}{r}*({hq}{r}-{t})*({t}+({hq}{r}-{t})/2-%s)^2' % (ybT, ybT), 'calc', '#,##0')
        Ig = '{Ig}{r}'
        AsT = 'IF({cond}{r}="VOLADO",{nN}{r}*%s,{nP}{r}*%s)' % (ab('MAX({vN}{r},2)'), ab('MAX({vP}{r},2)'))
        bc = 'IF({cond}{r}="VOLADO",{bq}{r},{bfq}{r})'
        rn = '({Es}/{Ec}*%s/(%s*{d}{r}))' % (AsT, bc)
        kd = '((SQRT(2*%s+%s^2)-%s)*{d}{r})' % (rn, rn, rn)
        Icr = '(%s*%s^3/3+{Es}/{Ec}*%s*({d}{r}-%s)^2)' % (bc, kd, AsT, kd)
        Mag = '({ffr}*%s/IF({cond}{r}="VOLADO",%s,{hq}{r}-%s))' % (Ig, ybT, ybT)
        Mse = '({Ms}{r}*1E5)'
        vals['Ie'] = ('=IF({Ms}{r}<=0,%s,IF(%s<=2/3*%s,%s,IF(%s<=0,0,%s/(1-(2/3*%s/%s)^2*(1-%s/%s)))))' % (Ig, Mse, Mag, Ig, AsT, Icr, Mag, Mse, Icr, Ig), 'calc', '#,##0')
        pcomp = 'IF({cond}{r}="VOLADO",{nP}{r}*%s,{nmin_np}*%s)/({bq}{r}*{d}{r})' % (ab('MAX({vP}{r},2)'), ab('MAX({vN}{r},2)'))
        rs = '(({wD}{r}+{wW}{r})/MAX({wD}{r}+{wL}{r},1E-6))'
        di = '(IF({cond}{r}="VOLADO",1/4,5/48)*%s*({L}{r}*100)^2/({Ec}*MAX({Ie}{r},1)))' % Mse
        vals['dlt'] = ('=IF(OR({cond}{r}="APOYADA",{cond}{r}="VOLADO"),%s*(1+2/(1+50*%s)*%s),0)' % (di, pcomp, rs), 'key', '0.00')
        vals['dlim'] = ('=IF({Lm}{r}>0,0.3+{L}{r}*100/480,0.5+{L}{r}*100/240)', 'calc', '0.00')
        # NTC-Mampostería 4.1.18: muro sobre viga
        ATm = '({tm}*100*{Lm}{r}*100)'
        a_ = 'MIN(1,{Lm}{r}/{L}{r})'
        Wp = '({Pu}{r}*1000/%s*{tm}*100*(%s^4-4*%s^3+8*%s))' % (ATm, a_, a_, a_)
        wu4 = '(({wt}{r}+{wl}{r}+{wpp}{r})*10)'
        fI = '(%s/(%s+5*%s))' % (Wp, Wp, wu4)
        vals['Pu'] = ('=IF({Lm}{r}>0,({wm}{r}+{waz}{r})*{L}{r}+{SPm}{r},0)', 'calc', '0.00')
        vals['Kf'] = ('=IF({Lm}{r}>0,({Em}*{tm}*100*({L}{r}*100)^3/({Ec}*%s*%s))^0.25,0)' % (Ig, fI), 'calc', '0.00')
        fce1 = '(1.6*{Kf}{r}-3.9)'; fce9 = '((3.2*{Kf}{r}-7.8)*1)'
        fce = 'MAX(1,IF(%s>=1,%s,IF(%s<0.9,%s,%s+(%s-%s)*(%s-0.9)/0.1)))' % (a_, fce1, a_, fce9, fce9, fce1, fce9, a_)
        vals['fcm'] = ('=IF({Lm}{r}>0,%s*{Pu}{r}*1000/%s,0)' % (fce, ATm), 'calc', '0.0')
        PD = '(({Lm}{r}*{wmuro}+(%s)*{L}{r}+{SPm}{r}/{FCcm})*1000)' % (wzD + '+' + wz('M'))
        WpD = '(%s/%s*{tm}*100*(%s^4-4*%s^3+8*%s))' % (PD, ATm, a_, a_, a_)
        wuD = '(({wD}{r}-{Lm}{r}*{wmuro}/{L}{r}-(%s))*10)' % wzD
        h4183 = '({L}{r}*100/((1/15*{Ec}*{bq}{r}/(%s+5*%s))^(1/3)))' % (WpD, wuD)
        I4183 = '(5/4*({L}{r}*100)^3/{Ec}*(%s+5*%s))' % (WpD, wuD)
        # h que da la misma I bruta con sección rectangular (para proponer peralte): h = (12·I/b)^(1/3)
        vals['h418'] = ('=IF({Lm}{r}>0,IF({Ig}{r}>=%s,0,MIN(%s,(12*%s/{bq}{r})^(1/3))),0)' % (I4183, h4183, I4183), 'calc', '0.0')
        vals['m418'] = ('=IF({Lm}{r}=0,"—",IF({hm}/{Lm}{r}<0.6,"REVISAR: H/L < 0.6, interacción muro-viga",IF({fcm}{r}<={FRm}*{fm},"CUMPLE","NO CUMPLE f_CM > F_R·f\'m")))', 'calc', '@')
        vals['htab'] = ('=IF({Lm}{r}>0,MAX(25,IF({Ig}{r}>=({L}{r}*100)^3/1600,0,MIN({L}{r}*100/14,(12*({L}{r}*100)^3/1600/{bq}{r})^(1/3)))),IF({cond}{r}="APOYADA",MAX(25,{L}{r}*100/16),IF({cond}{r}="VOLADO",MAX(25,{L}{r}*100/8),0)))', 'calc', '0.0')
        qm = '({pmax}*{fy}/{fpc})'
        vals['hflex'] = ('=SQRT(MAX({Mp}{r},{Mn}{r})*1E5/({FRf}*{bq}{r}*{fpc}*%s*(1-%s/2)))+({hq}{r}-{d}{r})' % (qm, qm), 'calc', '0.0')
        vals['hrec'] = ('=IF(OR({cond}{r}="BORDE DE VOLADO",{cond}{r}="SIN APOYO"),0,CEILING(MAX({htab}{r},{h418}{r},{hflex}{r},{h}),5))', 'key', '0')
        fitP = 'IF({nP}{r}=0,TRUE,{nP}{r}<=%s)' % nfit('MAX({vP}{r},2)', '{nmax}{r}'); fitN = 'IF({nN}{r}=0,TRUE,{nN}{r}<=%s)' % nfit('MAX({vN}{r},2)', '{nmax}{r}')
        etok = 'AND(OR(NOT(ISNUMBER({eP}{r})),{eP}{r}>=0.005),OR(NOT(ISNUMBER({eN}{r})),{eN}{r}>=0.005))'
        defok = '{dlt}{r}<={dlim}{r}'
        hok = 'AND({hq}{r}>={htab}{r},{hq}{r}>={h418}{r},{hq}{r}/{bq}{r}<=6)'
        mok = 'OR({Lm}{r}=0,LEFT({m418}{r},6)<>"NO CUM")'
        ok = 'AND({FUP}{r}<=1,{FUN}{r}<=1,%s,%s,%s,{cort}{r}<>"SECCIÓN INSUFICIENTE",%s,%s,%s)' % (fitP, fitN, etok, hok, defok, mok)
        vals['estado'] = ('=IF({cond}{r}="BORDE DE VOLADO","BORDE DE VOLADO — armado de orilla",IF({cond}{r}="SIN APOYO","SIN APOYO — REVISAR",IF(%s,IF(LEFT({m418}{r},7)="REVISAR","CUMPLE — REVISAR interacción muro-viga","CUMPLE")&IF({cond}{r}="VOLADO"," (en VOLADO)",""),'
                          '"NO CUMPLE: "&IF(OR({AsP}{r}<0,{AsN}{r}<0),"sección insuficiente en flexión ",IF(OR({FUP}{r}>1,{FUN}{r}>1),"flexión ",""))&IF(NOT(AND(%s,%s)),"no caben ","")&IF(NOT(%s),"ε_t ","")'
                          '&IF({cort}{r}="SECCIÓN INSUFICIENTE","cortante ","")&IF(NOT(%s),"peralte ","")&IF(NOT(%s),"flecha ","")&IF(NOT(%s),"muro sobre viga ","")&"→ h ≥ "&TEXT({hrec}{r},"0")&" cm")))'
                          % (ok, fitP, fitN, etok, hok, defok, mok), 'key', '@')
        notes = []
        for k, e in enumerate(ext):
            if e.get('prop'):
                notes.append('Extremo %s en castillo propuesto %s.' % ('i' if k == 0 else 'j', e['prop']))
            if e['tipo'] == 'NERVADURA':
                notes.append('Extremo %s se apoya en %s.' % ('i' if k == 0 else 'j', e['recibe']))
            if e['tipo'] == 'PUNTA':
                notes.append('Extremo %s se apoya en la punta del volado corto %s.' % ('i' if k == 0 else 'j', e['recibe']))
        if incoming:
            notes.append('Recibe a ' + ', '.join(x['de'] + (' (por el muro de arriba)' if x['via'] == 'muro' else '') for _, x in incoming) + '.')
        if t['np'] in PERIM:
            notes.append('Viga perimetral: refuerzo de integridad continuo arriba y abajo (≥ 2 barras) y estribos cerrados en todo el claro, más cerrados en 2h junto a cada apoyo (NTC-Concreto 6.3.7.7.1, 6.3.7.7.7 y 6.3.7.7.8, p. 134).')
        if t['muro_arriba'] > 0:
            notes.append('Muro de planta alta encima (%.2f m)%s: viga que soporta muro estructural (NTC-Mampostería 4.1.18): la flecha se calcula y el límite es 0.3 cm + L/480.' % (t['muro_arriba'], (' con la azotea de ' + ', '.join(dm for dm, _ in waz)) if waz else ''))
        if t['np'] == 'NP-11':
            notes.append('PROPUESTO: la dala bajo el muro de y = 14.247 se prolonga hasta la nervadura del eje F (x = 5.44); en el modelo remata suelta dentro de F80009.')
        if t['level'] == 'N+3.00' and 8.8 < t['y0'] < 10 and 5.3 < t['x0'] < 7.4 and t['np'] in ('NP-07', 'NP-30', 'NP-25'):
            notes.append('Llegada de la escalera: revisar junto con el diseño de la escalera.')
        if base == 'BORDE DE VOLADO':
            notes.append('Orilla libre de tablero en volado: lleva el armado de la orilla, no trabaja como trabe.')
        vals['nota'] = (' '.join(notes) or '—', 'calc', '@')
        for nm in _pcols:
            v, kind, fmt = vals[nm]
            if isinstance(v, str) and v.startswith('='):
                v = F(v, r)
            x = cell(ws, r, CI(PC[nm]), v, kind, fmt, al='left' if nm in ('nota', 'estado', 'eje', 'cort', 'cambio', 'm418') else 'center')
            if nm in ('nota', 'estado', 'cambio', 'm418'):
                x.alignment = Alignment(wrap_text=True, vertical='center')
        ws.row_dimensions[r].height = 30
    ws.freeze_panes = 'C9'
    add_cf(ws, 'B9:%s%d' % (PC['estado'], P1))
    for c_ in ('cambio', 'm418', 'cort'):
        add_cf(ws, '%s9:%s%d' % (PC[c_], PC[c_], P1))
    ws.sheet_view.zoomScale = 80
    return ws


def sheet_trib(wb):
    ws = wb.create_sheet('TRIB'); ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 1.78
    ws.merge_cells('B1:M1'); c = ws['B1']; c.value = 'TRIB — ÁREAS TRIBUTARIAS DEL DXF POR TRAMO DE NERVADURA PRINCIPAL O DE MURO'
    c.font = font(13, True, color='FFFFFF'); c.fill = fill(TITLE_BG); ws.row_dimensions[1].height = 26
    ws.merge_cells('B2:M2'); ws['B2'].value = ('Cada polígono de la capa TRIB_AREA de ' + DXFT + ' se asignó al tramo sobre el que descansa su base (repartido por proyección cuando cae sobre dos tramos). '
                                              'Las áreas son datos; la carga se recalcula con el w del tablero. Las bajadas por áreas tributarias son admisibles (NTC-Mampostería 3.1.2.2.1, p. 31).')
    ws['B2'].font = font(9, i=True, color=GRAY); ws['B2'].alignment = Alignment(wrap_text=True); ws.row_dimensions[2].height = 30
    for k, (t, w) in enumerate([('Tramo', 9), ('Tablero', 15), ('Nivel', 8), ('Área (m²)', 9), ('w_u del tablero (kg/m²)', 10), ('Carga última (t)', 9), ('w serv. (kg/m²)', 9),
                                ('Carga servicio (t)', 9), ('Tipo', 9), ('Estado del tramo', 34), ('Carga viva W_m (t)', 9), ('Carga viva media W (t)', 9)]):
        hdr_cell(ws, 5, 2 + k, t, w)
    ws.row_dimensions[5].height = 30
    for i, tr in enumerate(TRIBROWS):
        r = R0 + i; k = '%s %s' % (tr['tablero'], tr['level']); tr_ = TROW[k]
        tipo = 'NP' if tr['tramo'].startswith('NP') else ('DM' if tr['tramo'].startswith('DM') else 'SIN DALA')
        cell(ws, r, 2, tr['tramo']); cell(ws, r, 3, k); cell(ws, r, 4, tr['level']); cell(ws, r, 5, tr['area'], 'data', '0.000')
        cell(ws, r, 6, '=TABLEROS!$%s$%d' % (TC['wu'], tr_), fmt='0'); cell(ws, r, 7, '=E%d*F%d/1000' % (r, r), fmt='0.000')
        cell(ws, r, 8, '=TABLEROS!$%s$%d' % (TC['ws'], tr_), fmt='0'); cell(ws, r, 9, '=E%d*H%d/1000' % (r, r), fmt='0.000')
        cell(ws, r, 10, tipo)
        cell(ws, r, 11, ('=PRINCIPALES!$%s$%d' % (PC['estado'], PROW[tr['tramo']])) if tipo == 'NP' else ('sobre muro' if tipo == 'DM' else 'borde sin dala en el E2K'), al='left')
        cell(ws, r, 12, '=E%d*TABLEROS!$%s$%d/1000' % (r, TC['cv'], tr_), fmt='0.000')
        cell(ws, r, 13, '=E%d*IF(TABLEROS!$%s$%d="HABITACIÓN",%s,%s)/1000' % (r, TC['dest'], tr_, N['cvw_hab'], N['cvw_az']), fmt='0.000')
    ws.freeze_panes = 'B6'
    add_cf(ws, 'K6:K%d' % R1)
    return ws


def sheet_transfer(wb):
    ws = wb.create_sheet('TRANSFER'); ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 1.78
    ws.merge_cells('B1:Q1'); c = ws['B1']; c.value = 'TRANSFER — NERVADURAS PRINCIPALES QUE SE APOYAN EN OTRAS (DIRECTO O A TRAVÉS DEL MURO DE ARRIBA)'
    c.font = font(13, True, color='FFFFFF'); c.fill = fill(TITLE_BG); ws.row_dimensions[1].height = 26
    ws.merge_cells('B2:Q2'); ws['B2'].value = ('La reacción del extremo de un tramo entra como carga puntual en el tramo que lo recibe, a la distancia «pos» de su extremo i. «Por el muro»: la principal de azotea descarga en un castillo o muro del primer nivel que está sobre una principal del N+3.00. '
                                              'Las referencias son directas (sin búsquedas) para que Excel no vea referencias circulares.')
    ws['B2'].font = font(9, i=True, color=GRAY); ws['B2'].alignment = Alignment(wrap_text=True); ws.row_dimensions[2].height = 30
    for k, (t, w) in enumerate([('Tramo que entrega', 10), ('Extremo', 8), ('Tramo que recibe', 10), ('pos desde i (m)', 9), ('P_u (t)', 8), ('L del que recibe (m)', 9),
                                ('Momento que aporta (t·m)', 10), ('Cortante que aporta (t)', 10), ('A R_i (t)', 8), ('A R_j (t)', 8), ('Condición del que recibe', 14),
                                ('P serv. (t)', 8), ('M serv. que aporta (t·m)', 10), ('A R_i serv. (t)', 8), ('A R_j serv. (t)', 8), ('Vía', 9)]):
        hdr_cell(ws, 5, 2 + k, t, w)
    ws.row_dimensions[5].height = 42
    for i, x in enumerate(D['transfer']):
        r = XROW[i]; rd = PROW[x['de']]; ra = PROW[x['a']]
        cell(ws, r, 2, x['de']); cell(ws, r, 3, 'i' if x['extremo'] == 0 else 'j'); cell(ws, r, 4, x['a']); cell(ws, r, 5, x['pos'], 'data', '0.000')
        cell(ws, r, 6, '=PRINCIPALES!$%s$%d' % (PC['Ri'] if x['extremo'] == 0 else PC['Rj'], rd), fmt='0.000')
        cell(ws, r, 7, '=PRINCIPALES!$%s$%d' % (PC['L'], ra), fmt='0.000')
        cond = 'PRINCIPALES!$%s$%d' % (PC['cond'], ra); raiz = 'PRINCIPALES!$%s$%d' % (PC['raiz'], ra)
        R = lambda f: f.replace('{r}', str(r))
        cell(ws, r, 8, R('=IF(%s="APOYADA",F{r}*E{r}*(G{r}-E{r})/G{r},IF(%s="VOLADO",F{r}*IF(%s="i",E{r},G{r}-E{r}),0))') % (cond, cond, raiz), fmt='0.000')
        cell(ws, r, 9, R('=IF(%s="APOYADA",F{r}*MAX(E{r},G{r}-E{r})/G{r},IF(%s="VOLADO",F{r},0))') % (cond, cond), fmt='0.000')
        cell(ws, r, 10, R('=IF(%s="APOYADA",F{r}*(G{r}-E{r})/G{r},0)') % cond, fmt='0.000')
        cell(ws, r, 11, R('=IF(%s="APOYADA",F{r}*E{r}/G{r},0)') % cond, fmt='0.000')
        cell(ws, r, 12, '=' + cond)
        cell(ws, r, 13, '=PRINCIPALES!$%s$%d' % (PC['Ris'] if x['extremo'] == 0 else PC['Rjs'], rd), fmt='0.000')
        cell(ws, r, 14, R('=IF(%s="APOYADA",M{r}*E{r}*(G{r}-E{r})/G{r},IF(%s="VOLADO",M{r}*IF(%s="i",E{r},G{r}-E{r}),0))') % (cond, cond, raiz), fmt='0.000')
        cell(ws, r, 15, R('=IF(%s="APOYADA",M{r}*(G{r}-E{r})/G{r},0)') % cond, fmt='0.000')
        cell(ws, r, 16, R('=IF(%s="APOYADA",M{r}*E{r}/G{r},0)') % cond, fmt='0.000')
        cell(ws, r, 17, 'por el muro' if x['via'] == 'muro' else 'directo')
    return ws


def sheet_muros(wb):
    ws = wb.create_sheet('MUROS'); S = Narr(ws, N, {'A': 1.78, 'B': 22, 'C': 9, 'D': 22, 'E': 9, 'F': 10, 'G': 11, 'H': 11, 'I': 40}, last='I')
    S.title('MUROS — LO QUE LA LOSA LE ENTREGA A LOS MUROS Y A LOS CASTILLOS')
    S.subtitle('Para la revisión de muros con la NTC-Mampostería 2023: carga lineal de la losa sobre cada tramo de dala que sí tiene muro abajo, y reacciones concentradas de las nervaduras principales en sus apoyos.', 30)
    S.gap()
    S.section('A. CARGA LINEAL DE LOSA SOBRE MUROS (por tramo de dala sobre muro)')
    S.secnote('Sólo la losa (áreas tributarias del DXF). No incluye el peso del propio muro, ni el muro de arriba, ni las cargas lineales del E2K: eso se suma en la revisión del muro.')
    r = S.r
    for k, t in enumerate(['Tramo · eje', 'Nivel', 'De – a', 'L (m)', 'Área trib. (m²)', 'w servicio (t/m)', 'w última (t/m)', 'Nota']):
        hdr_cell(ws, r, 2 + k, t)
    S.r += 1
    for t in DMS:
        r = S.r
        cell(ws, r, 2, '%s · %s' % (t['np'], t['eje']), al='left'); cell(ws, r, 3, t['level'])
        cell(ws, r, 4, '(%.2f, %.2f) – (%.2f, %.2f)' % (t['x0'], t['y0'], t['x1'], t['y1']), al='left', sz=8)
        cell(ws, r, 5, t['L'], fmt='0.000')
        cell(ws, r, 6, '=SUMIFS(TRIB!$E$%d:$E$%d,TRIB!$B$%d:$B$%d,"%s")' % (R0, R1, R0, R1, t['np']), fmt='0.00')
        cell(ws, r, 7, '=SUMIFS(TRIB!$I$%d:$I$%d,TRIB!$B$%d:$B$%d,"%s")/E%d' % (R0, R1, R0, R1, t['np'], r), fmt='0.000')
        cell(ws, r, 8, '=SUMIFS(TRIB!$G$%d:$G$%d,TRIB!$B$%d:$B$%d,"%s")/E%d' % (R0, R1, R0, R1, t['np'], r), 'key', '0.000')
        cell(ws, r, 9, 'Muro del entrepiso de abajo (planta baja)' if t['level'] == 'N+3.00' else 'Muro del primer nivel', al='left', sz=8)
        S.r += 1
    S.gap()
    S.section('B. REACCIONES DE LAS NERVADURAS PRINCIPALES EN SUS APOYOS FIRMES')
    S.secnote('Carga concentrada que baja por el castillo o el muro del extremo. Las que caen en un castillo propuesto (C1…C8) son las que ese castillo tiene que llevar a la cimentación.')
    r = S.r
    for k, t in enumerate(['Tramo · extremo', 'Nivel', 'Apoyo (x, y)', 'Castillo', 'Tipo', 'R última (t)', 'R servicio aprox. (t)', 'Nota']):
        hdr_cell(ws, r, 2 + k, t)
    S.r += 1
    for t in NPS:
        for k in (0, 1):
            e = t['ext'][k]
            if e['tipo'] != 'FIRME':
                continue
            r = S.r; pr = PROW[t['np']]
            cell(ws, r, 2, '%s · %s' % (t['np'], 'i' if k == 0 else 'j'), al='left'); cell(ws, r, 3, t['level'])
            cell(ws, r, 4, '(%.3f, %.3f)' % ((t['x0'], t['y0']) if k == 0 else (t['x1'], t['y1'])), sz=8)
            cell(ws, r, 5, e.get('prop') or '—')
            cell(ws, r, 6, '=PRINCIPALES!$%s$%d' % (PC['ei'] if k == 0 else PC['ej'], pr))
            cell(ws, r, 7, '=PRINCIPALES!$%s$%d' % (PC['Ri'] if k == 0 else PC['Rj'], pr), 'key', '0.000')
            cell(ws, r, 8, '=PRINCIPALES!$%s$%d' % (PC['Ris'] if k == 0 else PC['Rjs'], pr), fmt='0.000')
            cell(ws, r, 9, 'Servicio = cargas sin factorizar (D + W_m).', al='left', sz=8)
            S.r += 1
    S.gap()
    S.notice('Para el muro: P_u total = (w_losa + peso propio del muro + muro de arriba)·L + reacciones concentradas, y se compara con la resistencia a compresión del muro confinado de la NTC-Mampostería 2023. '
             'Eso no está en este libro.', 30)
    add_cf(ws, 'B1:I%d' % S.r)
    return ws
