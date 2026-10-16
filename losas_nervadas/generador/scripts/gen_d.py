"""Parte D: RESUMEN, REFERENCIAS y armado del libro."""
import sys
from gen_c import *
from openpyxl import Workbook


def sheet_resumen(wb):
    ws = wb.create_sheet('RESUMEN')
    widths = {'A': 1.78, 'B': 16, 'C': 12, 'D': 10, 'E': 9, 'F': 7, 'G': 7, 'H': 8, 'I': 8, 'J': 9, 'K': 8, 'L': 9, 'M': 9, 'N': 30}
    S = Narr(ws, N, widths, last='N')
    S.title('RESUMEN DE LAS LOSAS NERVADAS — CASA ROCHA')
    S.subtitle('Losa reticular LR-25 (casetón 60×60×20, nervadura de 10 cm, capa de 5 cm) en entrepiso N+3.00 y azotea N+6.00, apoyada en muros de block confinados y en nervaduras principales.', 30)
    S.gap()
    S.section('1. REVISIONES — TODO EN UN SOLO LUGAR')
    r = S.r
    for k, t in enumerate(['Revisión', '', '', 'Cuántos', 'de', '', '', '', '', '', '', '', 'Estado']):
        hdr_cell(ws, r, 2 + k, t)
    ws.merge_cells('B%d:D%d' % (r, r)); ws.merge_cells('G%d:M%d' % (r, r)); S.r += 1
    est = 'TABLEROS!$%s$%d:$%s$%d' % (TC['estado'], T0, TC['estado'], T1)
    pest = 'PRINCIPALES!$%s$%d:$%s$%d' % (PC['estado'], P0, PC['estado'], P1)
    rows = [('Tableros que CUMPLEN (nervaduras secundarias)', '=COUNTIF(%s,"CUMPLE*")' % est, '=COUNTA(%s)-COUNTIF(%s,"HUECO*")' % (est, est), 'ok'),
            ('Tableros que piden MACIZAR casetones junto al apoyo', '=COUNTIF(%s,"*MACIZAR*")' % est, '=COUNTA(%s)' % est, 'info'),
            ('Tableros que esperan a una nervadura principal', '=COUNTIF(%s,"VER PRINCIPALES*")' % est, '=COUNTA(%s)' % est, 'zero'),
            ('Tableros que NO CUMPLEN', '=COUNTIF(%s,"NO CUMPLE*")' % est, '=COUNTA(%s)' % est, 'zero'),
            ('Tableros donde h_eq < espesor mínimo (calcular flecha)', '=COUNTIF(%s,"REVISAR: calcular flecha*")' % est, '=COUNTA(%s)' % est, 'zero'),
            ('Tableros que no cumplen 6.7.8.1.4 (se diseñan como losa sobre vigas)', '=COUNTIF(TABLEROS!$%s$%d:$%s$%d,"NO")' % (TC['g678'], T0, TC['g678'], T1), '=COUNTA(%s)-COUNTIF(%s,"HUECO*")' % (est, est), 'info'),
            ('Nervaduras principales que CUMPLEN con la sección propuesta', '=COUNTIF(%s,"CUMPLE*")' % pest, '=COUNTA(%s)-COUNTIF(%s,"BORDE*")' % (pest, pest), 'ok'),
            ('… de ellas, las que hay que PERALTAR respecto a la dala 15×25 del modelo', '=COUNTIF(PRINCIPALES!$%s$%d:$%s$%d,"PERALTAR*")' % (PC['cambio'], P0, PC['cambio'], P1), '=COUNTA(%s)-COUNTIF(%s,"BORDE*")' % (pest, pest), 'info'),
            ('Nervaduras principales con muro encima donde hay que revisar la interacción muro-viga (H/L < 0.6)', '=COUNTIF(PRINCIPALES!$%s$%d:$%s$%d,"REVISAR*")' % (PC['m418'], P0, PC['m418'], P1), '=COUNTIF(PRINCIPALES!$%s$%d:$%s$%d,"<>—")' % (PC['m418'], P0, PC['m418'], P1), 'info'),
            ('Nervaduras principales que NO CUMPLEN', '=COUNTIF(%s,"NO CUMPLE*")' % pest, '=COUNTA(%s)-COUNTIF(%s,"BORDE*")' % (pest, pest), 'zero'),
            ('Capa de compresión: temperatura, diafragma, recubrimiento, punzonamiento y flexión con 1 000 kg', '=COUNTIF(NERVADURA!$C$1:$C$400,"NO CUMPLE*")', 5, 'zero'),
            ('Nervaduras principales sin apoyo', '=COUNTIF(%s,"SIN APOYO*")' % pest, '=COUNTA(%s)' % pest, 'zero'),
            ('Castillos propuestos aceptados', '=COUNTIF(APOYOS!$C$1:$C$400,"SÍ")', len(D['castillos_prop']), 'info')]
    for c, f1, f2, kind in rows:
        r = S.r
        ws.merge_cells('B%d:D%d' % (r, r)); cell(ws, r, 2, c, al='left', sz=10)
        cell(ws, r, 5, f1, 'key', '0'); cell(ws, r, 6, f2, 'calc', '0')
        ws.merge_cells('G%d:M%d' % (r, r)); cell(ws, r, 7, '', 'calc')
        if kind == 'ok':
            st = '=IF(E%d=F%d,"CUMPLE — todos","REVISAR — "&(F%d-E%d)&" pendiente(s)")' % (r, r, r, r)
        elif kind == 'zero':
            st = '=IF(E%d=0,"CUMPLE — ninguno","REVISAR — ver lista")' % r
        else:
            st = '="informativo"'
        cell(ws, r, 14, st, 'key', '@', al='left')
        S.r += 1
    S.gap()
    S.section('2. DECISIONES QUE TE TOCAN')
    for txt in ['1) CASTILLOS: la hoja propone bajar a cimentación los castillos C1…C8 (APOYOS sección B). Sin ellos, las nervaduras principales de la cochera, del vacío de doble altura y de la escalera quedan en volado de 4 a 8 m. Si una bajada queda aislada (sin muro en planta baja) es columna: mínimo 25×25 (NTC-Concreto 6.4.2.1.1, p. 138).',
                '2) PRINCIPALES CON MURO ENCIMA: la casa es estructura Tipo I de mampostería (muros desplantados sobre vigas, NTC-M 1.6.1 c, p. 15). Esas vigas deben cumplir L/h ≤ 14, la flecha de 4.1.18.3 y 0.3 cm + L/480: la dala de 15×25 no alcanza en claros de más de 3.50 m. La tabla 4 da la sección propuesta (trabe descolgada); en PRINCIPALES puedes cambiar b y h y ver el resultado al momento.',
                '3) CAPA DE COMPRESIÓN DE 5 cm: con la malla 6x6-6/6 cumple temperatura (0.0018), recubrimiento y flexión con 1 000 kg, pero queda 2% abajo de la cuantía de diafragma (0.0025, 6.8.6.4) y 1% abajo en punzonamiento con la carga directo sobre la capa. Lo resuelve una capa de 6 cm (casetón de 19 cm para seguir en h = 25) con malla 6x6-4/4, o contar el reparto del piso (NERVADURA sección 9).',
                '4) SOBRECARGA: la hoja suma 20 + 20 kg/m² (NTC-Criterios 5.1.2.1). Si la sobrecarga del modelo ya trae la «carga adicional 0.04 t/m²» de Naim, ponlos en cero en DATOS sección 5.',
                '5) MODELO: corregir en ETABS F80032 y F80010 de azotea a LR-25, ponerle cargas a F2 y meter las secciones propuestas de las principales; luego volver a correr y comparar reacciones contra la hoja MUROS.',
                '6) PLANOS: concreto Clase 1A f\'c = 250 con agregado calizo de 1/2" (TMA ≤ 2/3 del recubrimiento), nota de control del ancho de nervadura (3.9.2), malla especificada y apuntalamiento del entrepiso al colar la azotea (NTC-Criterios 6.1.3.1).']:
        S.bullet(txt)
    S.gap()
    S.section('3. TABLA DE NERVADURAS SECUNDARIAS (COMO LA DE NAIM)')
    r = S.r
    heads = ['Tablero', 'Nervadura', 'Lecho', 'Mu (t·m)', 'b (cm)', 'h (cm)', 'Cant. var.', 'Var. No.', 'M_R (t·m)', 'Mu/M_R', 'Vu (t)', 'V_cR (t)', 'Estado del tablero']
    for k, t in enumerate(heads):
        hdr_cell(ws, r, 2 + k, t)
    ws.row_dimensions[r].height = 30; S.r += 1
    for t in TABS:
        tr = TROW[key(t)]
        T_ = lambda c: 'TABLEROS!$%s$%d' % (TC[c], tr)
        for i, (lab, lecho, blk, M, b) in enumerate([('‖ claro corto', 'Inferior', 'Pc', 'MPc', '=' + T_('bfc')), ('‖ claro corto', 'Superior', 'Nc', 'MNc', '={bwc}'),
                                                     ('‖ claro largo', 'Inferior', 'Pl', 'MPl', '=' + T_('bfl')), ('‖ claro largo', 'Superior', 'Nl', 'MNl', '={bwc}')]):
            r = S.r
            cell(ws, r, 2, key(t) if i == 0 else '', al='left', bold=True)
            cell(ws, r, 3, lab); cell(ws, r, 4, lecho)
            cell(ws, r, 5, '=' + T_(M), fmt='0.000')
            cell(ws, r, 6, b.format(**N), fmt='0'); cell(ws, r, 7, '={h}'.format(**N), fmt='0')
            cell(ws, r, 8, '=' + T_('n' + blk), 'key', '0'); cell(ws, r, 9, '=' + T_('v' + blk), 'key', 'General')
            cell(ws, r, 10, '=' + T_('MR' + blk), fmt='0.000'); cell(ws, r, 11, '=' + T_('FU' + blk), fmt='0.00')
            cell(ws, r, 12, ('=' + T_('Vu')) if i == 1 else '', fmt='0.000'); cell(ws, r, 13, ('=' + T_('VcR')) if i == 1 else '', fmt='0.000')
            cell(ws, r, 14, ('=' + T_('estado')) if i == 0 else '', al='left', sz=8)
            S.r += 1
    S.gap()
    S.section('4. TABLA DE NERVADURAS PRINCIPALES')
    r = S.r
    heads = ['Tramo', 'Nivel', 'Condición', 'L (m)', 'b (cm)', 'h (cm)', 'Inferior', '', 'Superior', '', 'Estribos', 'h req. (cm)', 'Estado']
    for k, t in enumerate(heads):
        hdr_cell(ws, r, 2 + k, t)
    ws.merge_cells('H%d:I%d' % (r, r)); ws.merge_cells('J%d:K%d' % (r, r)); ws.row_dimensions[r].height = 30; S.r += 1
    for t in NPS:
        r = S.r; pr = PROW[t['np']]
        P_ = lambda c: 'PRINCIPALES!$%s$%d' % (PC[c], pr)
        cell(ws, r, 2, t['np'] + ' · ' + t['eje'], al='left', sz=8); cell(ws, r, 3, t['level'])
        cell(ws, r, 4, '=' + P_('cond'), sz=8); cell(ws, r, 5, '=' + P_('L'), fmt='0.00')
        cell(ws, r, 6, '=' + P_('bq'), 'key', '0'); cell(ws, r, 7, '=' + P_('hq'), 'key', '0')
        ws.merge_cells('H%d:I%d' % (r, r)); cell(ws, r, 8, '=IF(%s=0,"—",%s&" No. "&TEXT(%s,"0.#"))' % (P_('nP'), P_('nP'), P_('vP')))
        ws.merge_cells('J%d:K%d' % (r, r)); cell(ws, r, 10, '=IF(%s=0,"—",%s&" No. "&TEXT(%s,"0.#"))' % (P_('nN'), P_('nN'), P_('vN')))
        cell(ws, r, 12, '=' + P_('cort'), sz=8); cell(ws, r, 13, '=' + P_('hrec'), 'key', '0')
        cell(ws, r, 14, '=' + P_('estado'), al='left', sz=8)
        S.r += 1
    S.gap()
    S.section('5. CANTIDADES')
    for k, (c, u) in enumerate([('Concreto de losas (con desperdicio)', 'm³'), ('Casetones 60×60×20', 'pzas'), ('Malla electrosoldada', 'm²'),
                                ('Acero de nervaduras secundarias', 'kg'), ('Acero de nervaduras principales', 'kg')]):
        r = S.r
        ws.merge_cells('B%d:D%d' % (r, r)); cell(ws, r, 2, c, al='left', sz=10)
        cell(ws, r, 5, '=' + N['qT%d' % k], 'key', '#,##0'); cell(ws, r, 6, u)
        S.r += 1
    S.gap()
    S.section('6. CONDICIONES DE OBRA')
    for txt in ['• Castillos que se bajan a cimentación: refuerzo anclado en el cimiento con escuadra a 90° hacia adentro, empotramiento ≥ 40 cm en cimiento de piedra; traslapes sólo en el tercio medio de la planta baja con estribos @ 7.5 cm (NTC-M 4.1.16.3 a 4.1.16.5, p. 43; 4.2.7.5.2, p. 53). Dalas con 4 varillas y As ≥ 0.2(f\'c/fy)·b·h: la D-15x20 pide 3.57 cm² (2 No. 4 + 2 No. 3) (NTC-M 6.2.2.2, p. 62).',
                '• Nervaduras principales sin muro abajo: se arman como trabe con recubrimiento de 4 cm, 2 No. 4 corridas como mínimo arriba y abajo y estribos cerrados No. 3 a 135° (gancho 8 cm), más cerrados en 2h junto a cada apoyo (NTC-Concreto 6.3.5.1.1, p. 122; 6.3.7.7, p. 134; 14.7.3.3, p. 375). Las que piden más de 25 cm quedan descolgadas.',
                '• Toda nervadura secundaria lleva 1 No. 4 corrida abajo y 1 No. 4 corrida arriba; la inferior se ancla para 1.25 fy en el paño (6.7.6.1.3, p. 175; 6.7.8.1.7, p. 188). Bastones con el largo de TABLEROS medido desde el paño; en bordes discontinuos rematan con gancho de 90° dentro de la dala (6.7.7.4.1.1, p. 179).',
                '• Donde la hoja pide MACIZAR, se quita la hilera de casetones junto al apoyo en todo el ancho del tablero. En esquinas exteriores sobre muros se refuerza arriba y abajo en un quinto del claro largo (6.7.7.3.1, p. 178).',
                '• En el borde común de dos tableros se arma con el mayor de los dos bastones (envolvente de 3.4.3.1).',
                '• Traslapes clase B: No. 3 ≈ 50 cm, No. 4 ≈ 66 cm, fuera de las secciones de esfuerzo máximo y alternados; malla traslapada un cuadro + 5 cm (14.5.1.3, Tabla 14.5.2.1, 14.5.4.1).']:
        S.bullet(txt)
    add_cf(ws, 'C1:N%d' % S.r)
    ws.sheet_view.zoomScale = 85
    return ws


def sheet_refs(wb):
    ws = wb.create_sheet('REFERENCIAS')
    S = Narr(ws, N, {'A': 1.78, 'B': 58, 'C': 40, 'D': 46, 'E': 10, 'F': 10})
    S.title('REFERENCIAS — QUÉ DOCUMENTO Y QUÉ PARTE CHECAR')
    S.subtitle('Los archivos del proyecto se citan por su tabla o capa; los documentos de Naim por página impresa; las NTC 2023 por numeral y página impresa de cada norma, cotejados contra el texto oficial del compendio.', 30)

    def block(title, items):
        S.section(title)
        S.header(('Tema', 'Documento', 'Tabla / capa / página / sección', '—', '—'))
        for a, b, c in items:
            r = S.r
            cell(ws, r, 2, a, al='left', sz=9, wrap=True); cell(ws, r, 3, b, al='left', sz=9, wrap=True); cell(ws, r, 4, c, 'ref', al='left', wrap=True)
            cell(ws, r, 5, ''); cell(ws, r, 6, '')
            ws.row_dimensions[r].height = max(18, 13 * (1 + max(len(a) // 60, len(c) // 46)))
            S.r += 1
    block('MODELO ETABS — ' + E2K, [
        ('Geometría de tableros, dalas, castillos y muros por nivel', E2K, 'POINT COORDINATES / LINE CONNECTIVITIES / AREA CONNECTIVITIES / AREA ASSIGNS'),
        ('Losa LR-25: espesor 0.25 y modificadores F 0.6688, M 0.29915', E2K, 'SLAB PROPERTIES — SHELLPROP "LR-25"'),
        ("Concreto f'c = 250 kg/cm², E = 14,000√f'c, γ = 2440 kg/m³", E2K, "MATERIAL PROPERTIES — \"Concreto f'c=250kg/cm2. Clase 1A\""),
        ('Peso equivalente de la LR-25: 1503.75 kg/m³', E2K, 'MATERIAL "Concreto f\'c=250kg/cm2. Clase 1A (LR-25 equiv)"'),
        ('Varilla fy = 4200 kg/cm²', E2K, 'MATERIAL "Fy=4200"'),
        ('Sobrecarga y cargas vivas por tablero', E2K, 'SHELL OBJECT LOADS — AREALOAD'),
        ('Barandales, muros bajos, pretiles y reacción de la escalera', E2K, 'FRAME OBJECT LOADS — LINELOAD'),
        ('Combinaciones 1.3 CM + 1.5 CV (Comb2.2) y 1.5/1.7 (Comb2.1, grupo A)', E2K, 'LOAD COMBINATIONS'),
        ('Dalas D-15x20 y vigas VP-15x40 del pergolado', E2K, 'FRAME SECTIONS / LINE ASSIGNS'),
    ])
    block('PLANOS DXF', [
        ('Casetón 60×60×20, nervadura 10 cm, capa 5 cm (h = 0.25)', DXFA, 'NOTAS DE CORTES, nota 2 (capa 010_Notas)'),
        ('Azotea: losa reticular, peralte según cálculo', DXFA, 'NOTAS DE CORTES, nota 3'),
        ('Vacío de doble altura, vacío de escalera, tragaluces', DXFA, 'Capa 004_Vacios, plantas PRIMER NIVEL y AZOTEA'),
        ('Escalera curva (cubo de escalera)', DXFA, 'Capa 005_Escalera'),
        ('Cochera de 8.70 m libres (muros sólo en A y J)', DXFA, 'Capa 001_Muros, PLANTA BAJA'),
        ('Tableros y áreas tributarias a 45°', DXFT, 'Capas TRIB_TABLERO, TRIB_AREA, TRIB_HUECO'),
    ])
    block('DOCUMENTOS DE NAIM (página impresa)', [
        ('Normas usadas: NTC-2023 concreto, mampostería, acero; MOC-CFE sismo 2015', NAIM24, 'pág. 3'),
        ("Materiales: f'c 250, fy 4200, f'm 15 (valor del Cap. 12, sólo para Tipo II; esta casa es Tipo I)", NAIM24, 'pág. 4'),
        ('Cargas muertas de entrepiso (0.33 t/m²) y azotea (0.29 t/m²); carga viva de habitación', NAIM24, 'pág. 5'),
        ('Carga viva de azotea con pendiente ≤ 5%', NAIM24, 'pág. 6'),
        ('Factores 1.3 permanentes / 1.5 variables, estructura grupo B', NAIM24, 'pág. 8 (lista de combinaciones en pág. 9)'),
        ('Tabla de nervaduras: recubrimiento 2 cm, MR y VcR de cada nervadura (con eso se verificaron las fórmulas)', NAIM24, 'pág. 13'),
        ('FR = 0.9 en flexión y As mín = 0.7√f\'c/fy·b·d (fórmula anterior; la 2023 pide máx(0.80√f\'c, 14)/fy)', NAIM21, 'pág. 13'),
        ('Detalle de nervadura principal/secundaria, malla 6x6-10/10, tabla de nervaduras', NAIMPL, 'EST-03, EST-04 y EST-05'),
        ('Mallas con Fy = 5000 kg/cm²; recubrimiento de 2 cm en losas y dalas', 'Naim — EST CASA YÉPEZ', 'Notas 1.2 y 1.2.2'),
    ])
    block(NTCC + ' — NUMERAL Y PÁGINA IMPRESA', [
        ("Concreto Clase 1A, f'c ≥ 250; Ec = 14 000√f'c (agregado calizo); f̄f = 2√f'c", NTCC, '2.2.2 a), p. 30 · Tabla 2.2.1, p. 31 · Tabla 2.2.6.3, p. 33 · 2.2.8.4, p. 34'),
        ('Tamaño máximo del agregado ≤ 2/3 del recubrimiento y de la separación libre', NTCC, '2.2.5.2.1, p. 32'),
        ('Barras Grado 42 y malla NMX-B-290 (fy de diseño ≤ 7 000 en losas; estribos corrugados, fy ≤ 4 200)', NTCC, '2.4.2.1.1, p. 43 · 2.4.2.1.6 y 2.4.2.1.7, p. 44 · Tabla 2.4.2.2.4.a, p. 46'),
        ('Métodos reconocidos y líneas de fluencia; losa como diafragma', NTCC, '3.4.1.1 y 3.4.1.2, p. 64'),
        ('Método de coeficientes: condiciones a) a d) y distribución de momentos', NTCC, '3.4.2.1 y 3.4.3.1, p. 64'),
        ('Coeficientes de momento, franjas centrales, caso I (losa monolítica)', NTCC, 'Tabla 3.4.2.1, p. 65 (≈50% de valores cambiaron respecto a 2017, comentario p. 66)'),
        ('Cargas lineales de muros dentro de un tablero', NTCC, '3.4.4.1 a 3.4.4.3 y Tabla 3.4.4.2, pp. 65–66'),
        ("Bloque equivalente: f''c = 0.85f'c, β1 = 0.85 hasta f'c = 300, ε_cu = 0.003", NTCC, '3.6.1 d) y e), ec. 3.6.1, p. 68'),
        ('Factores de resistencia: flexión 0.65–0.90 según ε_t; cortante 0.75', NTCC, 'Tablas 3.8.2.1 y 3.8.2.2, p. 70 · 3.8.2.4, p. 71'),
        ('Reducción de 20 mm en anchos < 200 mm salvo control anotado en planos', NTCC, '3.9.1 y 3.9.2, p. 74'),
        ('Clase de exposición y recubrimientos (losas y viguetas 20 mm; vigas 40 mm; intemperie 40 mm)', NTCC, 'Tabla 4.3.1, p. 76 · 4.13.1.1 y 4.13.2.2, p. 83 · Tabla 4.13.3.1 y 4.13.3.2, p. 84 · Tabla 4.14.16.a, p. 85'),
        ('Cuantía balanceada; ancho efectivo del patín; MR de sección rectangular y T', NTCC, 'ec. 5.2.1.3.1, p. 88 · Tabla 5.2.1.4.2, p. 89 · ecs. 5.2.2.1.1.1.a y b, p. 89 · 5.2.2.1.3.1, p. 90'),
        ("Cortante: V_cR con A_v,mín (0.5√f'c o 2p^(1/3)√f'c ≤ 1.25√f'c); sin refuerzo en el alma (2λs·p^(1/3)√f'c); límite de la sección; V_sR", NTCC, '5.5.2.1 y 5.5.2.2, p. 94 · 5.5.3.1.1 y 5.5.3.1.2, pp. 94–95 · 5.5.3.2.1, p. 95 · 5.5.3.6.1 y 5.5.3.8.1, p. 97'),
        ('Punzonamiento de la capa con la carga concentrada', NTCC, 'Tabla 5.6.2.1.1, p. 100'),
        ('Vigas (incluye nervaduras): claro, h/b ≤ 6, peralte mínimo L/16–L/8 y h ≥ 25 cm', NTCC, '6.3.1.2 b), 6.3.2.1 y 6.3.2.2, p. 117 · Tabla 6.3.3.1.1, p. 120 · 6.3.3.2.1 y 6.3.3.3.1, p. 121'),
        ("Vigas: As mín = máx(0.80√f'c, 14)/fy·b_w·d, 2 barras corridas, excepción 4/3; As máx 0.90 As,b", NTCC, '6.3.5.1.1 a 6.3.5.1.3 y 6.3.5.2.1, p. 122'),
        ('Vigas: A_v,mín (casos exentos) y su valor', NTCC, '6.3.5.4.1 y Tabla 6.3.5.4.1, p. 123 · Tabla 6.3.5.4.4, p. 124'),
        ('Corte y anclaje del refuerzo de flexión; punto de inflexión', NTCC, '6.3.7.3.4 a 6.3.7.3.6, p. 127 · 6.3.7.3.10, p. 129'),
        ('Separación máxima de estribos; integridad estructural (2h en extremos)', NTCC, 'Tabla 6.3.7.6.2.2, p. 132 · 6.3.7.7, p. 134'),
        ('Losas en dos direcciones: alcance (nervaduras en dos direcciones)', NTCC, '6.7.1.1 e) y comentario de 6.7.1, p. 165'),
        ('Espesor mínimo: Tabla 6.7.3.1.1 (sin vigas) y Tabla 6.7.3.1.2 (α_fm), +10% en borde discontinuo; flechas si no cumple', NTCC, 'Tabla 6.7.3.1.1, p. 167 · Tabla 6.7.3.1.2, 6.7.3.1.2.1 y 6.7.3.2.1, p. 168'),
        ('Losas controladas por tensión', NTCC, '6.7.3.3.1, p. 169'),
        ('Sección crítica de cortante a d del paño', NTCC, '6.7.4.3.2, p. 172'),
        ('As mín de losa 0.0018·A_g; 1 No. 4 abajo y 1 No. 4 arriba corridas en toda nervadura', NTCC, '6.7.6.1.1 y 6.7.6.1.3, p. 175'),
        ('Refuerzo de esquinas exteriores; bordes discontinuos', NTCC, '6.7.7.3.1, p. 178 · 6.7.7.4.1.1, p. 179'),
        ('Losas encasetonadas: b_w ≥ 10 cm, h ≤ 3.5b_w, separación libre ≤ 80 cm y ≤ claro/6; V_cR × 1.1; carga de 1 000 kg en 10×10 cm', NTCC, '6.7.8.1.2 a 6.7.8.1.6, p. 187'),
        ('Losas encasetonadas: barra inferior continua para 1.25 fy; refuerzo de la capa; si no cumple, losa sobre vigas; capa ≥ máx(a_c/13, 50 mm)', NTCC, '6.7.8.1.7 a 6.7.8.1.9 y 6.7.8.3.1, p. 188'),
        ('Diafragma: firme ≥ 50 mm (≥ 60 mm con claros ≥ 6 m); V_R; cuantía mínima 0.0025; separación ≤ 5t y 45 cm', NTCC, '6.8.1.1 a), p. 188 · 6.8.3.1.2, p. 189 · ec. 6.8.5.3.2, p. 192 · 6.8.6.3, 6.8.6.4 y 6.8.7.3, p. 196'),
        ('Columnas de ductilidad baja (si una bajada queda aislada): dimensión ≥ 25 cm', NTCC, 'ec. 6.4.1.1, p. 137 · 6.4.2.1.1, p. 138'),
        ('Flechas: límites 0.5 cm + L/240 y 0.3 cm + L/480; inercia efectiva (Bischoff); diferida λΔ = ξ/(1 + 50p′)', NTCC, '13.4.1.1, p. 336 · Tabla 13.4.3.2 y ec. 13.4.3.3, p. 338 · ec. 13.4.4.1, p. 339'),
        ('Vibración de pisos que no cumplen el espesor mínimo', NTCC, '13.7.1.1, p. 343'),
        ('Contracción y temperatura: cuantía 0.0018 (ya no 660x₁/(fy(x₁+100)))', NTCC, '13.8.2.1.1, p. 345'),
        ('Separación libre mínima: máx(25 mm, d_b, 1.5·TMA)', NTCC, '14.2.1, p. 347'),
        ('Ganchos y dobleces (90°: 12d_b; estribos 135°: 6d_b ≥ 80 mm)', NTCC, 'Tabla 14.3.1, p. 349 · 14.3.4 y Tabla 14.3.2, p. 350'),
        ('Longitud de desarrollo L_d, gancho L_dh y traslapes clase B; traslape de malla', NTCC, '14.4.2.1, p. 352 · Tabla 14.4.2.4, p. 353 · ec. 14.4.3.1, p. 355 · Tabla 14.5.2.1, p. 367 · 14.5.4.1, p. 369'),
        ('Estribo cerrado mínimo No. 3; separación ≤ 16d_b long. y 48d_b del estribo', NTCC, '14.7.3.2 y 14.7.3.3, p. 375'),
    ])
    block(NTCCA + ' — NUMERAL Y PÁGINA IMPRESA', [
        ('Combinación de cargas permanentes + variables con W_m uniforme', NTCCA, '2.3.1 a), p. 5'),
        ('Factores de carga: 1.3 / 1.5 (grupo B), 1.5 / 1.7 (grupo A); 1.1 accidentales; 0.9 favorables; 1.0 servicio', NTCCA, '3.4.1 a) a d), p. 7'),
        ('Límites de desplazamiento: remite a la norma del material', NTCCA, '4.1.1, p. 8'),
        ('Incremento de 20 kg/m² por losa colada en sitio y 20 kg/m² por capa de mortero', NTCCA, '5.1.2.1 y 5.1.2.2, p. 9'),
        ('Uso de W_m, W_a y W; muros divisorios aparte; cargas transitorias de obra', NTCCA, '6.1.1.2, 6.1.1.3, 6.1.2.1 y 6.1.3.1, p. 10'),
        ('Cargas vivas unitarias: habitación 190/100/80; azotea ≤ 5%: 100/70/15; cargas concentradas', NTCCA, 'Tabla 6.1.2.2 incisos a) y h), p. 11 · notas [1], [7] y [8], pp. 11–12'),
        ('Temperatura: se puede omitir en estructuras de menos de 30 m con azotea protegida', NTCCA, '6.2.3.1, p. 13'),
        ('Granizo: W_g = 900·H_s·C_s = 112.5 kg/m² (grupo B); flecha ≤ L/240 con pendiente ≤ 2%', NTCCA, '7.3.1.1 y 7.3.2.1, p. 17 · Tablas 7.3.2.2 y 7.3.2.3, p. 18 · 7.3.4.2, p. 19'),
    ])
    block(NTCM + ' — NUMERAL Y PÁGINA IMPRESA', [
        ('Estructura Tipo I: tiene muros estructurales desplantados sobre vigas o sobre la losa', NTCM, '1.6.1 c), p. 15 · 1.6.2, p. 16'),
        ("f'm de block de concreto f'p = 100 con mortero tipo I = 50 kg/cm²; Em = 800 f'm", NTCM, 'Tabla 2.8.1.4.1, p. 24 · ec. 2.8.5.4.a, p. 27'),
        ('Bajada de cargas por áreas tributarias', NTCM, '3.1.2.2.1, p. 31'),
        ('Factor de resistencia en compresión F_R = 0.6', NTCM, '4.1.6.2, p. 39'),
        ('Anclaje de castillos en la cimentación; muros que parten de otro nivel; barras del piso ancladas en los muros', NTCM, '4.1.16.3 a 4.1.16.5 y comentario, 4.1.17.2, p. 43'),
        ('Muros sobre vigas: esfuerzo en los extremos (f_CM ≤ F_R f\'m, H/L ≥ 0.6), flecha (ec. 4.1.18.3.1.a), L/h ≤ 14 y h ≥ 25 cm', NTCM, '4.1.18.1 a 4.1.18.2, pp. 44–45 · 4.1.18.3 y 4.1.18.4, p. 46'),
        ('Traslapes de castillos en el primer nivel (tercio medio, estribos @ h_c/2)', NTCM, '4.2.7.5.2 y 4.2.7.5.3, p. 53'),
        ('Castillos y dalas: dimensiones, f\'c ≥ 200, 4 barras con As ≥ 0.2(f\'c/fy)·b·h, estribos', NTCM, '6.1.3 y 6.2.1.1 a 6.2.1.3, pp. 60–61 · 6.2.2.2, 6.2.2.3 y 6.2.3.2, p. 62'),
        ('Resistencia a compresión del muro P_R = F_R·F_E·(f\'m·A_T + ΣAs·fy)', NTCM, 'ec. 6.4.1.1, p. 65 · F_E: 3.1.2.3.1, p. 32'),
    ])
    S.notice('ACI 318-19 ya no se usa en esta hoja: todo lo que antes se tomaba de ahí (nervaduras, peraltes de vigas, estribos, separación de barras, bastones) lo resuelve la NTC-Concreto 2023 con los numerales de arriba.', 30)
    return ws


def build(path):
    wb = Workbook(); wb.remove(wb.active)
    sheet_datos(wb)
    sheet_apoyos(wb)
    sheet_cargas(wb)
    sheet_tableros(wb)
    sheet_nervadura(wb)
    sheet_principales(wb)
    sheet_muros(wb)
    wsK = sheet_croquis(wb)
    sheet_cuant(wb)
    sheet_resumen(wb)
    sheet_refs(wb)
    sheet_tablas(wb)
    sheet_trib(wb)
    sheet_transfer(wb)
    build_croquis_datos(wb, wsK)
    from openpyxl.worksheet.properties import PageSetupProperties
    for ws in wb.worksheets:
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    wsK.page_setup.fitToHeight = 1; wsK.page_setup.orientation = 'portrait'
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return path


if __name__ == '__main__':
    print(build(sys.argv[1]))
