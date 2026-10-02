"""PDF breve de lo que se hizo (plano R1.6 + modelo ETABS R1.0)."""
import os
import sys

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

sys.path.insert(0, os.path.dirname(__file__))
import build_e2k as B  # noqa: E402
import escalera as E  # noqa: E402
import graficas_modelo as GM  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, '001_CASA_ROCHA_RESUMEN_R1.0.pdf')
IMG = sys.argv[1] if len(sys.argv) > 1 else '.'
# uso: python3 scripts/build_pdf.py <carpeta_imagenes> <machote.e2k>

ss = getSampleStyleSheet()
H1 = ParagraphStyle('h1', parent=ss['Heading1'], fontSize=15, spaceAfter=4, textColor=colors.HexColor('#1f3b57'))
H2 = ParagraphStyle('h2', parent=ss['Heading2'], fontSize=11.5, spaceBefore=8, spaceAfter=3,
                    textColor=colors.HexColor('#1f3b57'))
P = ParagraphStyle('p', parent=ss['BodyText'], fontSize=8.8, leading=11.2, alignment=TA_LEFT)
PS = ParagraphStyle('ps', parent=P, fontSize=7.6, leading=9.4, textColor=colors.HexColor('#444444'))
BUL = ParagraphStyle('b', parent=P, leftIndent=10, bulletIndent=2)


def bl(items):
    return [Paragraph(t, BUL, bulletText='•') for t in items]


def tabla(data, widths, head=True, fs=7.8):
    t = Table(data, colWidths=widths, repeatRows=1 if head else 0)
    st = [('FONTSIZE', (0, 0), (-1, -1), fs), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
          ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#999999')),
          ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
          ('TOPPADDING', (0, 0), (-1, -1), 1.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5)]
    if head:
        st += [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dfe8f1')),
               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')]
    t.setStyle(TableStyle(st))
    return t


def cel(t):
    return Paragraph(t, ParagraphStyle('c', parent=P, fontSize=7.6, leading=9.2))


def img(name, w):
    from PIL import Image as PI
    path = os.path.join(IMG, name)
    with PI.open(path) as im:
        a = im.size[1] / im.size[0]
    return Image(path, width=w, height=w * a)


def main():
    import geometria_modelo as G
    global NPB, NPA, GR
    GR = B.main(sys.argv[2])[0]['grupos']
    pb, pa = G.leer()
    NPB, NPA = len(B.partir(pb)), len(B.partir(pa))
    R = E.resumen()
    W = GM.pesos()
    doc = SimpleDocTemplate(OUT, pagesize=letter, leftMargin=1.6 * cm, rightMargin=1.6 * cm,
                            topMargin=1.4 * cm, bottomMargin=1.4 * cm,
                            title='Casa Rocha - Resumen R1.6 / ETABS R1.0', author='Claude Code')
    s = []
    s.append(Paragraph('Casa Rocha — qué se hizo (plano R1.6 y modelo ETABS R1.0)', H1))
    s.append(Paragraph('Morelia, Michoacán · 2 de octubre de 2026 · Base: 001_PLANO_CROCHA_R1.5.dxf y machote '
                       '002_HUAYACAN_R2.2.e2k', PS))
    s.append(Spacer(1, 4))
    s.append(tabla([
        ['Archivo', 'Qué es'],
        ['001_PLANO_CROCHA_R1.6.dxf', cel('Plano corregido (plantas, cortes, escalera y 3 layouts).')],
        ['001_CASA_ROCHA_R1.0.e2k', cel('Modelo ETABS. Abrir con <b>File › Import › ETABS .e2k</b> y guardar como '
                                        '.EDB. Revisar la ventana de advertencias al importar.')],
        ['scripts/', cel('Generadores en Python (rehacen el DXF y el .e2k si cambia algún dato).')],
    ], [5.2 * cm, 12.6 * cm]))

    # ------------------------------------------------------------- plano
    s.append(Paragraph('1. Plano DXF (R1.6)', H2))
    s += bl([
        '<b>Comedor para 15:</b> mesa de 1.10 × 4.20 m con 7 + 7 sillas y 1 en cabecera (0.60 m por persona). '
        'Quedan 1.3 m libres al muro y a los bancos de la isla. La cota de 4.78 se bajó para que no cruce la mesa.',
        '<b>Escalera sin el primer escalón:</b> el arranque se recorre 0.29 m y se mantienen 17 peraltes. '
        'Se redibujó en planta baja, planta alta, cortes A-A\' y B-B\' (además ya asienta en N.P.T. ±0.00; '
        'antes estaba 4 cm arriba) y en la lámina A-03, con isométrico nuevo y tablas de geometría y trazo.',
        '<b>Recámara principal:</b> se agregó la línea de centro del muro baño/recámara (y la del muro del eje D, '
        'que tampoco la tenía). El vano del baño medía 0.85 con hoja de 0.70: el muro se prolongó 10 cm '
        '(vano 0.75 = hoja 0.70 + 0.05, como dice la nota 5). El mueble de TV se recorrió para no tapar la puerta.',
        '<b>Ejes a centro de muro:</b> el eje H estaba 17.5 cm fuera del muro (39.36 → 39.185) y el eje 10 '
        'estaba 10 cm fuera (13.821 → 13.721; los cortes ya lo traían bien). Se revisaron todos los demás: '
        'coinciden a menos de 3.4 mm. Al mover H, las cotas G-H / H-J quedan en 1.33 / 1.60. En PA hay tres '
        'divisorios de baño a 10 cm de los ejes B, D y G; no se movieron para no cambiar las medidas interiores.',
        '<b>Cotas:</b> las de 1.50 y 1.43 que estaban volteadas se rehicieron derechas. En las cadenas con '
        'claros de 0.50 m, el texto va abajo de la línea para que no se encime.',
        '<b>Bloques:</b> autos y pick-up modernos en planta y de perfil en el corte A-A\' (con líneas ocultas). '
        'La sala se cambió por un seccional en L, dos sillones lounge, mesas de centro anidadas, tapete y '
        'lámpara de arco. Los cortes A-A\' y B-B\' también muestran la sala y el comedor nuevos.',
        '<b>Layouts:</b> plantas a 1:75 (antes 1:100), cortes a 1:40 (antes 1:50) y escalera a 1:20 (antes 1:25). '
        'Las notas se reacomodaron sin encimarse. Se corrigieron los rótulos y cotas que se encimaban (máquinas, '
        'recámaras, área de lavado, baño de PA y cotas de 0.50) y un rótulo RECÁMARA repetido. En el membrete, '
        'el renglón ARCHIVO se salía del marco. Se agregó un cuadro de revisiones.',
        '<b>Notas:</b> la nota 3 ahora dice azotea reticular, para que coincida con el modelo. La nota 6 se '
        'actualizó con la escalera nueva y se agregó la referencia del Reglamento de Morelia.',
    ])
    s.append(Spacer(1, 3))
    s.append(Paragraph('<b>Escalera recalculada</b> (no cumplía con sólo quitar el escalón):', P))
    filas = [
        ['Opción', 'Peralte', 'Huella en línea de huella', 'Huella a 0.40 del barandal int.', '2P + H', '¿Cumple?'],
        ['16 peraltes (sólo quitar el escalón)', '0.1875', '0.29', '0.26', '0.665', 'No: P > 0.18 y 2P+H > 0.65'],
        ['17 peraltes, 11 huellas radiales en la curva', '0.1765', '0.264', '0.236', '0.617', 'No: H40 < 0.25'],
        ['<b>17 peraltes, 16 huellas compensadas (elegida)</b>', '<b>%.4f</b>' % R['peralte'],
         '<b>%.3f – %.3f</b>' % (R['hw_min'], R['hw_max']), '<b>%.3f – %.3f</b>' % (R['h40_min'], R['h40_max']),
         '<b>%.3f – %.3f</b>' % (R['2p_h_min'], R['2p_h_max']), '<b>Sí (criterio del plano)</b>'],
    ]
    s.append(tabla([[cel(c) for c in r] for r in filas],
                   [5.0 * cm, 1.5 * cm, 2.6 * cm, 3.0 * cm, 2.3 * cm, 3.4 * cm], fs=7.2))
    s.append(Paragraph('Mismo hueco y mismos radios (int. 0.425 / ext. 1.425). La huella interior mínima es de %.3f m. '
                       'Queda pendiente con el D.R.O.: el Reglamento de Morelia pide huella de 0.28 y '
                       'máximo 15 peraltes entre descansos, y la escalera original ya tenía 17.' % R['hi_min'], PS))
    s.append(Spacer(1, 4))
    s.append(img('F_esc.png', 17.6 * cm))
    s.append(Paragraph('Lámina A-03 (1:20): planta numerada, alzado, isométrico, tabla por peldaño y '
                       'trazo de narices.', PS))

    # ------------------------------------------------------------- modelo
    s.append(PageBreak())
    s.append(Paragraph('2. Modelo ETABS (001_CASA_ROCHA_R1.0)', H2))
    s += bl([
        '<b>Machote:</b> se conservaron materiales, secciones, patrones de carga, casos, combinaciones, fuente de '
        'masa y preferencias. Se borró todo lo de Huayacán (geometría, asignaciones, cargas, grupos y '
        'sobreescrituras). Se cambió el título y el nombre del proyecto.',
        '<b>Retícula y niveles:</b> ejes A–J y 1–15, con origen en A-1 y H ya corregido. Niveles N 0.00, '
        'N+3.00 (entrepiso) y N+6.00 (azotea).',
        '<b>Materiales nuevos (NTC Mampostería 2023):</b> tabique rojo recocido f\'m = 15 kg/cm², '
        'Em = 600 f\'m = 9,000 kg/cm², ν = 0.25 (Gm = 0.4 Em), γ = 1,500 kg/m³ (v\'m = 3.5 kg/cm² para '
        'revisión). Block de concreto tipo A f\'m = 20 kg/cm², Em = 800 f\'m = 16,000 kg/cm², '
        'γ = 1,650 kg/m³. Se trabajó todo con tabique rojo (muro de 14 cm).',
        '<b>Muros:</b> %d tramos (%d en PB y %d en PA), dibujados a ejes de muro desde el DXF y cortados en '
        'cada encuentro. Puertas y ventanales quedan como huecos y su peso entra como carga lineal '
        '"CANCELERIA" (%d kg/m). No se modelaron remates de menos de 0.50 m. Cada muro lleva etiqueta '
        'de pier (PB-X…, PA-Y…) para sacar fuerzas por muro. Apoyos empotrados en la base de los muros de PB.'
        % (NPB + NPA, NPB, NPA, B.CANCEL),
        '<b>Losa reticular</b> (casetón 60×60×20 + capa de 5, nervios de 10, h = 0.25): se modeló como losa '
        'maciza de rigidez equivalente, h = %.1f cm, con un material de peso ajustado (γ = %.0f kg/m³) para '
        'que pese lo real: %.0f kg/m². Lleva diafragma rígido (D1 y Azotea). Huecos: vacío, escalera, '
        'pergolado y domo.' % (B.H_EQ * 100, B.GAMMA_LOSA, B.GAMMA_LOSA * B.H_EQ),
    ])
    s.append(Spacer(1, 2))
    s.append(tabla([
        ['Carga (NTC Criterios y Acciones 2023)', 'Entrepiso N+3.00', 'Azotea N+6.00'],
        ['Sobrecarga muerta (Sobrecarga)', '%d kg/m²' % B.SCM_ENT, '%d kg/m²' % B.SCM_AZO],
        ['Carga viva máx. / instantánea / media', '%d / %d / %d kg/m²' % B.CV_ENT, '%d / %d / %d kg/m²' % B.CV_AZO],
        ['Lineales', cel('Pretil de terraza %d kg/m, barandal %d kg/m, cancelería %d kg/m, mitad de la '
                         'escalera en el descanso (%d kg/m CM)' % (B.PRETIL_TERR, B.BARANDAL, B.CANCEL, B.ESC_CM)),
         cel('Pretil %d kg/m, pergolado %d kg/m por apoyo (patrón nuevo "Pergola"), domo %d kg/m, tinacos '
             '%d kg en 2 apoyos' % (B.PRETIL_AZO, B.PERGOLA, B.DOMO, B.TINACOS_KG))],
    ], [6.4 * cm, 5.7 * cm, 5.7 * cm]))
    s.append(Spacer(1, 4))
    s.append(Paragraph('<b>Sismo (Morelia).</b> Se usó la forma del espectro del MDOC CFE 2015 (la misma que '
                       'aplica PRODISIS). Aquí no hay acceso a PRODISIS, así que los parámetros de sitio son '
                       'supuestos conservadores:', P))
    sp = B.SIS
    s.append(tabla([
        ['a0r (roca)', 'Terreno', 'Fsit / Fres', 'a0 / c', 'Ta / Tb / Tc', 'r / k', 'Q / R', 'Coef. estático'],
        ['%.2f g' % sp['a0r'], 'Tipo II (sin estudio), Ts ≈ %.2f s' % sp['Ts'], '%.2f / %.2f' % (sp['Fsit'], sp['Fres']),
         '%.3f / %.3f g' % (sp['a0'], sp['c']), '%.1f / %.1f / %.1f s' % (sp['Ta'], sp['Tb'], sp['Tc']),
         '%.1f / %.1f' % (sp['r'], sp['k']), '%.0f / %.0f' % (sp['Q'], sp['R0']), '%.3f' % B.sa_diseno(0.12)],
    ], [1.7 * cm, 4.1 * cm, 1.9 * cm, 2.3 * cm, 2.7 * cm, 1.3 * cm, 1.3 * cm, 2.5 * cm], fs=7.2))
    s += bl([
        'Función <b>"Espectro Base de Diseño"</b> = espectro elástico (Sa/g). La usan SXLD/SYLD con el SF del '
        'machote (9.81/6).',
        'Función nueva <b>"Espectro Diseño SV Morelia (Q=2)"</b> = Sa/(Q\'R) en función del periodo. La usan '
        'SXSV/SYSV con SF = 9.81 y los casos de cimentación con SF = 9.81/0.65. El patrón estático "Sismo" '
        'tiene coeficiente %.3f. Los modos subieron de 12 a 30.' % B.sa_diseno(0.12),
        'Para usar PRODISIS: correr con las coordenadas del predio, exportar el espectro transparente y '
        'sustituir los valores de "Espectro Base de Diseño" (o leerlo desde archivo, como hacía el machote).',
    ])
    s.append(KeepTogether([img('modelo_espectro.png', 11.5 * cm)]))

    s.append(PageBreak())
    s.append(Paragraph('Muros y losas del modelo sobre el plano', H2))
    s.append(img('modelo.png', 17.6 * cm))
    s.append(Paragraph('Rojo: muros que siguen hasta la cimentación. Naranja: muros de PA que no tienen muro abajo '
                       '(apoyados en la losa). Azul: losa del nivel superior.', PS))
    s.append(Spacer(1, 4))
    s.append(img('modelo_3d.png', 11.0 * cm))

    s.append(Paragraph('3. Revisar antes de diseñar', H2))
    s += bl([
        '<b>a0r y tipo de terreno:</b> confirmar con PRODISIS y, de preferencia, con un estudio de mecánica de suelos.',
        '<b>%d muros de PA están sobre losa</b> (grupo "MUROS PA SOBRE LOSA (SIN MURO ABAJO)"). En el modelo '
        'descansan directo en la losa reticular, así que en obra necesitan trabe o nervadura reforzada. '
        'Hay además %d tramos sin losa arriba (bardas de la terraza 13-15).'
        % (GR.get('MUROS PA SOBRE LOSA (SIN MURO ABAJO)', 0), GR.get('MUROS PA SIN LOSA ARRIBA (BARDAS)', 0)),
        '<b>Cochera:</b> la losa N+3.00 sobre la cochera (ejes 2-6, A-G) sólo apoya en el muro A y en el baño de PB, '
        'y el muro oriente del baño de PA (eje G) cae en el borde libre de esa losa. El plano no trae trabes ni '
        'columnas al frente (eje 2) ni en el eje G, y el modelo no las inventa; hay que definirlas.',
        '<b>Claros grandes:</b> en sala y comedor la losa salva hasta 8.85 m entre los muros A y J. Revisar si '
        'h = 25 alcanza o si hacen falta trabes.',
        '<b>PB casi sin muros en X:</b> %.1f m de muro en X contra %.1f m en Y (garage, sala y comedor abiertos). '
        'Revisar torsión y cortante en esa dirección.' % (W['Lx_PB'], W['Ly_PB']),
        '<b>Q = 2, R = 2</b> (mampostería confinada de piezas macizas). El factor de limitación de daños '
        '(elástico/6) es el que traía el machote.',
        'Pesos aproximados: muros %.0f t, losas %.0f t, sobrecarga %.0f t, CV instantánea %.0f t.'
        % (W['peso_muros_t'], W['peso_losas_t'], W['scm_t'], W['cv_inst_t']),
        'El .e2k se revisó con un validador propio: referencias entre puntos, áreas, secciones, materiales, '
        'patrones, casos, combinaciones y grupos, sin errores. <b>No se pudo abrir en ETABS desde aquí</b>; si al '
        'importar aparece alguna advertencia, hay que revisarla.',
    ])
    doc.build(s)
    print('OK ->', OUT)


if __name__ == '__main__':
    main()
