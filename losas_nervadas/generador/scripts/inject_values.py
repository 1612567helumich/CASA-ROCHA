"""Copia los valores calculados (de un archivo recalculado por LibreOffice) como valores en caché del archivo original de openpyxl,
sin tocar fórmulas, estilos ni gráficas."""
import sys, zipfile, re, shutil
from lxml import etree
import openpyxl
src, calc, dst = sys.argv[1:4]
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
vals = openpyxl.load_workbook(calc, data_only=True)
zin = zipfile.ZipFile(src)
wb = etree.fromstring(zin.read('xl/workbook.xml'))
rels = etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
rmap = {r.get('Id'): r.get('Target') for r in rels}
sheets = {}
for s in wb.find('m:sheets', NS):
    tgt = rmap[s.get('{%s}id' % NS['r'])]
    sheets['xl/' + tgt.lstrip('/').replace('xl/', '')] = s.get('name')
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
n = 0
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename in sheets:
        ws = vals[sheets[item.filename]]
        root = etree.fromstring(data)
        for c in root.iter('{%s}c' % NS['m']):
            f = c.find('m:f', NS)
            if f is None:
                continue
            v = ws[c.get('r')].value
            old = c.find('m:v', NS)
            if old is not None:
                c.remove(old)
            ve = etree.SubElement(c, '{%s}v' % NS['m'])
            if isinstance(v, bool):
                c.set('t', 'b'); ve.text = '1' if v else '0'
            elif isinstance(v, (int, float)):
                if 't' in c.attrib:
                    del c.attrib['t']
                ve.text = repr(float(v)) if isinstance(v, float) else str(v)
            elif v is None:
                c.set('t', 'str'); ve.text = ''
            else:
                sv = str(v)
                if sv.startswith('#') and sv.endswith(('!', '?', 'A', '0')) and len(sv) < 9:
                    c.set('t', 'e')
                else:
                    c.set('t', 'str')
                ve.text = sv
            n += 1
        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    zout.writestr(item, data)
zout.close()
print('valores inyectados:', n)
