"""Dibuja la hoja CROQUIS a PNG para revisarla a ojo (LibreOffice + pdftoppm).

Uso:
    python3 render_croquis.py SALIDA_DIR [MURO ...]

Por cada muro arma una copia del libro con ese muro escogido en DETALLE, deja visible sólo CROQUIS, la convierte a PDF
con `soffice --headless --convert-to pdf` y pasa sus páginas a PNG con `pdftoppm`. No toca el libro entregado.
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
EJEMPLOS = ['CAD_PA_21', 'CAD_PB_13', 'CAD_PA_35', 'CAD_PB_01']


def render(sel, outdir, tag, dpi=110):
    import gen_muros
    tmp = tempfile.mkdtemp()
    full = os.path.join(tmp, 'croquis_%s.xlsx' % tag)
    gen_muros.build(full, sel=sel, only='CROQUIS')
    dst = full
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, dst], check=True, capture_output=True, timeout=600)
    pdf = dst[:-5] + '.pdf'
    os.makedirs(outdir, exist_ok=True)
    n = int([ln for ln in subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout.splitlines() if ln.startswith('Pages:')][0].split()[1])
    out = []
    for p in range(1, n + 1):
        stem = os.path.join(outdir, '%s_p%d' % (tag, p))
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), '-f', str(p), '-l', str(p), '-singlefile', pdf, stem], check=True)
        out.append(stem + '.png')
    return out


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'out')
    sels = sys.argv[2:] or EJEMPLOS
    for s in sels:
        print(s, '->', render(s, out, s))
