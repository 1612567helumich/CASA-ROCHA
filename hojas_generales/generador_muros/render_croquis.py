"""Dibuja la hoja CROQUIS a PNG para revisarla a ojo (LibreOffice + pdftoppm).

Uso:
    python3 render_croquis.py SALIDA_DIR [MURO ...]                 # hoja general (M-1, M-2, M-3)
    python3 render_croquis.py SALIDA_DIR --casa_rocha [MURO ...]    # hoja de Casa Rocha

Por cada muro arma una copia del libro con ese muro escogido en DETALLE, deja visible sólo CROQUIS, la convierte a PDF
con `soffice --headless --convert-to pdf` y pasa sus páginas a PNG con `pdftoppm`. No toca el libro entregado.
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
EJEMPLOS = {'general': ['M-1', 'M-2', 'M-3'], 'casa_rocha': ['CAD_PA_21', 'CAD_PB_13', 'CAD_PA_35', 'CAD_PB_01']}


def render(sel, outdir, tag, dpi=110, caso='general'):
    import gen_muros
    tmp = tempfile.mkdtemp()
    full = os.path.join(tmp, 'croquis_%s.xlsx' % tag)
    gen_muros.build(full, sel=sel, only='CROQUIS', caso=caso)
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
    args = [a for a in sys.argv[1:] if a != '--casa_rocha']
    caso = 'casa_rocha' if '--casa_rocha' in sys.argv else 'general'
    out = args[0] if args else os.path.join(HERE, 'out')
    for s in args[1:] or EJEMPLOS[caso]:
        print(s, '->', render(s, out, s, caso=caso))
