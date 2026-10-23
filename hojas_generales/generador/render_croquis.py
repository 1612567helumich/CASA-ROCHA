"""Dibuja la hoja CROQUIS a PNG para revisarla a ojo (LibreOffice + pdftoppm).

Uso:
    python3 render_croquis.py SALIDA_DIR ["N · EJEMPLO F80012 (volado)" ...]

Por cada tablero arma una copia del libro con ese tablero escogido en DETALLE, la convierte a PDF con
`soffice --headless --convert-to pdf` y pasa a PNG (con `pdftoppm`) sólo las páginas de la hoja CROQUIS.
Sin tableros en la línea de comandos dibuja los siete EJEMPLO. No toca el libro entregado.
"""
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
EJEMPLOS = ['N · EJEMPLO F80014 (interior)', 'N · EJEMPLO F80016 (de borde)', 'N · EJEMPLO F80024 (esquina, azotea)',
            'N · EJEMPLO F80012 (volado)', 'M · EJEMPLO M-1 (interior 3.0 × 4.0)', 'M · EJEMPLO M-2 (esquina 2.0 × 3.4)',
            'M · EJEMPLO M-3 (volado 1.2 m)']


def croquis_pages(pdf):
    """Páginas del PDF que son de la hoja CROQUIS (su encabezado o el de sus grupos)."""
    n = int([ln for ln in subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout.splitlines() if ln.startswith('Pages:')][0].split()[1])
    pages = []
    for i in range(1, n + 1):
        txt = subprocess.run(['pdftotext', '-f', str(i), '-l', str(i), pdf, '-'], capture_output=True, text=True).stdout
        if txt.lstrip().startswith(('CROQUIS —', 'LOSA MACIZA —', '▼ LOSA MACIZA')):
            pages.append(i)
    return pages


def render(sel, outdir, tag, dpi=110):
    import gen_general
    import detalle
    detalle.DEFAULT_SEL = sel
    tmp = tempfile.mkdtemp()
    full = os.path.join(tmp, 'croquis_%s.xlsx' % tag)
    gen_general.build(full)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, full], check=True, capture_output=True, timeout=300)
    pdf = full[:-5] + '.pdf'
    os.makedirs(outdir, exist_ok=True)
    out = []
    for p in croquis_pages(pdf):
        stem = os.path.join(outdir, '%s_p%d' % (tag, p))
        subprocess.run(['pdftoppm', '-png', '-r', str(dpi), '-f', str(p), '-l', str(p), '-singlefile', pdf, stem], check=True)
        out.append(stem + '.png')
    return out


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'out')
    sels = sys.argv[2:] or EJEMPLOS
    for i, s in enumerate(sels):
        print(s, '->', render(s, out, 't%d' % (i + 1)))
