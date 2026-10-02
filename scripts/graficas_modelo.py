"""Gráficas de verificación del modelo (muros y losas sobre el plano) y del espectro."""
import os
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import ezdxf
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy

sys.path.insert(0, os.path.dirname(__file__))
import geometria_modelo as G  # noqa: E402
import build_e2k as B  # noqa: E402


def planta(ax, doc, lay, off, ws, losa, titulo, grupos=None):
    ctx = RenderContext(doc)
    for name in list(ctx.layers):
        if not name.upper().startswith(('001_MUROS', '002_EJES', '012_VENT')):
            ctx.layers[name].is_visible = False
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.MONOCHROME_LIGHT_BG,
                        lineweight_scaling=0.3)
    Frontend(ctx, MatplotlibBackend(ax), config=cfg).draw_layout(lay, finalize=False)
    if losa is not None:
        geoms = [losa] if losa.geom_type == 'Polygon' else list(losa.geoms)
        for g in geoms:
            xs, ys = g.exterior.xy
            ax.fill([x + off for x in xs], ys, color='#9ecae1', alpha=0.45, lw=0)
            for h in g.interiors:
                xs, ys = h.xy
                ax.fill([x + off for x in xs], ys, color='white', alpha=1.0, lw=0)
    for w in ws:
        o, c, a, b = w[:4]
        col = '#d62728'
        if grupos is not None:
            col = grupos.get(tuple(w[:4]), col)
        if o == 'H':
            ax.plot([a + off, b + off], [c, c], color=col, lw=3.2, solid_capstyle='butt')
        else:
            ax.plot([c + off, c + off], [a, b], color=col, lw=3.2, solid_capstyle='butt')
    ax.set_xlim(28.8 + off, 43.9 + off)
    ax.set_ylim(-1.0, 25.2)
    ax.set_aspect('equal')
    ax.set_axis_off()
    ax.set_title(titulo, fontsize=11)


def main():
    doc = ezdxf.readfile(G.DXF)
    lay = doc.modelspace()
    pb, pa = G.leer()
    pb_s, pa_s = B.partir(pb), B.partir(pa)

    def cont(w):
        o, c, a, b = w
        L = 0
        for q in pb_s:
            if q[0] == o and abs(q[1] - c) < 0.02:
                L += max(0, min(b, q[3]) - max(a, q[2]))
        return L >= 0.9 * (b - a)
    gr = {tuple(w): ('#d62728' if cont(w) else '#ff7f0e') for w in pa_s}
    fig, axs = plt.subplots(1, 2, figsize=(13, 11))
    planta(axs[0], doc, lay, 0.0, pb_s, G.losa_entrepiso(), 'Planta baja: muros (rojo) y losa N+3.00 (azul)')
    planta(axs[1], doc, lay, 30.0, pa_s, G.losa_azotea(),
           'Planta alta: muros continuos (rojo) / sobre losa (naranja) y losa N+6.00 (azul)', gr)
    fig.tight_layout()
    out = sys.argv[1] if len(sys.argv) > 1 else 'modelo_plantas.png'
    fig.savefig(out, dpi=110)
    # espectro
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    T = [i * 0.01 for i in range(0, 401)]
    ax.plot(T, [B.sa_elastico(t) for t in T], color='#1f77b4', lw=2, label='Elástico (transparente)')
    ax.plot(T, [B.sa_diseno(t) for t in T], color='#d62728', lw=2, label="Diseño SV: Sa/(Q'R), Q=2")
    ax.plot(T, [B.sa_elastico(t) / 6 for t in T], color='#2ca02c', lw=1.5, ls='--', label='Limitación de daños (elástico/6)')
    ax.set_xlabel('Periodo T (s)')
    ax.set_ylabel('Sa / g')
    ax.grid(alpha=0.3)
    ax.legend()
    ax.set_xlim(0, 4)
    ax.set_title('Espectro Morelia (forma MDOC CFE 2015) — a0r = %.2f g, terreno tipo II' % B.SIS['a0r'])
    fig.tight_layout()
    fig.savefig(out.replace('.png', '_espectro.png'), dpi=130)


if __name__ == '__main__':
    main()


def iso3d(out):
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    pb, pa = G.leer()
    pb_s, pa_s = B.partir(pb), B.partir(pa)
    fig = plt.figure(figsize=(9, 9))
    ax = fig.add_subplot(111, projection='3d')
    polys, cols = [], []
    for lst, z0 in ((pb_s, 0.0), (pa_s, 3.0)):
        for o, c, a, b in lst:
            p1 = B.plan_xy(o, c, a)
            p2 = B.plan_xy(o, c, b)
            x1, y1, x2, y2 = G.rx(p1[0]), G.ry(p1[1]), G.rx(p2[0]), G.ry(p2[1])
            polys.append([(x1, y1, z0), (x2, y2, z0), (x2, y2, z0 + 3), (x1, y1, z0 + 3)])
            cols.append('#c96a4a' if z0 == 0 else '#e0a080')
    for poly, z in ((G.losa_entrepiso(), 3.0), (G.losa_azotea(), 6.0)):
        for x0, y0, x1, y1 in G.rectangulos(poly, list(G.EJES_X.values()), list(G.EJES_Y.values())):
            polys.append([(G.rx(x0), G.ry(y0), z), (G.rx(x1), G.ry(y0), z), (G.rx(x1), G.ry(y1), z),
                          (G.rx(x0), G.ry(y1), z)])
            cols.append('#9ecae1')
    pc = Poly3DCollection(polys, facecolors=cols, edgecolors='#444444', linewidths=0.3, alpha=0.85)
    ax.add_collection3d(pc)
    ax.set_xlim(-2.5, 10)
    ax.set_ylim(-1, 25)
    ax.set_zlim(0, 7)
    ax.set_box_aspect((12.5, 26, 7))
    ax.view_init(elev=24, azim=-58)
    ax.set_xlabel('X (m)')
    ax.set_ylabel('Y (m)')
    ax.set_zlabel('Z (m)')
    ax.set_title('Modelo 001_CASA_ROCHA_R1.0: muros de tabique (rojo) y losas reticulares (azul)', fontsize=10)
    fig.tight_layout()
    fig.savefig(out, dpi=130, bbox_inches='tight', pad_inches=0.05)


def pesos():
    pb, pa = G.leer()
    pb_s, pa_s = B.partir(pb), B.partir(pa)
    Lpb = sum(b - a for o, c, a, b in pb_s)
    Lpa = sum(b - a for o, c, a, b in pa_s)
    Lx_pb = sum(b - a for o, c, a, b in pb_s if o == 'H')
    Ly_pb = sum(b - a for o, c, a, b in pb_s if o == 'V')
    w_muro = 0.14 * 1500 * 3.0
    A1, A2 = G.losa_entrepiso().area, G.losa_azotea().area
    wl = B.GAMMA_LOSA * B.H_EQ
    d = dict(L_muros_PB=Lpb, L_muros_PA=Lpa, Lx_PB=Lx_pb, Ly_PB=Ly_pb, A_losa_N3=A1, A_losa_N6=A2,
             peso_muros_t=(Lpb + Lpa) * w_muro / 1000, peso_losas_t=(A1 + A2) * wl / 1000,
             scm_t=(A1 * B.SCM_ENT + A2 * B.SCM_AZO) / 1000,
             cv_inst_t=(A1 * B.CV_ENT[1] + A2 * B.CV_AZO[1]) / 1000)
    return d


if __name__ == '__main__' and len(sys.argv) > 2:
    iso3d(sys.argv[2])
    for k, v in pesos().items():
        print(k, round(v, 2))
