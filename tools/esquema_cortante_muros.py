"""Esquema 3D (deformacion exagerada) de por que truenan a cortante los muros de planta alta:
muro sobre muro, muro sobre dala con claro y dala sobre dala.
Uso: python3 esquema_cortante_muros.py SALIDA.png"""
import math
import sys
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection, Line3DCollection

TAB_PB, TAB_PA, DALA, CONC = "#c8704b", "#e0956f", "#8f8f8f", "#6d6d6d"


def caja(ax, x0, x1, y0, y1, z0, z1, color, dz=lambda x, y, z: 0.0, n=12, alpha=1.0, eje="x"):
    """Caja subdividida a lo largo de su eje largo para que se pueda deformar (dz en cada vertice)."""
    caras = []
    if eje == "x":
        xs = np.linspace(x0, x1, n + 1)
        segs = [((a, b), (y0, y1)) for a, b in zip(xs[:-1], xs[1:])]
    else:
        ys = np.linspace(y0, y1, n + 1)
        segs = [((x0, x1), (a, b)) for a, b in zip(ys[:-1], ys[1:])]
    for (xa, xb), (ya, yb) in segs:
        V = [(x, y, z) for z in (z0, z1) for y in (ya, yb) for x in (xa, xb)]
        V = [(x, y, z - dz(x, y, z)) for x, y, z in V]
        idx = [(0, 1, 3, 2), (4, 5, 7, 6), (0, 1, 5, 4), (2, 3, 7, 6), (0, 2, 6, 4), (1, 3, 7, 5)]
        caras += [[V[i] for i in f] for f in idx]
    pc = Poly3DCollection(caras, facecolor=color, edgecolor=color, linewidths=0.3, alpha=alpha)
    ax.add_collection3d(pc)


def fantasma(ax, x0, x1, y0, y1, z0, z1):
    P = [(x, y, z) for z in (z0, z1) for y in (y0, y1) for x in (x0, x1)]
    E = [(0, 1), (2, 3), (4, 5), (6, 7), (0, 2), (1, 3), (4, 6), (5, 7), (0, 4), (1, 5), (2, 6), (3, 7)]
    ax.add_collection3d(Line3DCollection([[P[a], P[b]] for a, b in E], colors="#9a9a9a", linewidths=0.8,
                                         linestyles="--"))


ENCIMA = []      # (ax, tipo, datos) para dibujar en 2D al final


def grieta(ax, p0, p1, n=9, amp=0.12):
    p0, p1 = np.array(p0, float), np.array(p1, float)
    pts = [p0 + (p1 - p0) * t for t in np.linspace(0, 1, n)]
    pts = [p + (np.array([0, 0, amp]) if i % 2 else 0) for i, p in enumerate(pts)]
    ENCIMA.append((ax, "linea", pts))


def rotulo(ax, x, y, z, texto, **kw):
    ENCIMA.append((ax, "texto", ((x, y, z), texto, kw)))


def dibujar_encima(fig):
    from mpl_toolkits.mplot3d import proj3d
    from matplotlib.lines import Line2D
    fig.canvas.draw()
    inv = fig.transFigure.inverted()
    def f(ax, p):
        x2, y2, _ = proj3d.proj_transform(p[0], p[1], p[2], ax.get_proj())
        return inv.transform(ax.transData.transform((x2, y2)))
    for ax, tipo, datos in ENCIMA:
        if tipo == "linea":
            P = [f(ax, p) for p in datos]
            fig.add_artist(Line2D([q[0] for q in P], [q[1] for q in P], color="black", lw=2.6,
                                  transform=fig.transFigure))
        else:
            p, texto, kw = datos
            q = f(ax, p)
            fig.text(q[0], q[1], texto, transform=fig.transFigure,
                     bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="none", alpha=0.85), **kw)


def flecha(ax, x, y, z, largo=0.9, color="#c62828"):
    ax.quiver(x, y, z + largo, 0, 0, -largo, color=color, arrow_length_ratio=0.35, lw=2)


def base(ax, titulo, xl, yl, azim=-62):
    ax.set_xlim(*xl); ax.set_ylim(*yl); ax.set_zlim(0, 6.6)
    ax.set_box_aspect((xl[1] - xl[0], yl[1] - yl[0], 6.6))
    ax.view_init(elev=18, azim=azim)
    ax.set_axis_off()
    ax.set_title(titulo, fontsize=15, fontweight="bold", pad=0)


fig = plt.figure(figsize=(21, 9.4), dpi=110)

# ---------- 1. muro sobre muro
ax = fig.add_subplot(1, 3, 1, projection="3d")
base(ax, "1. MURO SOBRE MURO (bien)", (-0.5, 4.5), (-1.2, 1.2))
caja(ax, 0, 4, -0.075, 0.075, 0, 2.9, TAB_PB)
caja(ax, 0, 4, -0.075, 0.075, 2.9, 3.2, DALA)
caja(ax, 0, 4, -0.075, 0.075, 3.2, 6.0, TAB_PA)
for x in (0.6, 2.0, 3.4):
    flecha(ax, x, 0, 6.05)
rotulo(ax, 2, 0, 1.4, "muro PB", fontsize=12, ha="center")
rotulo(ax, 2, 0, 4.5, "muro PA", fontsize=12, ha="center")
rotulo(ax, 4.15, 0, 3.0, "dala ahogada\nen la losa", fontsize=11)

# ---------- 2. muro sobre dala con claro
A = 0.45                                           # flecha exagerada
L = 4.5
def dB(x, y, z):
    s = A * math.sin(math.pi * x / L) if 0 <= x <= L else 0.0
    return s if z <= 3.21 else s * (1 - 0.75 * (z - 3.2) / 2.8)
ax = fig.add_subplot(1, 3, 2, projection="3d")
base(ax, "2. MURO SOBRE DALA CON CLARO (óvalos)", (-1.7, 6.2), (-1.2, 1.2))
caja(ax, -1.5, 0, -0.075, 0.075, 0, 2.9, TAB_PB, n=2)
caja(ax, L, L + 1.5, -0.075, 0.075, 0, 2.9, TAB_PB, n=2)
fantasma(ax, -1.5, L + 1.5, -0.075, 0.075, 2.9, 3.2)
caja(ax, -1.5, L + 1.5, -0.075, 0.075, 2.9, 3.2, DALA, dz=dB, n=30)
fantasma(ax, 0.4, 4.1, -0.075, 0.075, 3.2, 6.0)
caja(ax, 0.4, 4.1, -0.075, 0.075, 3.2, 6.0, TAB_PA, dz=dB, n=24)
for x in (0.9, 2.25, 3.6):
    flecha(ax, x, 0, 6.05 - dB(x, 0, 6.0))
grieta(ax, (0.55, -0.09, 3.35), (1.5, -0.09, 5.6))
grieta(ax, (3.95, -0.09, 3.35), (3.0, -0.09, 5.6))
rotulo(ax, L / 2, 0, 0.9, "cuarto abajo\n(sin muro)", fontsize=12, ha="center")
rotulo(ax, L / 2, 0, 4.3, "muro PA", fontsize=12, ha="center")
rotulo(ax, L / 2 + 0.6, 0, 2.2, "la dala se flecha", fontsize=12, ha="center", color="#c62828", fontweight="bold")

# ---------- 3. dala sobre dala
A3 = 0.7
LF, yw, xJ = 5.5, 2.75, 3.6                        # claro de la dala larga, posicion de la dala corta, muro de PB
def dF(y):
    return A3 * math.sin(math.pi * y / LF) if 0 <= y <= LF else 0.0
def dY(x):                                          # la dala corta cae en la dala larga (x=0) y en el muro de PB (x=xJ)
    return dF(yw) * (1 - x / xJ)
def dCorta(x, y, z):
    s_ = dY(x)
    return s_ if z <= 3.21 else s_ * (1 - 0.6 * (z - 3.2) / 2.8)
ax = fig.add_subplot(1, 3, 3, projection="3d")
base(ax, "3. DALA SOBRE DALA (los amarillos)", (-1.0, 4.2), (-0.3, 5.8), azim=-38)
caja(ax, -0.7, 0.7, LF - 0.075, LF + 0.075, 0, 2.9, TAB_PB, n=1)                        # apoyo del fondo
caja(ax, xJ - 0.075, xJ + 0.075, yw - 0.6, yw + 2.2, 0, 2.9, TAB_PB, n=2, eje="y")       # muro PB (eje J)
fantasma(ax, -0.075, 0.075, 0, LF, 2.9, 3.2)
caja(ax, -0.075, 0.075, 0, LF, 2.9, 3.2, DALA, dz=lambda x, y, z: dF(y), n=30, eje="y")  # dala larga
fantasma(ax, 0, xJ, yw - 0.075, yw + 0.075, 2.9, 3.2)
caja(ax, 0, xJ, yw - 0.075, yw + 0.075, 2.9, 3.2, DALA, dz=dCorta, n=16)                # dala corta
fantasma(ax, 0.7, xJ, yw - 0.075, yw + 0.075, 3.2, 6.0)
caja(ax, 0.7, xJ, yw - 0.075, yw + 0.075, 3.2, 6.0, TAB_PA, dz=dCorta, n=14)           # muro PA
caja(ax, -0.7, 0.7, -0.075, 0.075, 0, 2.9, TAB_PB, n=1)                                 # apoyo del frente
for x in (1.1, 2.1, 3.1):
    flecha(ax, x, yw, 6.05 - dCorta(x, yw, 6.0))
grieta(ax, (2.7, yw - 0.1, 3.3 - dY(2.7)), (1.0, yw - 0.1, 5.75 - 0.4 * dY(1.0)))
rotulo(ax, 2.6, yw, 4.6, "muro PA", fontsize=12, ha="center")
rotulo(ax, -0.2, 0.9, 3.5, "dala larga (eje F)\nse flecha", fontsize=12, ha="right", color="#c62828",
       fontweight="bold")
rotulo(ax, 0.0, yw, 2.0 - dY(0), "se baja\n(cae sobre la dala larga)", fontsize=12, color="#c62828", fontweight="bold", ha="right")
rotulo(ax, xJ + 0.1, yw, 3.3, "no se baja", fontsize=12, color="#2e7d32", fontweight="bold")
rotulo(ax, xJ + 0.1, yw + 1.2, 1.0, "muro PB\n(eje J)", fontsize=12)
rotulo(ax, 1.9, yw, 1.55 - dY(1.9), "dala corta\n(su claro sí pasa)", fontsize=11, ha="left")

fig.text(0.17, 0.05, "La carga baja derecho de muro a muro.\nLa dala no trabaja como viga: no hay cortante de gravedad.",
         ha="center", fontsize=13)
fig.text(0.50, 0.05, "Abajo no hay muro: la dala se flecha, el muro se arquea\npara cargarse solo y se agrieta en diagonal (cortante).",
         ha="center", fontsize=13)
fig.text(0.83, 0.05, "La dala corta del muro es buena, pero un extremo cae sobre la\ndala larga del eje F, que se flecha. Ese lado se baja, el otro no\n→ el muro se tuerce en su plano y se agrieta (cortante).",
         ha="center", fontsize=13)
fig.text(0.5, 0.965, "¿Por qué truenan a cortante los muros de planta alta?   (deformación exagerada · línea punteada = sin cargar)",
         ha="center", fontsize=17, fontweight="bold")
fig.text(0.5, 0.93, "Grieta negra = cortante en el muro · flechas rojas = carga de la azotea", ha="center", fontsize=12,
         color="#444")
plt.subplots_adjust(left=0.0, right=1.0, top=0.92, bottom=0.12, wspace=0.0)
dibujar_encima(fig)
fig.savefig(sys.argv[1] if len(sys.argv) > 1 else "Esquema_cortante_muros_PA.png")
print("ok")
