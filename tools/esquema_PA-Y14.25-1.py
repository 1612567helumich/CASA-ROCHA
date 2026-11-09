"""Por que truena PA-Y14.25-1: planta de ubicacion y dos cortes (eje F y eje Y14.25) con la
deformacion exagerada. Geometria del modelo R1.2 y fuerzas de la corrida R1.2.
Uso: python3 esquema_PA-Y14.25-1.py SALIDA.png"""
import math
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle, FancyArrowPatch

TAB_PB, TAB_PA, DALA, CONC, ROJO, VERDE = "#c8704b", "#e0956f", "#8f8f8f", "#6d6d6d", "#c62828", "#2e7d32"

# geometria (m)
yF0, yF1, xF = 9.95, 15.447, 5.422          # dala del eje F: de la dala del eje 8 al muro de concreto PB-11-1
yw, xw0, xJ = 14.247, 6.25, 8.85            # PA-Y14.25-1 y su dala (de la dala del eje F al muro PB-J-1)
yPF0, yPF1 = 11.0, 15.85                    # PA-F-1
A = 0.45                                    # flecha maxima dibujada (exagerada)
zd0, zd1, zm1 = 2.9, 3.1, 5.9               # dala y muro de PA


def dF(y):
    return A * math.sin(math.pi * (y - yF0) / (yF1 - yF0)) if yF0 <= y <= yF1 else 0.0


dJ = dF(yw)                                  # lo que baja el apoyo izquierdo de la dala de PA-Y14.25-1


def dY(x):
    return dJ * (xJ - x) / (xJ - xF) if xF <= x <= xJ else 0.0


def muro_def(ax, h0, h1, d, color, n=40, alivio=0.7):
    """Muro de PA de h0 a h1 sobre una dala que baja d(h): abajo sigue a la dala, arriba menos."""
    hs = [h0 + (h1 - h0) * i / n for i in range(n + 1)]
    abajo = [(h, zd1 - d(h)) for h in hs]
    arriba = [(h, zm1 - (1 - alivio) * d(h)) for h in reversed(hs)]
    ax.add_patch(Polygon(abajo + arriba, closed=True, fc=color, ec="#7a4a35", lw=1))
    ax.add_patch(Rectangle((h0, zd1), h1 - h0, zm1 - zd1, fill=False, ec="#999", ls="--", lw=1))


def dala_def(ax, h0, h1, d, n=40):
    hs = [h0 + (h1 - h0) * i / n for i in range(n + 1)]
    ax.add_patch(Polygon([(h, zd0 - d(h)) for h in hs] + [(h, zd1 - d(h)) for h in reversed(hs)], closed=True,
                         fc=DALA, ec="#555", lw=1))
    ax.add_patch(Rectangle((h0, zd0), h1 - h0, zd1 - zd0, fill=False, ec="#999", ls="--", lw=1))


def grieta(ax, p0, p1, n=10, amp=0.07):
    xs, zs = [], []
    for i in range(n + 1):
        t = i / n
        xs.append(p0[0] + (p1[0] - p0[0]) * t + (amp if i % 2 else -amp))
        zs.append(p0[1] + (p1[1] - p0[1]) * t)
    ax.plot(xs, zs, color="black", lw=2.4)


def cargas(ax, hs, d):
    for h in hs:
        z = zm1 - 0.3 * d(h)
        ax.add_patch(FancyArrowPatch((h, z + 0.6), (h, z + 0.05), arrowstyle="-|>", mutation_scale=16, color=ROJO, lw=2))


fig = plt.figure(figsize=(17, 10.5), dpi=110)
gs = fig.add_gridspec(2, 2, width_ratios=[1, 2.1], height_ratios=[1, 1], wspace=0.08, hspace=0.32,
                      left=0.03, right=0.98, top=0.88, bottom=0.05)

# ------------------------------------------------------------------ planta de ubicacion
ax = fig.add_subplot(gs[:, 0])
ax.set_title("Dónde está (planta, nivel N+3.00)", fontsize=14, fontweight="bold")
ax.plot([xJ, xJ], [9.3, 16.4], color=TAB_PB, lw=6, solid_capstyle="butt")
ax.plot([4.925, xJ], [yF1, yF1], color=CONC, lw=6)
ax.plot([8.14, xJ], [yF0, yF0], color=TAB_PB, lw=6)
ax.plot([4.4, xJ], [yF0, yF0], color=DALA, lw=2.5)
ax.plot([xF, xF], [yF0, yF1], color=ROJO, lw=4)
ax.plot([xF, xJ], [yw, yw], color="#1565c0", lw=4)
ax.plot([xF, xF], [yPF0, yPF1], color=TAB_PA, lw=11, alpha=0.55)
ax.plot([xw0, xJ], [yw, yw], color=TAB_PA, lw=11, alpha=0.55)
ax.text(xJ + 0.15, 12.2, "PB-J-1", fontsize=11, rotation=90, va="center")
ax.text(6.9, yF1 + 0.2, "PB-11-1 (concreto)", fontsize=11, ha="center")
ax.text(8.5, yF0 - 0.45, "PB-8-1", fontsize=10, ha="center")
ax.text(4.6, yF0 + 0.15, "dala eje 8", fontsize=10)
ax.text(xF - 0.15, 12.7, "dala eje F (corte A)\nclaro 5.50 m", fontsize=11, color=ROJO, ha="right", va="center",
        fontweight="bold")
ax.text(7.1, yw - 0.45, "dala de PA-Y14.25-1\n(corte B) claro 3.41 m", fontsize=11, color="#1565c0", ha="center",
        va="top", fontweight="bold")
ax.text(7.55, yw + 0.3, "PA-Y14.25-1", fontsize=11, ha="center")
ax.text(xF + 0.2, 11.4, "PA-F-1", fontsize=11, rotation=90)
ax.plot([xF], [yw], "o", ms=12, mfc="none", mec="black", mew=2)
ax.annotate("aquí se apoya\nuna dala en la otra", (xF, yw), (4.2, 15.6), fontsize=11, ha="left",
            arrowprops=dict(arrowstyle="->", lw=1.5))
ax.text(4.3, 9.0, "rojo y azul = dalas sin muro abajo · naranja claro = muros de PA\n"
        "ladrillo = muros de PB · gris oscuro = concreto", fontsize=9.5, color="#444", va="top")
ax.set_xlim(4.0, 9.3); ax.set_ylim(8.4, 16.6); ax.set_aspect("equal"); ax.axis("off")

# ------------------------------------------------------------------ corte A: eje F
ax = fig.add_subplot(gs[0, 1])
ax.set_title("Corte A — eje F: 5.50 m sin muro abajo, carga a PA-F-1 y se flecha",
             fontsize=14, fontweight="bold", loc="left")
ax.add_patch(Rectangle((yF1 - 0.075, 0), 0.15, zd0, fc=CONC))                  # PB-11-1 (se ve en corte)
ax.add_patch(Rectangle((yF0 - 0.075, zd0), 0.15, zd1 - zd0, fc="#555"))        # dala eje 8 (en corte)
dala_def(ax, yF0, yF1 + 0.4, dF)
muro_def(ax, yPF0, yPF1, dF, TAB_PA)
cargas(ax, [11.6, 12.6, 13.6, 14.6, 15.5], dF)
grieta(ax, (11.2, zd1 - dF(11.2) + 0.1), (12.4, 5.5))
grieta(ax, (15.3, zd1 + 0.1), (14.2, 5.5))
ax.plot([yw], [zd0 - dF(yw) - 0.02], "o", ms=13, mfc="#1565c0", mec="black")
ax.annotate(f"aquí llega la dala de PA-Y14.25-1:\nse baja junto con la dala del eje F", (yw, zd0 - dF(yw) - 0.05),
            (12.4, 0.9), fontsize=12, color="#1565c0", fontweight="bold", ha="center",
            arrowprops=dict(arrowstyle="->", lw=1.8, color="#1565c0"))
ax.text(yF1 + 0.12, 1.4, "PB-11-1\n(concreto)", fontsize=11)
ax.text(yF0, zd1 + 0.15, "dala\neje 8", fontsize=10, ha="center")
ax.text(13.4, 4.3, "PA-F-1", fontsize=13, ha="center", fontweight="bold")
ax.text(10.6, 1.6, "sin muro abajo\n(cuarto)", fontsize=11, ha="center", color="#555")
ax.set_xlim(9.4, 16.6); ax.set_ylim(0, 6.7); ax.set_aspect("equal"); ax.axis("off")

# ------------------------------------------------------------------ corte B: eje Y14.25
ax = fig.add_subplot(gs[1, 1])
ax.set_title("Corte B — eje Y14.25: la dala del muro",
             fontsize=14, fontweight="bold", loc="left")
ax.add_patch(Rectangle((xJ - 0.075, 0), 0.15, zd0, fc=TAB_PB))                 # PB-J-1 (en corte)
ax.add_patch(Rectangle((xF - 0.075, zd0 - dJ), 0.15, zd1 - zd0, fc=ROJO))      # dala eje F (en corte), bajada
ax.add_patch(Rectangle((xF - 0.075, zd0), 0.15, zd1 - zd0, fill=False, ec="#999", ls="--"))
dala_def(ax, xF, xJ, dY)
muro_def(ax, xw0, xJ, dY, TAB_PA, alivio=0.6)
cargas(ax, [6.6, 7.3, 8.0, 8.6], dY)
grieta(ax, (8.1, zd1 - dY(8.1) + 0.1), (6.55, 5.6))
ax.add_patch(FancyArrowPatch((xF, zd0 + 0.9), (xF, zd0 - dJ + 0.25), arrowstyle="-|>", mutation_scale=22,
                             color=ROJO, lw=3))
ax.text(xF - 0.12, zd0 - dJ - 0.25, "se baja\n(dala eje F)", fontsize=13, color=ROJO, fontweight="bold", ha="right",
        va="top")
ax.text(xJ + 0.15, zd1 + 0.1, "no se baja\n(PB-J-1)", fontsize=13, color=VERDE, fontweight="bold", va="bottom")
ax.text(xJ + 0.15, 1.2, "PB-J-1", fontsize=11)
ax.text(5.83, 4.2, "vano", fontsize=11, ha="center", color="#777")
ax.text(7.95, 5.05, "PA-Y14.25-1", fontsize=13, ha="center", fontweight="bold")
ax.text(7.0, 1.3, "sin muro abajo", fontsize=11, ha="center", color="#555")
ax.text(9.75, 5.8, "V_u = 5.4 t\n(5.0 t de pura gravedad)\naguanta V_R = 3.1 t\nratio 1.75", fontsize=12.5,
        va="top", bbox=dict(boxstyle="round,pad=0.5", fc="#fff3f3", ec=ROJO))
ax.set_xlim(4.4, 11.9); ax.set_ylim(0, 6.7); ax.set_aspect("equal"); ax.axis("off")

fig.suptitle("PA-Y14.25-1: su dala está bien, pero su lado izquierdo cuelga de la dala del eje F",
             fontsize=19, fontweight="bold", y=0.975)
fig.text(0.5, 0.925, "Deformación exagerada · punteado = sin cargar · grieta negra = cortante · flechas rojas = carga "
         "de azotea · Arreglo: peraltar la dala del eje F (óvalo 1, 15×50) o darle apoyo abajo; se arreglan los dos muros.",
         ha="center", fontsize=12, color="#333")
fig.savefig(sys.argv[1] if len(sys.argv) > 1 else "PA-Y14.25-1.png")
print("ok")
