#!/usr/bin/env python3
"""Plantas de piers (PB y PA) con el ratio de cada pier y, si falla, por que.

- Muros de tabique: resultados de la hoja de muros (MUROS), NTC-Mamposteria 2023.
  Ratio = el mayor de P_u/P_R, M_u/M_R, V_u/V_R y, si desplanta sobre viga,
  L_v/h_v contra su limite y f_CM contra F_R*f'm (4.1.18).
- Muros de concreto (seccion de muro con material de concreto en el e2k): revision
  rapida NTC-Concreto 2023, ductilidad baja, con las fuerzas de la misma hoja:
    cortante 6.5.5.3.3  V_R = F_R (alpha_c sqrt(f'c) + p_t f_yt) A_cv, F_R = 0.75
    flexocompresion por compatibilidad de deformaciones (5.3), F_R segun 3.8.2.2
  Armado supuesto: una capa de #3 @ 18 cm vertical y horizontal (p = 0.0026 >= 0.0025,
  6.5.6.2) y castillos de los extremos con 4 #3.

Uso: python3 plantas_piers_ratios.py MODELO.e2k HOJA.xlsx SALIDA.pdf [PIER,PIER,...]
     La lista opcional de piers marca con un circulo negro el claro de su viga de apoyo
     (L_v de la hoja) con el peralte que pide la 4.1.18 o, si no es por claro, el f_CM.
"""
import math
import os
import re
import sys
import tempfile

import openpyxl
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from redibujar_muros import Model  # noqa: E402

FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
F = lambda s, b=False: ImageFont.truetype(BOLD if b else FONT, s)
PX = 78           # px por metro
M0 = 150          # margen
PANEL = 1250      # ancho de la tabla de la derecha
VERDE, ROJO, GRIS = (30, 140, 60), (200, 35, 35), (170, 170, 170)

# ------------------------------------------------------------------ concreto
FY, ES, ECU, B1 = 4200.0, 2.0e6, 0.003, 0.85
AS_CAST, A_BAR, S_BAR, P_T = 2.84, 0.71, 18.0, 0.0025


def _barras(L):
    bs = [(7.5, AS_CAST), (L - 7.5, AS_CAST)]
    x = 15 + S_BAR
    while x < L - 15 - 5:
        bs.append((x, A_BAR))
        x += S_BAR
    return bs


def _fuerza(c, L, t, fc, bs):
    a = min(B1 * c, L)
    C = 0.85 * fc * a * t
    N, M, et = C, C * (L / 2 - a / 2), 0.0
    for x, A in bs:
        e = ECU * (c - x) / c
        s = max(-FY, min(FY, e * ES))
        if x <= a and s > 0:
            s -= 0.85 * fc
        N += A * s
        M += A * s * (L / 2 - x)
        et = max(et, -e)
    return N, M, et


def m_resistente(P, L, t, fc):
    """F_R*M_n (t*m) con P (kg, compresion +)."""
    bs = _barras(L)
    lo, hi = 0.5, 5 * L
    for _ in range(200):
        c = (lo + hi) / 2
        if _fuerza(c, L, t, fc, bs)[0] > P:
            hi = c
        else:
            lo = c
    _, M, et = _fuerza(c, L, t, fc, bs)
    ety = FY / ES
    fr = 0.65 if et <= ety else (0.9 if et >= ety + 0.003 else 0.65 + 0.25 * (et - ety) / 0.003)
    return fr * M / 1e5


def v_resistente(L, t, fc, Hm, Pu):
    r = Hm / L
    ac = 0.8 if r <= 1.5 else (0.53 if r >= 2 else 0.8 - 0.27 * (r - 1.5) / 0.5)
    if Pu < 0:
        ac = max(0.0, 0.53 * (1 + Pu * 1000 / (35 * t * L)))
    return 0.75 * (ac * math.sqrt(fc) + P_T * FY) * t * L / 1000


def revisa_concreto(r, t, fc, Hm):
    L = r["F"] * 100
    vr = v_resistente(L, t, fc, Hm, r["BN"])
    mr = min(m_resistente(r["BN"] * 1000, L, t, fc), m_resistente(r["BO"] * 1000, L, t, fc))
    rv, rm = r["BK"] / vr, r["BL"] / mr
    no = []
    if rv > 1:
        no.append(f"cortante: V_u {r['BK']:.1f} t > V_R {vr:.1f} t")
    if rm > 1:
        no.append(f"flexión: M_u {r['BL']:.1f} > M_R {mr:.1f} t·m")
    txt = (f"CONCRETO f'c {fc:.0f}, {t:.0f} cm, #3 @ 18 una capa. Cortante {rv:.2f} (V_u {r['BK']:.1f} / "
           f"V_R {vr:.0f} t), flexión {rm:.2f} (M_u {r['BL']:.1f} / M_R {mr:.0f} t·m)")
    return max(rv, rm), no, txt


# ------------------------------------------------------------------ mamposteria
def revisa_tabique(r, frfm):
    """(ratio, [razones], texto_ratio)"""
    rat, no = [], []
    est = r["EW"] or ""
    fallas = [s.strip() for s in (r["EU"] or "").split(";") if s.strip()]
    if any(s.startswith("sin tramos") for s in fallas):
        return math.inf, [f"machón de {r['F'] * 100:.0f} cm, menor que L mín.: diséñalo como columna"], "col."
    for k in ("CQ", "CX", "EC"):
        v = r[k]
        if isinstance(v, (int, float)) and v < 99:
            rat.append(v)
    if "compresión" in fallas:
        no.append(f"compresión {r['CQ']:.2f}: P_u {r['BO']:.1f} > P_R {r['CP']:.1f} t")
    if "flexión" in fallas:
        no.append(f"flexión {r['CX']:.2f}: M_u {r['BL']:.1f} > M_R {r['CW']:.1f} t·m")
    if "cortante" in fallas or "refuerzo horizontal" in fallas:
        if r["EC"] >= 99:
            rat.append(math.inf)
            s = "cortante ∞: P en tensión, el tabique no aporta (V_R = 0) y todo lo debe tomar el refuerzo horizontal"
        else:
            s = f"cortante {r['EC']:.2f}: V_u {r['BK']:.1f} > V_R {r['EB']:.1f} t"
        if r["ED"] == 3:
            s += "; con alambres en las juntas no se puede, hay que probar con malla"
        elif r["ED"] == -1:
            s += "; ni con refuerzo horizontal alcanza: más largo, más grueso o de concreto"
        no.append(s)
    ep = r["EP"] or ""
    if isinstance(r["EN"], (int, float)) and r["EN"] > 0:
        lims = [14.0] + ([r["EO"]] if isinstance(r["EO"], (int, float)) and r["EO"] > 0 else [])
        rl = r["EN"] / min(lims)
        fcm = (r["EM"] or 0) / frfm if frfm else 0
        rat += [rl, fcm]
        if ep.startswith("NO"):
            partes = []
            if "L_v/h_v" in ep or "flecha" in ep:
                por = "por flecha" if min(lims) < 14 else "límite 14"
                partes.append(f"viga de poco peralte, L/h {r['EN']:.1f} contra {min(lims):.1f} máx. {por} ({rl:.2f})")
            if "f_CM" in ep:
                partes.append(f"compresión en la base del muro f_CM = {fcm:.2f} veces F_R·f'm")
            no.append("muro sobre viga: " + "; ".join(partes))
    return (max(rat) if rat else 0.0), no, None


def leer_hoja(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    M, D = wb["MUROS"], wb["DATOS"]
    frfm = (D["C81"].value or 0) * (D["C30"].value or 0)
    from openpyxl.utils import column_index_from_string as ci
    cols = "B F AM BK BL BN BO CP CQ CW CX EB EC ED EM EN EO EP EU EW".split()
    out = {}
    for row in range(1, M.max_row + 1):
        name = M.cell(row, 2).value
        if not isinstance(name, str) or not re.match(r"P[AB]-", name):
            continue
        out[name] = {c: M.cell(row, ci(c)).value for c in cols}
    return out, frfm


def materiales(txt):
    tipo = dict(re.findall(r'MATERIAL\s+"([^"]+)"\s+TYPE "([^"]+)"', txt))
    fc = {m: float(v) / 1e4 for m, v in re.findall(r'MATERIAL\s+"([^"]+)"\s+FC (\S+)', txt)}
    sec = {}
    for s, m, t in re.findall(r'SHELLPROP\s+"([^"]+)"\s+PROPTYPE\s+"Wall"\s+MATERIAL "([^"]+)".*?WALLTHICKNESS (\S+)', txt):
        sec[s] = dict(tipo=tipo.get(m), fc=fc.get(m), t=float(t) * 100)
    return sec


def envolver(d, texto, font, ancho):
    lineas, cur = [], ""
    for w in texto.split(" "):
        prueba = (cur + " " + w).strip()
        if d.textlength(prueba, font=font) <= ancho:
            cur = prueba
        else:
            lineas.append(cur)
            cur = w
    if cur:
        lineas.append(cur)
    return lineas


def vigas_del_piso(txt, story):
    pts = {m[1]: (float(m[2]), float(m[3]))
           for m in re.finditer(r'^\s*POINT "([^"]+)"\s+(-?[\d.E+-]+)\s+(-?[\d.E+-]+)', txt, re.M)}
    asig = {m[1] for m in re.finditer(r'LINEASSIGN\s+"([^"]+)"\s+"%s"' % re.escape(story), txt)}
    return [(pts[m[2]], pts[m[3]]) for m in re.finditer(r'^\s*LINE\s+"([^"]+)"\s+BEAM\s+"([^"]+)"\s+"([^"]+)"', txt, re.M)
            if m[1] in asig]


def claro_de(e, Lv, ax, co, walls, story_abajo, castillos, vigas):
    """Tramo (h0, h1) de la viga de apoyo: entre apoyos consecutivos (muros o castillos del piso de abajo
    o vigas que la cruzan) cuya distancia sea L_v y que mas se encime con el muro; si no hay, L_v centrado."""
    cand = set()
    for (s2, ax2, co2), ps in walls.items():
        if s2 != story_abajo:
            continue
        for p in ps:
            if ax2 == ax and abs(co2 - co) < 0.02:
                cand.update((p["h0"], p["h1"]))
            elif ax2 != ax and p["h0"] - 0.02 <= co <= p["h1"] + 0.02:
                cand.add(co2)
    for x, y, *_ in castillos:
        if (ax == "X" and abs(x - co) < 0.02) or (ax == "Y" and abs(y - co) < 0.02):
            cand.add(y if ax == "X" else x)
    for A, B in vigas:
        k = 0 if ax == "Y" else 1          # coordenada a lo largo del eje
        if abs(A[k] - B[k]) < 0.01 and min(A[1 - k], B[1 - k]) - 0.02 <= co <= max(A[1 - k], B[1 - k]) + 0.02:
            cand.add(A[k])
    cand = sorted(cand)
    mejor = None
    for i, a in enumerate(cand):
        for b in cand[i + 1:]:
            if abs((b - a) - Lv) < 0.06:
                ov = min(b, e["h1"]) - max(a, e["h0"])
                if mejor is None or ov > mejor[0]:
                    mejor = (ov, a, b)
    if mejor and mejor[0] > 0:
        return mejor[1], mejor[2]
    m = (e["h0"] + e["h1"]) / 2
    return m - Lv / 2, m + Lv / 2


def main(e2k, hoja, salida, claros=""):
    claros = [c.strip() for c in claros.split(",") if c.strip()]
    txt = open(e2k, encoding="utf-8", newline="").read()
    M = Model(txt)
    secs = materiales(txt)
    res, frfm = leer_hoja(hoja)
    grids = {"X": [], "Y": []}
    for m in re.finditer(r'GRID "[^"]+"\s+LABEL "([^"]+)"\s+DIR "(X|Y)"\s+COORD (\S+)', txt):
        grids[m[2]].append((m[1], float(m[3])))
    walls, _ = M.wall_pieces()
    xs = [p[0] for p in M.points.values()] + [c for _, c in grids["X"]]
    ys = [p[1] for p in M.points.values()] + [c for _, c in grids["Y"]]
    x0, x1, y0, y1 = min(xs) - 0.3, max(xs) + 0.3, min(ys) - 0.3, max(ys) + 0.3
    Wp = int((x1 - x0) * PX) + 2 * M0 + 120
    W = Wp + PANEL
    H = int((y1 - y0) * PX) + 2 * M0 + 160
    X = lambda x: M0 + (x - x0) * PX
    Y = lambda y: H - M0 - (y - y0) * PX
    paginas, resumen = [], []
    for story, titulo in ((M.order[-2], "PLANTA BAJA"), (M.order[-3], "PLANTA ALTA")):
        Hm = (M.elev[story] - M.elev[M.below(story)]) * 100
        img = Image.new("RGB", (W, H), "white")
        d = ImageDraw.Draw(img)
        d.text((30, 20), f"{titulo} — piers con su ratio (muros de {M.below(story)} a {story})", font=F(30, True),
               fill=(20, 20, 20))
        d.text((30, 62), "Modelo R1.2 · fuerzas de la corrida R1.2 · hoja Muros_CasaRocha_NTC2023 · "
               "verde = cumple (ratio ≤ 1) · rojo = no cumple · rayado = muro de concreto",
               font=F(17), fill=(60, 60, 60))
        d.text((30, 88), "Ratio = el mayor de P_u/P_R, M_u/M_R, V_u/V_R y, si desplanta sobre viga, L/h de la viga y "
               "f_CM/(F_R·f'm). ∞ = V_R vale 0. col. = machón que va como columna.", font=F(17), fill=(60, 60, 60))
        for lab, c in grids["X"]:
            d.line([(X(c), Y(y0) + 10), (X(c), Y(y1) - 10)], fill=(205, 215, 235), width=1)
            for yy in (Y(y1) - 40, Y(y0) + 40):
                d.ellipse([X(c) - 17, yy - 17, X(c) + 17, yy + 17], outline=(120, 140, 180), width=2, fill="white")
                d.text((X(c), yy), lab, font=F(17, True), fill=(60, 80, 130), anchor="mm")
        for lab, c in grids["Y"]:
            d.line([(X(x0) - 10, Y(c)), (X(x1) + 10, Y(c))], fill=(205, 215, 235), width=1)
            xx = X(x0) - 45
            d.ellipse([xx - 17, Y(c) - 17, xx + 17, Y(c) + 17], outline=(120, 140, 180), width=2, fill="white")
            d.text((xx, Y(c)), lab, font=F(15, True), fill=(60, 80, 130), anchor="mm")
        piers = {}
        for (s, ax, co), pcs in walls.items():
            if s != story:
                continue
            for p in pcs:
                a = (co, p["h0"]) if ax == "X" else (p["h0"], co)
                b = (co, p["h1"]) if ax == "X" else (p["h1"], co)
                if p["pier"]:
                    e = piers.setdefault(p["pier"], dict(ax=ax, co=co, h0=9e9, h1=-9e9, secs=set()))
                    e["h0"], e["h1"] = min(e["h0"], p["h0"]), max(e["h1"], p["h1"])
                    e["secs"].add(p["section"])
                else:
                    d.line([(X(a[0]), Y(a[1])), (X(b[0]), Y(b[1]))], fill=GRIS, width=5)
        orden = [n for n in res if n in piers] + sorted(n for n in piers if n not in res)
        filas = []
        for name in orden:
            e = piers[name]
            sec = secs.get(next(iter(e["secs"])), {})
            e["concreto"] = sec.get("tipo") == "Concrete"
            r = res.get(name)
            if r is None:
                ratio, no, extra = math.nan, ["no está en la hoja de muros"], None
            elif e["concreto"]:
                ratio, no, extra = revisa_concreto(r, sec["t"], sec["fc"], Hm)
            else:
                ratio, no, extra = revisa_tabique(r, frfm)
            ok = not no
            e["col"] = VERDE if ok else ROJO
            e["ratio"] = ratio
            e["rtxt"] = extra if extra == "col." else ("∞" if ratio == math.inf else f"{ratio:.2f}")
            nota = ""
            if ok and r is not None and "carga vertical" in (r["EW"] or "") and not e["concreto"]:
                nota = "cumple (ojo: casi todo su V_u viene de carga vertical)"
            filas.append((name, e, ok, no, extra if e["concreto"] else None, nota))
            if e["ax"] == "X":
                p0, p1 = (X(e["co"]), Y(e["h0"])), (X(e["co"]), Y(e["h1"]))
            else:
                p0, p1 = (X(e["h0"]), Y(e["co"])), (X(e["h1"]), Y(e["co"]))
            e["p0"], e["p1"] = p0, p1
            if e["concreto"]:
                wdt = 16
                if e["ax"] == "X":
                    box = [p0[0] - wdt / 2, min(p0[1], p1[1]), p0[0] + wdt / 2, max(p0[1], p1[1])]
                else:
                    box = [min(p0[0], p1[0]), p0[1] - wdt / 2, max(p0[0], p1[0]), p0[1] + wdt / 2]
                d.rectangle(box, fill="white", outline=e["col"], width=2)
                k = box[0] - (box[3] - box[1])
                while k < box[2]:
                    a0 = (max(k, box[0]), box[3] - max(0, box[0] - k))
                    a1 = (min(k + (box[3] - box[1]), box[2]), box[1] + max(0, k + (box[3] - box[1]) - box[2]))
                    d.line([a0, a1], fill=e["col"], width=2)
                    k += 9
            else:
                d.line([p0, p1], fill=e["col"], width=9)
        for x, y, pt, ln in M.castillos()[story]:
            d.rectangle([X(x) - 5, Y(y) - 5, X(x) + 5, Y(y) + 5], fill=(25, 25, 25))
        puestas = []

        def libre(bx):
            return (bx[2] <= Wp - 10) and all(bx[2] <= o[0] or bx[0] >= o[2] or bx[3] <= o[1] or bx[1] >= o[3]
                                              for o in puestas)

        for name, e, ok, no, _, _ in filas:
            (ax0, ay0), (ax1, ay1) = e["p0"], e["p1"]
            texto = f"{name} · {e['rtxt']}"
            f = F(16, True)
            tw, th = d.textbbox((0, 0), texto, font=f)[2:]
            lab = Image.new("RGBA", (tw + 10, th + 8), (255, 255, 255, 0))
            ld = ImageDraw.Draw(lab)
            ld.rounded_rectangle([0, 0, tw + 9, th + 7], radius=4, fill=(255, 255, 255, 240), outline=e["col"], width=2)
            ld.text((5, 3), texto, font=f, fill=e["col"])
            mx, my = (ax0 + ax1) / 2, (ay0 + ay1) / 2
            if e["ax"] == "X":
                lab = lab.rotate(90, expand=True)
                w, h = lab.size
                cands = [(mx + 10 + side, my - h / 2 + k) for k in (0, -40, 40, -80, 80, -120, 120, -160, 160)
                         for side in (0, -w - 20)]
            else:
                w, h = lab.size
                cands = [(mx - w / 2 + k, my + side) for k in (0, -50, 50, -100, 100, -150, 150)
                         for side in (-h - 9, 9, -2 * h - 14, h + 14)]
            for cx, cy in cands:
                bx = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
                if libre(bx):
                    break
            puestas.append(bx)
            img.paste(lab, (int(cx), int(cy)), lab)
        # claros de vigas de apoyo marcados con circulo negro
        marcados = [n for n in claros if n in piers and res.get(n) and isinstance(res[n]["AM"], (int, float))]
        if marcados:
            abajo = M.below(story)
            vigas = vigas_del_piso(txt, abajo)
            cast = M.castillos()[abajo]
            grupos, sobre_muro = {}, []
            for n in marcados:
                e, r = piers[n], res[n]
                ov = sum(max(0.0, min(e["h1"], p["h1"]) - max(e["h0"], p["h0"]))
                         for (s2, ax2, co2), ps in walls.items() if s2 == abajo and ax2 == e["ax"] and abs(co2 - e["co"]) < 0.02
                         for p in ps)
                if ov / (e["h1"] - e["h0"]) > 0.5:      # en realidad va encima de un muro del piso de abajo
                    sobre_muro.append(n)
                    continue
                Lv = r["AM"]
                a, b = claro_de(e, Lv, e["ax"], e["co"], walls, abajo, cast, vigas)
                eo = r["EO"] if isinstance(r["EO"], (int, float)) and r["EO"] > 0 else 14.0
                hreq = math.ceil(max(Lv / 14, Lv / eo) * 100 / 5 - 1e-9) * 5
                g = grupos.setdefault((e["ax"], round(e["co"], 3), round(a, 2), round(b, 2)),
                                      dict(ax=e["ax"], co=e["co"], a=a, b=b, Lv=Lv, h=0, fcm=0.0, piers=[]))
                g["h"] = max(g["h"], hreq)
                g["fcm"] = max(g["fcm"], (r["EM"] or 0) / frfm if frfm else 0)
                g["piers"].append(n)
            d.text((30, 114), "Círculo negro = claro de la dala de apoyo que no pasa la 4.1.18 (NTC-M) ni con vigas de "
                   "15×30. Dice el claro y el peralte que pide, o si es por f_CM (muro corto).",
                   font=F(17, True), fill=(0, 0, 0))
            if sobre_muro:
                nota = ("Ojo: " + ", ".join(sobre_muro) + " la hoja los trae sobre viga, pero en el modelo van encima "
                        "de muros de planta baja: hay que corregir esa captura (no llevan círculo).")
                for i, ln in enumerate(envolver(d, nota, F(17, True), Wp - 60)):
                    d.text((30, 140 + 22 * i), ln, font=F(17, True), fill=ROJO)
            for g in grupos.values():
                a, b, Lv = g["a"], g["b"], g["Lv"]
                if g["ax"] == "X":
                    box = [X(g["co"]) - 34, Y(b) - 16, X(g["co"]) + 34, Y(a) + 16]
                else:
                    box = [X(a) - 16, Y(g["co"]) - 34, X(b) + 16, Y(g["co"]) + 34]
                d.ellipse(box, outline=(0, 0, 0), width=4)
                texto = (f"claro {Lv:.2f} m → peralte {g['h']} cm" if g["h"] > 30
                         else f"claro {Lv:.2f} m · f_CM {g['fcm']:.1f}× (muro corto)")
                f = F(15, True)
                tw, th = d.textbbox((0, 0), texto, font=f)[2:]
                w, h = tw + 10, th + 8
                cx0, cy0 = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
                cands = [(cx0 - w / 2, box[3] + 4), (cx0 - w / 2, box[1] - h - 4), (box[2] + 6, cy0 - h / 2),
                         (box[0] - w - 6, cy0 - h / 2)]
                cands += [(x_ + dx, y_ + dy) for x_, y_ in cands[:4] for dx, dy in ((0, 30), (0, -30), (40, 0), (-40, 0),
                                                                                     (0, 60), (0, -60))]
                for cx, cy in cands:
                    bx = (cx - 2, cy - 2, cx + w + 2, cy + h + 2)
                    if libre(bx) and cx > M0 - 10:
                        break
                puestas.append(bx)
                d.rounded_rectangle([cx, cy, cx + w, cy + h], radius=4, fill=(255, 255, 255), outline=(0, 0, 0), width=2)
                d.text((cx + 5, cy + 3), texto, font=f, fill=(0, 0, 0))
        # tabla
        tx, ty = Wp + 10, 190
        n_ok = sum(1 for f_ in filas if f_[2])
        resumen.append((titulo, n_ok, len(filas) - n_ok))
        d.text((tx, ty), f"{titulo}: {n_ok} cumplen y {len(filas) - n_ok} no", font=F(24, True), fill=(20, 20, 20))
        ty += 44
        c1, c2, c3 = tx, tx + 175, tx + 260
        d.text((c1, ty), "Pier", font=F(16, True), fill=(80, 80, 80))
        d.text((c2, ty), "Ratio", font=F(16, True), fill=(80, 80, 80))
        d.text((c3, ty), "Por qué falla", font=F(16, True), fill=(80, 80, 80))
        ty += 26
        d.line([(tx, ty), (W - 30, ty)], fill=(180, 180, 180), width=1)
        ty += 6
        ancho = W - 40 - c3
        for name, e, ok, no, extra, nota in filas:
            fr, fb, fp = F(15), F(15, True), F(13)
            lineas = []
            for s in (no or [nota or "cumple"]):
                lineas += [(ln, fr, (30, 30, 30) if no else (90, 90, 90))
                           for ln in envolver(d, ("• " + s) if no else s, fr, ancho)]
            if extra:
                lineas += [(ln, fp, (90, 90, 90)) for ln in envolver(d, extra, fp, ancho)]
            d.text((c1, ty), name, font=fb, fill=e["col"])
            d.text((c2, ty), e["rtxt"], font=fb, fill=e["col"])
            for i, (ln, fnt, colr) in enumerate(lineas):
                d.text((c3, ty + i * 20), ln, font=fnt, fill=colr)
            ty += max(1, len(lineas)) * 20 + 8
            d.line([(tx, ty - 4), (W - 30, ty - 4)], fill=(232, 232, 232), width=1)
        if any(e["concreto"] for _, e, *_ in filas):
            ty += 10
            nota_c = ("Muros de concreto: revisión rápida NTC-Concreto 2023, ductilidad baja. Cortante con la ec. "
                      "6.5.5.3.3 (α_c = 0.8, p_t = 0.0025, F_R = 0.75). Flexocompresión por compatibilidad con P_u "
                      "mín. y máx., armado supuesto #3 @ 18 cm en una capa al centro (6.5.7.1.3: el cortante medio es "
                      "menor que 0.6·√f'c) y 4 #3 en cada castillo extremo. La hoja de muros todavía los revisa como "
                      "tabique: ese resultado no aplica.")
            for ln in envolver(d, nota_c, F(14), W - 40 - tx):
                d.text((tx, ty), ln, font=F(14), fill=(90, 90, 90))
                ty += 19
        paginas.append(img)
    paginas[0].save(salida, save_all=True, append_images=paginas[1:], resolution=150)
    for p in resumen:
        print("%s: %d cumplen, %d no" % p)


if __name__ == "__main__":
    main(*sys.argv[1:5])
