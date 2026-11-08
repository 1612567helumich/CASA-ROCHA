#!/usr/bin/env python3
"""Parte cada viga del .e2k en todos los nudos donde le llega otra barra:
castillos del mismo piso y vigas del mismo nivel que terminan sobre ella o la
cruzan. Equivale a Edit > Edit Frames > Divide Frames en las intersecciones.

- Cada tramo hereda la asignacion de la viga (seccion, punto cardinal, malla).
- Cargas UNIFF se copian a cada tramo; cargas TRAPF (distancias relativas) se
  reparten por tramo con sus valores interpolados, en el mismo lugar exacto.
- Una viga asignada en varios pisos se parte por separado en cada piso.

Uso: python3 partir_vigas.py ENTRADA.e2k SALIDA.e2k
"""
import math
import re
import sys
from collections import defaultdict

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, NEWLINE, fmt  # noqa: E402

TOLP = 0.003   # m, distancia para considerar que un nudo cae sobre la viga


def seg_param(p, a, b):
    """t de la proyeccion de p sobre a-b y distancia al segmento (planta)."""
    dx, dy = b[0] - a[0], b[1] - a[1]
    L2 = dx * dx + dy * dy
    t = ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / L2
    q = (a[0] + t * dx, a[1] + t * dy)
    return t, math.dist(p, q)


def cross(a, b, c, d):
    """Interseccion propia de los segmentos a-b y c-d en planta (o None)."""
    r = (b[0] - a[0], b[1] - a[1])
    s = (d[0] - c[0], d[1] - c[1])
    den = r[0] * s[1] - r[1] * s[0]
    if abs(den) < 1e-12:
        return None
    t = ((c[0] - a[0]) * s[1] - (c[1] - a[1]) * s[0]) / den
    u = ((c[0] - a[0]) * r[1] - (c[1] - a[1]) * r[0]) / den
    if 1e-6 < t < 1 - 1e-6 and 1e-6 < u < 1 - 1e-6:
        return (a[0] + t * r[0], a[1] + t * r[1])
    return None


def main(src, dst):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    L = M.lines

    line_idx, assign_idx, load_idx = {}, defaultdict(list), defaultdict(list)
    for i, ln in enumerate(L):
        m = re.match(r'\s*LINE\s+"([^"]+)"\s+BEAM\s', ln)
        if m:
            line_idx[m[1]] = i
        m = re.match(r'\s*LINEASSIGN\s+"([^"]+)"\s+"([^"]+)"', ln)
        if m and m[1] in M.beams:
            assign_idx[(m[1], m[2])].append(i)
        m = re.match(r'\s*LINELOAD\s+"([^"]+)"\s+"([^"]+)"', ln)
        if m and m[1] in M.beams:
            load_idx[(m[1], m[2])].append(i)

    # vigas por piso con su z
    beams = []
    for b, (p1, p2) in M.beams.items():
        for story, sec in M.line_assign.get(b, []):
            A, B = M.points[p1], M.points[p2]
            beams.append((b, story, A[:2], B[:2], A[2], p1, p2))
    cols = defaultdict(list)
    for c, p in M.columns.items():
        for story, sec in M.line_assign.get(c, []):
            cols[story].append(M.points[p][:2])

    # puntos existentes para reutilizar
    pts = [(n, x, y, dz) for n, (x, y, dz) in M.points.items()]
    num = [int(n) for n in M.points if n.isdigit()]
    nxt = [max(num) + 1]
    new_points = []

    def get_pt(x, y, dz):
        best = None
        for n, a, b2, c in pts:
            if abs(a - x) <= 1e-3 and abs(b2 - y) <= 1e-3 and abs(c - dz) <= 1e-6:
                d = abs(a - x) + abs(b2 - y)
                if best is None or d < best[0]:
                    best = (d, n)
        if best:
            return best[1]
        n = str(nxt[0])
        nxt[0] += 1
        pts.append((n, x, y, dz))
        new_points.append(f'  POINT "{n}"  {fmt(x)} {fmt(y)}' + (f" {fmt(dz)}" if abs(dz) > 1e-9 else " "))
        return n

    multi = {b for b in M.beams if len(M.line_assign.get(b, [])) > 1}
    tag = lambda s: s.replace("N+", "N").replace(".00", "").replace(" ", "")
    replace = {}            # idx -> lista de lineas nuevas (o [] para borrar)
    report, check = [], []
    new_line_defs = defaultdict(list)    # viga original -> nuevas lineas LINE
    split_any = defaultdict(bool)
    for b, story, a, bb, dz, p1, p2 in beams:
        Lb = math.dist(a, bb)
        ts = []
        for c in cols[story]:
            t, d = seg_param(c, a, bb)
            if d < TOLP and 1e-6 < t < 1 - 1e-6 and min(t, 1 - t) * Lb > TOLP:
                ts.append(t)
        for b2, s2, a2, bb2, dz2, q1, q2 in beams:
            if b2 == b or s2 != story or abs(dz2 - dz) > 1e-6:
                continue
            for e in (a2, bb2):
                t, d = seg_param(e, a, bb)
                if d < TOLP and 1e-6 < t < 1 - 1e-6 and min(t, 1 - t) * Lb > TOLP:
                    ts.append(t)
            x = cross(a, bb, a2, bb2)
            if x:
                ts.append(seg_param(x, a, bb)[0])
        ts = sorted(ts)
        cl = []
        for t in ts:
            if not cl or (t - cl[-1]) * Lb > TOLP:
                cl.append(t)
        if not cl:
            continue
        split_any[b] = True
        cuts = [0.0] + cl + [1.0]
        # nombres y puntos de los tramos
        names, ends = [], []
        for k, (t0, t1) in enumerate(zip(cuts, cuts[1:]), 1):
            nm = f"{b}_{tag(story)}_{k}" if b in multi else f"{b}_{k}"
            P0 = p1 if t0 == 0 else get_pt(a[0] + t0 * (bb[0] - a[0]), a[1] + t0 * (bb[1] - a[1]), dz)
            P1 = p2 if t1 == 1 else get_pt(a[0] + t1 * (bb[0] - a[0]), a[1] + t1 * (bb[1] - a[1]), dz)
            names.append(nm)
            ends.append((t0, t1))
            new_line_defs[b].append(f'  LINE  "{nm}"  BEAM  "{P0}"  "{P1}"  0')
        # asignacion
        for i in assign_idx[(b, story)]:
            replace[i] = [L[i].replace(f'"{b}"', f'"{nm}"', 1) for nm in names]
        # cargas
        for i in load_idx[(b, story)]:
            ln = L[i]
            m = re.search(r'TYPE "(\w+)"', ln)
            new = []
            if m[1] == "UNIFF":
                new = [ln.replace(f'"{b}"', f'"{nm}"', 1) for nm in names]
                fv = float(re.search(r"FVAL (\S+)", ln)[1])
                check.append((b, story, fv * Lb, sum(fv * (t1 - t0) * Lb for t0, t1 in ends)))
            elif m[1] == "TRAPF":
                f1, f2 = float(re.search(r"FSTART (\S+)", ln)[1]), float(re.search(r"FEND (\S+)", ln)[1])
                r1, r2 = float(re.search(r"RDSTART (\S+)", ln)[1]), float(re.search(r"RDEND (\S+)", ln)[1])
                fat = lambda r: f1 + (f2 - f1) * (r - r1) / (r2 - r1) if r2 > r1 else f1
                before = (f1 + f2) / 2 * (r2 - r1) * Lb
                after = 0.0
                for nm, (t0, t1) in zip(names, ends):
                    o0, o1 = max(r1, t0), min(r2, t1)
                    if (o1 - o0) * Lb <= 1e-6:
                        continue
                    g0, g1 = fat(o0), fat(o1)
                    q0, q1 = (o0 - t0) / (t1 - t0), (o1 - t0) / (t1 - t0)
                    after += (g0 + g1) / 2 * (o1 - o0) * Lb
                    s = ln.replace(f'"{b}"', f'"{nm}"', 1)
                    s = re.sub(r"FSTART \S+", f"FSTART {fmt(round(g0, 6))}", s)
                    s = re.sub(r"FEND \S+", f"FEND {fmt(round(g1, 6))}", s)
                    s = re.sub(r"RDSTART \S+", f"RDSTART {fmt(round(q0, 6))}", s)
                    s = re.sub(r"RDEND \S+", f"RDEND {fmt(round(q1, 6))}", s)
                    new.append(s)
                check.append((b, story, before, after))
            else:
                raise SystemExit(f"Tipo de carga no soportado: {ln}")
            replace[i] = new
        report.append((b, story, len(names), round(Lb, 3)))

    # LINE: la original se borra si ya no le queda ninguna asignacion sin partir
    for b in split_any:
        remaining = [s for s, sec in M.line_assign[b] if not any(r[0] == b and r[1] == s for r in report)]
        i = line_idx[b]
        replace[i] = ([L[i]] if remaining else []) + new_line_defs[b]

    hdrs = sorted((i, h) for h, i in M.sec_idx.items())
    end_pts = None
    for j, (i, h) in enumerate(hdrs):
        if h == "$ POINT COORDINATES":
            end_pts = max(t for t in range(i, hdrs[j + 1][0]) if L[t].strip())
    out = []
    for i, ln in enumerate(L):
        out.extend(replace.get(i, [ln]))
        if i == end_pts:
            out.extend(new_points)
    with open(dst, "w", encoding="utf-8", newline="") as fh:
        fh.write(NEWLINE.join(out))

    # verificacion
    bad = [c for c in check if abs(c[2] - c[3]) > 1e-3 * max(1.0, abs(c[2]))]
    M2 = Model(NEWLINE.join(out))
    lens = defaultdict(float)
    for b2, (q1, q2) in M2.beams.items():
        base = b2.split("_")[0]
        for s, sec in M2.line_assign.get(b2, []):
            lens[(base, s)] += math.dist(M2.points[q1][:2], M2.points[q2][:2])
    bad_len = [(b, s) for b, s, a0, b0, dz, p1, p2 in beams if abs(lens[(b, s)] - math.dist(a0, b0)) > 1e-4]
    print(f"vigas partidas: {len(report)} de {len(beams)} asignaciones;"
          f" tramos nuevos: {sum(r[2] for r in report)}; puntos nuevos: {len(new_points)}")
    print(f"cargas repartidas: {len(check)}; con diferencia de carga total: {len(bad)}")
    print(f"vigas cuyo largo total cambio: {bad_len}")
    assert not bad and not bad_len


if __name__ == "__main__":
    main(*sys.argv[1:3])
