#!/usr/bin/env python3
"""Revisa como quedan conectados los muros en el modelo original y en el
redibujado: esquinas de tablero, castillos, dalas/losas y apoyos.
Uso: python3 verificar_conexiones.py ORIGINAL.e2k NUEVO.e2k"""
import math
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL  # noqa: E402


def on_seg(p, a, b, tol=0.003):
    ab = [b[i] - a[i] for i in range(3)]
    ap = [p[i] - a[i] for i in range(3)]
    L2 = sum(v * v for v in ab)
    if L2 == 0:
        return False
    t = sum(ab[i] * ap[i] for i in range(3)) / L2
    if t < -1e-9 or t > 1 + 1e-9:
        return False
    return math.dist(p, [a[i] + t * ab[i] for i in range(3)]) < tol


def in_poly(p, P, tol=0.003):
    for i in range(len(P)):
        if on_seg((p[0], p[1], 0), (*P[i], 0), (*P[(i + 1) % len(P)], 0), tol):
            return True
    x, y = p
    inside = False
    for i in range(len(P)):
        (x1, y1), (x2, y2) = P[i], P[(i + 1) % len(P)]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            inside = not inside
    return inside


def analyze(path):
    txt = open(path, encoding="utf-8", newline="").read()
    M = Model(txt)
    base = M.order[-1]
    R = lambda p: tuple(round(v, 4) for v in p)

    # paneles de muro en 3D
    panels = []
    for idx, n, s, sec, r in M.area_assign:
        if not M.is_wall_section(sec):
            continue
        P = []
        for p, f in M.areas[n][1]:
            x, y, dz = M.points[p]
            P.append((x, y, M.elev[M.below(s, f)] - dz))
        panels.append((n, s, P))
    walls, _ = M.wall_pieces()

    # barras
    frames = []
    for ln, p in M.columns.items():
        x, y, _ = M.points[p]
        for s, sec in M.line_assign.get(ln, []):
            frames.append(("castillo", ln, s, (x, y, M.elev[M.below(s)]), (x, y, M.elev[s])))
    for ln, (p1, p2) in M.beams.items():
        a, b = M.points[p1], M.points[p2]
        for s, sec in M.line_assign.get(ln, []):
            frames.append(("viga", ln, s, (a[0], a[1], M.elev[s] - a[2]), (b[0], b[1], M.elev[s] - b[2])))
    frame_nodes = {R(f[3]) for f in frames} | {R(f[4]) for f in frames}

    # losas y aberturas
    floors, opens = [], []
    for idx, n, s, sec, r in M.area_assign:
        P = [M.points[p][:2] for p, f in M.areas[n][1]]
        if sec is not None and not M.is_wall_section(sec):
            floors.append((M.elev[s], P))
        elif 'OPENING "Yes"' in r:
            opens.append((M.elev[s], P))

    def on_floor(p):
        return any(abs(p[2] - z) < TOL and in_poly(p[:2], P) for z, P in floors)

    supports = set()
    for m in re.finditer(r'POINTASSIGN\s+"([^"]+)"\s+"([^"]+)"\s+RESTRAINT "U', txt):
        x, y, dz = M.points[m[1]]
        supports.add(R((x, y, M.elev[m[2]] - dz)))

    # 1) esquinas de tablero
    use = Counter(R(c) for n, s, P in panels for c in set(R(q) for q in P))
    corners = set(use)
    free, hanging = [], []
    for c in corners:
        shared = use[c] >= 2
        fr = c in frame_nodes or any(on_seg(c, f[3], f[4]) for f in frames)
        fl = on_floor(c)
        sp = c in supports
        if not (shared or fr or fl or sp):
            free.append(c)
        # nudo colgado: cae a media orilla de otro tablero sin ser su esquina
        for n, s, P in panels:
            if c in {R(q) for q in P}:
                continue
            if any(on_seg(c, P[i], P[(i + 1) % 4]) for i in range(4)):
                hanging.append(c)
                break

    # 2) castillos con muro: sus nudos de arriba y abajo son esquina de algun tablero
    cast_ok = cast_bad = 0
    bad_cast = []
    for kind, ln, s, a, b in frames:
        if kind != "castillo":
            continue
        has_wall = any(abs((a[0] if ax == "X" else a[1]) - co) <= TOL and st == s
                       and any(abs(p["h0"] - (a[1] if ax == "X" else a[0])) <= TOL or abs(p["h1"] - (a[1] if ax == "X" else a[0])) <= TOL
                               or p["h0"] < (a[1] if ax == "X" else a[0]) < p["h1"] for p in pcs)
                       for (st, ax, co), pcs in walls.items())
        if not has_wall:
            continue
        if R(a) in corners and R(b) in corners:
            cast_ok += 1
        else:
            cast_bad += 1
            bad_cast.append((ln, s))

    # 3) orilla de arriba y de abajo de cada tablero: sobre dala o losa
    def edge_cover(z_level, story_beams, h0, h1, axis, coord):
        L = h1 - h0
        iv = []
        for kind, ln, s, a, b in story_beams:
            if abs(a[2] - z_level) > TOL:
                continue
            if axis == "X" and abs(a[0] - coord) <= TOL and abs(b[0] - coord) <= TOL:
                iv.append(sorted((a[1], b[1])))
            if axis == "Y" and abs(a[1] - coord) <= TOL and abs(b[1] - coord) <= TOL:
                iv.append(sorted((a[0], b[0])))
        cov = 0.0
        for lo, hi in sorted(iv):
            cov += max(0.0, min(hi, h1) - max(lo, h0))
        cov = min(cov, L)
        # lo que no tiene dala: revisar si esta sobre losa (punto medio de cada 10 cm)
        n = max(1, int(L / 0.1))
        slab = 0.0
        for i in range(n):
            h = h0 + (i + 0.5) * L / n
            p = (coord, h, z_level) if axis == "X" else (h, coord, z_level)
            if on_floor(p):
                slab += L / n
        return L, cov, slab

    stats = defaultdict(float)
    beams = [f for f in frames if f[0] == "viga"]
    free_edges = []
    for (s, axis, coord), pcs in walls.items():
        zt, zb = M.elev[s], M.elev[M.below(s)]
        for p in pcs:
            for which, z in (("arriba", p["z1"]), ("abajo", p["z0"])):
                if abs(z - zt) > TOL and abs(z - zb) > TOL:
                    continue                     # orilla intermedia (antepecho/dintel)
                if which == "abajo" and M.below(s) == base:
                    stats["base"] += p["h1"] - p["h0"]
                    continue
                L, cov, slab = edge_cover(z, beams, p["h0"], p["h1"], axis, coord)
                stats["orilla total"] += L
                stats["con dala"] += cov
                stats["con losa"] += slab
                if cov < L - 0.01 and slab < L - 0.01:
                    free_edges.append((p["name"], s, which, f"{axis}={coord:g}", round(L - max(cov, slab), 2)))
    return dict(paneles=len(panels), esquinas=len(corners), sueltas=free, colgados=len(set(hanging)),
                cast_ok=cast_ok, cast_bad=bad_cast, stats=dict(stats), orillas_libres=free_edges)


A = analyze(sys.argv[1])
B = analyze(sys.argv[2])
row = lambda t, a, b: print(f"{t:62s}{a!s:>12s}{b!s:>12s}")
row("", "original", "redibujado")
row("tableros de muro", A["paneles"], B["paneles"])
row("esquinas de tablero", A["esquinas"], B["esquinas"])
row("esquinas sueltas (sin otro tablero, barra, losa ni apoyo)", len(A["sueltas"]), len(B["sueltas"]))
row("nudos colgados (esquina a media orilla de otro tablero)", A["colgados"], B["colgados"])
row("castillos con muro: nudo de arriba y abajo en esquina de muro", A["cast_ok"], B["cast_ok"])
row("castillos con muro que NO comparten esos nudos", len(A["cast_bad"]), len(B["cast_bad"]))
for k in ("orilla total", "con dala", "con losa"):
    row(f"orillas de muro en losa/entrepiso (m): {k}", f"{A['stats'].get(k, 0):.1f}", f"{B['stats'].get(k, 0):.1f}")
row("orillas de muro en la base (m)", f"{A['stats'].get('base', 0):.1f}", f"{B['stats'].get('base', 0):.1f}")
row("orillas sin dala ni losa", len(A["orillas_libres"]), len(B["orillas_libres"]))
for x in B["sueltas"]:
    print("   esquina suelta (nuevo):", x)
for x in B["cast_bad"]:
    print("   castillo sin nudo compartido (nuevo):", x)
for x in B["orillas_libres"]:
    print("   orilla libre (nuevo):", x)
