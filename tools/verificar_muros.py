#!/usr/bin/env python3
"""Compara el .e2k original contra el redibujado y verifica que no se haya
estropeado nada. Uso: python3 verificar_muros.py ORIGINAL.e2k NUEVO.e2k"""
import re
import sys
from collections import Counter, defaultdict

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from redibujar_muros import Model, TOL, NEWLINE  # noqa: E402

orig_txt = open(sys.argv[1], encoding="utf-8", newline="").read()
new_txt = open(sys.argv[2], encoding="utf-8", newline="").read()
A, B = Model(orig_txt), Model(new_txt)
ok = True


def check(cond, msg):
    global ok
    print(("OK    " if cond else "FALLA ") + msg)
    ok &= bool(cond)


# 1. Fuera de las 4 secciones tocadas, el archivo es identico
def strip(txt):
    keep, sec = [], None
    for ln in txt.split(NEWLINE):
        if ln.startswith("$ "):
            sec = ln.strip()
        if sec in ("$ POINT COORDINATES", "$ AREA CONNECTIVITIES", "$ AREA ASSIGNS", "$ POINT ASSIGNS"):
            continue
        keep.append(ln)
    return keep


check(strip(orig_txt) == strip(new_txt), "todo fuera de puntos/areas/asignaciones de area queda identico")
check(new_txt.count("\n") == new_txt.count("\r\n"), "saltos de linea CRLF como el original")
check(new_txt.rstrip().endswith("$ END OF MODEL FILE"), "archivo completo hasta END OF MODEL FILE")

# 2. Asignaciones que no son muro: identicas
def non_wall(M):
    return sorted((n, s, sec, r) for i, n, s, sec, r in M.area_assign if not M.is_wall_section(sec))


check(non_wall(A) == non_wall(B), "losas y aberturas de losa sin cambios")
pa = [ln for ln in orig_txt.split(NEWLINE) if ln.lstrip().startswith("POINTASSIGN")]
pb = [ln for ln in new_txt.split(NEWLINE) if ln.lstrip().startswith("POINTASSIGN")]
removed_pa = [ln for ln in pa if ln not in pb]
check(all(ln in pa for ln in pb), "no se agregaron asignaciones de punto")
print(f"      asignaciones de punto quitadas (puntos huerfanos de los muros viejos): {len(removed_pa)}")

# 3. Todos los puntos usados existen; sin areas duplicadas
used = {p for _, conn, _ in B.areas.values() for p, _ in conn}
check(used <= set(B.points), "todas las esquinas de area tienen POINT definido")
assigned = Counter(n for i, n, s, sec, r in B.area_assign)
check(all(n in B.areas for n in assigned), "toda asignacion de area apunta a un AREA existente")
check(all(n in assigned for n in B.areas), "todo AREA tiene asignacion")
geo = Counter()
for i, n, s, sec, r in B.area_assign:
    geo[(s, frozenset(B.areas[n][1]))] += 1
check(max(geo.values()) == 1, "sin tableros duplicados")
names = Counter(re.findall(r'(?m)^\s*POINT "([^"]+)"', new_txt))
check(max(names.values()) == 1, "nombres de punto unicos")

# 4. Geometria de muros
ga, _ = A.wall_pieces()
gb, skb = B.wall_pieces()
check(not skb, "todos los tableros nuevos son rectangulares y verticales")
cast = B.castillos()
inside, narrow, area_rows = [], [], []
for key in sorted(set(ga) | set(gb)):
    story, axis, coord = key
    pa_ = ga.get(key, [])
    pb_ = gb.get(key, [])
    a0 = sum((p["h1"] - p["h0"]) * (p["z1"] - p["z0"]) for p in pa_)
    a1 = sum((p["h1"] - p["h0"]) * (p["z1"] - p["z0"]) for p in pb_)
    area_rows.append((story, f"{axis}={coord}", a0, a1, len(pa_), len(pb_)))
    for x, y, pt, ln in cast[story]:
        off, h = (x, y) if axis == "X" else (y, x)
        if abs(off - coord) > TOL:
            continue
        for p in pb_:
            if p["h0"] + TOL < h < p["h1"] - TOL:
                inside.append((ln, story, key, p["name"]))
    for p in pb_:
        w = p["h1"] - p["h0"]
        if w < 0.30:
            narrow.append((story, f"{axis}={coord}", p["name"], round(w, 3), p["pier"]))
check(not inside, "ningun castillo queda en medio de un tablero")
for x in inside:
    print("      ", x)
ta = sum(r[2] for r in area_rows)
tb = sum(r[3] for r in area_rows)
print(f"      area de muro: antes {ta:.2f} m2, despues {tb:.2f} m2, diferencia {tb - ta:+.2f} m2 ({100 * (tb - ta) / ta:+.1f} %)")
for r in area_rows:
    if abs(r[3] - r[2]) > 0.01:
        print(f"         {r[0]} {r[1]:>12s}: {r[2]:6.2f} -> {r[3]:6.2f} m2 ({r[3] - r[2]:+.2f})")
print(f"      tableros de menos de 30 cm de ancho: {len(narrow)}")
for x in narrow:
    print("         ", x)

# 5. Pier labels: mismos nombres y misma area por pier salvo las tiras
def pier_area(g):
    d = defaultdict(float)
    for pcs in g.values():
        for p in pcs:
            d[p["pier"]] += (p["h1"] - p["h0"]) * (p["z1"] - p["z0"])
    return d


pa_, pb_ = pier_area(ga), pier_area(gb)
check(set(pa_) == set(pb_), f"mismos pier labels ({len([k for k in pa_ if k])} piers)")

# 6. Base: toda esquina de muro en la base tiene empotramiento
rest = {m[1] for m in re.finditer(r'POINTASSIGN\s+"([^"]+)"\s+"N 0\.00"\s+RESTRAINT "UX', new_txt)}
base_story = B.order[-2]
base_pts = set()
for i, n, s, sec, r in B.area_assign:
    if B.is_wall_section(sec) and s == base_story:
        for p, f in B.areas[n][1]:
            if f == 1:
                base_pts.add(p)
check(base_pts <= rest, f"las {len(base_pts)} esquinas de tablero en la base estan empotradas")
for p in sorted(base_pts - rest):
    print("       sin apoyo:", p, B.points[p])

# 7. Conectividad: esquinas de tablero compartidas
corner_use = Counter()
for i, n, s, sec, r in B.area_assign:
    if B.is_wall_section(sec):
        for p, f in B.areas[n][1]:
            corner_use[(p, B.below(s, f))] += 1
print(f"      esquinas de tablero: {len(corner_use)}")

# 8. Espacios muertos: lo que era muro y ya no es (y al reves) solo puede ser
#    una tira pegada a un castillo; y ningun tablero se encima con otro.
from redibujar_muros import SNAP, cluster  # noqa: E402


def cells(P, Q):
    hs = cluster([p[k] for p in P + Q for k in ("h0", "h1")])
    zs = cluster([p[k] for p in P + Q for k in ("z0", "z1")])
    for a, b in zip(hs, hs[1:]):
        for c, d in zip(zs, zs[1:]):
            yield a, b, c, d, (a + b) / 2, (c + d) / 2


def count_in(S, h, z):
    return sum(1 for p in S if p["h0"] - TOL < h < p["h1"] + TOL and p["z0"] - TOL < z < p["z1"] + TOL)


bad_gap, overlaps, strips = [], [], []
for key in sorted(set(ga) | set(gb)):
    story, axis, coord = key
    P, Q = ga.get(key, []), gb.get(key, [])
    ch = []
    for x, y, pt, ln in cast[story]:
        off, h = (x, y) if axis == "X" else (y, x)
        if abs(off - coord) <= TOL:
            ch.append(h)
    for a, b, c, d, hm, zm in cells(P, Q):
        n_old, n_new = count_in(P, hm, zm), count_in(Q, hm, zm)
        if n_new > 1:
            overlaps.append((story, axis, coord, a, b, c, d))
        if (n_old > 0) != (n_new > 0):
            touches = any(abs(a - c0) <= TOL or abs(b - c0) <= TOL for c0 in ch)
            rec = (story, f"{axis}={coord:g}", round(a, 4), round(b, 4), round(c, 2), round(d, 2),
                   "se quito" if n_old else "se agrego")
            if b - a <= SNAP + TOL and touches:
                strips.append(rec)
            else:
                bad_gap.append(rec)
check(not overlaps, "ningun tablero se encima con otro")
check(not bad_gap, "no hay huecos nuevos: todo lo que cambio es una tira pegada a un castillo")
for x in bad_gap:
    print("      ", x)
w = [r[3] - r[2] for r in strips]
print(f"      tiras que cambiaron: {len(strips)} rectangulos, ancho maximo {max(w) * 100:.1f} cm,"
      f" {sum(1 for r in strips if r[6] == 'se quito')} quitadas, {sum(1 for r in strips if r[6] == 'se agrego')} agregadas")

# 9. Nudos casi encimados (podrian dejar rendijas o elementos de 1 mm)
P3 = set()
for i, n, s, sec, r in B.area_assign:
    if B.is_wall_section(sec):
        for p, f in B.areas[n][1]:
            x, y, dz = B.points[p]
            P3.add((p, round(x, 5), round(y, 5), round(B.elev[B.below(s, f)] - dz, 5)))
P3 = sorted(P3, key=lambda t: t[1:])
near = []
for i, a in enumerate(P3):
    for b in P3[i + 1:]:
        if b[1] - a[1] > 0.01:
            break
        dd = ((a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2 + (a[3] - b[3]) ** 2) ** 0.5
        if 0 < dd < 0.01 and a[0] != b[0]:
            near.append((a, b, round(dd * 1000, 2)))
check(not near, "no hay nudos de muro a menos de 1 cm de otro nudo distinto")
for x in near:
    print("      ", x)

print("\nRESULTADO:", "TODO BIEN" if ok else "HAY FALLAS")
sys.exit(0 if ok else 1)
