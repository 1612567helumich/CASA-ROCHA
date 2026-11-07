#!/usr/bin/env python3
"""Redibuja los muros de un .e2k de ETABS para que cada tablero quede
delimitado por ejes de castillo (y por los cortes de dala/dintel/antepecho).

Que hace, por cada eje de muro y por cada entrepiso:
  1. Lleva al eje del castillo toda orilla de muro que quede a <= SNAP de un
     castillo (las "tiras" junto a los huecos). No mueve orillas que esten en
     un cruce con otro muro.
  2. Parte cada pedazo de muro solo en los ejes de castillo que le caen
     adentro. No agrega cortes a lo alto: antepechos, dinteles y macizos
     quedan partidos a la altura que ya tenian en el modelo.
  3. Cada tablero hereda seccion, pier label y opciones de malla del pedazo
     original del que sale. Puntos nuevos solo donde hacen falta.

No toca: castillos, dalas, losas, cargas, combinaciones ni nada fuera de
las secciones POINT COORDINATES, AREA CONNECTIVITIES, POINT ASSIGNS
(solo limpia puntos que quedan huerfanos) y AREA ASSIGNS de muros.

Uso:  python3 redibujar_muros.py ENTRADA.e2k SALIDA.e2k [reporte.json]
"""
import json
import re
import sys
from collections import defaultdict

SNAP = 0.10   # m. Orilla de muro a <= SNAP de un eje de castillo se lleva al castillo
TOL = 0.002   # m. Tolerancia geometrica
NEWLINE = "\r\n"


def fmt(v):
    s = f"{v:.6f}".rstrip("0").rstrip(".")
    return "0" if s in ("-0", "") else s


# --------------------------------------------------------------------------
# Lectura
# --------------------------------------------------------------------------
class Model:
    def __init__(self, text):
        self.lines = text.split(NEWLINE)
        self.parse()

    def section_of(self, idx):
        sec = None
        for i in range(idx, -1, -1):
            if self.lines[i].startswith("$ "):
                return self.lines[i]
        return sec

    def parse(self):
        L = self.lines
        self.stories = []          # top-down
        self.points = {}           # name -> (x, y, dz)
        self.point_line = {}       # name -> line idx
        self.columns = {}          # line name -> plan point
        self.beams = {}            # line name -> (p1, p2)
        self.line_assign = defaultdict(list)   # name -> [(story, section)]
        self.areas = {}            # name -> (kind, [(pt, flag)], idx)
        self.area_assign = []      # (idx, name, story, section or None, rest)
        self.shell_type = {}       # shell prop -> PROPTYPE
        self.sec_idx = {}          # "$ HEADER" -> idx
        for i, ln in enumerate(L):
            if ln.startswith("$ "):
                self.sec_idx[ln.strip()] = i
            m = re.match(r'\s*STORY "([^"]+)"\s+(HEIGHT|ELEV)\s+(\S+)', ln)
            if m:
                self.stories.append((m[1], m[2], float(m[3])))
                continue
            m = re.match(r'\s*POINT "([^"]+)"\s+(\S+)\s+(\S+)(?:\s+(\S+))?\s*$', ln)
            if m:
                self.points[m[1]] = (float(m[2]), float(m[3]), float(m[4] or 0))
                self.point_line[m[1]] = i
                continue
            m = re.match(r'\s*LINE\s+"([^"]+)"\s+(\w+)\s+"([^"]+)"\s+"([^"]+)"', ln)
            if m:
                if m[2] == "COLUMN":
                    self.columns[m[1]] = m[3]
                else:
                    self.beams[m[1]] = (m[3], m[4])
                continue
            m = re.match(r'\s*LINEASSIGN\s+"([^"]+)"\s+"([^"]+)"\s+SECTION "([^"]+)"', ln)
            if m:
                self.line_assign[m[1]].append((m[2], m[3]))
                continue
            m = re.match(r'\s*AREA "([^"]+)"\s+(\w+)\s+(\d+)\s+(.*)$', ln)
            if m:
                n = int(m[3])
                toks = re.findall(r'"([^"]+)"|(\S+)', m[4])
                pts = [a for a, _ in toks[:n]]
                flags = [int(b) for _, b in toks[n:2 * n]]
                self.areas[m[1]] = (m[2], list(zip(pts, flags)), i)
                continue
            m = re.match(r'(\s*AREAASSIGN\s+)"([^"]+)"\s+"([^"]+)"(.*)$', ln)
            if m:
                sec = re.match(r'\s*SECTION "([^"]+)"', m[4])
                self.area_assign.append((i, m[2], m[3], sec[1] if sec else None, m[4]))
                continue
            m = re.match(r'\s*SHELLPROP\s+"([^"]+)"\s+PROPTYPE\s+"([^"]+)"', ln)
            if m:
                self.shell_type[m[1]] = m[2]
        # story elevations
        self.order = [s[0] for s in self.stories]
        self.elev = {}
        z = None
        for name, kind, v in reversed(self.stories):
            z = v if kind == "ELEV" else z + v
            self.elev[name] = z

    def below(self, story, k=1):
        return self.order[self.order.index(story) + k]

    def is_wall_section(self, sec):
        return sec is not None and self.shell_type.get(sec) == "Wall"

    def wall_pieces(self):
        """Devuelve {(story, eje, coord): [pieza]} con piezas rectangulares en (h, z)."""
        out = defaultdict(list)
        skipped = []
        for idx, name, story, sec, rest in self.area_assign:
            if not self.is_wall_section(sec):
                continue
            kind, conn, _ = self.areas[name]
            P = []
            for p, f in conn:
                x, y, dz = self.points[p]
                P.append((x, y, self.elev[self.below(story, f)] - dz))
            xs = [p[0] for p in P]
            ys = [p[1] for p in P]
            if max(xs) - min(xs) < TOL:
                axis, coord, H = "X", sum(xs) / len(xs), ys
            elif max(ys) - min(ys) < TOL:
                axis, coord, H = "Y", sum(ys) / len(ys), xs
            else:
                skipped.append((name, story, "muro no alineado a X o Y"))
                continue
            Z = [p[2] for p in P]
            hs = sorted(set(round(h, 4) for h in H))
            zs = sorted(set(round(z, 4) for z in Z))
            if len(conn) != 4 or (max(hs) - min(hs) < TOL) or not _two_clusters(H) or not _two_clusters(Z):
                skipped.append((name, story, "muro no rectangular"))
                continue
            # orientacion (signo del area en el plano h-z)
            s = 0.0
            for i in range(4):
                h1, z1 = H[i], Z[i]
                h2, z2 = H[(i + 1) % 4], Z[(i + 1) % 4]
                s += h1 * z2 - h2 * z1
            pier = re.search(r'PIER\s+"([^"]+)"', rest)
            out[(story, axis, round(coord, 3))].append(dict(
                name=name, story=story, axis=axis, coord=coord,
                h0=min(H), h1=max(H), z0=min(Z), z1=max(Z),
                ccw=s > 0, rest=rest, section=sec, pier=pier[1] if pier else None,
                flags=[f for _, f in conn]))
        return out, skipped

    def castillos(self):
        """{story: [(x, y, plan_point, line_name)]}"""
        out = defaultdict(list)
        for ln, p in self.columns.items():
            for story, sec in self.line_assign.get(ln, []):
                x, y, _ = self.points[p]
                out[story].append((x, y, p, ln))
        return out


def _two_clusters(vals):
    v = sorted(vals)
    lo = [a for a in v if a - v[0] < TOL]
    hi = [a for a in v if v[-1] - a < TOL]
    return len(lo) + len(hi) == len(v) and v[-1] - v[0] > TOL


def cluster(vals):
    """Agrupa valores a menos de TOL y devuelve representantes ordenados."""
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= TOL:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(c) / len(c) for c in out]


# --------------------------------------------------------------------------
# Redibujo de un eje de muro
# --------------------------------------------------------------------------
def rebuild_line(key, pieces, cast_h, perp_junctions, zb, zt, log):
    story, axis, coord = key
    # 1) orillas en h -> llevar al castillo si esta a <= SNAP
    edges = cluster([p["h0"] for p in pieces] + [p["h1"] for p in pieces])
    cast_h = cluster(cast_h)
    snap = {}
    for e in edges:
        if any(abs(e - c) <= TOL for c in cast_h):
            snap[e] = min(cast_h, key=lambda c: abs(c - e))
            continue
        near = [c for c in cast_h if abs(c - e) <= SNAP]
        if not near:
            snap[e] = e
            continue
        c = min(near, key=lambda c: abs(c - e))
        if any(abs(e - j) <= TOL for j in perp_junctions):
            snap[e] = e
            log["no_movidas_por_cruce"].append(dict(eje=_eje(axis, coord), piso=story, orilla=round(e, 4), castillo=round(c, 4)))
            continue
        snap[e] = c
        log["orillas_movidas"].append(dict(eje=_eje(axis, coord), piso=story, de=round(e, 4), a=round(c, 4), mueve_cm=round(abs(c - e) * 100, 1)))

    def S(h):
        e = min(edges, key=lambda e: abs(e - h))
        return snap[e]

    sp = []
    for p in pieces:
        h0, h1 = S(p["h0"]), S(p["h1"])
        if h1 - h0 > TOL:
            sp.append(dict(p, h0=h0, h1=h1))
        else:
            log["pedazos_eliminados"].append(dict(eje=_eje(axis, coord), piso=story, muro=p["name"], ancho_cm=round((p["h1"] - p["h0"]) * 100, 1)))
    # monotonia
    for e1, e2 in zip(edges, edges[1:]):
        if snap[e2] < snap[e1] - TOL:
            raise RuntimeError(f"Orden invertido al mover orillas en {key}")

    # 2) cada pedazo se parte SOLO en los ejes de castillo que le caen adentro;
    #    los cortes a lo alto quedan exactamente como venian en el modelo
    cells = []
    for p in sp:
        hb = cluster([p["h0"], p["h1"]] + [c for c in cast_h if p["h0"] + TOL < c < p["h1"] - TOL])
        for a, b in zip(hb, hb[1:]):
            cells.append(dict(p, h0=a, h1=b, parent=p["name"]))
    return cells


def _eje(axis, coord):
    return f"{axis}={fmt(coord)}"


# --------------------------------------------------------------------------
# Principal
# --------------------------------------------------------------------------
def main(src, dst, report_path=None):
    text = open(src, encoding="utf-8", newline="").read()
    M = Model(text)
    groups, skipped = M.wall_pieces()
    cast = M.castillos()
    log = defaultdict(list)
    log["omitidos"] = [dict(muro=a, piso=b, motivo=c) for a, b, c in skipped]

    # muros perpendiculares por piso para detectar cruces
    by_story_axis = defaultdict(list)
    for (story, axis, coord), pcs in groups.items():
        for p in pcs:
            by_story_axis[(story, axis)].append(p)

    new_cells = {}
    for key, pcs in sorted(groups.items()):
        story, axis, coord = key
        zt, zb = M.elev[story], M.elev[M.below(story)]
        ch = []
        for x, y, p, ln in cast[story]:
            off, h = (x, y) if axis == "X" else (y, x)
            if abs(off - coord) <= TOL:
                ch.append(h)
        other = "Y" if axis == "X" else "X"
        junc = [q["coord"] for q in by_story_axis[(story, other)] if q["h0"] - TOL <= coord <= q["h1"] + TOL]
        new_cells[key] = rebuild_line(key, pcs, ch, junc, zb, zt, log)

    # ---------------- escribir ----------------
    L = list(M.lines)
    rebuilt_assign_idx = {idx for idx, name, story, sec, rest in M.area_assign if M.is_wall_section(sec)}
    rebuilt_names = {name for idx, name, story, sec, rest in M.area_assign if idx in rebuilt_assign_idx}
    # un AREA solo se borra si todas sus asignaciones son de muro
    keep_names = {name for idx, name, story, sec, rest in M.area_assign if idx not in rebuilt_assign_idx}
    skipped_names = {s[0] for s in skipped}
    drop_area = {n for n in rebuilt_names if n not in keep_names and n not in skipped_names}
    drop_assign = {idx for idx, name, story, sec, rest in M.area_assign
                   if idx in rebuilt_assign_idx and name not in skipped_names}

    # puntos: reutilizar si existe (x, y, dz) igual
    pt_list = [(n, x, y, dz) for n, (x, y, dz) in M.points.items()]

    def find_pt(x, y, dz):
        best = None
        for n, a, b, c in pt_list:
            if abs(a - x) <= 1e-3 and abs(b - y) <= 1e-3 and abs(c - dz) <= 1e-3:
                d = abs(a - x) + abs(b - y) + abs(c - dz)
                if best is None or d < best[0]:
                    best = (d, n)
        return best[1] if best else None

    num = [int(n) for n in M.points if n.isdigit()]
    next_pt = [max(num) + 1 if num else 1]
    new_point_lines = []

    def get_pt(x, y, dz):
        n = find_pt(x, y, dz)
        if n:
            return n
        n = str(next_pt[0])
        next_pt[0] += 1
        pt_list.append((n, x, y, dz))
        tail = f" {fmt(dz)}" if abs(dz) > 1e-9 else " "
        new_point_lines.append(f'  POINT "{n}"  {fmt(x)} {fmt(y)}{tail}')
        log["puntos_nuevos"].append(n)
        return n

    existing_area_names = set(M.areas)
    k = 1
    new_area_lines, new_assign_lines = [], []
    summary = []
    for key in sorted(new_cells):
        story, axis, coord = key
        zt, zb = M.elev[story], M.elev[M.below(story)]
        for c in sorted(new_cells[key], key=lambda c: (c["h0"], c["z0"])):
            while f"WC{k}" in existing_area_names:
                k += 1
            name = f"WC{k}"
            k += 1
            corners = [(c["h0"], c["z0"]), (c["h1"], c["z0"]), (c["h1"], c["z1"]), (c["h0"], c["z1"])]
            if not c["ccw"]:
                corners = [(c["h1"], c["z0"]), (c["h0"], c["z0"]), (c["h0"], c["z1"]), (c["h1"], c["z1"])]
            names, flags = [], []
            for h, z in corners:
                x, y = (c["coord"], h) if axis == "X" else (h, c["coord"])
                # usar coordenada exacta del eje para que coincida con los puntos existentes
                if abs(z - zb) <= TOL:
                    names.append(get_pt(x, y, 0.0))
                    flags.append(1)
                else:
                    names.append(get_pt(x, y, round(zt - z, 6)))
                    flags.append(0)
            new_area_lines.append(
                f'  AREA "{name}"  PANEL  4  ' + "  ".join(f'"{n}"' for n in names) + "  " + "  ".join(str(f) for f in flags) + "  ")
            new_assign_lines.append(f'  AREAASSIGN  "{name}"  "{story}"{c["rest"]}')
            summary.append(dict(nuevo=name, piso=story, eje=_eje(axis, coord), h0=round(c["h0"], 4), h1=round(c["h1"], 4),
                                z0=round(c["z0"], 3), z1=round(c["z1"], 3), viene_de=c["parent"], pier=c["pier"]))

    # construir archivo
    out = []
    sec = None
    pending_insert = {}
    # localizar fin de cada seccion (ultima linea no vacia antes del siguiente "$ ")
    hdrs = sorted((i, h) for h, i in M.sec_idx.items())
    end_of = {}
    for j, (i, h) in enumerate(hdrs):
        nxt = hdrs[j + 1][0] if j + 1 < len(hdrs) else len(L)
        last = i
        for t in range(i + 1, nxt):
            if L[t].strip():
                last = t
        end_of[h] = last
    insert_after = {
        end_of["$ POINT COORDINATES"]: new_point_lines,
        end_of["$ AREA CONNECTIVITIES"]: new_area_lines,
        end_of["$ AREA ASSIGNS"]: new_assign_lines,
    }
    drop_area_idx = {M.areas[n][2] for n in drop_area}
    for i, ln in enumerate(L):
        if i in drop_area_idx or i in drop_assign:
            pass
        else:
            out.append(ln)
        if i in insert_after:
            out.extend(insert_after[i])
    new_text = NEWLINE.join(out)

    # limpiar puntos que quedaron huerfanos (solo los que usaban los muros borrados)
    used_before = {p for n in drop_area for p, _ in M.areas[n][1]}
    M2 = Model(new_text)
    used = set()
    for n, (kind, conn, _) in M2.areas.items():
        used.update(p for p, _ in conn)
    for ln_name, p in M2.columns.items():
        used.add(p)
    for p1, p2 in M2.beams.values():
        used.update((p1, p2))
    used.update(re.findall(r'POINTLOAD\s+"([^"]+)"', new_text))
    userjoint = set(re.findall(r'POINTASSIGN\s+"([^"]+)"\s+"[^"]+"[^\r\n]*USERJOINT\s+"Yes"', new_text))
    orphans = {p for p in used_before if p not in used and p not in userjoint}
    if orphans:
        keep = []
        for ln in new_text.split(NEWLINE):
            m = re.match(r'\s*(?:POINT|POINTASSIGN)\s+"([^"]+)"', ln)
            if m and m[1] in orphans:
                continue
            keep.append(ln)
        new_text = NEWLINE.join(keep)
    log["puntos_huerfanos_eliminados"] = sorted(orphans, key=lambda s: (len(s), s))

    with open(dst, "w", encoding="utf-8", newline="") as f:
        f.write(new_text)

    rep = dict(
        muros_originales=len(drop_assign),
        tableros_nuevos=len(new_assign_lines),
        orillas_movidas=log["orillas_movidas"],
        no_movidas_por_cruce=log["no_movidas_por_cruce"],
        pedazos_eliminados=log["pedazos_eliminados"],
        puntos_nuevos=len(log["puntos_nuevos"]),
        puntos_huerfanos_eliminados=log["puntos_huerfanos_eliminados"],
        omitidos=log["omitidos"],
        tableros=summary,
    )
    if report_path:
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(rep, f, ensure_ascii=False, indent=1)
    return rep


if __name__ == "__main__":
    r = main(*sys.argv[1:4])
    print(f"muros originales: {r['muros_originales']}  ->  tableros nuevos: {r['tableros_nuevos']}")
    print(f"orillas movidas al castillo: {len(r['orillas_movidas'])}   no movidas por cruce: {len(r['no_movidas_por_cruce'])}")
    print(f"puntos nuevos: {r['puntos_nuevos']}   puntos huerfanos eliminados: {len(r['puntos_huerfanos_eliminados'])}")
    print(f"omitidos: {r['omitidos']}")
