"""Utilidades comunes para editar el plano DXF de Casa Rocha con ezdxf."""
import math
import ezdxf
from ezdxf.math import Vec2

# Override de cotas idéntico al que traen las cotas del plano original
DIM_OVR = {
    'dimtxt': 0.17, 'dimfxl': 1.0, 'dimjogang': 0.785398163397,
    'dimpost': '<>m', 'dimdsep': 46, 'dimzin': 0, 'dimtix': 1,
    'dimtfill': 1, 'dimrnd': 0.001,
}


class Cad:
    def __init__(self, path):
        self.doc = ezdxf.readfile(path)
        self.msp = self.doc.modelspace()
        self.reindex()

    def reindex(self):
        self.H = {e.dxf.handle: e for e in self.msp}

    def e(self, h):
        return self.H[h]

    def kill(self, *handles):
        for h in handles:
            ent = self.H.get(h)
            if ent is not None and ent.is_alive:
                self.msp.delete_entity(ent)

    # ---------------- geometría ----------------
    def line(self, p1, p2, layer, **kw):
        a = {'layer': layer}
        a.update(kw)
        return self.msp.add_line(tuple(p1)[:2], tuple(p2)[:2], dxfattribs=a)

    def pline(self, pts, layer, closed=False, block=None, **kw):
        a = {'layer': layer}
        a.update(kw)
        lay = block if block is not None else self.msp
        return lay.add_lwpolyline([tuple(p)[:2] for p in pts], close=closed, dxfattribs=a)

    def rect(self, x0, y0, x1, y1, layer, block=None, **kw):
        return self.pline([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], layer, closed=True, block=block, **kw)

    def circle(self, c, r, layer, block=None, **kw):
        a = {'layer': layer}
        a.update(kw)
        lay = block if block is not None else self.msp
        return lay.add_circle(tuple(c)[:2], r, dxfattribs=a)

    def arc(self, c, r, a0, a1, layer, block=None, **kw):
        a = {'layer': layer}
        a.update(kw)
        lay = block if block is not None else self.msp
        return lay.add_arc(tuple(c)[:2], r, a0, a1, dxfattribs=a)

    def mtext(self, text, ins, h, layer='007_Textos', style='OpenSans', attach=1, width=None, rot=0.0):
        a = {'layer': layer, 'style': style, 'char_height': h, 'attachment_point': attach}
        if width:
            a['width'] = width
        m = self.msp.add_mtext(text, dxfattribs=a)
        m.set_location(tuple(ins)[:2] + (0,), rotation=rot, attachment_point=attach)
        return m

    def _norm_dim(self, d, layer):
        """Deja el bloque de la cota igual que las cotas originales (color y grosor POR BLOQUE)."""
        dim = d.dimension
        blk = self.doc.blocks.get(dim.dxf.geometry)
        for ent in blk:
            if ent.dxftype() == 'POINT':
                continue
            ent.dxf.layer = layer
            ent.dxf.color = 0
            if ent.dxftype() in ('LINE', 'INSERT'):
                ent.dxf.lineweight = -2
        return dim

    def ovr_de(self, handle):
        """Override de estilo de una cota original (para rehacerla idéntica)."""
        return dict(self.H[handle].override().dimstyle_attribs)

    def hdim(self, x1, x2, y_def, y_line, layer='008_Cotas', ovr=None, base=None):
        """Cota horizontal de x1 a x2 (x1<x2) con línea de cota en y_line."""
        o = dict(base if base is not None else DIM_OVR)
        o.update(ovr or {})
        d = self.msp.add_linear_dim(base=(x1, y_line), p1=(x1, y_def), p2=(x2, y_def), angle=0,
                                    dimstyle='EZDXF', override=o, dxfattribs={'layer': layer})
        d.render()
        return self._norm_dim(d, layer)

    def vdim(self, y1, y2, x_def, x_line, layer='008_Cotas', ovr=None, base=None):
        o = dict(base if base is not None else DIM_OVR)
        o.update(ovr or {})
        d = self.msp.add_linear_dim(base=(x_line, y1), p1=(x_def, y1), p2=(x_def, y2), angle=90,
                                    dimstyle='EZDXF', override=o, dxfattribs={'layer': layer})
        d.render()
        return self._norm_dim(d, layer)

    def save(self, path):
        self.doc.saveas(path)


def pts_of(e):
    t = e.dxftype()
    if t == 'LINE':
        return [Vec2(e.dxf.start), Vec2(e.dxf.end)]
    if t == 'LWPOLYLINE':
        return [Vec2(p[0], p[1]) for p in e.get_points()]
    if t in ('CIRCLE', 'ARC'):
        return [Vec2(e.dxf.center)]
    if t in ('INSERT', 'MTEXT', 'TEXT'):
        return [Vec2(e.dxf.insert)]
    return []


def arc_pts(c, r, a0, a1, n):
    """Puntos de un arco (ángulos en grados)."""
    return [(c[0] + r * math.cos(math.radians(a0 + (a1 - a0) * i / n)),
             c[1] + r * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]
