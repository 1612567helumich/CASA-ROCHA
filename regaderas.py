# Regaderas de muro a muro: 1.45 (ancho libre del baño) x 1.05 de fondo
import ezdxf, math
F = '/home/user/CASA-ROCHA/001_PLANO_CROCHA_R1_4.dxf'
d = ezdxf.readfile(F); m = d.modelspace()
L, P = 1.45 / 2, 1.05 / 2
name = 'MOB_REGADERA_145x105'
b = d.blocks.new(name)
a = {'layer': '0'}
b.add_lwpolyline([(L, -P), (L, P), (-L, P), (-L, -P)], close=True, dxfattribs=a)
b.add_lwpolyline([(-L, -P), (L, P)], dxfattribs=a)
b.add_lwpolyline([(-L, P), (L, -P)], dxfattribs=a)
b.add_lwpolyline([(0.05 * math.cos(t * math.pi / 32), 0.05 * math.sin(t * math.pi / 32)) for t in range(64)], close=True, dxfattribs=a)
b.add_lwpolyline([(L, -P), (L, -P + 0.02), (-L, -P + 0.02), (-L, -P)], close=True, dxfattribs=a)
# handle: (centro nuevo) — lado largo entre muros, fondo 1.05 desde el muro del fondo
target = {
    '52A6': (32.010 + P, (5.296 + 6.746) / 2),   # PB baño de entrada, muro fondo x=32.01
    '52B6': (39.985, 24.171 - P),               # PB baño de atrás, muro fondo y=24.171
    '52B8': (62.010 + P, 14.521),               # PN1 baño izq, muro fondo x=62.01
    '52BE': (70.710 - P, 14.521),               # PN1 baño der, muro fondo x=70.71
    '52E1': (67.061, 2.871 + P),                # PN1 baño recámara principal, muro fondo y=2.871
}
for h, (x, y) in target.items():
    e = d.entitydb[h]
    e.dxf.name = name
    e.dxf.insert = (x, y, 0)
    e.dxf.xscale = e.dxf.yscale = 1
d.saveas(F)
print('ok')
