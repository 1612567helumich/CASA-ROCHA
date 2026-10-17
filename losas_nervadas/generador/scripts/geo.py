"""Analisis geometrico: tableros, bordes, apoyos, vacios, nervaduras principales y areas tributarias."""
import ezdxf, json, math, re
from shapely.geometry import Polygon, LineString, Point, box
from shapely.ops import unary_union
import os
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + '/'
m=json.load(open(W+'work/model.json'))
TOL=0.02
LEVELS={'N+3.00':0.0,'N+6.00':14.0}   # desplazamiento en X del DXF de tributarias
# ---------- tableros desde DXF de areas tributarias ----------
dt=ezdxf.readfile(W+'in/c566697e-AREAS_TRIBUTARIAS_R3.2.dxf'); mt=dt.modelspace()
labels=[(t.dxf.text,(t.dxf.insert[0],t.dxf.insert[1])) for t in mt.query('TEXT[layer=="TRIB_TABLERO_M2"]')]
tabs=[]
for p in mt.query('LWPOLYLINE[layer=="TRIB_TABLERO"]'):
    pts=[(float(x),float(y)) for x,y,*_ in p.get_points()]
    if pts[0]==pts[-1]: pts=pts[:-1]
    poly=Polygon(pts)
    lvl='N+6.00' if min(x for x,y in pts)>12 else 'N+3.00'
    dx=LEVELS[lvl]
    # etiqueta: texto dentro del poligono (cerca de esquina sup izq)
    lab=None
    for t,(x,y) in labels:
        if poly.buffer(0.01).contains(Point(x,y)) and not t.startswith('HUECO'):
            lab=t; break
    name=lab.split()[0]; area_lbl=float(re.search(r'([\d.]+) m',lab).group(1))
    tabs.append(dict(name=name,level=lvl,pts=[(round(x-dx,3),round(y,3)) for x,y in pts],area_dxf=area_lbl,label=lab))
huecos={'N+3.00':[],'N+6.00':[]}
for p in mt.query('LWPOLYLINE[layer=="TRIB_HUECO"]'):
    pts=[(float(x)-14.0,float(y)) for x,y,*_ in p.get_points()]
    huecos['N+6.00'].append(Polygon(pts))
# ---------- modelo ----------
areas={(a['name'],a['story']):a for a in m['areas']}
for t in tabs:
    a=areas.get((t['name'],t['level']))
    t['etabs_section']=a['section'] if a else None
    t['loads']=a['loads'] if a else {}
def walls_of(level):
    out=[]
    for a in m['areas']:
        if a['story']==level and a['kind']=='PANEL':
            pts=[]
            for c in a['coords']:
                c=(round(c[0],3),round(c[1],3))
                if c not in pts: pts.append(c)
            xs=[c[0] for c in pts]; ys=[c[1] for c in pts]
            out.append(LineString([(min(xs),min(ys)),(max(xs),max(ys))]))
    return unary_union(out)
def beams_of(level):
    return [LineString([tuple(b['i']),tuple(b['j'])]) for b in m['beams'] if b['story']==level and b['kind']=='BEAM' and b['section']=='002_D-15x20']
def cols_of(level):
    return sorted(set((round(b['i'][0],3),round(b['i'][1],3)) for b in m['beams'] if b['story']==level and b['kind']=='COLUMN'))
# muros que soportan cada losa = muros del entrepiso de abajo (story = nivel de la losa)
SUPPORT={'N+3.00':'N+3.00','N+6.00':'N+6.00'}
# muros que cargan sobre cada losa = muros del entrepiso de arriba
ABOVE={'N+3.00':'N+6.00','N+6.00':None}
# ---------- vacios del plano arquitectonico ----------
VAC={'N+3.00':[dict(nombre='Vacío doble altura (sala)',poly=box(0.075,7.946,4.247,13.946),fuente='001_PLANO_CROCHA_R3.0.dxf capa 004_Vacios, PLANTA PRIMER NIVEL'),
               dict(nombre='Escalera curva (cubo)',poly=Polygon([(5.46,6.948),(8.85,6.948),(8.85,9.95),(5.422,9.95)]).difference(box(5.46,8.91,7.35,9.95)),fuente='001_PLANO_CROCHA_R3.0.dxf capa 005_Escalera; hueco sin tablero en el E2K')],
     'N+6.00':[dict(nombre='Pergolado sobre doble altura (vigas VP-15x40 @0.75 + cristal)',poly=box(0.075,7.946,4.427,13.946),fuente='001_PLANO_CROCHA_R3.0.dxf capa 004_Vacios, PLANTA DE AZOTEA; E2K vigas VP-15x40'),
               dict(nombre='Vacío sobre escalera',poly=box(5.998,7.023,8.775,9.875),fuente='001_PLANO_CROCHA_R3.0.dxf capa 004_Vacios, PLANTA DE AZOTEA')]}
for h in huecos['N+6.00']:
    VAC['N+6.00'].append(dict(nombre='Hueco 0.30 m (tragaluz)',poly=h,fuente='AREAS_TRIBUTARIAS_R3.2.dxf capa TRIB_HUECO / E2K áreas A1-A3'))
res=dict(tableros=[],vacios={},principales=[])
for lvl in LEVELS:
    res['vacios'][lvl]=[dict(nombre=v['nombre'],coords=[(round(x,3),round(y,3)) for x,y in v['poly'].exterior.coords],fuente=v['fuente'],area=round(v['poly'].area,2)) for v in VAC[lvl]]
for t in tabs:
    lvl=t['level']; P=Polygon(t['pts'])
    x0,y0,x1,y1=P.bounds
    t['bbox']=[x0,y0,x1,y1]; t['Lx']=round(x1-x0,3); t['Ly']=round(y1-y0,3)
    t['area_poly']=round(P.area,3); t['rect']=abs(P.area-(x1-x0)*(y1-y0))<0.01
    sup=walls_of(SUPPORT[lvl]); others=[Polygon(o['pts']) for o in tabs if o['level']==lvl and o is not t]
    names=[o['name'] for o in tabs if o['level']==lvl and o is not t]
    edges={'inferior':LineString([(x0,y0),(x1,y0)]),'superior':LineString([(x0,y1),(x1,y1)]),'izquierdo':LineString([(x0,y0),(x0,y1)]),'derecho':LineString([(x1,y0),(x1,y1)])}
    out={}
    for k,e in edges.items():
        L=e.length
        fm=e.intersection(sup.buffer(TOL)).length/L
        nb=[]; fn=0
        for o,nm in zip(others,names):
            ov=e.intersection(o.buffer(TOL)).length
            if ov>0.05: nb.append((nm,round(ov,3))); fn+=ov
        fv=0; vn=[]
        for v in VAC[lvl]:
            ov=e.intersection(v['poly'].buffer(0.12)).length
            if ov>0.05: fv+=ov; vn.append(v['nombre'])
        out[k]=dict(L=round(L,3),frac_muro=round(min(fm,1),3),frac_vecino=round(min(fn/L,1),3),vecinos=nb,frac_vacio=round(min(fv/L,1),3),vacios=vn)
    t['bordes']=out
    res['tableros'].append(t)
json.dump(res,open(W+'work/geo.json','w'),indent=1,ensure_ascii=False)
for t in res['tableros']:
    print(t['name'],t['level'],t['Lx'],'x',t['Ly'],'rect' if t['rect'] else 'IRREG',t['etabs_section'],t['loads'].get('Sobrecarga'))
    for k,e in t['bordes'].items():
        print('    %-9s L=%.3f muro=%.2f vecino=%.2f %s vacio=%.2f %s'%(k,e['L'],e['frac_muro'],e['frac_vecino'],[n for n,_ in e['vecinos']],e['frac_vacio'],e['vacios']))
