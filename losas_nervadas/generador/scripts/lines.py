"""Nervaduras principales: lineas de dala del E2K, tramos sobre muro y tramos libres entre apoyos firmes."""
import json, math
from shapely.geometry import LineString, Point, Polygon, box
from shapely.ops import unary_union
import ezdxf
W='/tmp/claude-0/-home-user-CASA-ROCHA/faa149e0-9f48-5f0d-b5ee-28b4c7156d8c/scratchpad/'
m=json.load(open(W+'work/model.json')); geo=json.load(open(W+'work/geo.json'))
T=0.03
def merge_lines(segs):
    """segs: list of ((x0,y0),(x1,y1)) axis-aligned. returns merged list of (ori, c, a, b)."""
    H={};V={}
    for (p,q) in segs:
        if abs(p[1]-q[1])<1e-3:
            H.setdefault(round(p[1],3),[]).append((min(p[0],q[0]),max(p[0],q[0])))
        elif abs(p[0]-q[0])<1e-3:
            V.setdefault(round(p[0],3),[]).append((min(p[1],q[1]),max(p[1],q[1])))
    out=[]
    for ori,D in (('H',H),('V',V)):
        for c,iv in D.items():
            iv.sort(); cur=list(iv[0])
            for a,b in iv[1:]:
                if a<=cur[1]+T: cur[1]=max(cur[1],b)
                else: out.append((ori,c,cur[0],cur[1])); cur=[a,b]
            out.append((ori,c,cur[0],cur[1]))
    return out
def walls(story):
    segs=[]
    for a in m['areas']:
        if a['story']==story and a['kind']=='PANEL':
            xs=[c[0] for c in a['coords']]; ys=[c[1] for c in a['coords']]
            segs.append(((min(xs),min(ys)),(max(xs),max(ys))))
    return merge_lines(segs)
def cols(story):
    return sorted(set((round(b['i'][0],3),round(b['i'][1],3)) for b in m['beams'] if b['story']==story and b['kind']=='COLUMN'))
def dalas(story):
    return merge_lines([(tuple(b['i']),tuple(b['j'])) for b in m['beams'] if b['story']==story and b['kind']=='BEAM' and b['section']=='002_D-15x20'])
ABOVE={'N+3.00':'N+6.00','N+6.00':None}
res=[]
for lvl in ['N+3.00','N+6.00']:
    WL=walls(lvl); CL=cols(lvl); DL=dalas(lvl)
    WA=walls(ABOVE[lvl]) if ABOVE[lvl] else []
    for (ori,c,a,b) in DL:
        # intervalos sobre muro colineal
        cov=[]
        for (o2,c2,a2,b2) in WL:
            if o2==ori and abs(c2-c)<T:
                lo,hi=max(a,a2),min(b,b2)
                if hi>lo+T: cov.append((lo,hi))
        cov.sort()
        # puntos de apoyo firme: castillos sobre la linea, muros perpendiculares que cruzan o llegan a la linea
        pts=set()
        for (x,y) in CL:
            s,cc=(x,y) if ori=='H' else (y,x)
            if abs(cc-c)<T and a-T<=s<=b+T: pts.add(round(s,3))
        for (o2,c2,a2,b2) in WL:
            if o2!=ori and a2-T<=c<=b2+T and a-T<=c2<=b+T: pts.add(round(c2,3))
        for lo,hi in cov: pts.add(round(lo,3)); pts.add(round(hi,3))
        # cruces con otras dalas (apoyo en otra nervadura, no firme)
        jn=set()
        for (o2,c2,a2,b2) in DL:
            if o2!=ori and a2-T<=c<=b2+T and a-T<=c2<=b+T: jn.add(round(c2,3))
        # tramos
        cuts=sorted(set([round(a,3),round(b,3)]+list(pts)))
        segs=[]
        for s0,s1 in zip(cuts[:-1],cuts[1:]):
            if s1-s0<0.05: continue
            mid=(s0+s1)/2
            onwall=any(lo-T<=mid<=hi+T for lo,hi in cov)
            segs.append([s0,s1,'MURO' if onwall else 'LIBRE'])
        # fusiona libres contiguos (cuando el corte no es apoyo firme no hay corte; ya estan cortados solo en firmes)
        for s0,s1,tp in segs:
            e0='FIRME' if any(abs(s0-p)<T for p in pts) else ('NERV' if any(abs(s0-p)<T for p in jn) else 'LIBRE')
            e1='FIRME' if any(abs(s1-p)<T for p in pts) else ('NERV' if any(abs(s1-p)<T for p in jn) else 'LIBRE')
            # muro arriba sobre el tramo
            ov=0
            for (o2,c2,a2,b2) in WA:
                if o2==ori and abs(c2-c)<T:
                    ov+=max(0,min(s1,b2)-max(s0,a2))
            interior=[p for p in sorted(jn) if s0+T<p<s1-T]
            res.append(dict(level=lvl,ori=ori,c=c,s0=s0,s1=s1,L=round(s1-s0,3),tipo=tp,e0=e0,e1=e1,muro_arriba=round(ov,3),cruces_interiores=interior))
json.dump(res,open(W+'work/lines.json','w'),indent=1)
for r in res:
    if r['tipo']=='LIBRE':
        eje=('y=%.3f'%r['c']) if r['ori']=='H' else ('x=%.3f'%r['c'])
        print(r['level'],eje,'%6.3f→%6.3f L=%.2f'%(r['s0'],r['s1'],r['L']),r['e0'],r['e1'],'muro_arriba=%.2f'%r['muro_arriba'],'cruces',r['cruces_interiores'])
print(sum(1 for r in res if r['tipo']=='MURO'),'tramos sobre muro')
