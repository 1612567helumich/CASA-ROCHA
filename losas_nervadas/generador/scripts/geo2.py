"""Datos de proyecto para la hoja: tableros, bordes, vacios, tramos de nervadura principal, tributarias y transferencias."""
import ezdxf, json, re
from shapely.geometry import Polygon, LineString, Point, box
from shapely.ops import unary_union
W='/tmp/claude-0/-home-user-CASA-ROCHA/faa149e0-9f48-5f0d-b5ee-28b4c7156d8c/scratchpad/'
m=json.load(open(W+'work/model.json')); geo=json.load(open(W+'work/geo.json'))
T=0.03; TC=0.05
LV=['N+3.00','N+6.00']; DX={'N+3.00':0.0,'N+6.00':14.0}
ABOVE={'N+3.00':'N+6.00','N+6.00':None}
def merge(segs,tc=TC):
    H={};V={}
    for (p,q) in segs:
        if abs(p[1]-q[1])<1e-3: H.setdefault(p[1],[]).append((min(p[0],q[0]),max(p[0],q[0])))
        elif abs(p[0]-q[0])<1e-3: V.setdefault(p[0],[]).append((min(p[1],q[1]),max(p[1],q[1])))
    out=[]
    for ori,D in (('H',H),('V',V)):
        cs=sorted(D)
        groups=[]
        for c in cs:
            if groups and abs(c-groups[-1][0])<tc: groups[-1][1].extend(D[c]); groups[-1][2].append(c)
            else: groups.append([c,list(D[c]),[c]])
        for c,iv,clist in groups:
            cc=round(sum(clist)/len(clist),3)
            iv.sort(); cur=list(iv[0])
            for a,b in iv[1:]:
                if a<=cur[1]+T: cur[1]=max(cur[1],b)
                else: out.append([ori,cc,round(cur[0],3),round(cur[1],3)]); cur=[a,b]
            out.append([ori,cc,round(cur[0],3),round(cur[1],3)])
    return out
def wall_segs(st):
    s=[]
    for a in m['areas']:
        if a['story']==st and a['kind']=='PANEL':
            xs=[c[0] for c in a['coords']]; ys=[c[1] for c in a['coords']]
            s.append(((min(xs),min(ys)),(max(xs),max(ys))))
    return s
PROP={'N+3.00':[(4.327,1.919),(4.327,2.921),(5.923,2.921),(4.327,7.871),(4.327,14.021),(5.422,9.95)],'N+6.00':[]}
import os
USE_PROP=os.environ.get('PROP','1')=='1'
def cols(st,prop=True):
    base=sorted(set((round(b['i'][0],3),round(b['i'][1],3)) for b in m['beams'] if b['story']==st and b['kind']=='COLUMN'))
    if prop and USE_PROP: base=sorted(set(base)|set(PROP.get(st,[])))
    return base
def beam_segs(st,sec='002_D-15x20'): return [(tuple(b['i']),tuple(b['j'])) for b in m['beams'] if b['story']==st and b['kind']=='BEAM' and b['section']==sec]
def pt(ori,c,s): return (s,c) if ori=='H' else (c,s)
out=dict(tableros=[],vacios=geo['vacios'],tramos=[],transfer=[],trib=[],castillos_sin_continuidad=[],lineloads=[])
# ---------- E2K line loads (por viga) ----------
beams={(b['name'],b['story']):b for b in m['beams']}
LL=[]
for t in m['lineloads']:
    b=beams[(t[1],t[2])]; lc=t[t.index('LC')+1]
    p,q=b['i'],b['j']; Lb=((q[0]-p[0])**2+(q[1]-p[1])**2)**0.5
    if 'FVAL' in t: f0=f1=float(t[t.index('FVAL')+1]); r0,r1=0,1
    else: f0=float(t[t.index('FSTART')+1]); f1=float(t[t.index('FEND')+1]); r0=float(t[t.index('RDSTART')+1]); r1=float(t[t.index('RDEND')+1])
    a=(p[0]+(q[0]-p[0])*r0,p[1]+(q[1]-p[1])*r0); bb=(p[0]+(q[0]-p[0])*r1,p[1]+(q[1]-p[1])*r1)
    LL.append(dict(beam=t[1],level=t[2],lc=lc,p=a,q=bb,f=(f0+f1)/2,L=round(Lb*(r1-r0),3)))
out['lineloads']=[dict(beam=l['beam'],level=l['level'],lc=l['lc'],x0=round(l['p'][0],3),y0=round(l['p'][1],3),x1=round(l['q'][0],3),y1=round(l['q'][1],3),w=l['f'],L=l['L']) for l in LL]
# ---------- tramos ----------
tid=0
for lvl in LV:
    WL=merge(wall_segs(lvl)); CL=cols(lvl); DL=merge(beam_segs(lvl))
    WA=merge(wall_segs(ABOVE[lvl])) if ABOVE[lvl] else []
    lines=[]
    for (ori,c,a,b) in DL:
        cov=[(max(a,a2),min(b,b2)) for (o2,c2,a2,b2) in WL if o2==ori and abs(c2-c)<TC and min(b,b2)>max(a,a2)+T]
        firm=set()
        for (x,y) in CL:
            s,cc=(x,y) if ori=='H' else (y,x)
            if abs(cc-c)<TC and a-T<=s<=b+T: firm.add(round(s,3))
        for (o2,c2,a2,b2) in WL:
            if o2!=ori and a2-TC<=c<=b2+TC and a-T<=c2<=b+T: firm.add(round(c2,3))
        for lo,hi in cov: firm.add(round(lo,3)); firm.add(round(hi,3))
        cuts=sorted(set([a,b]+[f for f in firm if a-T<=f<=b+T]))
        cuts2=[]
        for x in cuts:
            if not cuts2 or x-cuts2[-1]>0.05: cuts2.append(x)
        lines.append(dict(ori=ori,c=c,a=a,b=b,cov=cov,firm=sorted(firm),cuts=cuts2))
    tr=[]
    for L_ in lines:
        for s0,s1 in zip(L_['cuts'][:-1],L_['cuts'][1:]):
            mid=(s0+s1)/2
            onwall=any(lo-T<=mid<=hi+T for lo,hi in L_['cov'])
            tid+=1
            tr.append(dict(id=tid,level=lvl,ori=L_['ori'],c=L_['c'],s0=s0,s1=s1,L=round(s1-s0,3),tipo='MURO' if onwall else 'LIBRE',
                           f0=any(abs(s0-f)<0.05 for f in L_['firm']),f1=any(abs(s1-f)<0.05 for f in L_['firm'])))
    # nodos y clasificacion de extremos
    def at(t,k): return pt(t['ori'],t['c'],t['s0'] if k==0 else t['s1'])
    def interior_of(P,excl):
        res=[]
        for t in tr:
            if t is excl: continue
            s,cc=(P[0],P[1]) if t['ori']=='H' else (P[1],P[0])
            if abs(cc-t['c'])<TC and t['s0']+0.05<s<t['s1']-0.05: res.append((t,s-t['s0']))
        return res
    def ends_at(P,excl):
        res=[]
        for t in tr:
            if t is excl: continue
            for k in (0,1):
                Q=at(t,k)
                if abs(Q[0]-P[0])<TC and abs(Q[1]-P[1])<TC: res.append((t,k))
        return res
    for t in tr:
        t['ext']=[None,None]
        if t['tipo']!='LIBRE': continue
        for k in (0,1):
            if (t['f0'] if k==0 else t['f1']): t['ext'][k]=dict(tipo='FIRME'); continue
            P=at(t,k)
            inn=[x for x in interior_of(P,t) ]
            if inn:
                rt,a=inn[0]
                t['ext'][k]=dict(tipo='MURO' if rt['tipo']=='MURO' else 'NERVADURA',recibe=rt['id'],a=round(a,3)); continue
            t['ext'][k]=dict(tipo='NUDO',P=P)
    # cantilevers & punta
    for t in tr:
        if t['tipo']!='LIBRE': continue
        k0,k1=t['ext']
        sup=[e['tipo'] in ('FIRME','MURO','NERVADURA') for e in (k0,k1)]
        if all(sup): t['cond']='APOYADA'
        elif any(sup): t['cond']='VOLADO'
        else: t['cond']='?'
    for t in tr:
        if t['tipo']=='LIBRE' and t['cond']=='?':
            ok=[]
            for k in (0,1):
                P=at(t,k); hit=None
                for (o,ko) in ends_at(P,t):
                    if o['tipo']=='LIBRE' and o.get('cond')=='VOLADO' and o['ext'][ko]['tipo']=='NUDO':
                        hit=(o,ko); break
                if hit:
                    t['ext'][k]=dict(tipo='PUNTA',recibe=hit[0]['id'],a=round(hit[0]['L'],3))
                    ok.append(True)
                else: ok.append(False)
            t['cond']='APOYADA EN VOLADOS' if all(ok) else 'SIN APOYO'
    for t in tr:
        if t['tipo']=='LIBRE':
            for k in (0,1):
                e=t['ext'][k]
                if e['tipo']=='NUDO': e['tipo']='LIBRE'; e.pop('P',None)
    # muro arriba
    for t in tr:
        ov=0
        for (o2,c2,a2,b2) in WA:
            if o2==t['ori'] and abs(c2-t['c'])<TC: ov+=max(0,min(t['s1'],b2)-max(t['s0'],a2))
        t['muro_arriba']=round(ov,3)
        # cargas lineales e2k sobre el tramo (promedio en el tramo, kg/m) por caso
        acc={}
        for l in LL:
            if l['level']!=lvl: continue
            p,q=l['p'],l['q']
            if t['ori']=='H' and abs(p[1]-q[1])<1e-3 and abs(p[1]-t['c'])<TC:
                lo,hi=min(p[0],q[0]),max(p[0],q[0])
            elif t['ori']=='V' and abs(p[0]-q[0])<1e-3 and abs(p[0]-t['c'])<TC:
                lo,hi=min(p[1],q[1]),max(p[1],q[1])
            else: continue
            ov=max(0,min(t['s1'],hi)-max(t['s0'],lo))
            if ov>0: acc[l['lc']]=acc.get(l['lc'],0)+l['f']*ov
        t['e2k_lin']={k:round(v/t['L'],1) for k,v in acc.items()}
    out['tramos'].extend(tr)
    # castillos de arriba sin continuidad abajo
    if ABOVE[lvl]:
        up=cols(ABOVE[lvl],False); dn=set(cols(lvl,False))
        WLs=[LineString([pt(o,c,a),pt(o,c,b)]) for (o,c,a,b) in WL]
        for P in up:
            if P in dn: continue
            if any(l.distance(Point(P))<TC for l in WLs): continue
            out['castillos_sin_continuidad'].append(dict(level=lvl,x=P[0],y=P[1]))
# ---------- tributarias ----------
dt=ezdxf.readfile(W+'in/c566697e-AREAS_TRIBUTARIAS_R3.2.dxf'); mt=dt.modelspace()
tabpolys={(t['name'],t['level']):Polygon(t['pts']) for t in geo['tableros']}
for p in mt.query('LWPOLYLINE[layer=="TRIB_AREA"]'):
    pts=[(float(x),float(y)) for x,y,*_ in p.get_points()]
    lvl='N+6.00' if min(x for x,y in pts)>12-0.01 and max(x for x,y in pts)>13 else 'N+3.00'
    pts=[(x-DX[lvl],y) for x,y in pts]
    P=Polygon(pts).buffer(0)
    if P.area<1e-4: continue
    owner=None
    for (nm,lv),TP in tabpolys.items():
        if lv==lvl and TP.buffer(0.01).contains(P.representative_point()) and TP.buffer(0.02).contains(P): owner=nm; break
    # base: lado del poligono mas largo que cae sobre el borde del tablero
    best=None
    cs=list(P.exterior.coords)
    TP=tabpolys[(owner,lvl)]
    for (a,b) in zip(cs[:-1],cs[1:]):
        seg=LineString([a,b])
        if seg.length<0.02: continue
        if TP.exterior.buffer(0.01).contains(seg):
            if best is None or seg.length>best.length: best=seg
    (ax,ay),(bx,by)=best.coords
    ori='H' if abs(ay-by)<1e-3 else 'V'
    c=round(ay,3) if ori=='H' else round(ax,3)
    lo,hi=(min(ax,bx),max(ax,bx)) if ori=='H' else (min(ay,by),max(ay,by))
    # reparte el area entre tramos colineales por proyeccion
    for t in out['tramos']:
        if t['level']!=lvl or t['ori']!=ori or abs(t['c']-c)>TC: continue
        lo2,hi2=max(lo,t['s0']),min(hi,t['s1'])
        if hi2<=lo2+1e-3: continue
        band=box(lo2,-100,hi2,100) if ori=='H' else box(-100,lo2,100,hi2)
        A=P.intersection(band).area
        if A>1e-4: out['trib'].append(dict(tramo=t['id'],tablero=owner,level=lvl,area=round(A,4)))
# transferencias (reacciones de un tramo que se apoyan en otro)
for t in out['tramos']:
    if t['tipo']!='LIBRE': continue
    for k in (0,1):
        e=t['ext'][k]
        if e['tipo'] in ('NERVADURA','PUNTA'):
            out['transfer'].append(dict(de=t['id'],extremo=k,a=e['recibe'],pos=e['a']))
# tableros
for tb in geo['tableros']:
    out['tableros'].append(tb)
json.dump(out,open(W+'work/proyecto%s.json'%('' if USE_PROP else '_modelo'),'w'),indent=1,ensure_ascii=False)
# reporte
for t in out['tramos']:
    if t['tipo']=='LIBRE':
        eje=('y=%.3f'%t['c']) if t['ori']=='H' else ('x=%.3f'%t['c'])
        A=sum(r['area'] for r in out['trib'] if r['tramo']==t['id'])
        print('%3d %s %-9s %6.3f→%6.3f L=%5.2f %-18s %-30s %-30s Atrib=%.2f muroArr=%.2f e2k=%s'%(t['id'],t['level'],eje,t['s0'],t['s1'],t['L'],t['cond'],t['ext'][0],t['ext'][1],A,t['muro_arriba'],t['e2k_lin']))
print('transfer',out['transfer'])
print('castillos sin continuidad',out['castillos_sin_continuidad'])
tot={}
for r in out['trib']: tot[(r['tablero'],r['level'])]=tot.get((r['tablero'],r['level']),0)+r['area']
for tb in geo['tableros']:
    print(tb['name'],tb['level'],tb['area_dxf'],round(tot.get((tb['name'],tb['level']),0),2),round(tb['area_poly'],2))
