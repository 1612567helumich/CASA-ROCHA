import ezdxf, json, sys, math
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + '/'
d=ezdxf.readfile(W+'in/264d7eb7-001_PLANO_CROCHA_R3.0.dxf')
msp=d.modelspace()
m=json.load(open(W+'work/model.json'))
y0=193.61
cfg={'N+3.00':('PN1',61.935,'PB',31.936),'N+6.00':('AZ',91.935,'PN1',61.935)}
def segs(x0, layers):
    out=[]
    for e in msp:
        if e.dxf.layer not in layers: continue
        t=e.dxftype()
        if t=='LINE':
            p=[tuple(e.dxf.start)[:2],tuple(e.dxf.end)[:2]]
        elif t=='LWPOLYLINE':
            p=[(x,y) for x,y,*_ in e.get_points()]
            if e.closed: p.append(p[0])
        elif t=='ARC' or t=='CIRCLE':
            c=e.dxf.center; r=e.dxf.radius
            a0,a1=(e.dxf.start_angle,e.dxf.end_angle) if t=='ARC' else (0,360)
            if a1<a0: a1+=360
            n=24; p=[(c[0]+r*math.cos(math.radians(a0+(a1-a0)*k/n)),c[1]+r*math.sin(math.radians(a0+(a1-a0)*k/n))) for k in range(n+1)]
        else: continue
        if not (x0-4<p[0][0]<x0+12 and y0-3<p[0][1]<y0+28): continue
        out.append([(x-x0,y-y0) for x,y in p])
    return out
for st,(top,xt,bot,xb) in cfg.items():
    fig,ax=plt.subplots(figsize=(11,26))
    for s in segs(xb,['001_Muros']): ax.plot(*zip(*s),color='#bbbbbb',lw=0.6)
    for s in segs(xt,['004_Vacios']): ax.plot(*zip(*s),color='magenta',lw=1.2)
    for s in segs(xt,['005_Escalera']): ax.plot(*zip(*s),color='orange',lw=0.5)
    for a in m['areas']:
        if a['story']!=st: continue
        P=a['coords']+[a['coords'][0]]
        if a['kind']=='PANEL':
            continue
        col={'LR-25':'blue','Slab1':'red',None:'green'}.get(a['section'],'k')
        ax.plot(*zip(*P),color=col,lw=1.5)
        cx=sum(c[0] for c in a['coords'])/len(a['coords']); cy=sum(c[1] for c in a['coords'])/len(a['coords'])
        ax.text(cx,cy,a['name']+'\n'+str(a['section']),fontsize=7,ha='center',color=col)
    for a in m['areas']:
        if a['story']==st and a['kind']=='PANEL':
            xs=[c[0] for c in a['coords']]; ys=[c[1] for c in a['coords']]
            ax.plot([min(xs),max(xs)],[min(ys),max(ys)] if xs[0]==xs[-1] or True else None,color='black',lw=4,alpha=0.35)
    for b in m['beams']:
        if b['story']==st and b['kind']=='BEAM':
            ax.plot([b['i'][0],b['j'][0]],[b['i'][1],b['j'][1]],color='red' if b['section']=='002_D-15x20' else 'cyan',lw=1,ls='--')
    ax.set_aspect('equal'); ax.grid(True,lw=0.3); ax.set_title(f'{st}: tableros (azul LR-25, rojo Slab1), muros abajo (gris grueso), dalas (rojo punteado), vacios (magenta), arq {top}')
    ax.set_xlim(-3,10); ax.set_ylim(-1,25.5)
    fig.savefig(W+f'work/overlay_{st}.png',dpi=70,bbox_inches='tight')
    print('ok',st)
