import ezdxf, json, sys, math
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
W='/tmp/claude-0/-home-user-CASA-ROCHA/faa149e0-9f48-5f0d-b5ee-28b4c7156d8c/scratchpad/'
d=ezdxf.readfile(W+'in/264d7eb7-001_PLANO_CROCHA_R3.0.dxf')
m=json.load(open(W+'work/model.json'))
y0=193.61
st=sys.argv[1]; planta=sys.argv[2]; x0=float(sys.argv[3])
fig=plt.figure(figsize=(12,30)); ax=fig.add_axes([0.03,0.02,0.95,0.96])
ctx=RenderContext(d)
from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy
cfg=Configuration(background_policy=BackgroundPolicy.WHITE, color_policy=ColorPolicy.BLACK)
Frontend(ctx, MatplotlibBackend(ax), config=cfg).draw_layout(d.modelspace(), finalize=False)
# overlay in arch coords
for a in m['areas']:
    if a['story']!=st or a['kind']=='PANEL': continue
    P=[(x+x0,y+y0) for x,y in a['coords']]; P.append(P[0])
    ax.plot(*zip(*P),color='blue',lw=1.2,alpha=0.7)
    cx=sum(p[0] for p in P[:-1])/len(P[:-1]); cy=sum(p[1] for p in P[:-1])/len(P[:-1])
    ax.text(cx,cy,a['name'],fontsize=8,color='blue',ha='center')
for b in m['beams']:
    if b['story']!=st: continue
    if b['kind']=='COLUMN':
        ax.plot(b['i'][0]+x0,b['i'][1]+y0,'s',color='red',ms=6)
for a in m['areas']:
    if a['story']==st and a['kind']=='PANEL':
        xs=[c[0]+x0 for c in a['coords']]; ys=[c[1]+y0 for c in a['coords']]
        ax.plot([min(xs),max(xs)],[min(ys),max(ys)],color='green',lw=5,alpha=0.35)
ax.set_xlim(x0-3,x0+10); ax.set_ylim(y0-1.5,y0+25.5); ax.set_aspect('equal')
fig.savefig(W+f'work/arch_{planta}.png',dpi=60)
print('ok')
