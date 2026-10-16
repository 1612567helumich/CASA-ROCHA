import re, json, sys, shlex
path=sys.argv[1]
L=open(path,encoding='utf-8').read().splitlines()
pts={}; lines={}; areas={}; aassign={}; lassign={}; aload={}; lload=[]
stories={}
sec=None
for l in L:
    if l.startswith('$ '): sec=l[2:].strip(); continue
    s=l.strip()
    if not s: continue
    try: t=shlex.split(s)
    except Exception: t=s.split()
    if sec=='POINT COORDINATES' and t[0]=='POINT':
        pts[t[1]]=(float(t[2]),float(t[3]))
    elif sec=='LINE CONNECTIVITIES' and t[0]=='LINE':
        lines[t[1]]=dict(kind=t[2],i=t[3],j=t[4])
    elif sec=='AREA CONNECTIVITIES' and t[0]=='AREA':
        n=int(t[3]); areas[t[1]]=dict(kind=t[2],pts=t[4:4+n])
    elif sec=='AREA ASSIGNS' and t[0]=='AREAASSIGN':
        d=aassign.setdefault((t[1],t[2]),{})
        if 'SECTION' in t: d['section']=t[t.index('SECTION')+1]
    elif sec=='LINE ASSIGNS' and t[0]=='LINEASSIGN':
        d=lassign.setdefault((t[1],t[2]),{})
        if 'SECTION' in t: d['section']=t[t.index('SECTION')+1]
    elif sec=='SHELL OBJECT LOADS' and t[0]=='AREALOAD':
        aload.setdefault((t[1],t[2]),{})[t[t.index('LC')+1]]=float(t[t.index('FVAL')+1])
    elif sec=='FRAME OBJECT LOADS' and t[0]=='LINELOAD':
        lload.append(t)
out=dict(areas=[],beams=[])
for (name,story),d in aassign.items():
    a=areas[name]
    out['areas'].append(dict(name=name,story=story,section=d.get('section'),kind=a['kind'],coords=[pts[p] for p in a['pts']],loads=aload.get((name,story),{})))
for (name,story),d in lassign.items():
    ln=lines[name]
    out['beams'].append(dict(name=name,story=story,section=d.get('section'),kind=ln['kind'],i=pts[ln['i']],j=pts[ln['j']]))
out['lineloads']=lload
json.dump(out,open(sys.argv[2],'w'),indent=1,ensure_ascii=False)
