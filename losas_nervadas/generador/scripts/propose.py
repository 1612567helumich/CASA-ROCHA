"""Itera la sección propuesta (b × h) de cada nervadura principal hasta que todas cumplan.
Usa el propio libro (LibreOffice recalcula) como calculadora: no replica fórmulas."""
import json, subprocess, sys, glob, os, time
import openpyxl
from openpyxl.utils import column_index_from_string as CI
W = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) + '/'
RECALC = glob.glob('/root/.claude/skills/synced/*/xlsx/scripts/recalc.py')[0]
PF = W + 'work/propuesta_np.json'
prop = json.load(open(PF)) if os.path.exists(PF) else {}
data = json.load(open(W + 'work/data.json'))
punta = []
for t in data['tramos']:
    for e in t.get('ext', []):
        if e and e.get('tipo') == 'PUNTA' and e.get('recibe'):
            punta.append((t['np'], e['recibe']))
for it in range(int(sys.argv[1]) if len(sys.argv) > 1 else 12):
    t0 = time.time()
    subprocess.run(['python3', 'gen_d.py', '../out/LosasNervadas_CasaRocha.xlsx'], cwd=W + 'scripts', check=True, capture_output=True)
    subprocess.run(['cp', W + 'out/LosasNervadas_CasaRocha.xlsx', W + 'out/rc.xlsx'], check=True)
    out = subprocess.run(['python3', RECALC, W + 'out/rc.xlsx', '300'], capture_output=True, text=True)
    res = json.loads(out.stdout[out.stdout.index('{'):])
    sys.path.insert(0, W + 'scripts')
    import importlib, gen_b
    importlib.reload(gen_b)
    PC = gen_b.PC
    wb = openpyxl.load_workbook(W + 'out/rc.xlsx', data_only=True)
    ws = wb['PRINCIPALES']
    changed = 0; bad = []
    for r in range(gen_b.P0, gen_b.P1 + 1):
        g = lambda k: ws.cell(r, CI(PC[k])).value
        np_, est = g('np'), str(g('estado'))
        b, h, hrec = int(g('bq')), int(g('hq')), g('hrec') or 0
        nx = int(g('nPx') or 0)
        if not est.startswith('NO CUMPLE'):
            continue
        bad.append((np_, b, h, nx, est))
        nb, nh, nnx = b, h, nx
        reasons = est.split(':', 1)[1]
        strength = any(k in reasons for k in ('flexión', 'cortante', 'peralte', 'muro sobre viga', 'ε_t'))
        if 'no caben' in reasons or h / b > 6:
            nb = min(b + 5, 40)
        if strength:
            nh = min(max(h + 5, int(hrec)), 90)
        elif 'flecha' in reasons:
            # primero más acero inferior (sube la I agrietada) si cabe otra varilla; si no, más peralte
            v = g('vP') or 4; dbv = {2: 0.64, 2.5: 0.79, 3: 0.95, 4: 1.27, 5: 1.59, 6: 1.91, 8: 2.54}[v]
            fit = int((g('nmax') + max(2.5, dbv)) / (dbv + max(2.5, dbv)))
            if (g('nP') or 0) + 1 <= fit and nx < 2:
                nnx = nx + 1
            else:
                nh = min(h + 5, 90); nnx = 0
        if nh / nb > 6:
            nb = min(nb + 5, 40)
        if np_ in gen_b.VOLLAT:   # viga de orilla del volado: b es la franja maciza (fórmula en PRINCIPALES); sólo se puede peraltar
            nb = b
        if (nb, nh, nnx) != (b, h, nx):
            prop[np_] = [nb, nh, nnx]; changed += 1
    for a, c in punta:   # el tramo que se apoya en la punta de un volado corto y ese volado son la misma trabe
        sa, sc = (prop.get(a, []) + [15, 25, 0])[:3], (prop.get(c, []) + [15, 25, 0])[:3]
        m = [max(sa[0], sc[0]), max(sa[1], sc[1])]
        if m != sa[:2] or m != sc[:2]:
            prop[a] = m + [sa[2]]; prop[c] = m + [sc[2]]; changed += 1
    json.dump(prop, open(PF, 'w'), indent=1, sort_keys=True)
    print('iter %d: errores %s, no cumplen %d, cambios %d, %.0f s' % (it, res.get('total_errors'), len(bad), changed, time.time() - t0), flush=True)
    for x in bad:
        print('   ', x, flush=True)
    if changed == 0:
        break
print(json.dumps(prop, sort_keys=True))
