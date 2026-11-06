"""Genera costos_ml.js (comisión y envío de referencia) a partir del reporte Ventas AR de ML.

Uso: python herramientas/generar_costos_ml.py "Ventas_AR.xlsx" costos_ml.js
"""
import openpyxl, sys, re, json, statistics as st, warnings, datetime
warnings.filterwarnings('ignore')
src, dst = sys.argv[1], sys.argv[2]
DIAS = 90
MESES = {m: i + 1 for i, m in enumerate(['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'])}
MESES['setiembre'] = 9

wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
rows = list(wb.worksheets[0].iter_rows(values_only=True))
h = next(i for i, r in enumerate(rows[:20]) if r and r[0] == '# de venta')
H = [str(c or '').strip() for c in rows[h]]
def col(pred):
    return next(i for i, x in enumerate(H) if pred(x))
C = dict(fecha=col(lambda x: x.startswith('Fecha de venta')), estado=col(lambda x: x == 'Estado'),
         uni=col(lambda x: x == 'Unidades'), ing=col(lambda x: x.startswith('Ingresos por productos')),
         cargo=col(lambda x: x.startswith('Cargo por venta')), ingenv=col(lambda x: x.startswith('Ingresos por env')),
         env=col(lambda x: x.startswith('Costos de env')), sku=col(lambda x: x == 'SKU'),
         mla=col(lambda x: x.startswith('# de publicaci')), tit=col(lambda x: x.startswith('Título de la publicaci')),
         cuotas=col(lambda x: x.startswith('Tiene cuotas')))

def num(v):
    try: return float(v)
    except (TypeError, ValueError): return 0.0
def fecha(v):
    m = re.search(r'(\d{1,2}) de ([a-záéíóú]+) de (\d{4})', str(v or '').lower())
    return datetime.date(int(m[3]), MESES[m[2]], int(m[1])) if m and m[2] in MESES else None
def medida(t):
    m = re.search(r'\b(\d{3})\s*[/\s]\s*(\d{2})\s*(?:Z?R|-)?\s*(\d{2})(?!\d)', t or '', re.I)
    return (f'{m[1]}/{m[2]}R{m[3]}', 'R' + m[3]) if m else (None, None)

ventas = []
for r in rows[h + 1:]:
    if not r or not r[0]: continue
    if re.search(r'cancel|devoluci|reembols|mediaci', str(r[C['estado']] or ''), re.I): continue
    uni, ing = num(r[C['uni']]), num(r[C['ing']])
    if uni <= 0 or ing <= 0: continue
    sku = str(r[C['sku']] or '').strip().upper()
    k = re.match(r'^KITX(\d+)-(.+)$', sku)
    kitn = int(k[1]) if k else 1
    md, rod = medida(str(r[C['tit']] or ''))
    envneto = -num(r[C['env']]) - num(r[C['ingenv']])
    ventas.append(dict(f=fecha(r[C['fecha']]), mla=str(r[C['mla']] or '').strip().upper(), sku=k[2] if k else sku,
                       md=md, rod=rod, cuotas=str(r[C['cuotas']] or '').strip().lower().startswith('s'),
                       com=-num(r[C['cargo']]) / ing, env=envneto / (uni * kitn) if envneto > 0 else None,
                       unidades=uni * kitn))

hasta = max(v['f'] for v in ventas if v['f'])
desde = hasta - datetime.timedelta(days=DIAS)

def agrupar(key, campo, min_n):
    grupos = {}
    for v in ventas:
        kv = key(v)
        if kv and v[campo] is not None:
            grupos.setdefault(kv, []).append(v)
    out = {}
    for kv, vs in grupos.items():
        rec = [v for v in vs if v['f'] and v['f'] >= desde]
        usar = rec if len(rec) >= min_n else vs     # recientes; si no alcanzan, todo el historial
        if len(usar) >= min_n:
            out[kv] = [round(st.median(v[campo] for v in usar), 4 if campo == 'com' else 0), len(usar)]
    return out

rec_sin_cuotas = [v['com'] for v in ventas if not v['cuotas'] and v['f'] and v['f'] >= desde]
rec_env = [v['env'] for v in ventas if v['rod'] and v['env'] and v['f'] and v['f'] >= desde]
data = {
    'generado': hasta.isoformat(),
    'dias': DIAS,
    'ventas': len(ventas),
    'comisionDefault': round(st.median(rec_sin_cuotas), 4),
    'envioDefault': round(st.median(rec_env)),
    'comisionPorMla': agrupar(lambda v: v['mla'], 'com', 1),
    'envioPorSku': agrupar(lambda v: v['sku'], 'env', 3),
    'envioPorMedida': agrupar(lambda v: v['md'], 'env', 3),
    'envioPorRodado': agrupar(lambda v: v['rod'], 'env', 1),
}
# Neumáticos vendidos por SKU en los últimos DIAS días (los kits cuentan cada neumático)
ventas_sku = {}
for v in ventas:
    if v['sku'] and v['f'] and v['f'] >= desde:
        ventas_sku[v['sku']] = ventas_sku.get(v['sku'], 0) + v['unidades']
data['ventasPorSku'] = {k: int(round(x)) for k, x in ventas_sku.items()}
with open(dst, 'w', encoding='utf-8', newline='\n') as f:
    f.write('// Comisión y envío de referencia de Mercado Libre, calculados del reporte "Ventas AR"\n')
    f.write(f'// (mediana de las ventas de los últimos {DIAS} días al {hasta.isoformat()}; si no alcanzan, todo el historial).\n')
    f.write('// Formato de cada entrada: [valor, cantidad de ventas usadas]. Generado automáticamente.\n')
    f.write('window.COSTOS_ML = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n')
print({k: (v if not isinstance(v, dict) else len(v)) for k, v in data.items()})
print(sorted(data['envioPorRodado'].items()))
