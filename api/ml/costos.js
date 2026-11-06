// Costos reales de venta de NUESTRAS publicaciones, según Mercado Libre.
//
// POST { mlas: ['MLA...', ...] }   (hasta 40 por llamada)
// Responde { costos: { MLA: { precio, comisionPct, cargoFijo, envio, logistica, tipo, categoria } }, errores: {...} }
//   - comisionPct: % del cargo por vender (listing_prices de la categoría y tipo de publicación)
//   - cargoFijo:   cargo fijo por unidad que ML suma en precios bajos (0 en neumáticos)
//   - envio:       lo que ML le cobra al vendedor por el envío gratis de esa publicación
const ml = require('../_lib/ml');

const MAX_MLAS = 40;

async function leerBody(req) {
    if (req.body && typeof req.body === 'object') return req.body;
    if (typeof req.body === 'string') return JSON.parse(req.body || '{}');
    const chunks = [];
    for await (const c of req) chunks.push(c);
    return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
}
const normMla = (s) => String(s || '').trim().toUpperCase().replace(/^MLA-?/, 'MLA');
const partir = (arr, n) => Array.from({ length: Math.ceil(arr.length / n) }, (_, i) => arr.slice(i * n, i * n + n));

// Cargo por vender para un precio, categoría y tipo de publicación
async function comision(precio, categoria, tipo, token) {
    const q = new URLSearchParams({ price: String(precio), category_id: categoria, listing_type_id: tipo, currency_id: 'ARS' });
    const r = await ml.mlGet(`/sites/MLA/listing_prices?${q}`, token);
    const d = Array.isArray(r.data) ? r.data.find(x => x.listing_type_id === tipo) || r.data[0] : r.data;
    if (!r.ok || !d || d.sale_fee_amount == null) return null;
    const det = d.sale_fee_details || {};
    const pct = det.percentage_fee != null ? det.percentage_fee / 100
        : (det.meli_percentage_fee != null ? det.meli_percentage_fee / 100 : d.sale_fee_amount / precio);
    return { pct, fijo: det.fixed_fee || 0, monto: d.sale_fee_amount };
}

// Costo del envío gratis que paga el vendedor por esa publicación.
// Se prueban variantes del recurso de ML; "intentos" guarda el resultado de cada una (diagnóstico).
async function envio(item, uid, token) {
    const intentos = [];
    const sh = item.shipping || {};
    const base = { item_price: String(item.price), listing_type_id: item.listing_type_id, mode: sh.mode || 'me2', condition: 'new' };
    if (sh.logistic_type) base.logistic_type = sh.logistic_type;
    const rutas = [
        `/users/${uid}/shipping_options/free?` + new URLSearchParams({ item_id: item.id, ...base }),
        `/users/${uid}/shipping_options/free?item_id=${item.id}`,
        `/items/${item.id}/shipping_options/free`
    ];
    if (sh.dimensions) rutas.push(`/users/${uid}/shipping_options/free?` + new URLSearchParams({ dimensions: sh.dimensions, ...base }));
    for (const path of rutas) {
        const r = await ml.mlGet(path, token);
        const d = r.data || {};
        const c = d.coverage && d.coverage.all_country;
        intentos.push({ ruta: path.split('?')[0] + (path.includes('dimensions=') ? '?dimensions' : path.includes('item_price') ? '?item_id+precio' : ''), estado: r.status, mensaje: d.message || d.error || (c ? 'ok' : '') });
        if (c && c.list_cost != null) return { costo: c.list_cost, intentos };
    }
    return { costo: null, intentos };
}

module.exports = async (req, res) => {
    if (req.method !== 'POST') return ml.json(res, 405, { error: 'Usá POST' });
    const s = await ml.sesionValida(req, res);
    if (!s) return ml.json(res, 401, { error: 'No hay conexión con Mercado Libre.' });
    let body;
    try { body = await leerBody(req); } catch (e) { return ml.json(res, 400, { error: 'Body inválido' }); }
    const mlas = [...new Set((body.mlas || []).map(normMla).filter(m => /^MLA\d+$/.test(m)))].slice(0, MAX_MLAS);
    const token = s.at;
    const costos = {}, errores = {};
    try {
        const items = {};
        await ml.mapLimit(partir(mlas, 20), 3, async (grupo) => {
            const r = await ml.mlGet(`/items?ids=${grupo.join(',')}&attributes=id,price,category_id,listing_type_id,shipping,seller_id`, token);
            if (r.ok && Array.isArray(r.data)) r.data.forEach(x => { if (x && x.code === 200 && x.body) items[x.body.id] = x.body; });
        });
        await ml.mapLimit(mlas, 6, async (id) => {
            const it = items[id];
            if (!it) { errores[id] = 'no_encontrada'; return; }
            const uid = it.seller_id || s.uid;
            const [com, env] = await Promise.all([
                comision(it.price, it.category_id, it.listing_type_id, token),
                envio(it, uid, token)
            ]);
            costos[id] = {
                precio: it.price,
                comisionPct: com ? com.pct : null,
                cargoFijo: com ? com.fijo : null,
                envio: env.costo,
                ...(env.costo == null ? { envioIntentos: env.intentos } : {}),
                logistica: it.shipping && it.shipping.logistic_type || null,
                envioGratis: !!(it.shipping && it.shipping.free_shipping),
                tipo: it.listing_type_id,
                categoria: it.category_id
            };
        });
        return ml.json(res, 200, { costos, errores });
    } catch (e) {
        return ml.json(res, 502, { error: 'Error consultando Mercado Libre: ' + e.message });
    }
};
