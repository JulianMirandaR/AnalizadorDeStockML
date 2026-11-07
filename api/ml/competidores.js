// Busca los competidores de nuestras publicaciones en Mercado Libre.
//
// POST { skus: [{ sku, mlas: ['MLA...', ...] }], seguir: ['MLA...'] }
//   - skus: nuestras publicaciones de cada SKU. Se busca el producto de catálogo al que
//     pertenecen y se traen TODAS las publicaciones de otros vendedores en ese producto.
//   - seguir: publicaciones de competidores cargadas a mano, para actualizarles el precio.
// Responde { resultados: { SKU: { catalogo, competidores: [...] , motivo? } }, seguidos: { MLA: {...} } }
const ml = require('../_lib/ml');

const MAX_SKUS = 15;
const MAX_POR_PRODUCTO = 60;

async function leerBody(req) {
    if (req.body && typeof req.body === 'object') return req.body;
    if (typeof req.body === 'string') return JSON.parse(req.body || '{}');
    const chunks = [];
    for await (const c of req) chunks.push(c);
    return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
}
const normMla = (s) => String(s || '').trim().toUpperCase().replace(/^MLA-?/, 'MLA');
const partir = (arr, n) => Array.from({ length: Math.ceil(arr.length / n) }, (_, i) => arr.slice(i * n, i * n + n));

// /items?ids= admite hasta 20 publicaciones por llamada
async function multiItems(ids, atributos, token) {
    const mapa = {};
    await ml.mapLimit(partir(ids, 20), 4, async (grupo) => {
        const r = await ml.mlGet(`/items?ids=${grupo.join(',')}&attributes=${atributos}`, token);
        if (!r.ok || !Array.isArray(r.data)) return;
        r.data.forEach(x => { if (x && x.code === 200 && x.body) mapa[x.body.id] = x.body; });
    });
    return mapa;
}

// Todas las publicaciones activas de un producto de catálogo
async function itemsDeProducto(productId, token) {
    const out = [];
    let offset = 0;
    while (out.length < MAX_POR_PRODUCTO) {
        const r = await ml.mlGet(`/products/${productId}/items?limit=20&offset=${offset}`, token);
        if (!r.ok || !r.data || !Array.isArray(r.data.results)) {
            if (!out.length) return { error: r.status === 404 ? 'sin_ganador' : `error_${r.status}` };
            break;
        }
        out.push(...r.data.results);
        const total = r.data.paging ? r.data.paging.total : out.length;
        offset += r.data.results.length;
        if (!r.data.results.length || offset >= total) break;
    }
    return { items: out };
}

// Apodos de vendedores (para mostrar quién es cada competidor)
async function apodos(ids, token) {
    const mapa = {};
    await ml.mapLimit(ids, 6, async (id) => {
        const r = await ml.mlGet(`/users/${id}`, token);
        if (r.ok && r.data) mapa[id] = r.data.nickname || '';
    });
    return mapa;
}

// Búsqueda en el catálogo por título. Se usa para publicaciones que ML no deja leer
// (las de nuestra otra cuenta, que no es la conectada): se buscan productos de catálogo con
// el título y se quedan los que tienen la misma medida y el mismo modelo.
const sinAcentos = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
function medidaDe(t) {
    const m = sinAcentos(t).match(/(\d{3})\s*[\/\s]\s*(\d{2})\s*(?:Z?R|-)?\s*(\d{2})(?!\d)/);
    return m ? m[1] + '/' + m[2] + 'R' + m[3] : null;
}
function modelosDe(t) {
    const w = sinAcentos(t).replace(/[^A-Z0-9]+/g, ' ').split(' ');
    return new Set(w.filter(x => /^[A-Z]{1,4}\d{1,4}$/.test(x) && !/^(Z?R\d{2}|P\d{3}|LT\d*|X\d+)$/.test(x)));
}
// Primero con el título; si no aparece, con una búsqueda corta: marca + modelo + medida
const RELLENO = new Set(['NEUMATICO', 'NEUMATICOS', 'CUBIERTA', 'CUBIERTAS', 'LLANTA', 'NUEVO', 'AUTO', 'CAMIONETA', 'UTILITARIO',
    'SEGURIDAD', 'MAYOR', 'MAXIMO', 'RENDIMIENTO', 'DEPORTIVO', 'INDICE', 'VELOCIDAD', 'PARA', 'CARGA', 'KIT']);
async function productosPorTitulo(titulo, token) {
    const r1 = await productosPorConsulta(titulo, token);
    if (r1.ids) return r1;
    const md = medidaDe(titulo);
    if (!md) return r1;
    const marca = sinAcentos(titulo).replace(/[^A-Z]+/g, ' ').split(' ').filter(w => w.length >= 4 && !RELLENO.has(w));
    const corta = [...marca.slice(0, 2), ...modelosDe(titulo), md].join(' ');
    const r2 = await productosPorConsulta(corta, token);
    return r2.ids ? r2 : r1;
}
async function productosPorConsulta(titulo, token) {
    const q = sinAcentos(titulo).replace(/\bKIT\s*(X|DE)?\s*\d+\b/g, ' ').replace(/\s+/g, ' ').trim();
    const r = await ml.mlGet(`/products/search?status=active&site_id=MLA&limit=10&q=${encodeURIComponent(q)}`, token);
    if (!r.ok || !r.data || !Array.isArray(r.data.results)) return { error: `busqueda_${r.status}` };
    const md = medidaDe(q), mods = modelosDe(q);
    const ok = r.data.results.filter(p => {
        const n = p.name || '';
        if (md && medidaDe(n) !== md) return false;
        if (mods.size && ![...modelosDe(n)].some(x => mods.has(x))) return false;
        return true;
    });
    return ok.length ? { ids: ok.slice(0, 3).map(p => p.id) } : { error: r.data.results.length ? 'busqueda_sin_coincidencia' : 'busqueda_vacia' };
}

// Precio de venta real (con promociones) de una publicación
async function precioVenta(id, token) {
    const r = await ml.mlGet(`/items/${id}/sale_price?context=channel_marketplace`, token);
    return r.ok && r.data && r.data.amount ? { precio: r.data.amount, precioLista: r.data.regular_amount || null } : null;
}

module.exports = async (req, res) => {
    if (req.method !== 'POST') return ml.json(res, 405, { error: 'Usá POST' });
    const s = await ml.sesionValida(req, res);
    if (!s) return ml.json(res, 401, { error: 'No hay conexión con Mercado Libre. Tocá "Conectar Mercado Libre".' });

    let body;
    try { body = await leerBody(req); } catch (e) { return ml.json(res, 400, { error: 'Body inválido' }); }
    const skus = (Array.isArray(body.skus) ? body.skus : []).slice(0, MAX_SKUS)
        .map(x => ({ sku: String(x.sku || ''), mlas: [...new Set((x.mlas || []).map(normMla).filter(m => /^MLA\d+$/.test(m)))], titulo: String(x.titulo || '').slice(0, 150) }))
        .filter(x => x.sku && (x.mlas.length || x.titulo));
    const seguir = [...new Set((Array.isArray(body.seguir) ? body.seguir : []).map(normMla).filter(m => /^MLA\d+$/.test(m)))].slice(0, 100);
    const token = s.at;

    try {
        // 1) A qué producto de catálogo pertenece cada una de nuestras publicaciones
        const propias = await multiItems([...new Set(skus.flatMap(x => x.mlas))], 'id,catalog_product_id,seller_id', token);
        const nuestrasMlas = new Set(skus.flatMap(x => x.mlas));
        const vendedoresPropios = new Set([String(s.uid)]);
        Object.values(propias).forEach(it => it.seller_id && vendedoresPropios.add(String(it.seller_id)));

        const catalogoDeSku = {};
        const errorBusqueda = {};
        skus.forEach(x => {
            catalogoDeSku[x.sku] = [...new Set(x.mlas.map(m => propias[m] && propias[m].catalog_product_id).filter(Boolean))].slice(0, 3);
        });
        // Sin catálogo por las publicaciones (p. ej. cuenta secundaria): se busca por el título
        await ml.mapLimit(skus.filter(x => !catalogoDeSku[x.sku].length && x.titulo), 4, async (x) => {
            const r = await productosPorTitulo(x.titulo, token);
            if (r.ids) catalogoDeSku[x.sku] = r.ids;
            else errorBusqueda[x.sku] = r.error;
        });

        // 2) Publicaciones de todos los vendedores en esos productos
        const productos = [...new Set(Object.values(catalogoDeSku).flat())];
        const porProducto = {};
        await ml.mapLimit(productos, 4, async (p) => { porProducto[p] = await itemsDeProducto(p, token); });

        // 3) Apodos de los vendedores que aparecen
        const vendedores = new Set();
        Object.values(porProducto).forEach(r => (r.items || []).forEach(it => it.seller_id && vendedores.add(String(it.seller_id))));
        const nick = await apodos([...vendedores].filter(v => !vendedoresPropios.has(v)), token);

        const resultados = {};
        skus.forEach(x => {
            const cats = catalogoDeSku[x.sku];
            if (!cats.length) {
                resultados[x.sku] = { catalogo: [], competidores: [], motivo: errorBusqueda[x.sku] || (Object.keys(propias).some(m => x.mlas.includes(m)) ? 'sin_catalogo' : 'no_encontrada') };
                return;
            }
            const vistos = new Set();
            const competidores = [];
            let motivo = null;
            cats.forEach(p => {
                const r = porProducto[p] || {};
                if (r.error) { motivo = motivo || r.error; return; }
                r.items.forEach(it => {
                    const id = normMla(it.item_id);
                    const seller = String(it.seller_id || '');
                    if (vistos.has(id) || nuestrasMlas.has(id) || vendedoresPropios.has(seller)) return;
                    if (!(it.price > 0)) return;
                    vistos.add(id);
                    competidores.push({
                        mla: id,
                        precio: it.price,
                        precioLista: it.original_price || null,
                        sellerId: seller,
                        vendedor: nick[seller] || '',
                        envioGratis: !!(it.shipping && it.shipping.free_shipping),
                        full: !!(it.shipping && it.shipping.logistic_type === 'fulfillment'),
                        tienda: !!it.official_store_id,
                        producto: p
                    });
                });
            });
            competidores.sort((a, b) => a.precio - b.precio);
            resultados[x.sku] = { catalogo: cats, competidores, ...(competidores.length ? {} : { motivo: motivo || 'sin_competidores' }) };
        });

        // 4) Precio actual de los competidores cargados a mano
        const seguidos = {};
        if (seguir.length) {
            const info = await multiItems(seguir, 'id,price,seller_id,status', token);
            const nickSeg = await apodos([...new Set(Object.values(info).map(i => String(i.seller_id)).filter(v => v && !nick[v]))], token);
            await ml.mapLimit(seguir, 6, async (id) => {
                const it = info[id];
                if (!it) { seguidos[id] = { error: 'no_encontrada' }; return; }
                const pv = await precioVenta(id, token);
                seguidos[id] = {
                    precio: pv ? pv.precio : it.price,
                    precioLista: pv ? pv.precioLista : null,
                    estado: it.status,
                    vendedor: nick[String(it.seller_id)] || nickSeg[String(it.seller_id)] || ''
                };
            });
        }

        return ml.json(res, 200, { resultados, seguidos });
    } catch (e) {
        return ml.json(res, 502, { error: 'Error consultando Mercado Libre: ' + e.message });
    }
};
