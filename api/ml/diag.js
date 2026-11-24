// Diagnóstico de la búsqueda de competidores: qué responde Mercado Libre en cada paso.
// GET /api/ml/diag?mla=MLA123   (una publicación nuestra)
// Solo devuelve códigos de estado y datos públicos del catálogo; nunca el token.
const ml = require('../_lib/ml');

const resumen = (r) => ({ estado: r.status, mensaje: r.data && (r.data.message || r.data.error) || (r.ok ? 'ok' : '') });

module.exports = async (req, res) => {
    const s = await ml.sesionValida(req, res);
    if (!s) return ml.json(res, 401, { error: 'No hay conexión con Mercado Libre.' });
    const url = new URL(req.url, 'http://x');
    const mla = String(url.searchParams.get('mla') || '').trim().toUpperCase().replace(/^MLA-?/, 'MLA');
    if (!/^MLA\d+$/.test(mla)) return ml.json(res, 400, { error: 'Pasá ?mla=MLA...' });
    const t = s.at;
    const pasos = {};
    try {
        const yo = await ml.mlGet('/users/me', t);
        pasos.usuario = { ...resumen(yo), id: yo.data && yo.data.id, nickname: yo.data && yo.data.nickname, estado_cuenta: yo.data && yo.data.status };

        const multi = await ml.mlGet(`/items?ids=${mla}&attributes=id,catalog_product_id,seller_id,status,catalog_listing`, t);
        const it = multi.ok && Array.isArray(multi.data) && multi.data[0] ? multi.data[0] : null;
        pasos.publicacion_multiget = { ...resumen(multi), codigo: it && it.code, cuerpo: it && it.body };

        const uno = await ml.mlGet(`/items/${mla}?attributes=id,catalog_product_id,seller_id,status`, t);
        pasos.publicacion_directa = { ...resumen(uno), cuerpo: uno.ok ? uno.data : undefined };

        const cat = (it && it.body && it.body.catalog_product_id) || (uno.data && uno.data.catalog_product_id);
        pasos.catalogo = cat || null;
        if (cat) {
            const prod = await ml.mlGet(`/products/${cat}/items?limit=5`, t);
            pasos.productos_items = {
                ...resumen(prod),
                total: prod.data && prod.data.paging ? prod.data.paging.total : undefined,
                ejemplos: prod.ok && prod.data && Array.isArray(prod.data.results)
                    ? prod.data.results.map(x => ({ item: x.item_id, vendedor: x.seller_id, precio: x.price })) : undefined
            };
            const p = await ml.mlGet(`/products/${cat}`, t);
            pasos.producto = { ...resumen(p), nombre: p.data && p.data.name, ganador: p.data && p.data.buy_box_winner ? p.data.buy_box_winner.item_id : null };
        }
        return ml.json(res, 200, { mla, pasos });
    } catch (e) {
        return ml.json(res, 502, { error: e.message, pasos });
    }
};
