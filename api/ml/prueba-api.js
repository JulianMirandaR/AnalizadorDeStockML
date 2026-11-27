// Prueba TEMPORAL de la API de ML: ¿esta app puede publicar? Este endpoint NO publica nada:
// las escrituras solo van a POST /items/validate (ML valida el cuerpo y no crea la publicación).
//
// POST /api/ml/prueba-api   { "get": ["/users/me", "/categories/MLA1234"], "validate": { ...cuerpo de un item... } }
//   get      lista de rutas de lectura de la API de ML (solo las permitidas abajo)
//   validate cuerpo para /items/validate
// Responde por cada ruta el código de estado y los datos; nunca devuelve el token.
const ml = require('../_lib/ml');

// Lecturas permitidas (solo consultas de datos propios y públicos del catálogo / categorías).
const GET_OK = [
    /^\/users\/me(\?|$)/,
    /^\/users\/\d+\/items\/search\?/,
    /^\/users\/\d+\/shipping_options\/free\?/,
    /^\/items\/MLA\d+(\?|$)/,
    /^\/sites\/MLA\/(domain_discovery\/search|listing_prices|search|listing_types)(\?|$)/,
    /^\/categories\/MLA[\w-]+(\/attributes|\/technical_specs\/input)?(\?|$)/,
    /^\/products\/(search\?|MLA\d+(\?|$))/,
    /^\/catalog_domains\/MLA-[\w-]+(\/attributes)?(\?|$)/,
    /^\/domains\/MLA-[\w-]+(\?|$)/
];

async function mlPost(path, cuerpo, token) {
    const r = await fetch(ml.API + path, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify(cuerpo)
    });
    const texto = await r.text();
    let data = null;
    try { data = texto ? JSON.parse(texto) : null; } catch (e) { data = texto.slice(0, 2000); }
    return { status: r.status, ok: r.ok, data };
}

module.exports = async (req, res) => {
    if (req.method !== 'POST') return ml.json(res, 405, { error: 'Usar POST con JSON.' });
    const s = await ml.sesionValida(req, res);
    if (!s) return ml.json(res, 401, { error: 'No hay conexión con Mercado Libre.' });
    let pedido = req.body;
    if (typeof pedido === 'string') { try { pedido = JSON.parse(pedido); } catch (e) { pedido = null; } }
    if (!pedido || typeof pedido !== 'object') return ml.json(res, 400, { error: 'Cuerpo JSON inválido.' });

    const salida = { get: {} };
    const lecturas = Array.isArray(pedido.get) ? pedido.get.slice(0, 20) : [];
    for (const ruta of lecturas) {
        const p = String(ruta);
        if (p.includes('..') || !GET_OK.some(re => re.test(p))) { salida.get[p] = { estado: 0, error: 'ruta no permitida' }; continue; }
        const r = await ml.mlGet(p, s.at);
        salida.get[p] = { estado: r.status, datos: r.data };
    }
    if (pedido.validate && typeof pedido.validate === 'object') {
        const r = await mlPost('/items/validate', pedido.validate, s.at);
        salida.validate = { estado: r.status, datos: r.data };
    }
    return ml.json(res, 200, salida);
};
