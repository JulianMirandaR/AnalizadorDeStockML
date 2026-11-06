// Inicia el login con Mercado Libre (OAuth). Se abre en una ventana emergente desde la página.
const ml = require('../_lib/ml');

module.exports = (req, res) => {
    if (!ml.configurado()) {
        res.statusCode = 500;
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        return res.end(`<p style="font-family:sans-serif">Falta configurar <b>ML_CLIENT_ID</b> y <b>ML_CLIENT_SECRET</b> en Vercel.<br>
            URL de retorno a registrar en la app de Mercado Libre: <code>${ml.redirectUri(req)}</code></p>`);
    }
    const state = ml.nuevoState();
    ml.agregarCookie(res, ml.cookie(req, ml.STATE_COOKIE, state, 600));
    const url = `${ml.AUTH_URL}?` + new URLSearchParams({
        response_type: 'code',
        client_id: ml.CLIENT_ID(),
        redirect_uri: ml.redirectUri(req),
        state
    }).toString();
    res.statusCode = 302;
    res.setHeader('Location', url);
    res.end();
};
