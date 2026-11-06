// Vuelta del login de Mercado Libre: canjea el código por los tokens, los guarda en la cookie
// y avisa a la página que abrió la ventana emergente.
const ml = require('../_lib/ml');

const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function pagina(res, ok, mensaje) {
    res.statusCode = ok ? 200 : 400;
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    const payload = JSON.stringify({ tipo: 'ml-login', ok, mensaje: String(mensaje || '') }).replace(/</g, '\\u003c');
    res.end(`<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Mercado Libre</title></head>
<body style="font-family:sans-serif;background:#0f172a;color:#f8fafc;display:flex;align-items:center;justify-content:center;height:100vh;margin:0">
<p>${ok ? '✅ Conectado con Mercado Libre. Ya podés cerrar esta ventana.' : '❌ ' + esc(mensaje)}</p>
<script>
  try { if (window.opener) { window.opener.postMessage(${payload}, location.origin); ${ok ? 'setTimeout(function(){ window.close(); }, 800);' : ''} } } catch (e) {}
</script></body></html>`);
}

module.exports = async (req, res) => {
    const url = new URL(req.url, 'http://x');
    const code = url.searchParams.get('code');
    const state = url.searchParams.get('state');
    const error = url.searchParams.get('error');
    ml.agregarCookie(res, ml.cookie(req, ml.STATE_COOKIE, '', 0));
    if (error) return pagina(res, false, 'Mercado Libre rechazó la autorización: ' + error);
    if (!code) return pagina(res, false, 'Falta el código de autorización.');
    if (!state || state !== ml.leerCookies(req)[ml.STATE_COOKIE]) return pagina(res, false, 'La sesión de login venció. Volvé a intentar.');
    if (!ml.configurado()) return pagina(res, false, 'Faltan ML_CLIENT_ID / ML_CLIENT_SECRET en Vercel.');
    try {
        const t = await ml.pedirToken({ grant_type: 'authorization_code', code, redirect_uri: ml.redirectUri(req) });
        const yo = await ml.mlGet('/users/me', t.access_token);
        ml.guardarSesion(req, res, ml.sesionDeToken(t, { nick: yo.ok && yo.data ? yo.data.nickname : '' }));
        return pagina(res, true);
    } catch (e) {
        return pagina(res, false, 'No se pudo completar el login: ' + e.message);
    }
};
