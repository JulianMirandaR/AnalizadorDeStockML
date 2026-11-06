// Estado de la conexión con Mercado Libre (para mostrar en la página).
// Además de mirar la cookie, se verifica con ML que el token siga vigente: si la
// autorización fue revocada (por ejemplo desde "Administrar permisos"), se borra la sesión.
const ml = require('../_lib/ml');

module.exports = async (req, res) => {
    const s = await ml.sesionValida(req, res);
    if (s) {
        const yo = await ml.mlGet('/users/me', s.at).catch(() => ({ ok: false, status: 0 }));
        if (yo.ok || (yo.status !== 401 && yo.status !== 403)) {
            return ml.json(res, 200, { conectado: true, nickname: s.nick || '', userId: s.uid });
        }
        ml.borrarSesion(req, res);
    }
    return ml.json(res, 200, { conectado: false, configurado: ml.configurado(), redirectUri: ml.redirectUri(req) });
};
