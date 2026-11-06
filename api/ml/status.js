// Estado de la conexión con Mercado Libre (para mostrar en la página).
const ml = require('../_lib/ml');

module.exports = async (req, res) => {
    const s = await ml.sesionValida(req, res);
    if (s) return ml.json(res, 200, { conectado: true, nickname: s.nick || '', userId: s.uid });
    return ml.json(res, 200, { conectado: false, configurado: ml.configurado(), redirectUri: ml.redirectUri(req) });
};
