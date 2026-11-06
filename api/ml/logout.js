// Desconecta Mercado Libre (borra la cookie con los tokens).
const ml = require('../_lib/ml');

module.exports = (req, res) => {
    ml.borrarSesion(req, res);
    ml.json(res, 200, { ok: true });
};
