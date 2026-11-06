// Helpers compartidos para hablar con la API de Mercado Libre desde las funciones de Vercel.
// La sesión (tokens de ML) se guarda en una cookie httpOnly del navegador del usuario:
// no hay base de datos y el token nunca queda expuesto al JavaScript de la página.
const crypto = require('crypto');

const API = 'https://api.mercadolibre.com';
const AUTH_URL = 'https://auth.mercadolibre.com.ar/authorization';
const COOKIE = 'ml_auth';
const STATE_COOKIE = 'ml_state';
const MAX_AGE = 180 * 24 * 3600; // el refresh token de ML dura 6 meses

const CLIENT_ID = () => process.env.ML_CLIENT_ID;
const CLIENT_SECRET = () => process.env.ML_CLIENT_SECRET;
const configurado = () => !!(CLIENT_ID() && CLIENT_SECRET());

function baseUrl(req) {
    const proto = (req.headers['x-forwarded-proto'] || 'https').split(',')[0].trim();
    const host = (req.headers['x-forwarded-host'] || req.headers.host || '').split(',')[0].trim();
    return `${proto}://${host}`;
}
const redirectUri = (req) => process.env.ML_REDIRECT_URI || `${baseUrl(req)}/api/ml/callback`;

function leerCookies(req) {
    const out = {};
    String(req.headers.cookie || '').split(';').forEach(par => {
        const i = par.indexOf('=');
        if (i > 0) out[par.slice(0, i).trim()] = decodeURIComponent(par.slice(i + 1).trim());
    });
    return out;
}
const segura = (req) => baseUrl(req).startsWith('https') ? '; Secure' : '';
function cookie(req, nombre, valor, maxAge) {
    return `${nombre}=${encodeURIComponent(valor)}; Path=/api/ml; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${segura(req)}`;
}
function agregarCookie(res, valor) {
    const prev = res.getHeader('Set-Cookie');
    res.setHeader('Set-Cookie', [].concat(prev || [], valor));
}

function leerSesion(req) {
    const raw = leerCookies(req)[COOKIE];
    if (!raw) return null;
    try { return JSON.parse(Buffer.from(raw, 'base64url').toString('utf8')); } catch (e) { return null; }
}
function guardarSesion(req, res, s) {
    agregarCookie(res, cookie(req, COOKIE, Buffer.from(JSON.stringify(s)).toString('base64url'), MAX_AGE));
}
function borrarSesion(req, res) {
    agregarCookie(res, cookie(req, COOKIE, '', 0));
}

async function pedirToken(params) {
    const r = await fetch(`${API}/oauth/token`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
        body: new URLSearchParams({ client_id: CLIENT_ID(), client_secret: CLIENT_SECRET(), ...params }).toString()
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok || !data.access_token) {
        const e = new Error(data.message || data.error || `ML respondió ${r.status} al pedir el token`);
        e.status = r.status;
        throw e;
    }
    return data;
}
const sesionDeToken = (t, extra) => ({
    at: t.access_token, rt: t.refresh_token, exp: Date.now() + (t.expires_in || 21600) * 1000,
    uid: t.user_id, ...extra
});

// Devuelve una sesión con access token vigente (lo renueva si venció) o null si no hay login.
async function sesionValida(req, res) {
    const s = leerSesion(req);
    if (!s || !s.rt) return null;
    if (s.at && s.exp > Date.now() + 120000) return s;
    try {
        const t = await pedirToken({ grant_type: 'refresh_token', refresh_token: s.rt });
        const nueva = sesionDeToken(t, { uid: t.user_id || s.uid, nick: s.nick });
        guardarSesion(req, res, nueva);
        return nueva;
    } catch (e) {
        borrarSesion(req, res); // refresh token vencido o revocado: hay que volver a conectar
        return null;
    }
}

async function mlGet(path, token) {
    const r = await fetch(API + path, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
    const data = await r.json().catch(() => null);
    return { status: r.status, ok: r.ok, data };
}

// Ejecuta fn sobre cada elemento con como máximo "limite" llamadas en paralelo.
async function mapLimit(items, limite, fn) {
    const out = new Array(items.length);
    let i = 0;
    const workers = Array.from({ length: Math.min(limite, items.length) }, async () => {
        while (i < items.length) {
            const k = i++;
            out[k] = await fn(items[k], k);
        }
    });
    await Promise.all(workers);
    return out;
}

function json(res, status, obj) {
    res.statusCode = status;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    res.end(JSON.stringify(obj));
}

const nuevoState = () => crypto.randomBytes(16).toString('hex');

module.exports = {
    API, AUTH_URL, COOKIE, STATE_COOKIE, CLIENT_ID, configurado, redirectUri, leerCookies, cookie, agregarCookie,
    leerSesion, guardarSesion, borrarSesion, pedirToken, sesionDeToken, sesionValida, mlGet, mapLimit, json, nuevoState
};
