// Mismo esquema de seguridad que services/_security.js en api-unificada (CORS por
// lista blanca, headers de seguridad, trust proxy + rate limit por IP). Servicio
// aparte, repo aparte → se duplica el módulo en vez de compartirlo entre repos.
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');

const origenesPermitidos = (process.env.ALLOWED_ORIGINS || '')
  .split(',').map(s => s.trim()).filter(Boolean);
const origenesDev = ['http://localhost:4200', 'http://127.0.0.1:4200'];
const listaOrigenes = origenesPermitidos.length ? origenesPermitidos : origenesDev;

const corsOptions = {
  origin(origin, cb) {
    if (!origin || listaOrigenes.includes(origin)) return cb(null, true);
    cb(null, false);
  },
};

const limiteGeneral = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 300,
  standardHeaders: true,
  legacyHeaders: false,
  message: { success: false, message: 'Demasiadas solicitudes. Intenta de nuevo en unos minutos.' },
});

function aplicarSeguridad(app) {
  app.set('trust proxy', 1);
  app.use(helmet({
    contentSecurityPolicy: false,
    crossOriginResourcePolicy: { policy: 'cross-origin' },
  }));
  app.use(cors(corsOptions));
  app.use(limiteGeneral);
}

module.exports = { aplicarSeguridad };
