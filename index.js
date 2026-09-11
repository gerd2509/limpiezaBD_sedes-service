// ─────────────────────────────────────────────────────────────────────────────
// limpiezaBD_sedes-service — Microservicio Leoncito
// Procesos pesados que antes corrían en el navegador (dashboard). Empieza con la
// LIMPIEZA/DEDUP de teléfonos por DNI (módulo Limpieza BBDD). Luego crecerá con el
// cruce cartera↔gestión por DNI+teléfono (avance-cartera, embudos, comparativo).
// ─────────────────────────────────────────────────────────────────────────────
const express = require('express');
const cors = require('cors');
const compression = require('compression');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 4002;

app.use(compression());   // gzip: comprime las respuestas JSON (5-10× menos bytes)
app.use(cors());
app.use(express.json({ limit: '25mb' }));

// Subida de archivos en memoria (hasta 60 MB por Excel).
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 60 * 1024 * 1024 } });

// ── Config (igual que el módulo Limpieza BBDD) ──────────────────────────────
const COLUMNAS_TELEFONO = [
  'CEL actual', 'Telefono', 'Fax', 'Movil', 'Nextel',
  'ProvedorExterno1', 'ProvedorExterno2', 'ProvedorExterno3',
];
// Fragmentos de cabecera que identifican una columna de TELÉFONOS (además de las conocidas).
const FRAG_TEL = [
  'celular', 'telefono', 'movil', 'fono', 'whatsapp', 'nextel', 'fax',
  'provedorexterno', 'proveedorexterno', 'cel', 'numero', 'numeros',
];
// Cabeceras que NUNCA son teléfono (evita falsos positivos con DNI/documento).
const FRAG_NO_TEL = ['documento', 'dni', 'docidentidad'];
// Mínimo de dígitos para considerar un número como teléfono (2-3 díg. = basura; celular=9, fijo=7).
const MIN_TEL_DIGITOS = 6;
const COLUMNA_DNI = 'Dni';

// ── Helpers ─────────────────────────────────────────────────────────────────
const norm = (s) => (s ?? '').toString().toLowerCase().trim().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, '');
function buscarHeader(headers, objetivo) {
  const o = norm(objetivo);
  return headers.find((h) => norm(h) === o);
}
// Detecta TODAS las columnas de teléfono: conocidas + cabecera con fragmento de teléfono; excluye DNI.
function detectarTelHeaders(headers, dniHeader) {
  const conocidas = new Set(COLUMNAS_TELEFONO.map((c) => norm(c)));
  return headers.filter((h) => {
    if (dniHeader && h === dniHeader) return false;
    const n = norm(h);
    if (FRAG_NO_TEL.some((f) => n.includes(f))) return false;
    if (conocidas.has(n)) return true;
    return FRAG_TEL.some((f) => n.includes(f));
  });
}
// Separa dígitos pegados sin separador: celulares Perú de 9 dígitos que empiezan en 9.
function separarConcatenados(d) {
  if (d.length <= 11) return [d];
  const bloques = [];
  let i = 0;
  while (i + 9 <= d.length && d[i] === '9') { bloques.push(d.slice(i, i + 9)); i += 9; }
  return (i === d.length && bloques.length > 1) ? bloques : [d];
}
function extraerNumeros(valor) {
  if (valor === null || valor === undefined) return [];
  const t = valor.toString().trim();
  if (!t) return [];
  const out = [];
  for (const parte of t.split(/[\s/,;|&·\-]+/)) {
    const d = parte.replace(/\D/g, '');
    if (!d) continue;
    for (const n of separarConcatenados(d)) {
      if (n.length >= MIN_TEL_DIGITOS) out.push(n);   // descarta números demasiado cortos (2-3 díg.)
    }
  }
  return out;
}

// ── Núcleo: dedup de teléfonos por DNI (modo "sedes") ───────────────────────
function limpiarSedes(filas) {
  if (!filas.length) throw new Error('El archivo no tiene filas.');

  const headersOriginales = Object.keys(filas[0]);
  const dniHeader = buscarHeader(headersOriginales, COLUMNA_DNI);
  const telHeaders = detectarTelHeaders(headersOriginales, dniHeader);

  if (!telHeaders.length) throw new Error('No se encontró ninguna columna de teléfonos para limpiar.');

  const headersBase = headersOriginales.filter((h) => !telHeaders.includes(h));

  const grupos = new Map(); // clave -> { base, numeros[] }
  let totalNumerosAntes = 0;

  filas.forEach((fila, idx) => {
    const dniRaw = dniHeader ? (fila[dniHeader] ?? '').toString().trim() : '';
    const clave = dniRaw !== '' ? `dni:${dniRaw}` : `fila:${idx}`;
    let g = grupos.get(clave);
    if (!g) { g = { base: { ...fila }, numeros: [] }; grupos.set(clave, g); }
    for (const h of telHeaders) {
      const nums = extraerNumeros(fila[h]);
      totalNumerosAntes += nums.length;
      for (const n of nums) if (!g.numeros.includes(n)) g.numeros.push(n);
    }
  });

  let maxNumeros = 1;
  grupos.forEach((g) => (maxNumeros = Math.max(maxNumeros, g.numeros.length)));

  const colsNumero = [];
  for (let i = 0; i < maxNumeros; i++) colsNumero.push(telHeaders[i] ?? `Telefono ${i + 1}`);

  const headersSalida = [...headersBase, ...colsNumero];
  const salida = [];
  let totalUnicos = 0;
  grupos.forEach((g) => {
    const fila = {};
    for (const h of headersBase) fila[h] = g.base[h] ?? '';
    colsNumero.forEach((col, i) => (fila[col] = g.numeros[i] ?? ''));
    salida.push(fila);
    totalUnicos += g.numeros.length;
  });

  return {
    headers: headersSalida,
    filas: salida,
    dniHeader,
    colsNumero,
    stats: {
      totalFilasLeidas: filas.length,
      totalDnis: grupos.size,
      numerosAntes: totalNumerosAntes,
      numerosUnicos: totalUnicos,
      duplicadosEliminados: Math.max(0, totalNumerosAntes - totalUnicos),
      maxNumerosPorDni: maxNumeros,
    },
  };
}

// Lee las filas del cuerpo (JSON {filas:[...]}) o del Excel subido (campo `archivo`).
function filasDe(req) {
  if (Array.isArray(req.body?.filas)) return req.body.filas;
  if (req.file?.buffer) {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const hoja = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(hoja, { defval: '', raw: false });
  }
  return null;
}

// ── Endpoints ───────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ ok: true, service: 'limpiezaBD_sedes-service', ts: new Date().toISOString() }));

// POST /limpieza/sedes  (multipart `archivo`  o  JSON {filas})  [?format=xlsx]
app.post('/limpieza/sedes', upload.single('archivo'), (req, res) => {
  try {
    const filas = filasDe(req);
    if (!filas || !filas.length) {
      return res.status(400).json({ ok: false, message: 'Envía un Excel (campo "archivo") o un JSON { filas: [...] }.' });
    }
    const r = limpiarSedes(filas);

    // Descarga directa del Excel limpio.
    if ((req.query.format || '').toString().toLowerCase() === 'xlsx') {
      const ws = XLSX.utils.json_to_sheet(r.filas, { header: r.headers });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'BASE LIMPIA');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="base_limpia.xlsx"');
      return res.send(buf);
    }

    // Preview (primeras 8) para mostrar en la UI.
    const preview = r.filas.slice(0, 8).map((f) => ({
      dni: (f[r.dniHeader] ?? '').toString(),
      numeros: r.colsNumero.map((c) => (f[c] ?? '').toString()).filter((v) => v !== ''),
    }));

    res.json({ ok: true, stats: r.stats, headers: r.headers, filas: r.filas, preview });
  } catch (e) {
    console.error('❌ /limpieza/sedes:', e.message);
    res.status(500).json({ ok: false, message: e.message || 'No se pudo procesar el archivo.' });
  }
});

app.listen(PORT, () => console.log(`✅ limpiezaBD_sedes-service escuchando en http://localhost:${PORT}`));
