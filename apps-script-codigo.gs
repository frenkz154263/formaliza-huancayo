// Formaliza Fácil Huancayo · servidor de registros, cuentas y avance (versión 2)
// Pega este código en script.google.com (reemplaza todo) y vuelve a implementar.
const SHEET_ID = '1f_loxN6qfviiynSPeKaWlVNs3R3eei9AaSlI7xE5JY8';
const COLS = ['fecha','correo','nombre','rubro','distrito','donde','ruc','licencia','comprobantes','marca','trabajadores','plan','estado','ultimo','proximo','notas','paso1','paso2','paso3','paso4','paso5','paso6','paso7','paso8','mensaje'];
// Lo único que la página puede escribir. Plan, estado, contacto, notas, pasos y mensaje son solo del equipo (se editan en la hoja).
const PERMITIDOS = ['nombre','rubro','distrito','donde','ruc','licencia','comprobantes','marca','trabajadores'];

const CONTENIDO_BASE = [
 [
  "aviso",
  "",
  "Aviso destacado arriba de la página. Déjalo vacío para ocultarlo."
 ],
 [
  "precio",
  "49.90",
  "Precio mensual de la suscripción. Solo el número, por ejemplo 49.90."
 ],
 [
  "yape",
  "",
  "Número de Yape para el cobro."
 ],
 [
  "plin",
  "",
  "Número de Plin para el cobro."
 ],
 [
  "pago_nombre",
  "",
  "Nombre al que sale el pago."
 ],
 [
  "pago_link",
  "",
  "Enlace de pago en línea (opcional). Debe empezar con https."
 ],
 [
  "whatsapp",
  "https://chat.whatsapp.com/FvrUnBIk67r0WWnknmwoha",
  "Enlace del grupo o chat de WhatsApp."
 ],
 [
  "video1",
  "1SD3qv0ypatdMAz-XdjYeZY3CBkM8dEo3",
  "Video 1: enlace de Drive o su ID. El archivo debe estar compartido con «Cualquier persona con el enlace»."
 ],
 [
  "video2",
  "1n2nBQC89SJ0S0uDVGSj_k6WghAnWLPtL",
  "Video 2: enlace de Drive o su ID."
 ],
 [
  "inicio_texto",
  "Una plataforma que diagnostica a pequeños emprendedores de Huancayo y arma su ruta personal para formalizarse.",
  "Párrafo de Inicio, debajo del botón «Nuestro proceso»."
 ],
 [
  "casos_titulo",
  "No te expongas a estos casos",
  "Título de la sección de videos y fotos en Inicio."
 ],
 [
  "ben_titulo",
  "Beneficios de formalizar tu negocio",
  "Beneficios: título de la página."
 ],
 [
  "ben_intro",
  "Formalizar tu negocio es dar un paso hacia un futuro más seguro y con mayores oportunidades. No importa si tu negocio es pequeño, mediano o grande: estar formalizado te permite trabajar con mayor seguridad, acceder a nuevas oportunidades y hacer crecer tu emprendimiento.",
  "Beneficios: párrafo de introducción."
 ],
 [
  "ben_frase1",
  "Formalizarte hoy es abrir nuevas oportunidades para mañana.",
  "Beneficios: primera frase destacada."
 ],
 [
  "ben_frase2",
  "Tu negocio puede empezar pequeño, pero tus oportunidades no tienen por qué serlo.",
  "Beneficios: segunda frase destacada."
 ],
 [
  "ben1_t",
  "Acceder a créditos y financiamiento",
  "Beneficio 1: título."
 ],
 [
  "ben1_d",
  "para invertir y hacer crecer tu negocio.",
  "Beneficio 1: descripción."
 ],
 [
  "ben2_t",
  "Trabajar con empresas y acceder a nuevos mercados",
  "Beneficio 2: título."
 ],
 [
  "ben2_d",
  "que requieren proveedores formales.",
  "Beneficio 2: descripción."
 ],
 [
  "ben3_t",
  "Generar mayor confianza",
  "Beneficio 3: título."
 ],
 [
  "ben3_d",
  "en tus clientes y proveedores.",
  "Beneficio 3: descripción."
 ],
 [
  "ben4_t",
  "Proteger tu negocio",
  "Beneficio 4: título."
 ],
 [
  "ben4_d",
  "y desarrollar tus actividades dentro del marco legal.",
  "Beneficio 4: descripción."
 ],
 [
  "ben5_t",
  "Acceder a programas, capacitaciones y beneficios",
  "Beneficio 5: título."
 ],
 [
  "ben5_d",
  "destinados a emprendedores y empresas.",
  "Beneficio 5: descripción."
 ],
 [
  "ben6_t",
  "Construir una imagen profesional",
  "Beneficio 6: título."
 ],
 [
  "ben6_d",
  "y fortalecer la reputación de tu negocio.",
  "Beneficio 6: descripción."
 ],
 [
  "ben7_t",
  "Planificar tu crecimiento",
  "Beneficio 7: título."
 ],
 [
  "ben7_d",
  "con mayor orden y seguridad.",
  "Beneficio 7: descripción."
 ],
 [
  "ben_cierre",
  "¡La formalidad impulsa tus sueños!",
  "Beneficios: frase de cierre."
 ],
 [
  "pie_aviso",
  "Proyecto provisional. Información orientativa.",
  "Texto pequeño al final de la página."
 ]
];

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function main_() { return ss_().getSheets()[0]; }
function oculta_(nombre, encabezado) {
  let sh = ss_().getSheetByName(nombre);
  if (!sh) { sh = ss_().insertSheet(nombre); sh.appendRow(encabezado); if (nombre === 'Cuentas') sh.hideSheet(); }
  return sh;
}
function cuentas_() { return oculta_('Cuentas', ['correo', 'sal', 'huella', 'creada']); }
function pagos_() { return oculta_('Pagos', ['fecha', 'correo', 'código de operación', 'monto', 'estado', 'vigente hasta']); }
function fechaDe_(x) {
  if (x instanceof Date) return x;
  const m = String(x).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
}
function avisos_() { return oculta_('Avisos', ['fecha', 'correo', 'tipo']); }
function huella_(sal, clave) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + clave);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}
// Llave de sesión: se calcula con la huella guardada, no revela la contraseña.
function token_(huella) { return huella_('tk', huella); }
function filaDe_(sh, correo, col) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const v = sh.getRange(2, col, last - 1, 1).getValues().flat().map(x => String(x).toLowerCase());
  const i = v.indexOf(correo);
  return i < 0 ? -1 : i + 2;
}
function guardar_(p, correo, nuevo) {
  const sh = main_();
  let row = filaDe_(sh, correo, 2);
  if (row === -1) {
    row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setValue(new Date());
    sh.getRange(row, 2).setNumberFormat('@').setValue(correo);
    sh.getRange(row, 12).setValue(nuevo ? 'Gratis' : '');
    sh.getRange(row, 13).setValue('Nuevo');
  }
  PERMITIDOS.forEach(c => {
    const v = p[c];
    if (v === undefined || v === null || v === '') return;
    const max = 150;
    // Formato de texto: evita que algo escrito como "=..." se ejecute como fórmula.
    sh.getRange(row, COLS.indexOf(c) + 1).setNumberFormat('@').setValue(String(v).slice(0, max));
  });
  return row;
}
function salida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Pestaña «Contenido»: lo que el equipo escribe ahí se muestra en la página (se actualiza en ~1 minuto).
function contenido_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('contenido');
  if (hit) return JSON.parse(hit);
  let sh = ss_().getSheetByName('Contenido');
  if (!sh) {
    sh = ss_().insertSheet('Contenido');
    sh.getRange(1, 1, 1, 3).setValues([['clave (no la cambies)', 'valor (edítalo aquí)', 'qué es']]);
    sh.getRange(2, 1, CONTENIDO_BASE.length, 3).setNumberFormat('@').setValues(CONTENIDO_BASE);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 520); sh.setColumnWidth(3, 380); sh.setFrozenRows(1);
  }
  const n = sh.getLastRow(), c = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(r => {
    const k = String(r[0]).trim(), x = String(r[1]).trim();
    if (k && x) c[k] = x.slice(0, 1500);
  });
  const out = { ok: true, c: c };
  cache.put('contenido', JSON.stringify(out), 60);
  return out;
}

function doPost(e) {
  // El contenido público se entrega sin esperar turno ni pedir sesión.
  let p0 = {};
  try { p0 = JSON.parse(e.postData.contents); } catch (x) { return salida_({ ok: false }); }
  if (p0.accion === 'contenido') return salida_(contenido_());
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    const correo = String(p.correo || '').trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo) || correo.length > 120) return salida_({ ok: false, error: 'correo' });
    const a = p.accion;

    // Avisos de interés (pie de página y suscripción): van a otra hoja, no tocan cuentas.
    if (a === 'aviso') {
      avisos_().appendRow([new Date(), correo, String(p.tipo || '').slice(0, 40)]);
      return salida_({ ok: true });
    }

    const ct = cuentas_();
    if (a === 'registro') {
      if (filaDe_(ct, correo, 1) !== -1) return salida_({ ok: false, error: 'existe' });
      if (!p.clave || String(p.clave).length < 20) return salida_({ ok: false, error: 'sin-clave' });
      const sal = Utilities.getUuid();
      const h = huella_(sal, p.clave);
      ct.appendRow([correo, sal, h, new Date()]);
      guardar_(p, correo, true);
      return salida_({ ok: true, token: token_(h) });
    }

    const r = filaDe_(ct, correo, 1);
    if (r === -1) return salida_({ ok: false, error: 'nocuenta' });
    const [, sal, huella] = ct.getRange(r, 1, 1, 4).getValues()[0];

    if (a === 'login') {
      // Máximo 5 intentos fallidos por correo cada 15 minutos.
      const cache = CacheService.getScriptCache(), kf = 'f:' + correo, n = Number(cache.get(kf) || 0);
      if (n >= 5) return salida_({ ok: false, error: 'bloqueado' });
      if (huella_(sal, p.clave || '') !== huella) { cache.put(kf, String(n + 1), 900); return salida_({ ok: false, error: 'clave' }); }
      cache.remove(kf);
      const sh = main_();
      const fila = filaDe_(sh, correo, 2);
      const datos = {};
      if (fila !== -1) {
        const vals = sh.getRange(fila, 1, 1, COLS.length).getValues()[0];
        COLS.forEach((c, k) => { if (PERMITIDOS.indexOf(c) !== -1) datos[c] = vals[k] === '' ? '' : String(vals[k]); });
      }
      return salida_({ ok: true, token: token_(huella), datos: datos });
    }

    // Cualquier otro guardado (diagnóstico, avance, pago, estado) exige la llave de sesión de esa cuenta.
    if (String(p.token || '') !== token_(huella)) return salida_({ ok: false, error: 'auth' });

    // Estado de los 8 pasos: lo escribe el equipo en la hoja (columnas paso1 a paso8); aquí solo se lee.
    if (a === 'pasos') {
      const sh = main_(), fila = filaDe_(sh, correo, 2);
      const v = fila === -1 ? [] : sh.getRange(fila, COLS.indexOf('paso1') + 1, 1, 8).getValues()[0].map(x => String(x || '').trim());
      return salida_({ ok: true, pasos: v });
    }

    // La persona avisa que pagó: queda "Por verificar" hasta que el equipo lo apruebe en la hoja Pagos.
    if (a === 'pago') {
      const pg = pagos_(), fila = pg.getLastRow() + 1;
      pg.getRange(fila, 1, 1, 6).setNumberFormat('@').setValues([[Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm'), correo, String(p.codigo || '').slice(0, 40), '49.90', 'Por verificar', '']]);
      return salida_({ ok: true });
    }

    // Estado de consultas: solo se entrega si hay un pago Aprobado con fecha vigente.
    if (a === 'estado') {
      const pg = pagos_(), n = pg.getLastRow(), hoy = new Date();
      hoy.setHours(0, 0, 0, 0);
      let pago = '', paid = false, vig = '';
      if (n >= 2) {
        const v = pg.getRange(2, 1, n - 1, 6).getValues();
        for (let i = v.length - 1; i >= 0; i--) {
          if (String(v[i][1]).toLowerCase() !== correo) continue;
          const est = String(v[i][4]).trim();
          if (!pago) pago = est;
          const d = fechaDe_(v[i][5]);
          if (est.toLowerCase() === 'aprobado' && d && d >= hoy) { paid = true; vig = Utilities.formatDate(d, 'America/Lima', 'dd/MM/yyyy'); break; }
        }
      }
      if (!paid) return salida_({ ok: true, paid: false, pago: pago });
      const sh = main_(), fila = filaDe_(sh, correo, 2), out = { ok: true, paid: true, vigente: vig };
      if (fila !== -1) {
        const vals = sh.getRange(fila, 1, 1, COLS.length).getValues()[0];
        ['estado', 'ultimo', 'proximo', 'mensaje'].forEach(c => {
          const x = vals[COLS.indexOf(c)];
          out[c] = x instanceof Date ? Utilities.formatDate(x, 'America/Lima', 'dd/MM/yyyy') : String(x || '');
        });
      }
      return salida_(out);
    }
    guardar_(p, correo, false);
    return salida_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
