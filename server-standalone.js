const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { URL } = require('url');
const { exec, spawn } = require('child_process');
const XLSX = require('xlsx');

// La versión sale de package.json (pkg lo empaqueta junto con el resto). El
// front la usa para marcar con un punto verde cuando hay novedades sin leer:
// al subir la versión acá, agregar la entrada en la solapa "Novedades" de
// public/index.html.
const { version: APP_VERSION } = require('./package.json');

// --- Persistencia en disco (escritura atómica + respaldo) ---
// La PC del depósito se puede apagar de golpe. Un writeFileSync interrumpido
// deja el JSON cortado a la mitad, y si además arrancáramos con lista vacía,
// el primer guardado siguiente pisaría el archivo dañado y se perdería todo
// el historial. Por eso: se escribe en un temporal, se conserva la última
// versión buena en .bak, y recién ahí se reemplaza el archivo real.
function escribirJsonAtomico(rutaFinal, datos) {
  const rutaTmp = `${rutaFinal}.tmp`;
  const rutaBak = `${rutaFinal}.bak`;
  const contenido = JSON.stringify(datos, null, 2);

  const fd = fs.openSync(rutaTmp, 'w');
  try {
    fs.writeFileSync(fd, contenido, 'utf8');
    fs.fsyncSync(fd); // que llegue al disco antes de reemplazar nada
  } finally {
    fs.closeSync(fd);
  }

  try {
    if (fs.existsSync(rutaFinal)) fs.copyFileSync(rutaFinal, rutaBak);
  } catch (err) {
    console.log(`  ⚠ No se pudo respaldar ${path.basename(rutaFinal)}: ${err.message}`);
  }

  fs.renameSync(rutaTmp, rutaFinal); // atómico dentro del mismo volumen
}

// Lee un JSON tolerando que esté dañado: si el archivo principal no parsea,
// prueba con el .bak; si tampoco, aparta el corrupto con otro nombre (nunca
// lo pisa en silencio) y devuelve el valor por defecto.
function leerJsonSeguro(ruta, valorPorDefecto, validar) {
  for (const candidato of [ruta, `${ruta}.bak`]) {
    let crudo;
    try {
      crudo = fs.readFileSync(candidato, 'utf8');
    } catch {
      continue; // no existe: probamos el siguiente
    }
    try {
      const datos = JSON.parse(crudo);
      if (validar && !validar(datos)) throw new Error('el contenido no tiene el formato esperado');
      if (candidato !== ruta) {
        console.log(`  ⚠ ${path.basename(ruta)} estaba dañado; se recuperó desde el respaldo .bak`);
      }
      return datos;
    } catch (err) {
      const apartado = `${candidato}.dañado-${new Date().toISOString().replace(/[:.]/g, '-')}`;
      try {
        fs.renameSync(candidato, apartado);
        console.log(`  ⚠ No se pudo leer ${path.basename(candidato)} (${err.message}). Se guardó como ${path.basename(apartado)} para no perder los datos.`);
      } catch { /* si no se puede mover, al menos no lo sobreescribimos sin avisar */ }
    }
  }
  return valorPorDefecto;
}

// --- Catálogo de productos (EAN / DUN / Factor) ---
// Se parte del catálogo por defecto embebido en la app; si el usuario sube
// una tabla actualizada desde la interfaz, esa queda guardada junto al
// ejecutable (catalogo.json) y tiene prioridad a partir de ahí.
const catalogoDefault = require('./products.json');

function normalizar(s) {
  return (s || '').toString().trim().toUpperCase().replace(/\s+/g, ' ');
}

let productos = [];
let catalogoInfo = { origen: 'default', actualizado: null, archivo: null, cantidad: 0 };
let indicePorSku = new Map();
let indicePorNombre = new Map();

function reconstruirIndices() {
  indicePorSku = new Map();
  indicePorNombre = new Map();
  for (const p of productos) {
    if (p.sku) indicePorSku.set(normalizar(p.sku), p);
    if (p.nombre) {
      const key = normalizar(p.nombre);
      if (indicePorNombre.has(key)) indicePorNombre.set(key, 'AMBIGUO');
      else indicePorNombre.set(key, p);
    }
  }
}

function cargarCatalogoInicial() {
  const catalogoGuardadoPath = path.join(APP_DIR, 'catalogo.json');
  const raw = leerJsonSeguro(
    catalogoGuardadoPath,
    null,
    d => d && Array.isArray(d.productos) && d.productos.length > 0
  );
  if (raw) {
    productos = raw.productos;
    catalogoInfo = { origen: 'subido', actualizado: raw.actualizado, archivo: raw.archivo, cantidad: productos.length };
  } else {
    productos = catalogoDefault;
    catalogoInfo = { origen: 'default', actualizado: null, archivo: null, cantidad: productos.length };
  }
  reconstruirIndices();
}

function guardarCatalogo(nuevosProductos, nombreArchivo) {
  productos = nuevosProductos;
  catalogoInfo = {
    origen: 'subido',
    actualizado: new Date().toISOString(),
    archivo: nombreArchivo,
    cantidad: productos.length
  };
  reconstruirIndices();
  const catalogoGuardadoPath = path.join(APP_DIR, 'catalogo.json');
  escribirJsonAtomico(catalogoGuardadoPath, {
    productos,
    actualizado: catalogoInfo.actualizado,
    archivo: nombreArchivo
  });
}

function limpiarCodigo(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return String(Math.trunc(v));
  return String(v).trim() || null;
}

function limpiarFactor(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : Math.trunc(n);
}

// Encuentra la columna correcta sin importar mayúsculas/espacios exactos
function buscarColumna(row, nombresPosibles) {
  const keys = Object.keys(row);
  for (const nombre of nombresPosibles) {
    const match = keys.find(k => normalizar(k) === normalizar(nombre));
    if (match) return row[match];
  }
  return null;
}

function parsearTablaProductos(buffer, nombreArchivo) {
  const esCSV = /\.csv$/i.test(nombreArchivo);
  let filas;

  if (esCSV) {
    const texto = buffer.toString('utf8');
    const lineas = texto.split(/\r?\n/).filter(l => l.trim() !== '');
    if (!lineas.length) throw new Error('El archivo CSV está vacío.');
    const headers = lineas[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    filas = lineas.slice(1).map(linea => {
      const valores = linea.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
      const obj = {};
      headers.forEach((h, i) => { obj[h] = valores[i]; });
      return obj;
    });
  } else {
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    filas = XLSX.utils.sheet_to_json(sheet, { defval: null });
  }

  const nuevosProductos = [];
  for (const row of filas) {
    const sku = buscarColumna(row, ['SKU']);
    if (!sku) continue;
    nuevosProductos.push({
      sku: String(sku).trim(),
      nombre: (buscarColumna(row, ['Nombre']) || '').toString().trim() || null,
      ean: limpiarCodigo(buscarColumna(row, ['EAN'])),
      dun: limpiarCodigo(buscarColumna(row, ['DUN'])),
      factor: limpiarFactor(buscarColumna(row, ['Factor']))
    });
  }

  if (!nuevosProductos.length) {
    throw new Error('No se encontraron filas válidas (revisá que exista una columna "SKU").');
  }

  return nuevosProductos;
}

function buscarProducto(codigo, concepto) {
  if (codigo) {
    const p = indicePorSku.get(normalizar(codigo));
    if (p) return p;
  }
  if (concepto) {
    const p = indicePorNombre.get(normalizar(concepto));
    if (p && p !== 'AMBIGUO') return p;
  }
  return null;
}

// Cuando se empaqueta con pkg, __dirname apunta al snapshot; usamos la carpeta
// real del ejecutable para poder leer/escribir config.json y para servir /public.
const APP_DIR = process.pkg ? path.dirname(process.execPath) : __dirname;
const CONFIG_PATH = path.join(APP_DIR, 'config.json');
const ARMADOS_PATH = path.join(APP_DIR, 'armados.json');
const CLIENTES_CACHE_PATH = path.join(APP_DIR, 'clientes_cache.json');
const PUBLIC_DIR = path.join(__dirname, 'public');
const LOGO_PATH = path.join(APP_DIR, 'assets', 'logo_suprabond.png');
const TEMP_DIR = path.join(os.tmpdir(), 'picking-contabilium-etiquetas');

const BASE_URL = 'https://rest.contabilium.com.uy';
const PORT = 3000;

// Por defecto la app escucha SOLO en la propia PC. Si en el depósito la
// quieren usar desde otra máquina de la red, agregar "host": "0.0.0.0" en
// config.json — pero ojo, en ese caso cualquiera que llegue al puerto 3000
// entra a la pantalla de login.
const HOST_DEFECTO = '127.0.0.1';

// Logo del encabezado de la etiqueta (se carga una vez al arrancar)
let logoBase64 = null;
try {
  logoBase64 = fs.readFileSync(LOGO_PATH).toString('base64');
} catch {
  console.log('  ⚠ No se encontró el logo en assets/logo_suprabond.png; se usará el encabezado de texto.');
}

// --- Pedidos armados (persistencia + consulta) ---
let armados = [];

function cargarArmadosInicial() {
  armados = leerJsonSeguro(ARMADOS_PATH, [], Array.isArray);
}

function guardarArmadosDisco() {
  escribirJsonAtomico(ARMADOS_PATH, armados);
}

// Normaliza el detalle de lo escaneado que manda el navegador. Guardamos una
// línea por producto con lo pedido y lo realmente contado, porque es lo que va
// a necesitar el facturador para emitir por lo que salió y no por lo que decía
// la orden. Las líneas de combo quedan marcadas: en la orden son un solo
// renglón, acá están abiertas en sus componentes.
function normalizarItemsArmado(items) {
  if (!Array.isArray(items)) return [];
  return items.slice(0, 500).map(it => ({
    codigo: it && it.codigo ? String(it.codigo) : null,
    concepto: it && it.concepto ? String(it.concepto).slice(0, 200) : '',
    pedido: Number(it && it.pedido) || 0,
    escaneado: Number(it && it.escaneado) || 0,
    combo: !!(it && it.combo)
  }));
}

function registrarArmado({ orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente, usuarioArmado, verificado, items }) {
  orderId = String(orderId);
  const timestamp = new Date().toISOString();
  const existente = armados.findIndex(a => a.orderId === orderId);
  const detalle = normalizarItemsArmado(items);
  const registro = {
    orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente,
    usuarioArmado, timestamp, verificado: !!verificado,
    items: detalle,
    // Se despachó exactamente lo pedido. Hoy es siempre true porque la app no
    // deja confirmar de otra forma; queda calculado para cuando se habiliten
    // los parciales y el facturador tenga que distinguirlos.
    completo: detalle.length > 0 && detalle.every(i => i.escaneado === i.pedido),
    // Todavía no llegó al buzón del facturador. Lo pone en true
    // `sincronizarArmados()` cuando el envío se confirma.
    enviado: false
  };
  if (existente >= 0) armados[existente] = registro;
  else armados.push(registro);
  guardarArmadosDisco();
  return registro;
}

function registrarImpresion(orderId, usuarioImpresion) {
  orderId = String(orderId);
  const idx = armados.findIndex(a => a.orderId === orderId);
  if (idx >= 0) {
    armados[idx].usuarioImpresion = usuarioImpresion;
    armados[idx].fechaImpresion = new Date().toISOString();
    guardarArmadosDisco();
  }
}

function eliminarArmado(orderId) {
  orderId = String(orderId);
  armados = armados.filter(a => a.orderId !== orderId);
  guardarArmadosDisco();
}

// El timestamp se guarda en UTC (ISO), pero el depósito razona en hora de
// Uruguay (UTC−3). Si filtráramos por la fecha del ISO, un armado del viernes
// 21:30 quedaría archivado como del sábado — y encima el Excel muestra la
// hora local, así que el reporte se contradecía con su propio filtro.
// toLocaleString('es-AR') formatea en 12 horas y, en las versiones actuales de
// Node, se come el "p. m.": un armado de las 19:02 salía impreso como
// "07:02:45" en el Excel. Formateamos a mano para no depender del ICU de turno.
function formatearFechaHoraLocal(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '-';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mi = String(d.getMinutes()).padStart(2, '0');
  return `${dd}/${mm}/${d.getFullYear()} ${hh}:${mi}`;
}

function fechaLocalISO(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '';
  const local = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10); // YYYY-MM-DD en hora local
}

function consultarArmados({ numeroOrden, fechaDesde, fechaHasta }) {
  let resultado = [...armados];
  if (numeroOrden) {
    const buscado = numeroOrden.trim().toLowerCase();
    resultado = resultado.filter(a => (a.numeroOrden || '').toLowerCase().includes(buscado));
  }
  if (fechaDesde) {
    resultado = resultado.filter(a => fechaLocalISO(a.timestamp) >= fechaDesde);
  }
  if (fechaHasta) {
    resultado = resultado.filter(a => fechaLocalISO(a.timestamp) <= fechaHasta);
  }
  return resultado.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

// --- Envío al facturador (buzón en Google Sheets) ---
//
// El armado se guarda SIEMPRE primero en disco y recién después se intenta
// mandar. Si el envío falla, el registro queda marcado `enviado: false` y se
// reintenta solo: al arrancar, cada 5 minutos, y cada vez que se confirma un
// armado nuevo. El depósito nunca queda trabado por un problema de internet.
//
// Reenviar de más es inofensivo: el buzón es un log de eventos y el
// facturador se queda con uno solo por orden. Perder un envío no lo es —
// ese pedido no le aparece nunca a quien factura.
//
// La autenticación es con un Service Account propio, cuyo archivo va al lado
// del ejecutable (`buzon-sa.json`). Ese service account tiene acceso a UNA
// sola planilla, la del buzón: si esta PC se pierde, no se llega desde acá a
// las planillas de comisiones ni a los logs de facturación.
//
// Se firma el token a mano con el módulo `crypto` de Node en vez de usar las
// librerías de Google, para no sumarle dependencias al .exe. Son 40 líneas y
// el protocolo (JWT firmado con RS256 → access token) no cambia nunca.
const SYNC_INTERVALO_MS = 5 * 60 * 1000;
const SYNC_LOTE_MAX = 50;
const BUZON_SA_PATH = path.join(APP_DIR, 'buzon-sa.json');
const BUZON_TAB = 'armados';
const BUZON_COLUMNAS = [
  'timestamp', 'fecha_local', 'evento', 'id_orden', 'numero_orden',
  'fecha_orden', 'id_cliente', 'bultos', 'lineas', 'unidades', 'usuario',
  'completo', 'verificado', 'items_json', 'id_comprobante', 'numero_factura',
  'cae', 'observacion'
];

let sincronizando = false;
let sincronizacionUltimoError = null;
let googleToken = null;
let googleTokenExpira = 0;

function leerServiceAccountBuzon() {
  try {
    const sa = JSON.parse(fs.readFileSync(BUZON_SA_PATH, 'utf8'));
    if (!sa.client_email || !sa.private_key) return null;
    return sa;
  } catch {
    return null;
  }
}

function base64url(buf) {
  return Buffer.from(buf).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// Intercambia el Service Account por un access token de Google (OAuth2
// "JWT bearer"). El token dura una hora; lo cacheamos con un minuto de
// margen para no pedir uno por cada envío.
async function obtenerTokenGoogle(sa) {
  const ahora = Math.floor(Date.now() / 1000);
  if (googleToken && ahora < googleTokenExpira - 60) return googleToken;

  const header = { alg: 'RS256', typ: 'JWT' };
  const claims = {
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: sa.token_uri || 'https://oauth2.googleapis.com/token',
    iat: ahora,
    exp: ahora + 3600
  };

  const sinFirma = `${base64url(JSON.stringify(header))}.${base64url(JSON.stringify(claims))}`;
  const firma = crypto.createSign('RSA-SHA256').update(sinFirma).sign(sa.private_key);
  const jwt = `${sinFirma}.${base64url(firma)}`;

  const resp = await fetch(claims.aud, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    })
  });

  const datos = await resp.json().catch(() => ({}));
  if (!resp.ok || !datos.access_token) {
    throw new Error(
      `Google rechazó las credenciales del buzón (${resp.status}): ` +
      `${datos.error_description || datos.error || 'sin detalle'}`
    );
  }

  googleToken = datos.access_token;
  googleTokenExpira = ahora + (datos.expires_in || 3600);
  return googleToken;
}

async function googleFetch(url, opciones, token) {
  const resp = await fetch(url, {
    ...opciones,
    headers: {
      ...(opciones && opciones.headers),
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });
  const datos = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    const msg = (datos.error && datos.error.message) || `HTTP ${resp.status}`;
    throw new Error(msg);
  }
  return datos;
}

// Crea la pestaña y su encabezado si todavía no existen, para que la planilla
// del buzón se pueda estrenar vacía sin que nadie prepare nada a mano.
async function asegurarTabBuzon(sheetId, token) {
  const meta = await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title`,
    { method: 'GET' }, token
  );
  const existe = (meta.sheets || []).some(
    s => s.properties && s.properties.title === BUZON_TAB
  );

  if (!existe) {
    await googleFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}:batchUpdate`,
      { method: 'POST', body: JSON.stringify({ requests: [{ addSheet: { properties: { title: BUZON_TAB } } }] }) },
      token
    );
  }

  const encabezado = await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${BUZON_TAB}!A1:R1`,
    { method: 'GET' }, token
  );
  if (!encabezado.values || !encabezado.values.length) {
    await googleFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${BUZON_TAB}!A1?valueInputOption=RAW`,
      { method: 'PUT', body: JSON.stringify({ values: [BUZON_COLUMNAS] }) },
      token
    );
  }
}

async function enviarFilasAlBuzon(sheetId, sa, filas) {
  const token = await obtenerTokenGoogle(sa);
  await asegurarTabBuzon(sheetId, token);

  const valores = filas.map(f => BUZON_COLUMNAS.map(c => {
    const v = f[c];
    return (v === null || v === undefined) ? '' : String(v);
  }));

  // RAW y no USER_ENTERED: el número de orden viene con ceros a la izquierda
  // ("00012036") y USER_ENTERED lo guardaría como 12036, rompiendo el cruce
  // contra Contabilium del otro lado.
  await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${BUZON_TAB}!A1:append` +
    `?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { method: 'POST', body: JSON.stringify({ values: valores }) },
    token
  );
}

function armadoAFilaBuzon(a) {
  return {
    timestamp: a.timestamp,
    fecha_local: fechaLocalISO(a.timestamp),
    evento: 'armado',
    id_orden: a.orderId,
    numero_orden: a.numeroOrden || '',
    fecha_orden: a.fechaOrden || '',
    id_cliente: a.idCliente == null ? '' : String(a.idCliente),
    bultos: a.bultos == null ? '' : String(a.bultos),
    lineas: a.lineas == null ? '' : String(a.lineas),
    unidades: a.unidades == null ? '' : String(a.unidades),
    usuario: a.usuarioArmado || '',
    completo: a.completo ? 'SI' : 'NO',
    verificado: a.verificado ? 'SI' : 'NO',
    items_json: JSON.stringify(a.items || []).slice(0, 45000),
    id_comprobante: '',
    numero_factura: '',
    cae: '',
    observacion: ''
  };
}

function armadosPendientesDeEnvio() {
  return armados.filter(a => a.enviado !== true);
}

// Devuelve {ok, enviados, error}. Nunca lanza: la llaman handlers que no
// se pueden romper por esto.
async function sincronizarArmados() {
  const cfg = leerConfig() || {};
  const sa = leerServiceAccountBuzon();
  if (!cfg.sheetId || !sa) {
    return { ok: false, enviados: 0, error: 'sin_configurar' };
  }
  if (sincronizando) return { ok: true, enviados: 0, error: null };

  const pendientes = armadosPendientesDeEnvio().slice(0, SYNC_LOTE_MAX);
  if (!pendientes.length) {
    sincronizacionUltimoError = null;
    return { ok: true, enviados: 0, error: null };
  }

  sincronizando = true;
  try {
    await enviarFilasAlBuzon(cfg.sheetId, sa, pendientes.map(armadoAFilaBuzon));

    const idsEnviados = new Set(pendientes.map(p => p.orderId));
    for (const a of armados) {
      if (idsEnviados.has(a.orderId)) a.enviado = true;
    }
    guardarArmadosDisco();
    sincronizacionUltimoError = null;
    console.log(`  ✔ Enviados al facturador: ${pendientes.length} armado(s).`);
    return { ok: true, enviados: pendientes.length, error: null };
  } catch (err) {
    sincronizacionUltimoError = err.message;
    console.log(`  ⚠ No se pudo enviar al facturador (${err.message}). Queda pendiente y se reintenta solo.`);
    return { ok: false, enviados: 0, error: err.message };
  } finally {
    sincronizando = false;
  }
}

function estadoSincronizacion() {
  const cfg = leerConfig() || {};
  const sa = leerServiceAccountBuzon();
  return {
    configurado: !!(cfg.sheetId && sa),
    sheetId: cfg.sheetId || '',
    credencialCargada: !!sa,
    credencialEmail: sa ? sa.client_email : '',
    pendientes: armadosPendientesDeEnvio().length,
    ultimoError: sincronizacionUltimoError
  };
}

// --- Config (credenciales + usuarios) ---
// Ojo: acá viven también los usuarios de la app, así que perder este archivo
// obliga a reconfigurar todo. Va con el mismo respaldo que el resto.
function leerConfig() {
  return leerJsonSeguro(CONFIG_PATH, null, d => d && typeof d === 'object');
}
function guardarConfig(cfg) {
  escribirJsonAtomico(CONFIG_PATH, cfg);
}

// --- Usuarios y sesiones ---
const crypto = require('crypto');

function hashPassword(password, salt) {
  salt = salt || crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt, hash };
}

function verificarPassword(password, salt, hashGuardado) {
  const { hash } = hashPassword(password, salt);
  return hash === hashGuardado;
}

function leerUsuarios() {
  const cfg = leerConfig();
  return (cfg && cfg.usuarios) || [];
}

function guardarUsuarios(usuarios) {
  const cfg = leerConfig() || {};
  guardarConfig({ ...cfg, usuarios });
}

function crearUsuario(usuario, password) {
  const usuarios = leerUsuarios();
  if (usuarios.some(u => u.usuario.toLowerCase() === usuario.toLowerCase())) {
    throw new Error('Ya existe un usuario con ese nombre.');
  }
  const { salt, hash } = hashPassword(password);
  usuarios.push({ usuario, salt, hash });
  guardarUsuarios(usuarios);
}

function eliminarUsuario(usuario) {
  const usuarios = leerUsuarios().filter(u => u.usuario.toLowerCase() !== usuario.toLowerCase());
  guardarUsuarios(usuarios);
}

function autenticar(usuario, password) {
  const usuarios = leerUsuarios();
  const u = usuarios.find(x => x.usuario.toLowerCase() === (usuario || '').toLowerCase());
  if (!u) return false;
  return verificarPassword(password, u.salt, u.hash);
}

// Sesiones en memoria: token -> { usuario, creado, ultimoUso }
const sesiones = new Map();

// Caducan a las 12 h SIN USO (no desde el login), para que no se corte a
// mitad de un turno pero tampoco quede una sesión abierta para siempre.
const SESION_TTL_MS = 12 * 60 * 60 * 1000;

function crearSesion(usuario) {
  const token = crypto.randomBytes(24).toString('hex');
  const ahora = Date.now();
  sesiones.set(token, { usuario, creado: ahora, ultimoUso: ahora });
  return token;
}

function usuarioDeSesion(token) {
  const s = sesiones.get(token);
  if (!s) return null;
  if (Date.now() - s.ultimoUso > SESION_TTL_MS) {
    sesiones.delete(token);
    return null;
  }
  s.ultimoUso = Date.now();
  return s.usuario;
}

function cerrarSesion(token) {
  sesiones.delete(token);
}

// --- Token (cache en memoria) ---
let cachedToken = null;
let tokenExpiresAt = 0;

async function getToken() {
  const cfg = leerConfig();
  if (!cfg || !cfg.clientId || !cfg.clientSecret) {
    throw new Error('Faltan las credenciales de Contabilium. Configuralas primero.');
  }

  const now = Date.now();
  if (cachedToken && now < tokenExpiresAt - 5000) return cachedToken;

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: cfg.clientId,
    client_secret: cfg.clientSecret
  });

  const resp = await fetch(`${BASE_URL}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  if (!resp.ok) {
    throw new Error(`No se pudo autenticar con Contabilium (${resp.status}). Revisá tus credenciales.`);
  }

  const data = await resp.json();
  cachedToken = data.access_token;
  tokenExpiresAt = now + (data.expires_in ? data.expires_in * 1000 : 60 * 1000);
  return cachedToken;
}

const REINTENTOS_MAX = 3;
const ESPERA_BASE_MS = 800;

function esperar(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Contabilium responde 429 cuando se le pega muy seguido (ya nos pasó en los
// otros proyectos contra la misma API). Reintenta respetando el Retry-After
// que manda el servidor, y también ante 5xx o cortes de red, que son
// transitorios. Un error definitivo sigue levantando excepción.
async function authedFetch(url, { intentos = REINTENTOS_MAX } = {}) {
  let ultimoError = null;

  for (let intento = 1; intento <= intentos; intento++) {
    let resp;
    try {
      const token = await getToken();
      resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    } catch (err) {
      ultimoError = new Error(`No se pudo contactar a Contabilium: ${err.message}`);
      if (intento === intentos) throw ultimoError;
      await esperar(ESPERA_BASE_MS * intento);
      continue;
    }

    if (resp.ok) return resp.json();

    // Token vencido antes de lo previsto: lo tiramos y pedimos uno nuevo.
    if (resp.status === 401 && intento < intentos) {
      cachedToken = null;
      tokenExpiresAt = 0;
      continue;
    }

    if ((resp.status === 429 || resp.status >= 500) && intento < intentos) {
      const retryAfter = parseInt(resp.headers.get('retry-after') || '', 10);
      const espera = Number.isFinite(retryAfter)
        ? retryAfter * 1000
        : ESPERA_BASE_MS * Math.pow(2, intento - 1);
      console.log(`  … Contabilium respondió ${resp.status}; reintento ${intento + 1}/${intentos} en ${Math.round(espera / 1000)}s`);
      await esperar(espera);
      continue;
    }

    const text = await resp.text();
    throw new Error(`Error consultando Contabilium (${resp.status}): ${text}`);
  }

  throw ultimoError || new Error('No se pudo consultar Contabilium.');
}

async function fetchAllOrdenes(fechaDesde, fechaHasta) {
  const ordenes = [];
  let page = 1;
  while (true) {
    const url = `${BASE_URL}/api/ordenesVenta/search?fechaDesde=${fechaDesde}&fechaHasta=${fechaHasta}&page=${page}`;
    const data = await authedFetch(url);
    const items = data.Items || data.items || [];
    ordenes.push(...items);
    if (items.length < 50) break;
    page += 1;
  }
  return ordenes;
}

async function fetchOrdenDetalle(id) {
  return authedFetch(`${BASE_URL}/api/ordenesVenta/?id=${id}`);
}

// --- Combos (expansión de SKU compuestos para picking) ---
const conceptosCache = new Map();

async function obtenerConcepto(idConcepto) {
  if (conceptosCache.has(idConcepto)) return conceptosCache.get(idConcepto);
  const data = await authedFetch(`${BASE_URL}/api/conceptos/?id=${idConcepto}`);
  conceptosCache.set(idConcepto, data);
  return data;
}

function construirItemPicking(codigo, concepto, cantidad, extra = {}) {
  const prod = buscarProducto(codigo, concepto);
  return {
    codigo,
    concepto,
    cantidad,
    ean: prod ? prod.ean : null,
    dun: prod ? prod.dun : null,
    factor: prod ? prod.factor : null,
    reconocido: !!prod,
    ...extra
  };
}

// Expande los items de una orden: si un item es un Combo ("C"), lo reemplaza
// por sus componentes reales (SKU y cantidad = cantidad del combo × cantidad
// del componente), para que se pueda pickear cada producto físico.
async function expandirItemsOrden(items) {
  const resultado = [];

  for (const it of items) {
    if (it.Tipo === 'C') {
      try {
        const combo = await obtenerConcepto(it.IdConcepto);
        const componentes = combo.Items || [];

        if (componentes.length) {
          for (const comp of componentes) {
            const cantidadTotal = (comp.Cantidad || 0) * (it.Cantidad || 0);
            const prodInfo = buscarProducto(comp.Codigo, null);
            const nombreComponente = prodInfo && prodInfo.nombre ? prodInfo.nombre : comp.Codigo;
            resultado.push(construirItemPicking(
              comp.Codigo,
              `${nombreComponente} (combo: ${it.Concepto})`,
              cantidadTotal,
              { esComponenteCombo: true, comboOrigen: it.Concepto }
            ));
          }
          continue;
        }
        console.log(`  ⚠ El combo "${it.Concepto}" (idConcepto=${it.IdConcepto}) no trajo componentes; se deja como línea simple.`);
      } catch (err) {
        console.log(`  ⚠ No se pudo expandir el combo "${it.Concepto}" (idConcepto=${it.IdConcepto}): ${err.message}. Se deja como línea simple.`);
      }
    }

    resultado.push(construirItemPicking(it.Codigo, it.Concepto, it.Cantidad));
  }

  return resultado;
}

// --- Clientes (para etiquetas de despacho) ---
let clientesPorId = new Map();
let clientesCargados = false;
let clientesUltimaActualizacion = 0;
let clientesCargando = null; // Promise en curso, para no disparar cargas en paralelo duplicadas

function cargarClientesDesdeDiscoSiExiste() {
  const raw = leerJsonSeguro(CLIENTES_CACHE_PATH, null, d => d && Array.isArray(d.clientes));
  if (!raw) return; // todavía no hay caché; se carga desde la API cuando haga falta
  clientesPorId = new Map();
  for (const c of raw.clientes) clientesPorId.set(String(c.Id), c);
  clientesCargados = true;
  clientesUltimaActualizacion = new Date(raw.actualizado).getTime() || 0;
  console.log(`  Clientes cargados desde caché en disco: ${clientesPorId.size} (actualizado ${raw.actualizado})`);
}

function guardarClientesDisco() {
  try {
    escribirJsonAtomico(CLIENTES_CACHE_PATH, {
      actualizado: new Date().toISOString(),
      clientes: Array.from(clientesPorId.values())
    });
  } catch (err) {
    console.log(`  ⚠ No se pudo guardar la caché de clientes en disco: ${err.message}`);
  }
}

const CLIENTES_CONCURRENCIA = 6;

async function cargarClientesInterno() {
  const nuevoMapa = new Map();
  let page = 1;
  let seguir = true;

  while (seguir) {
    const paginas = Array.from({ length: CLIENTES_CONCURRENCIA }, (_, i) => page + i);

    // IMPORTANTE: acá no se puede tragar el error de una página. Antes, si una
    // fallaba (un 429, un corte) se la trataba como página vacía, el barrido
    // cortaba ahí y se guardaba en disco un padrón incompleto como si estuviera
    // completo. El síntoma no dice nada útil: etiquetas con "Cliente no
    // identificado". Ahora, si una página falla después de sus reintentos,
    // abortamos el refresco entero y conservamos la caché anterior.
    const resultados = await Promise.all(
      paginas.map(p =>
        authedFetch(`${BASE_URL}/api/clientes/search?pageSize=500&page=${p}`)
          .then(data => data.Items || data.items || (Array.isArray(data) ? data : []))
          .catch(err => {
            throw new Error(`falló la página ${p} del padrón de clientes: ${err.message}`);
          })
      )
    );

    let algunaVacia = false;
    for (let i = 0; i < resultados.length; i++) {
      const items = resultados[i];
      for (const c of items) nuevoMapa.set(String(c.Id), c);
      console.log(`  Clientes: página ${paginas[i]} → ${items.length} registro(s)`);
      if (items.length === 0) algunaVacia = true; // fin real del padrón
    }

    if (algunaVacia) seguir = false;
    else page += CLIENTES_CONCURRENCIA;

    if (page > 2000) break; // salvaguarda para no loopear indefinidamente
  }

  clientesPorId = nuevoMapa;
  clientesCargados = true;
  clientesUltimaActualizacion = Date.now();
  console.log(`  Total de clientes cargados en caché: ${clientesPorId.size}`);
  guardarClientesDisco();
}

// Evita que dos pedidos simultáneos disparen dos cargas completas en paralelo
function cargarClientes() {
  if (!clientesCargando) {
    clientesCargando = cargarClientesInterno().finally(() => { clientesCargando = null; });
  }
  return clientesCargando;
}

// IDs que ya buscamos con la caché fresca y no aparecieron (típicamente
// contactos cargados como Proveedor). Sin esto, cada etiqueta de esos
// contactos vuelve a paginar el padrón entero.
const clientesNoEncontrados = new Set();

async function obtenerCliente(idCliente) {
  if (!idCliente) return null;
  if (!clientesCargados) await cargarClientes(); // sin caché no hay nada que hacer: que el error suba

  const clave = String(idCliente);
  let c = clientesPorId.get(clave);

  // Si no aparece, puede ser un cliente nuevo: solo re-consultamos si la
  // caché ya tiene un rato (para no volver a paginar todo por cada intento)
  if (!c && !clientesNoEncontrados.has(clave) && Date.now() - clientesUltimaActualizacion > 5 * 60 * 1000) {
    try {
      await cargarClientes();
    } catch (err) {
      // Refrescar es best effort: seguimos con la caché que ya teníamos.
      console.log(`  ⚠ No se pudo refrescar el padrón de clientes: ${err.message}`);
    }
    c = clientesPorId.get(clave);
    if (!c) clientesNoEncontrados.add(clave);
  }

  return c || null;
}

async function obtenerProveedor(id) {
  try {
    const data = await authedFetch(`${BASE_URL}/api/proveedores/obtener?id=${id}`);
    return data && data.Id ? data : null;
  } catch (err) {
    return null;
  }
}

// Algunos contactos están registrados como Proveedor en Contabilium con el
// mismo ID que se usa como IDCliente en la orden, en vez de como Cliente.
async function obtenerClienteOProveedor(idCliente) {
  const cliente = await obtenerCliente(idCliente);
  if (cliente) return { datos: cliente, origen: 'cliente' };

  const proveedor = await obtenerProveedor(idCliente);
  if (proveedor) return { datos: proveedor, origen: 'proveedor' };

  return null;
}

function escaparHtml(s) {
  return (s ?? '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function datosEtiquetaDesdeCliente(cliente) {
  const razonSocial = cliente ? (cliente.RazonSocial || '') : '';
  const nombreFantasia = cliente ? (cliente.NombreFantasia || '') : '';
  const destinatario = nombreFantasia
    ? `<b>${escaparHtml(nombreFantasia)}</b> (${escaparHtml(razonSocial)})`
    : escaparHtml(razonSocial) || 'Cliente no identificado';

  const direccionPartes = cliente
    ? [cliente.Domicilio, cliente.Ciudad, cliente.Provincia].filter(Boolean)
    : [];
  const direccion = direccionPartes.length ? direccionPartes.join(', ') : '-';

  const observaciones = cliente ? (cliente.Observaciones || '') : '';
  const contacto = cliente ? (cliente.Telefono || '-') : '-';

  return { destinatario, direccion, observaciones, contacto };
}

// Igual que datosEtiquetaDesdeCliente, pero en texto plano (sin HTML), para
// armar comandos ZPL en vez de una etiqueta HTML.
function datosEtiquetaPlano(cliente) {
  const razonSocial = cliente ? (cliente.RazonSocial || '') : '';
  const nombreFantasia = cliente ? (cliente.NombreFantasia || '') : '';
  const destinatario = nombreFantasia
    ? `${nombreFantasia} (${razonSocial})`
    : razonSocial || 'Cliente no identificado';

  const direccionPartes = cliente
    ? [cliente.Domicilio, cliente.Ciudad, cliente.Provincia].filter(Boolean)
    : [];
  const direccion = direccionPartes.length ? direccionPartes.join(', ') : '-';

  const observaciones = cliente ? (cliente.Observaciones || '') : '';
  const contacto = cliente ? (cliente.Telefono || '-') : '-';

  return { destinatario, direccion, observaciones, contacto };
}

function zplEscape(s) {
  return (s ?? '').toString()
    .replace(/\^/g, '')
    .replace(/~/g, '')
    .replace(/[\r\n]+/g, ' ');
}

// Genera el ZPL de una sola etiqueta (1200x800 dots = 15x10cm a 203dpi),
// para mandarlo directo a la impresora sin pasar por el driver GDI de
// Windows (que venía escalando mal el contenido).
// La impresora físicamente solo puede imprimir 800 dots (10cm) de ancho
// por pasada. Windows compensa esto para impresiones normales con una
// rotación de 90° configurada en el driver, pero esa rotación NO se aplica
// al ZPL enviado directo (bypass de Windows). Por eso diseñamos el
// contenido en un lienzo "virtual" acostado de 1200x800 (15x10cm, como se
// quiere leer) y lo rotamos 90° nosotros mismos hacia el lienzo físico
// real de la impresora (800x1200), usando texto con orientación rotada.
const VIRT_ANCHO = 1200; // 15cm, como se lee
const VIRT_ALTO = 800;   // 10cm, como se lee
const FISICO_ANCHO = 800;  // 10cm, ancho real del cabezal
const FISICO_ALTO = 1200;  // 15cm, largo real de avance

// Rota un punto del lienzo virtual (acostado) al lienzo físico real,
// 90° en sentido horario. El texto/línea rotado se extiende hacia
// mayor X físico a partir del origen, así que restamos también su
// propio alto/grosor para que no se pase del borde ni pise al vecino.
function rotarXY(x, y, alto) {
  return { x: FISICO_ANCHO - y - alto, y: x };
}

function campoTextoRotado(x, y, alto, ancho, blockWidth, lineas, spacing, justif, texto) {
  // Un bloque de varias líneas rotado ocupa, en el eje que usamos para
  // separar campos, el alto de UNA línea multiplicado por la cantidad de
  // líneas (más el espaciado entre ellas) — no solo el alto de una línea.
  const huella = alto * lineas + spacing * Math.max(0, lineas - 1);
  const p = rotarXY(x, y, huella);
  return `^FO${p.x},${p.y}^A0R,${alto},${ancho}^FB${blockWidth},${lineas},${spacing},${justif}^FD${zplEscape(texto)}^FS`;
}

function lineaHorizontalRotada(x, y, largo, grosor) {
  // Una línea horizontal en el diseño virtual se convierte en una línea
  // vertical en el lienzo físico real.
  const p = rotarXY(x, y, grosor);
  return `^FO${p.x},${p.y}^GB0,${largo},${grosor}^FS`;
}

function datosDestinatarioSeparados(cliente) {
  const razonSocial = cliente ? (cliente.RazonSocial || '') : '';
  const nombreFantasia = cliente ? (cliente.NombreFantasia || '') : '';
  if (nombreFantasia) {
    return { principal: nombreFantasia, secundario: `(${razonSocial})` };
  }
  return { principal: razonSocial || 'Cliente no identificado', secundario: '' };
}

function generarZplEtiqueta({ numeroOrden, i, totalBultos, cliente }) {
  const { direccion, observaciones, contacto } = datosEtiquetaPlano(cliente);
  const { principal, secundario } = datosDestinatarioSeparados(cliente);

  const lineaObs = observaciones ? `Obs.: ${observaciones}` : '';

  const MARGEN = 30;
  const BLOQUE = VIRT_ANCHO - MARGEN * 2;

  const partes = [
    '^XA',
    '^CI28',
    `^PW${FISICO_ANCHO}`,
    `^LL${FISICO_ALTO}`,
    '^LH0,0',
    // Encabezado: título centrado (60pt), celular más chico debajo (34pt)
    campoTextoRotado(MARGEN, 25, 60, 60, BLOQUE, 1, 0, 'C', 'GRUPO SUPRABOND URUGUAY SAS'),
    campoTextoRotado(850, 105, 34, 34, 320, 1, 0, 'R', 'Cel 093 900 536'),
    lineaHorizontalRotada(MARGEN, 160, BLOQUE, 4),
    // Nombre fantasía: 54pt, sin negrita simulada
    campoTextoRotado(MARGEN, 185, 54, 54, BLOQUE, 1, 0, 'L', principal)
  ];

  if (secundario) {
    partes.push(campoTextoRotado(MARGEN, 259, 40, 40, BLOQUE, 1, 0, 'L', secundario));
  }

  partes.push(
    campoTextoRotado(MARGEN, 319, 40, 40, BLOQUE, 2, 10, 'L', 'Direccion: ' + direccion),
    campoTextoRotado(MARGEN, 429, 40, 40, BLOQUE, 2, 8, 'L', lineaObs),
    campoTextoRotado(MARGEN, 537, 40, 40, BLOQUE, 1, 0, 'L', 'Contacto: ' + contacto),
    lineaHorizontalRotada(MARGEN, 655, BLOQUE, 4),
    // O/V N°: 54pt, número de bulto: 60pt
    campoTextoRotado(MARGEN, 690, 54, 54, 700, 1, 0, 'L', 'O/V N°: ' + numeroOrden),
    campoTextoRotado(750, 690, 60, 60, 420, 1, 0, 'R', `${i}/${totalBultos}`),
    '^XZ'
  );

  return partes.join('\n');
}

// El logo se lee de assets/ al arrancar. Si falta, la etiqueta cae al
// encabezado de texto (que es además lo que usa siempre la versión ZPL).
// Va como fondo por CSS y no como <img> para que el PNG viaje UNA sola vez
// aunque el pedido tenga 8 bultos: embebido por etiqueta, el HTML de una
// impresión pasaba del megabyte.
function claseEncabezado() {
  return logoBase64 ? 'encabezado con-logo' : 'encabezado';
}

function encabezadoHtml() {
  return logoBase64 ? '' : 'GRUPO SUPRABOND URUGUAY SAS';
}

function estilosEtiqueta() {
  if (!logoBase64) return ESTILOS_ETIQUETA;
  return `${ESTILOS_ETIQUETA}
  .encabezado.con-logo {
    background-image: url("data:image/png;base64,${logoBase64}");
    background-repeat: no-repeat;
    background-position: center;
    background-size: contain;
  }`;
}

const ESTILOS_ETIQUETA = `
  * { box-sizing: border-box; }
  body { margin: 0; font-family: Arial, Helvetica, sans-serif; }
  .etiqueta {
    width: 15cm;
    height: 10cm;
    padding: 0.9cm;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    page-break-after: always;
  }
  .etiqueta:last-child { page-break-after: auto; }
  .encabezado {
    font-size: 20pt;
    font-weight: bold;
    text-align: center;
    border-bottom: 2px solid #000;
    padding-bottom: 0.3cm;
    min-height: 1.6cm;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .encabezado img.logo { max-width: 100%; max-height: 1.6cm; }
  .cuerpo { flex: 1; margin-top: 0.4cm; }
  .campo { font-size: 15pt; margin-bottom: 0.3cm; }
  .campo .label { font-weight: bold; }
  .observaciones { font-size: 12pt; }
  .pie {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 2px solid #000;
    padding-top: 0.3cm;
    font-size: 13pt;
  }
  .bulto, .ov { font-size: 22pt; font-weight: bold; }
`;

function contenidoEtiqueta({ numeroOrden, i, totalBultos, destinatario, direccion, observaciones, contacto }) {
  return `
    <div class="etiqueta">
      <div class="${claseEncabezado()}">${encabezadoHtml()}</div>
      <div class="cuerpo">
        <div class="campo"><span class="label">Destinatario:</span> ${destinatario}</div>
        <div class="campo"><span class="label">Dirección:</span> ${escaparHtml(direccion)}</div>
        ${observaciones ? `<div class="campo observaciones"><span class="label">Obs.:</span> ${escaparHtml(observaciones)}</div>` : ''}
        <div class="campo"><span class="label">Contacto:</span> ${escaparHtml(contacto)}</div>
      </div>
      <div class="pie">
        <span class="ov">O/V N°: ${escaparHtml(numeroOrden)}</span>
        <span class="bulto">${i}/${totalBultos}</span>
      </div>
    </div>
  `;
}

// HTML con todas las etiquetas de una orden, para vista previa / impresión
// manual desde el navegador (respaldo si la impresión directa no está
// disponible en esta PC).
function generarHtmlEtiquetas({ numeroOrden, bultos, cliente }) {
  const datos = datosEtiquetaDesdeCliente(cliente);
  const totalBultos = bultos || 1;
  let etiquetasHtml = '';

  for (let i = 1; i <= totalBultos; i++) {
    etiquetasHtml += contenidoEtiqueta({ numeroOrden, i, totalBultos, ...datos });
  }

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8" />
<title>Etiquetas - Pedido ${escaparHtml(numeroOrden)}</title>
<style>
  @page { size: 15cm 10cm landscape; margin: 0; }
  ${estilosEtiqueta()}
  @media screen {
    body { background: #eee; padding: 20px; }
    .etiqueta { background: white; margin-bottom: 20px; box-shadow: 0 0 6px rgba(0,0,0,0.3); }
  }
</style>
</head>
<body>
${etiquetasHtml}
<script>
  window.onload = () => { window.print(); };
</script>
</body>
</html>`;
}

// --- Impresión directa (sin diálogo) ---
function ejecutarComando(comando, args, timeoutMs = 20000) {
  return new Promise((resolve, reject) => {
    const proceso = spawn(comando, args, { windowsHide: true });
    let stderr = '';
    let stdout = '';
    let terminado = false;

    const timer = setTimeout(() => {
      if (terminado) return;
      terminado = true;
      proceso.kill();
      reject(new Error(`"${comando}" no respondió a tiempo (posible ventana bloqueada en Windows). Se canceló automáticamente.`));
    }, timeoutMs);

    proceso.stdout && proceso.stdout.on('data', d => { stdout += d.toString(); });
    proceso.stderr && proceso.stderr.on('data', d => { stderr += d.toString(); });
    proceso.on('error', err => {
      if (terminado) return;
      terminado = true;
      clearTimeout(timer);
      reject(err);
    });
    proceso.on('close', code => {
      if (terminado) return;
      terminado = true;
      clearTimeout(timer);
      if (stdout.trim()) console.log(`  [${comando}] ${stdout.trim().replace(/\n/g, '\n  ')}`);
      if (code === 0) resolve(stdout);
      else reject(new Error(`"${comando}" terminó con código ${code}. ${stderr}`));
    });
  });
}

function generarScriptImpresionZPL() {
  return `
param(
  [Parameter(Mandatory=$true)][string]$ZplPath,
  [Parameter(Mandatory=$true)][string]$PrinterName
)

$codigoRaw = @"
using System;
using System.Runtime.InteropServices;

public class RawPrinterHelper {
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
  public class DOCINFOA {
    [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
    [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
    [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
  }

  [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string src, out IntPtr hPrinter, IntPtr pDefault);

  [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool ClosePrinter(IntPtr hPrinter);

  [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

  [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool EndDocPrinter(IntPtr hPrinter);

  [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool StartPagePrinter(IntPtr hPrinter);

  [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool EndPagePrinter(IntPtr hPrinter);

  [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
  public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

  public static bool SendBytesToPrinter(string printerName, byte[] bytes) {
    IntPtr hPrinter;
    DOCINFOA di = new DOCINFOA();
    di.pDocName = "Etiqueta ZPL";
    di.pDataType = "RAW";
    bool ok = false;

    if (OpenPrinter(printerName, out hPrinter, IntPtr.Zero)) {
      if (StartDocPrinter(hPrinter, 1, di)) {
        if (StartPagePrinter(hPrinter)) {
          IntPtr pUnmanagedBytes = Marshal.AllocCoTaskMem(bytes.Length);
          Marshal.Copy(bytes, 0, pUnmanagedBytes, bytes.Length);
          int dwWritten;
          ok = WritePrinter(hPrinter, pUnmanagedBytes, bytes.Length, out dwWritten);
          Marshal.FreeCoTaskMem(pUnmanagedBytes);
          EndPagePrinter(hPrinter);
        }
        EndDocPrinter(hPrinter);
      }
      ClosePrinter(hPrinter);
    }
    return ok;
  }
}
"@

Add-Type -TypeDefinition $codigoRaw -Language CSharp

$zplTexto = Get-Content -Path $ZplPath -Raw -Encoding UTF8
$bytes = [System.Text.Encoding]::UTF8.GetBytes($zplTexto)

$resultado = [RawPrinterHelper]::SendBytesToPrinter($PrinterName, $bytes)

if (-not $resultado) {
  Write-Error "No se pudo enviar el ZPL a la impresora '$PrinterName' (revisa el nombre exacto o que este encendida/conectada)."
  exit 1
}

Write-Output "ZPL enviado correctamente a '$PrinterName' ($($bytes.Length) bytes)."
`;
}

async function enviarZplAImpresora(zplTexto, nombreImpresora) {
  const rutaZpl = path.join(TEMP_DIR, `etiqueta_${Date.now()}.zpl`);
  const rutaScript = path.join(TEMP_DIR, 'imprimir_zpl.ps1');

  fs.writeFileSync(rutaZpl, zplTexto, 'utf8');
  fs.writeFileSync(rutaScript, generarScriptImpresionZPL());

  try {
    const stdout = await ejecutarComando('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-File', rutaScript,
      '-ZplPath', rutaZpl,
      '-PrinterName', nombreImpresora
    ]);
    return stdout.trim();
  } finally {
    fs.unlink(rutaZpl, () => {});
  }
}

async function imprimirEtiquetasDirecto({ numeroOrden, bultos, cliente, nombreImpresora }) {
  if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

  const totalBultos = bultos || 1;
  const diagnostico = [];

  for (let i = 1; i <= totalBultos; i++) {
    const zpl = generarZplEtiqueta({ numeroOrden, i, totalBultos, cliente });
    const salida = await enviarZplAImpresora(zpl, nombreImpresora);
    if (salida) diagnostico.push(`Bulto ${i}/${totalBultos}: ${salida}`);
  }

  return { totalBultos, diagnostico };
}

async function resolverClienteParaEtiqueta(registro) {
  if (!registro.idCliente) {
    console.log(`  ⚠ La orden ${registro.numeroOrden} no tiene idCliente guardado en el registro de armado (pedido armado con una versión anterior de la app?).`);
    return null;
  }
  try {
    const encontrado = await obtenerClienteOProveedor(registro.idCliente);
    if (encontrado) {
      if (encontrado.origen === 'proveedor') {
        console.log(`  ℹ El contacto idCliente=${registro.idCliente} de la orden ${registro.numeroOrden} se encontró como Proveedor, no como Cliente.`);
      }
      return encontrado.datos;
    }
    console.log(`  ⚠ No se encontró el contacto idCliente=${registro.idCliente} ni en Clientes ni en Proveedores para la orden ${registro.numeroOrden} (clientes en caché: ${clientesPorId.size})`);
    return null;
  } catch (err) {
    console.log(`  ⚠ Error al obtener el contacto idCliente=${registro.idCliente}: ${err.message}`);
    return null;
  }
}

// --- Helpers HTTP ---
function sendJSON(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

function serveStatic(req, res, urlPath) {
  let filePath = urlPath === '/' ? '/index.html' : urlPath;
  filePath = path.join(PUBLIC_DIR, filePath);

  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    return res.end('Forbidden');
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      return res.end('No encontrado');
    }
    const ext = path.extname(filePath);
    const types = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css' };
    res.writeHead(200, { 'Content-Type': (types[ext] || 'application/octet-stream') + '; charset=utf-8' });
    res.end(data);
  });
}

// ¿El pedido viene de la propia PC donde corre la app? Se usa para que la
// creación del primer usuario (la cuenta que después crea a las demás) no la
// pueda hacer alguien de la red si abren el puerto.
function esPedidoLocal(req) {
  const ip = req.socket.remoteAddress || '';
  return ip === '127.0.0.1' || ip === '::1' || ip === '::ffff:127.0.0.1';
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => (data += chunk));
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

function readBodyBuffer(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

const server = http.createServer(async (req, res) => {
  const parsed = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = parsed.pathname;

  try {
    // --- Autenticación ---
    if (pathname === '/api/auth/estado' && req.method === 'GET') {
      const hayUsuarios = leerUsuarios().length > 0;
      const token = req.headers['x-session-token'];
      const usuario = token ? usuarioDeSesion(token) : null;
      return sendJSON(res, 200, { hayUsuarios, autenticado: !!usuario, usuario, version: APP_VERSION });
    }

    if (pathname === '/api/auth/registrar-primero' && req.method === 'POST') {
      if (leerUsuarios().length > 0) {
        return sendJSON(res, 400, { error: 'Ya existen usuarios. Pedile a un usuario existente que te cree una cuenta.' });
      }
      if (!esPedidoLocal(req)) {
        return sendJSON(res, 403, { error: 'El primer usuario se tiene que crear desde la PC donde corre la app.' });
      }
      const raw = await readBody(req);
      const { usuario, password } = JSON.parse(raw || '{}');
      if (!usuario || !password) return sendJSON(res, 400, { error: 'Faltan datos' });
      crearUsuario(usuario, password);
      const token = crearSesion(usuario);
      return sendJSON(res, 200, { ok: true, token, usuario });
    }

    if (pathname === '/api/auth/login' && req.method === 'POST') {
      const raw = await readBody(req);
      const { usuario, password } = JSON.parse(raw || '{}');
      if (!autenticar(usuario, password)) {
        return sendJSON(res, 401, { error: 'Usuario o contraseña incorrectos.' });
      }
      const token = crearSesion(usuario);
      return sendJSON(res, 200, { ok: true, token, usuario });
    }

    if (pathname === '/api/auth/logout' && req.method === 'POST') {
      const token = req.headers['x-session-token'];
      if (token) cerrarSesion(token);
      return sendJSON(res, 200, { ok: true });
    }

    // A partir de acá, todas las rutas /api requieren sesión iniciada
    if (pathname.startsWith('/api/') && !pathname.startsWith('/api/auth/')) {
      const token = req.headers['x-session-token'];
      const usuarioActual = token ? usuarioDeSesion(token) : null;
      if (!usuarioActual) {
        return sendJSON(res, 401, { error: 'Sesión inválida o expirada. Volvé a iniciar sesión.' });
      }
      req.usuarioActual = usuarioActual;
    }

    // --- Usuarios (gestión) ---
    if (pathname === '/api/usuarios' && req.method === 'GET') {
      const usuarios = leerUsuarios().map(u => ({ usuario: u.usuario }));
      return sendJSON(res, 200, { usuarios });
    }

    if (pathname === '/api/usuarios' && req.method === 'POST') {
      const raw = await readBody(req);
      const { usuario, password } = JSON.parse(raw || '{}');
      if (!usuario || !password) return sendJSON(res, 400, { error: 'Faltan datos' });
      if (password.length < 4) return sendJSON(res, 400, { error: 'La contraseña debe tener al menos 4 caracteres.' });
      try {
        crearUsuario(usuario, password);
        return sendJSON(res, 200, { ok: true });
      } catch (err) {
        return sendJSON(res, 400, { error: err.message });
      }
    }

    if (pathname === '/api/usuarios' && req.method === 'DELETE') {
      const usuario = parsed.searchParams.get('usuario');
      if (!usuario) return sendJSON(res, 400, { error: 'Falta usuario' });
      if (usuario.toLowerCase() === req.usuarioActual.toLowerCase()) {
        return sendJSON(res, 400, { error: 'No podés eliminar tu propio usuario mientras tenés la sesión iniciada.' });
      }
      eliminarUsuario(usuario);
      return sendJSON(res, 200, { ok: true });
    }

    // --- Estado de configuración ---
    if (pathname === '/api/config' && req.method === 'GET') {
      const cfg = leerConfig();
      return sendJSON(res, 200, {
        configurado: !!(cfg && cfg.clientId && cfg.clientSecret),
        printerName: cfg ? (cfg.printerName || '') : '',
        sheetId: cfg ? (cfg.sheetId || '') : ''
      });
    }

    if (pathname === '/api/config' && req.method === 'POST') {
      const raw = await readBody(req);
      const { clientId, clientSecret, printerName } = JSON.parse(raw || '{}');
      if (!clientId || !clientSecret) {
        return sendJSON(res, 400, { error: 'Faltan datos' });
      }
      const cfgAnterior = leerConfig() || {};
      guardarConfig({
        clientId,
        clientSecret,
        printerName: printerName !== undefined ? printerName : cfgAnterior.printerName
      });
      cachedToken = null;
      return sendJSON(res, 200, { ok: true });
    }

    // --- Envío al facturador ---
    if (pathname === '/api/sincronizacion' && req.method === 'GET') {
      return sendJSON(res, 200, estadoSincronizacion());
    }

    if (pathname === '/api/sincronizacion' && req.method === 'POST') {
      const resultado = await sincronizarArmados();
      if (!resultado.ok && resultado.error === 'sin_configurar') {
        return sendJSON(res, 400, {
          error: 'Todavía no configuraste el envío al facturador (falta la dirección del buzón y su clave).'
        });
      }
      if (!resultado.ok) {
        return sendJSON(res, 502, { error: `No se pudo enviar: ${resultado.error}` });
      }
      return sendJSON(res, 200, { ok: true, enviados: resultado.enviados, ...estadoSincronizacion() });
    }

    if (pathname === '/api/config/buzon' && req.method === 'POST') {
      const raw = await readBody(req);
      let { sheetId } = JSON.parse(raw || '{}');
      const cfgActual = leerConfig();
      if (!cfgActual) return sendJSON(res, 400, { error: 'Primero configurá tus credenciales de Contabilium.' });

      sheetId = (sheetId || '').trim();
      // Se acepta la dirección entera de la planilla, no solo el ID: es lo
      // que uno tiene a mano cuando la está mirando en el navegador.
      const desdeUrl = sheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if (desdeUrl) sheetId = desdeUrl[1];
      if (sheetId && !/^[a-zA-Z0-9_-]{20,}$/.test(sheetId)) {
        return sendJSON(res, 400, {
          error: 'Eso no parece el ID de una planilla de Google. Pegá la dirección completa de la planilla del buzón.'
        });
      }

      guardarConfig({ ...cfgActual, sheetId });
      const resultado = await sincronizarArmados();
      return sendJSON(res, 200, {
        ok: true,
        enviados: resultado.enviados,
        errorEnvio: resultado.ok ? null : resultado.error,
        ...estadoSincronizacion()
      });
    }

    if (pathname === '/api/config/impresora' && req.method === 'POST') {
      const raw = await readBody(req);
      const { printerName } = JSON.parse(raw || '{}');
      const cfgActual = leerConfig();
      if (!cfgActual) return sendJSON(res, 400, { error: 'Primero configurá tus credenciales de Contabilium.' });
      guardarConfig({ ...cfgActual, printerName: printerName || '' });
      return sendJSON(res, 200, { ok: true, printerName: printerName || '' });
    }

    // --- Catálogo EAN/DUN ---
    if (pathname === '/api/catalogo' && req.method === 'GET') {
      return sendJSON(res, 200, catalogoInfo);
    }

    if (pathname === '/api/catalogo' && req.method === 'POST') {
      const nombreArchivoRaw = req.headers['x-filename'] || 'archivo';
      const nombreArchivo = decodeURIComponent(nombreArchivoRaw);
      const buffer = await readBodyBuffer(req);
      try {
        const nuevosProductos = parsearTablaProductos(buffer, nombreArchivo);
        guardarCatalogo(nuevosProductos, nombreArchivo);
        return sendJSON(res, 200, catalogoInfo);
      } catch (err) {
        return sendJSON(res, 400, { error: err.message });
      }
    }

    // --- Picking ---
    if (pathname === '/api/picking' && req.method === 'GET') {
      const fechaDesde = parsed.searchParams.get('fechaDesde');
      const fechaHasta = parsed.searchParams.get('fechaHasta');

      if (!fechaDesde || !fechaHasta) {
        return sendJSON(res, 400, { error: 'fechaDesde y fechaHasta son obligatorios' });
      }

      const ordenes = await fetchAllOrdenes(fechaDesde, fechaHasta);
      // Solo interesan las órdenes pendientes de picking. Una orden puede
      // quedar en Pendiente pero ya tener comprobante asociado (facturada
      // desde la web de Contabilium): esa no hay que volver a armarla.
      const filtradas = ordenes.filter(
        o => o.Estado === 'Pendiente' && !((o.IDComprobante || 0) > 0)
      );

      const picking = [];
      for (const orden of filtradas) {
        try {
          const detalle = await fetchOrdenDetalle(orden.ID);
          const itemsOriginales = detalle.Items || [];
          const items = await expandirItemsOrden(itemsOriginales);
          const cantidadProductos = items.reduce((sum, it) => sum + (it.cantidad || 0), 0);

          // Si esta orden ya fue armada antes (aunque Contabilium todavía la
          // muestre como Pendiente), avisamos al frontend para que no se
          // pueda volver a armar por error.
          const armadoExistente = armados.find(a => a.orderId === String(orden.ID)) || null;

          picking.push({
            id: orden.ID,
            idCliente: detalle.IDCliente,
            numeroOrden: detalle.NumeroOrden || orden.NumeroOrden,
            fecha: detalle.FechaCreacion || orden.FechaCreacion,
            estado: orden.Estado,
            cantidadProductos,
            items,
            armadoExistente
          });
        } catch (err) {
          picking.push({
            id: orden.ID,
            numeroOrden: orden.NumeroOrden,
            fecha: orden.FechaCreacion,
            estado: orden.Estado,
            cantidadProductos: null,
            items: [],
            error: 'No se pudo obtener el detalle de esta orden'
          });
        }
      }

      return sendJSON(res, 200, { total: picking.length, ordenes: picking });
    }

    // --- Pedidos armados ---
    if (pathname === '/api/armado' && req.method === 'POST') {
      const raw = await readBody(req);
      const datos = JSON.parse(raw || '{}');
      const { orderId, bultos, lineas, unidades } = datos;
      if (!orderId || !bultos) {
        return sendJSON(res, 400, { error: 'Faltan datos (orderId, bultos)' });
      }

      // Releemos la orden en Contabilium en vez de confiar en lo que manda el
      // navegador: una pestaña abierta hace rato puede estar mostrando una
      // orden que mientras tanto se canceló o se facturó. Si la API no
      // responde, igual registramos el armado (el depósito no puede quedar
      // parado por eso) pero lo dejamos marcado como no verificado.
      let numeroOrden = datos.numeroOrden;
      let fechaOrden = datos.fechaOrden;
      let idCliente = datos.idCliente;
      let verificado = false;

      try {
        const detalle = await fetchOrdenDetalle(orderId);
        const estado = (detalle.Estado || '').trim();
        const yaFacturada = (detalle.IDComprobante || 0) > 0;
        if (estado !== 'Pendiente' || yaFacturada) {
          const motivo = yaFacturada ? 'ya tiene factura asociada' : `figura como "${estado || 'sin estado'}"`;
          return sendJSON(res, 409, {
            error: `La orden ${detalle.NumeroOrden || numeroOrden || orderId} ${motivo} en Contabilium. Volvé a consultar el listado antes de armarla.`
          });
        }
        numeroOrden = detalle.NumeroOrden || numeroOrden;
        fechaOrden = detalle.FechaCreacion || fechaOrden;
        idCliente = detalle.IDCliente || idCliente;
        verificado = true;
      } catch (err) {
        console.log(`  ⚠ No se pudo verificar la orden ${orderId} contra Contabilium (${err.message}). Se registra el armado igual, sin verificar.`);
      }

      const registro = registrarArmado({
        orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente,
        usuarioArmado: req.usuarioActual,
        verificado,
        items: datos.items
      });

      // El envío al facturador no bloquea la respuesta: el armado ya está
      // guardado en disco y, si el envío falla, se reintenta solo.
      sincronizarArmados().catch(() => {});

      return sendJSON(res, 200, registro);
    }

    if (pathname === '/api/armado' && req.method === 'DELETE') {
      const orderId = parsed.searchParams.get('orderId');
      if (!orderId) return sendJSON(res, 400, { error: 'Falta orderId' });
      eliminarArmado(orderId);
      return sendJSON(res, 200, { ok: true });
    }

    if (pathname === '/api/armado' && req.method === 'GET') {
      const numeroOrden = parsed.searchParams.get('numeroOrden');
      const fechaDesde = parsed.searchParams.get('fechaDesde');
      const fechaHasta = parsed.searchParams.get('fechaHasta');
      const resultado = consultarArmados({ numeroOrden, fechaDesde, fechaHasta });
      return sendJSON(res, 200, { total: resultado.length, armados: resultado });
    }

    if (pathname === '/api/etiquetas/imprimir' && req.method === 'GET') {
      const orderId = String(parsed.searchParams.get('orderId') || '');
      const registro = armados.find(a => a.orderId === orderId);

      if (!registro) {
        res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
        return res.end('<h1>No se encontró el pedido armado</h1><p>Confirmá el armado antes de imprimir las etiquetas.</p>');
      }

      const cliente = await resolverClienteParaEtiqueta(registro);

      const html = generarHtmlEtiquetas({
        numeroOrden: registro.numeroOrden,
        bultos: registro.bultos,
        cliente
      });

      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(html);
    }

    if (pathname === '/api/etiquetas/imprimir-directo' && req.method === 'POST') {
      const raw = await readBody(req);
      const { orderId } = JSON.parse(raw || '{}');
      const registro = armados.find(a => a.orderId === String(orderId || ''));

      if (!registro) {
        return sendJSON(res, 404, { error: 'No se encontró el pedido armado. Confirmá el armado antes de imprimir.' });
      }

      const cfg = leerConfig();
      const nombreImpresora = cfg && cfg.printerName;
      if (!nombreImpresora) {
        return sendJSON(res, 400, { error: 'Todavía no configuraste el nombre de la impresora Zebra. Configurala arriba, en el recuadro de la impresora.' });
      }

      const cliente = await resolverClienteParaEtiqueta(registro);

      try {
        const { totalBultos, diagnostico } = await imprimirEtiquetasDirecto({
          numeroOrden: registro.numeroOrden,
          bultos: registro.bultos,
          cliente,
          nombreImpresora
        });
        registrarImpresion(registro.orderId, req.usuarioActual);
        return sendJSON(res, 200, { ok: true, impresas: totalBultos, diagnostico });
      } catch (err) {
        console.log(`  ⚠ Error al imprimir etiquetas de la orden ${registro.numeroOrden}: ${err.message}`);
        return sendJSON(res, 500, { error: `No se pudo imprimir: ${err.message}` });
      }
    }

    if (pathname === '/api/armado/reporte' && req.method === 'GET') {
      const numeroOrden = parsed.searchParams.get('numeroOrden');
      const fechaDesde = parsed.searchParams.get('fechaDesde');
      const fechaHasta = parsed.searchParams.get('fechaHasta');
      const resultado = consultarArmados({ numeroOrden, fechaDesde, fechaHasta });

      const filas = resultado.map(a => ({
        'N° Orden': a.numeroOrden,
        'Fecha de la Orden': a.fechaOrden,
        'Líneas': a.lineas,
        'Unidades': a.unidades,
        'Bultos': a.bultos,
        'Armado por': a.usuarioArmado || '-',
        'Armado el': formatearFechaHoraLocal(a.timestamp),
        'Impreso por': a.usuarioImpresion || '-',
        'Impreso el': a.fechaImpresion ? formatearFechaHoraLocal(a.fechaImpresion) : '-',
        // "-" en los armados anteriores a esta versión, que no se verificaban.
        'Orden verificada': a.verificado === undefined ? '-' : (a.verificado ? 'Sí' : 'No')
      }));

      const ws = XLSX.utils.json_to_sheet(filas);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Pedidos Armados');
      const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', bookSST: true });

      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="reporte_pedidos_armados.xlsx"',
        'Content-Length': buffer.length
      });
      return res.end(buffer);
    }

    // --- Archivos estáticos (la interfaz) ---
    if (req.method === 'GET') {
      return serveStatic(req, res, pathname);
    }

    res.writeHead(404);
    res.end('No encontrado');
  } catch (err) {
    sendJSON(res, 500, { error: err.message });
  }
});

const HOST = (leerConfig() || {}).host || HOST_DEFECTO;

server.listen(PORT, HOST, () => {
  cargarCatalogoInicial();
  cargarArmadosInicial();
  cargarClientesDesdeDiscoSiExiste();

  console.log('==================================================');
  console.log(' Lista de Picking - Contabilium');
  console.log(` Abrí tu navegador en: http://localhost:${PORT}`);
  console.log(` Catálogo cargado: ${catalogoInfo.cantidad} productos (${catalogoInfo.origen})`);
  console.log(` Pedidos armados en el historial: ${armados.length}`);
  if (HOST !== HOST_DEFECTO) {
    console.log(` ⚠ Escuchando en ${HOST}: la app queda accesible desde la red.`);
  }
  console.log(' Dejá esta ventana abierta mientras usás la app.');
  console.log('==================================================');

  // Reintento de los armados que quedaron sin llegar al facturador: uno al
  // arrancar (por si la PC estuvo apagada) y después cada 5 minutos.
  const pendientesAlArrancar = armadosPendientesDeEnvio().length;
  if (pendientesAlArrancar) {
    console.log(`  ${pendientesAlArrancar} armado(s) todavía sin enviar al facturador; se reintenta ahora.`);
  }
  sincronizarArmados().catch(() => {});
  setInterval(() => { sincronizarArmados().catch(() => {}); }, SYNC_INTERVALO_MS);

  // Actualiza la caché de clientes en segundo plano (no bloquea el arranque).
  // Si todavía no hay credenciales configuradas, simplemente no hace nada
  // hasta que se impriman etiquetas por primera vez.
  cargarClientes().catch(err => {
    console.log(`  (No se pudo refrescar la caché de clientes al arrancar: ${err.message})`);
  });

  // Abre el navegador automáticamente
  const url = `http://localhost:${PORT}`;
  const cmd = process.platform === 'win32'
    ? `start ${url}`
    : process.platform === 'darwin'
      ? `open ${url}`
      : `xdg-open ${url}`;
  exec(cmd, () => {});
});
