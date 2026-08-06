const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { URL } = require('url');
const { exec, spawn } = require('child_process');
const XLSX = require('xlsx');

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
  try {
    const raw = JSON.parse(fs.readFileSync(catalogoGuardadoPath, 'utf8'));
    productos = raw.productos;
    catalogoInfo = { origen: 'subido', actualizado: raw.actualizado, archivo: raw.archivo, cantidad: productos.length };
  } catch {
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
  fs.writeFileSync(catalogoGuardadoPath, JSON.stringify({
    productos,
    actualizado: catalogoInfo.actualizado,
    archivo: nombreArchivo
  }, null, 2));
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
const WKHTMLTOIMAGE_PATH = path.join(APP_DIR, 'bin', 'wkhtmltoimage.exe');
const LOGO_PATH = path.join(APP_DIR, 'assets', 'logo_suprabond.png');
const TEMP_DIR = path.join(os.tmpdir(), 'picking-contabilium-etiquetas');

const BASE_URL = 'https://rest.contabilium.com.uy';
const PORT = 3000;

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
  try {
    armados = JSON.parse(fs.readFileSync(ARMADOS_PATH, 'utf8'));
  } catch {
    armados = [];
  }
}

function guardarArmadosDisco() {
  fs.writeFileSync(ARMADOS_PATH, JSON.stringify(armados, null, 2));
}

function registrarArmado({ orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente, usuarioArmado }) {
  orderId = String(orderId);
  const timestamp = new Date().toISOString();
  const existente = armados.findIndex(a => a.orderId === orderId);
  const registro = { orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente, usuarioArmado, timestamp };
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

function fechaISOaDate(iso) {
  return iso.slice(0, 10); // YYYY-MM-DD
}

function consultarArmados({ numeroOrden, fechaDesde, fechaHasta }) {
  let resultado = [...armados];
  if (numeroOrden) {
    const buscado = numeroOrden.trim().toLowerCase();
    resultado = resultado.filter(a => (a.numeroOrden || '').toLowerCase().includes(buscado));
  }
  if (fechaDesde) {
    resultado = resultado.filter(a => fechaISOaDate(a.timestamp) >= fechaDesde);
  }
  if (fechaHasta) {
    resultado = resultado.filter(a => fechaISOaDate(a.timestamp) <= fechaHasta);
  }
  return resultado.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

// --- Config (credenciales) ---
function leerConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  } catch {
    return null;
  }
}
function guardarConfig(cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2));
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

// Sesiones en memoria: token -> { usuario, creado }
const sesiones = new Map();

function crearSesion(usuario) {
  const token = crypto.randomBytes(24).toString('hex');
  sesiones.set(token, { usuario, creado: Date.now() });
  return token;
}

function usuarioDeSesion(token) {
  const s = sesiones.get(token);
  return s ? s.usuario : null;
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

async function authedFetch(url) {
  const token = await getToken();
  const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Error consultando Contabilium (${resp.status}): ${text}`);
  }
  return resp.json();
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
  try {
    const raw = JSON.parse(fs.readFileSync(CLIENTES_CACHE_PATH, 'utf8'));
    clientesPorId = new Map();
    for (const c of raw.clientes) clientesPorId.set(String(c.Id), c);
    clientesCargados = true;
    clientesUltimaActualizacion = new Date(raw.actualizado).getTime() || 0;
    console.log(`  Clientes cargados desde caché en disco: ${clientesPorId.size} (actualizado ${raw.actualizado})`);
  } catch {
    // No hay caché en disco todavía; se cargará desde la API cuando haga falta
  }
}

function guardarClientesDisco() {
  try {
    fs.writeFileSync(CLIENTES_CACHE_PATH, JSON.stringify({
      actualizado: new Date().toISOString(),
      clientes: Array.from(clientesPorId.values())
    }));
  } catch (err) {
    console.log(`  ⚠ No se pudo guardar la caché de clientes en disco: ${err.message}`);
  }
}

async function cargarClientesInterno() {
  const nuevoMapa = new Map();
  const CONCURRENCIA = 8;
  let page = 1;
  let seguir = true;

  while (seguir) {
    const paginas = Array.from({ length: CONCURRENCIA }, (_, i) => page + i);
    const resultados = await Promise.all(
      paginas.map(p =>
        authedFetch(`${BASE_URL}/api/clientes/search?pageSize=500&page=${p}`)
          .then(data => data.Items || data.items || (Array.isArray(data) ? data : []))
          .catch(() => [])
      )
    );

    let algunaVacia = false;
    for (let i = 0; i < resultados.length; i++) {
      const items = resultados[i];
      for (const c of items) nuevoMapa.set(String(c.Id), c);
      console.log(`  Clientes: página ${paginas[i]} → ${items.length} registro(s)`);
      if (items.length === 0) algunaVacia = true;
    }

    if (algunaVacia) seguir = false;
    else page += CONCURRENCIA;

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

async function obtenerCliente(idCliente) {
  if (!idCliente) return null;
  if (!clientesCargados) await cargarClientes();

  let c = clientesPorId.get(String(idCliente));

  // Si no aparece, puede ser un cliente nuevo: solo re-consultamos si la
  // caché ya tiene un rato (para no volver a paginar todo por cada intento)
  if (!c && Date.now() - clientesUltimaActualizacion > 5 * 60 * 1000) {
    await cargarClientes();
    c = clientesPorId.get(String(idCliente));
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

function encabezadoHtml() {
  return `GRUPO SUPRABOND URUGUAY SAS`;
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
      <div class="encabezado">${encabezadoHtml()}</div>
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
  ${ESTILOS_ETIQUETA}
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

// HTML de UNA sola etiqueta, en tamaño de píxeles exacto, para renderizar a
// PNG con wkhtmltoimage y enviarla directo a la impresora sin diálogo.
const ETIQUETA_PX_ANCHO = 1772; // 15cm a 300dpi
const ETIQUETA_PX_ALTO = 1181;  // 10cm a 300dpi

function generarHtmlEtiquetaUnica({ numeroOrden, i, totalBultos, cliente }) {
  const datos = datosEtiquetaDesdeCliente(cliente);
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8" />
<style>
  html, body { width: ${ETIQUETA_PX_ANCHO}px; height: ${ETIQUETA_PX_ALTO}px; }
  ${ESTILOS_ETIQUETA}
  .etiqueta { width: ${ETIQUETA_PX_ANCHO}px; height: ${ETIQUETA_PX_ALTO}px; padding: 60px; }
  .cuerpo { flex: 1; margin-top: 0.4cm; display: flex; flex-direction: column; justify-content: space-between; }
  .campo { font-size: 46pt; margin-bottom: 0; }
  .observaciones { font-size: 34pt; }
  .pie { font-size: 36pt; }
  .bulto, .ov { font-size: 58pt; }
  .encabezado { min-height: 170px; font-size: 54pt; }
  .encabezado img.logo { max-height: 170px; }
</style>
</head>
<body>
${contenidoEtiqueta({ numeroOrden, i, totalBultos, ...datos })}
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
      return sendJSON(res, 200, { hayUsuarios, autenticado: !!usuario, usuario });
    }

    if (pathname === '/api/auth/registrar-primero' && req.method === 'POST') {
      if (leerUsuarios().length > 0) {
        return sendJSON(res, 400, { error: 'Ya existen usuarios. Pedile a un usuario existente que te cree una cuenta.' });
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
        printerName: cfg ? (cfg.printerName || '') : ''
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
      // Solo interesan las órdenes pendientes de picking
      const filtradas = ordenes.filter(o => o.Estado === 'Pendiente');

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
      const { orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente } = JSON.parse(raw || '{}');
      if (!orderId || !bultos) {
        return sendJSON(res, 400, { error: 'Faltan datos (orderId, bultos)' });
      }
      const registro = registrarArmado({ orderId, numeroOrden, fechaOrden, bultos, lineas, unidades, idCliente, usuarioArmado: req.usuarioActual });
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
        'Armado el': new Date(a.timestamp).toLocaleString('es-AR'),
        'Impreso por': a.usuarioImpresion || '-',
        'Impreso el': a.fechaImpresion ? new Date(a.fechaImpresion).toLocaleString('es-AR') : '-'
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

server.listen(PORT, () => {
  cargarCatalogoInicial();
  cargarArmadosInicial();
  cargarClientesDesdeDiscoSiExiste();

  console.log('==================================================');
  console.log(' Lista de Picking - Contabilium');
  console.log(` Abrí tu navegador en: http://localhost:${PORT}`);
  console.log(` Catálogo cargado: ${catalogoInfo.cantidad} productos (${catalogoInfo.origen})`);
  console.log(' Dejá esta ventana abierta mientras usás la app.');
  console.log('==================================================');

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
