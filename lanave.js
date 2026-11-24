// Integración con La Nave CargoPostal (transporte del interior).
//
// Este módulo no sabe nada del servidor: recibe datos y devuelve datos, y el
// cliente HTTP recibe las credenciales por parámetro. Así se puede probar solo
// (test/lanave.test.js) y el servidor decide qué guardar y cuándo.
//
// Lo que se aprendió probando contra el ambiente de prueba (nov-2026), y que
// la documentación de La Nave no dice:
//   - Un envío (remito) lleva UN código de producto y la cantidad de bultos de
//     ese rango. Un pedido con bultos de distinto peso son varios envíos.
//   - El precio sale del código, NO de los kilos: un 241 con 18 kg cotiza
//     igual que con 4 kg, y un bulto de 30 kg pasa sin error. Elegir bien el
//     rango y frenar los de más de 25 kg es responsabilidad nuestra.
//   - `kilos` tiene que ser ENTERO: 4.5 da "El dato: kilos no es válido".
//   - La clave externa no se puede repetir mientras el envío esté vivo (da
//     errorCodigo 2, genérico, con el texto "ya se utilizó"). Una clave de un
//     envío ANULADO sí se puede reusar, y la consulta pasa a devolver el más
//     nuevo: por eso la clave lleva el número de intento.
//   - Con una clave de 80 caracteres el envío se crea pero después la
//     consulta no lo encuentra. Con 50 anda: nos quedamos bien abajo.
//   - Las comillas dobles en el nombre o la dirección se meten crudas en la
//     etiqueta EPL y la rompen.
//   - La etiqueta viene en EPL, con un "<ENTER>" literal pegado a cada salto
//     de línea, y trae códigos internos de clasificación ("R10-979", "CF2")
//     que no están en ningún otro campo. Hay que guardarla al recibirla: la
//     consulta posterior devuelve el código de barras pero no la etiqueta.

'use strict';

const https = require('https');

// --- Rangos de peso -------------------------------------------------------
//
// Códigos de producto del cliente 2497 (SUPRABOND), tipo 02 crédito. El 243
// está dado de alta en producción pero no en el ambiente de prueba (Gabriel,
// 25/11/2026): si lo rechaza, el error tiene que verse, no reemplazarse en
// silencio por el 244, que es más caro.
const RANGOS = [
  { producto: 241, hasta: 5, nombre: 'Bulto hasta 5 kg' },
  { producto: 242, hasta: 10, nombre: 'Bulto hasta 10 kg' },
  { producto: 243, hasta: 15, nombre: 'Bulto hasta 15 kg' },
  { producto: 244, hasta: 20, nombre: 'Bulto hasta 20 kg' },
  { producto: 280, hasta: 25, nombre: 'Bulto hasta 25 kg' }
];

// Normativa ergonómica (GSU y transportes): ningún bulto de más de 25 kg.
const KILOS_MAX_BULTO = 25;

function rangoParaKilos(kilos) {
  return RANGOS.find(r => kilos <= r.hasta) || null;
}

// Agrupa los bultos del armado en envíos, uno por rango de peso. Devuelve
// { envios } o { error, bultosExcedidos } si algún bulto pasa de 25 kg.
//
// `kilos` del envío se redondea HACIA ARRIBA: La Nave solo acepta enteros, y
// declarar de menos es lo que después aparece como corrección en el pesaje.
function agruparBultos(kilosBultos) {
  if (!Array.isArray(kilosBultos) || !kilosBultos.length) {
    return { error: 'El armado no tiene el peso de los bultos.' };
  }
  const excedidos = [];
  kilosBultos.forEach((k, i) => { if (!(k > 0) || k > KILOS_MAX_BULTO) excedidos.push({ bulto: i + 1, kilos: k }); });
  if (excedidos.length) {
    const cuales = excedidos.map(e => `bulto ${e.bulto} (${String(e.kilos).replace('.', ',')} kg)`).join(', ');
    return {
      error: `No se puede mandar por La Nave un bulto de más de ${KILOS_MAX_BULTO} kg: ${cuales}. Hay que dividirlo y volver a armar.`,
      bultosExcedidos: excedidos
    };
  }

  const porProducto = new Map();
  kilosBultos.forEach((k, i) => {
    const r = rangoParaKilos(k);
    if (!porProducto.has(r.producto)) {
      porProducto.set(r.producto, { producto: r.producto, nombre: r.nombre, bultos: 0, kilosBultos: [], indicesBultos: [] });
    }
    const g = porProducto.get(r.producto);
    g.bultos += 1;
    g.kilosBultos.push(k);
    g.indicesBultos.push(i + 1); // número de bulto del armado, base 1
  });

  const envios = [...porProducto.values()]
    .sort((a, b) => a.producto - b.producto)
    .map(g => {
      const suma = g.kilosBultos.reduce((s, k) => s + k, 0);
      // El redondeo de coma flotante no tiene que sumar un kilo de más:
      // 4.4 + 5.6 = 10.000000000000002 → 10, no 11.
      return { ...g, kilos: Math.max(1, Math.ceil(Math.round(suma * 1000) / 1000)) };
    });
  return { envios };
}

// --- Clave externa --------------------------------------------------------
//
// Única forma de ligar un envío de La Nave con nuestro pedido. Formato:
// <orden>-<producto>-<intento>, por ejemplo "12773-244-1". Solo letras,
// números y guiones, y corta: bien por debajo de los 50 que se probaron.
function claveExterna(numeroOrden, producto, intento) {
  const orden = String(numeroOrden == null ? '' : numeroOrden).replace(/[^0-9A-Za-z]/g, '').slice(-16) || 'SN';
  return `${orden}-${producto}-${intento}`;
}

// --- Localidad --------------------------------------------------------------
//
// El desplegable de Ciudad de Contabilium tiene menos localidades que La Nave
// (no existen Shangrilá, Neptunia, Santa Ana…). Decisión del 25/11/2026: la
// localidad exacta se escribe en las Observaciones del cliente, en una línea
// "Localidad: SHANGRILA". Esa línea manda; si no está, se usa la Ciudad.
function sinTildes(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function normalizarLocalidad(s) {
  return sinTildes(s).toUpperCase().replace(/["']/g, '').replace(/\s+/g, ' ').trim();
}

function localidadDeObservaciones(observaciones) {
  const lineas = String(observaciones || '').split(/\r?\n|\|/);
  for (const l of lineas) {
    // Tolerante: "Localidad: X", "LOCALIDAD:X", "localidad - X". Lo que
    // sigue a un punto y coma o a un " / " ya es otro dato.
    const m = sinTildes(l).match(/^\s*localidad\s*[:\-]\s*(.+)$/i);
    if (m) {
      const valor = normalizarLocalidad(m[1].split(/;|\s\/\s/)[0]);
      if (valor) return valor;
    }
  }
  return null;
}

function esMontevideo(texto) {
  return normalizarLocalidad(texto) === 'MONTEVIDEO';
}

// Devuelve { localidad, origen } o { problema, montevideo? }.
function localidadDelCliente(cliente) {
  if (!cliente) return { problema: 'No se encontró el cliente de la orden en Contabilium.' };
  const deObs = localidadDeObservaciones(cliente.Observaciones);
  if (deObs) {
    if (esMontevideo(deObs)) return { problema: 'El cliente es de Montevideo: va con camión propio, no por La Nave.', montevideo: true };
    return { localidad: deObs, origen: 'observaciones' };
  }
  const ciudad = normalizarLocalidad(cliente.Ciudad);
  if (esMontevideo(ciudad) || normalizarLocalidad(cliente.Provincia) === 'MONTEVIDEO') {
    return { problema: 'El cliente es de Montevideo: va con camión propio, no por La Nave.', montevideo: true };
  }
  if (!ciudad || ciudad === 'SIN IDENTIFICAR') {
    return {
      problema: 'El cliente no tiene la ciudad cargada en Contabilium. Hay que corregirla en la ficha del cliente (Ciudad, o una línea "Localidad: …" en Observaciones) y volver a intentar.'
    };
  }
  return { localidad: ciudad, origen: 'ciudad' };
}

// --- Datos del destinatario -------------------------------------------------
//
// Lo que va a la etiqueta de La Nave: sin comillas dobles (rompen el EPL),
// sin caracteres de control, espacios colapsados y con tope de largo.
function limpiarTexto(s, max) {
  return String(s == null ? '' : s)
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/"/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, max);
}

function destinatarioDelCliente(cliente) {
  const nombre = cliente.NombreFantasia || cliente.RazonSocial || '';
  const direccion = [cliente.Domicilio, cliente.Cp ? `CP ${cliente.Cp}` : ''].filter(Boolean).join(' ');
  return {
    destinatario: limpiarTexto(nombre, 60),
    direccion: limpiarTexto(direccion, 100),
    telefono: limpiarTexto(cliente.Telefono, 30)
  };
}

// --- Etiqueta EPL -------------------------------------------------------------
//
// Se guarda la etiqueta "desarmada": los textos con su posición, el código de
// barras y los recuadros. Con eso armamos nuestra propia etiqueta ZPL de
// 10x15 (la que el depósito ya imprime bien), conservando todo lo que trae
// la de La Nave, incluidos los códigos de clasificación.
function separarCampos(s) {
  // Campos separados por coma; el último puede ser un texto entre comillas
  // que a su vez trae comas.
  const out = [];
  let actual = '';
  let enComillas = false;
  for (const ch of s) {
    if (ch === '"') { enComillas = !enComillas; actual += ch; continue; }
    if (ch === ',' && !enComillas) { out.push(actual); actual = ''; continue; }
    actual += ch;
  }
  out.push(actual);
  return out;
}

function desComillar(s) {
  const t = String(s || '').trim();
  return t.startsWith('"') && t.endsWith('"') && t.length >= 2 ? t.slice(1, -1) : t;
}

function parsearEtiquetaEPL(epl) {
  const lineas = String(epl || '').replace(/<ENTER>/g, '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const textos = [];
  const barras = [];
  const recuadros = [];
  let ancho = null;
  let alto = null;
  for (const l of lineas) {
    if (/^q\d+/.test(l)) { ancho = Number(l.slice(1)); continue; }
    if (/^Q\d+/.test(l)) { alto = Number(l.slice(1).split(',')[0]); continue; }
    if (l[0] === 'A' && /^A\d/.test(l)) {
      const c = separarCampos(l.slice(1));
      if (c.length >= 8) {
        textos.push({
          x: Number(c[0]), y: Number(c[1]), rot: Number(c[2]), fuente: c[3],
          multH: Number(c[4]) || 1, multV: Number(c[5]) || 1, inverso: c[6] === 'R',
          texto: desComillar(c.slice(7).join(','))
        });
      }
      continue;
    }
    if (l[0] === 'B' && /^B\d/.test(l)) {
      const c = separarCampos(l.slice(1));
      if (c.length >= 9) {
        barras.push({
          x: Number(c[0]), y: Number(c[1]), rot: Number(c[2]), tipo: c[3],
          angosta: Number(c[4]), ancha: Number(c[5]), alto: Number(c[6]),
          legible: c[7] === 'B', dato: desComillar(c.slice(8).join(','))
        });
      }
      continue;
    }
    if (l[0] === 'X' && /^X\d/.test(l)) {
      const c = l.slice(1).split(',').map(Number);
      if (c.length >= 5) recuadros.push({ x1: c[0], y1: c[1], grosor: c[2], x2: c[3], y2: c[4] });
    }
  }
  return { ancho, alto, textos, barras, recuadros };
}

// --- Cliente HTTP -------------------------------------------------------------
//
// La contraseña viaja en el cuerpo de cada pedido: este cliente NUNCA
// escribe el cuerpo ni la respuesta cruda en la consola, y los mensajes de
// error que devuelve no la contienen.
function crearCliente({ cliente, pwd, ambiente }) {
  const prueba = ambiente !== 'produccion';
  const base = prueba
    ? 'https://cargopostal.uy:8443/PostalNetClientes/restApi/'
    : 'https://cargopostal.uy/PostalNetClientes/restApi/';

  function tapar(texto) {
    return pwd ? String(texto).split(pwd).join('***') : String(texto);
  }

  function post(endpoint, cuerpo, timeoutMs = 45000) {
    const datos = Buffer.from(JSON.stringify({ cliente: String(cliente), pwd, ...cuerpo }), 'utf8');
    return new Promise((resolve, reject) => {
      const req = https.request(base + endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8', 'Content-Length': datos.length },
        timeout: timeoutMs
      }, res => {
        const partes = [];
        res.on('data', d => partes.push(d));
        res.on('end', () => {
          const texto = Buffer.concat(partes).toString('utf8');
          let json = null;
          try { json = JSON.parse(texto); } catch { /* se informa abajo */ }
          if (!json) return reject(new Error(`La Nave respondió algo que no se entiende (HTTP ${res.statusCode}).`));
          // Los errores de validación de campos vienen como HTTP 500 con
          // {status:"ERROR", status_message}; los de negocio, como 200 con
          // errorCodigo. Se devuelven los dos como respuesta, no como caída.
          resolve({ http: res.statusCode, json });
        });
      });
      req.on('timeout', () => req.destroy(new Error('La Nave no respondió a tiempo.')));
      req.on('error', err => reject(new Error(tapar(`No se pudo conectar con La Nave: ${err.message}`))));
      req.write(datos);
      req.end();
    });
  }

  function errorDe(json) {
    if (!json) return 'Respuesta vacía de La Nave.';
    if (json.status === 'ERROR') return tapar(json.status_message || 'Error de La Nave.');
    if (json.errorCodigo && json.errorCodigo !== 0) return tapar(`${json.errorDescripcion || 'Error'} (código ${json.errorCodigo})`);
    return null;
  }

  function cuerpoEnvio({ generarRemito, clave, producto, bultos, kilos, destinatario, direccion, localidad, telefono }) {
    return {
      generarRemito,
      agenciaOrigen: '001',
      claveExterna: clave,
      tipo: 'C', // todo GSU va a crédito (confirmado por Mariano, 25/11/2026)
      producto,
      bultos,
      kilos,
      destinatario,
      direccion,
      localidad,
      rut: '',
      valorCR: 0,
      facturaCR: '',
      telefonoDest: telefono || '',
      // Vacío a propósito: con un mail, La Nave le escribe al cliente con
      // cada cambio de estado. Nadie decidió eso todavía.
      mailDest: '',
      latitudDest: '',
      longitudDest: ''
    };
  }

  async function enviar(datos, generarRemito) {
    const { json } = await post('envio', cuerpoEnvio({ ...datos, generarRemito }));
    const error = errorDe(json);
    if (error) return { ok: false, error, codigo: json.errorCodigo };
    return {
      ok: true,
      remito: json.remito,
      agencia: json.agencia,
      precio: json.precio,
      localidadLaNave: json.localidadDestinoTiempost || '',
      codigoBarra: json.codigoBarra || [],
      etiquetas: json.etiqueta || [],
      fechaHora: json.fechaHora || ''
    };
  }

  return {
    prueba,
    cotizar: datos => enviar({ ...datos, clave: datos.clave || 'COTIZACION' }, 'N'),
    despachar: datos => enviar(datos, 'S'),
    async consultar(clave) {
      const { json } = await post('consulta', { claveExterna: clave });
      // "No existe el remito" no es una falla: es la respuesta que esperamos
      // antes de despachar por primera vez.
      if (json && json.errorCodigo === 4) return { ok: true, existe: false };
      const error = errorDe(json);
      if (error) return { ok: false, error };
      if (!json.remito) return { ok: true, existe: false };
      return {
        ok: true,
        existe: true,
        remito: json.remito,
        agencia: json.agencia,
        estadoCodigo: json.estadoCodigo,
        estado: json.estadoDescripcion,
        codigoBarra: json.codigoBarra || [],
        claveExterna: json.claveExterna
      };
    },
    async anular(agencia, remito) {
      const { json } = await post('anularEnvio', { agenciaOrigen: agencia || '001', remito });
      if (json && json.status === 'OK') return { ok: true, mensaje: json.status_message || '' };
      return { ok: false, error: tapar((json && (json.status_message || json.errorDescripcion)) || 'La Nave no pudo anular el envío.') };
    }
  };
}

module.exports = {
  RANGOS,
  KILOS_MAX_BULTO,
  rangoParaKilos,
  agruparBultos,
  claveExterna,
  localidadDeObservaciones,
  localidadDelCliente,
  normalizarLocalidad,
  destinatarioDelCliente,
  limpiarTexto,
  parsearEtiquetaEPL,
  crearCliente
};
