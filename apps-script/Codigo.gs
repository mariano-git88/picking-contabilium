/**
 * Buzón de armados — Apps Script
 *
 * La app de Picking corre local en la PC del depósito y el facturador corre
 * en la nube: no se ven entre sí. Este script es el punto de encuentro.
 * Recibe un armado por POST y lo apenda a la tab `armados` del Sheet
 * `gsu-facturacion-log`, de donde el facturador lo lee.
 *
 * Por qué esto y no la librería de Google adentro del .exe: para que la PC
 * del depósito no tenga que guardar las credenciales del service account,
 * que dan acceso a TODOS los Sheets de GSU (histórico, comisiones,
 * cobranzas). Acá lo único que viaja a esa máquina es una URL y un token
 * que no sirve para nada más que agregar una fila de armado.
 *
 * ---------------------------------------------------------------------
 * CÓMO SE INSTALA (una sola vez)
 * ---------------------------------------------------------------------
 * 1. Abrir el Sheet `gsu-facturacion-log` → Extensiones → Apps Script.
 * 2. Pegar este archivo entero, reemplazando lo que haya.
 * 3. Configuración → Propiedades del script → agregar dos propiedades:
 *      SPREADSHEET_ID  → el id del Sheet (está en la URL, entre /d/ y /edit)
 *      TOKEN           → una clave larga inventada, la misma que se carga
 *                        después en la app de Picking
 * 4. Implementar → Nueva implementación → tipo "Aplicación web":
 *      Ejecutar como:      Yo
 *      Quién tiene acceso: Cualquier usuario
 *    (Hace falta "cualquier usuario" porque la PC del depósito no inicia
 *     sesión de Google. Por eso el token: sin token no escribe nada.)
 * 5. Copiar la URL que queda (termina en /exec) y cargarla en la app de
 *    Picking junto con el token, en el recuadro "Envío al facturador".
 *
 * Si algún día se cambia el token, hay que cambiarlo en los dos lados.
 * Si se modifica este código, hay que volver a "Implementar" para que la
 * versión publicada se actualice — guardar no alcanza.
 */

// Mismo orden y mismos nombres que ARMADOS_COLUMNS en gsheets.py.
// Si se toca acá, hay que tocar allá.
var COLUMNAS = [
  'timestamp',
  'fecha_local',
  'evento',
  'id_orden',
  'numero_orden',
  'fecha_orden',
  'id_cliente',
  'bultos',
  'lineas',
  'unidades',
  'usuario',
  'completo',
  'verificado',
  'items_json',
  'id_comprobante',
  'numero_factura',
  'cae',
  'observacion'
];

var TAB = 'armados';

function doPost(e) {
  try {
    var props = PropertiesService.getScriptProperties();
    var tokenEsperado = props.getProperty('TOKEN');
    var spreadsheetId = props.getProperty('SPREADSHEET_ID');

    if (!tokenEsperado || !spreadsheetId) {
      return responder(false, 'El script no está configurado (faltan TOKEN o SPREADSHEET_ID en las propiedades).');
    }
    if (!e || !e.postData || !e.postData.contents) {
      return responder(false, 'Pedido vacío.');
    }

    var body = JSON.parse(e.postData.contents);

    // Comparación de largo constante, para no filtrar el token de a un
    // caracter por diferencia de tiempo de respuesta.
    if (!tokenIgual(body.token, tokenEsperado)) {
      return responder(false, 'Token inválido.');
    }

    var filas = body.filas;
    if (!Array.isArray(filas) || filas.length === 0) {
      return responder(false, 'No vinieron filas para agregar.');
    }
    if (filas.length > 200) {
      return responder(false, 'Demasiadas filas en un solo envío (máximo 200).');
    }

    // Sin candado, dos armados simultáneos pueden escribir sobre la misma
    // fila y perderse uno.
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return responder(false, 'El buzón está ocupado, reintentá en un momento.');
    }

    try {
      var ws = obtenerTab(spreadsheetId);
      var datos = filas.map(function (f) {
        return COLUMNAS.map(function (c) {
          var v = f[c];
          return (v === null || v === undefined) ? '' : String(v);
        });
      });

      // setValues en vez de appendRow: escribe todo el lote de una y no
      // reinterpreta los valores, así "00012036" no se convierte en 12036.
      var primeraLibre = ws.getLastRow() + 1;
      ws.getRange(primeraLibre, 1, datos.length, COLUMNAS.length).setValues(datos);

      return responder(true, null, { filas: datos.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return responder(false, String(err));
  }
}

function doGet() {
  // Sirve para probar desde el navegador que la implementación está viva.
  return responder(true, null, { mensaje: 'Buzón de armados operativo.' });
}

function obtenerTab(spreadsheetId) {
  var sh = SpreadsheetApp.openById(spreadsheetId);
  var ws = sh.getSheetByName(TAB);
  if (!ws) {
    ws = sh.insertSheet(TAB);
  }
  // Encabezado: si la tab está vacía, lo escribimos.
  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]);
    ws.setFrozenRows(1);
  }
  // La columna del número de orden va como texto, para que los ceros de la
  // izquierda sobrevivan aunque alguien edite el Sheet a mano.
  var colNumero = COLUMNAS.indexOf('numero_orden') + 1;
  ws.getRange(2, colNumero, ws.getMaxRows() - 1, 1).setNumberFormat('@');
  return ws;
}

function tokenIgual(recibido, esperado) {
  if (typeof recibido !== 'string' || recibido.length !== esperado.length) {
    return false;
  }
  var distinto = 0;
  for (var i = 0; i < esperado.length; i++) {
    distinto |= recibido.charCodeAt(i) ^ esperado.charCodeAt(i);
  }
  return distinto === 0;
}

/**
 * Apps Script siempre responde HTTP 200 desde una aplicación web: no se
 * puede devolver un 400 o un 401. Por eso el resultado real viaja en el
 * campo `ok` del cuerpo, y quien llama TIENE que mirarlo — si se guía por
 * el código HTTP, va a dar por bueno cualquier rechazo.
 */
function responder(ok, error, extra) {
  var payload = { ok: ok };
  if (error) payload.error = error;
  if (extra) {
    for (var k in extra) payload[k] = extra[k];
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
