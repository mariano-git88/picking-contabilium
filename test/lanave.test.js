// Pruebas de la lógica de La Nave que no necesita red.
// Correr con:  npm test
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const ln = require('../lanave.js');

test('agrupa por rango de peso: 4 kg + 18 kg + 17 kg = dos envíos', () => {
  const { envios, error } = ln.agruparBultos([4, 18, 17]);
  assert.equal(error, undefined);
  assert.deepEqual(envios.map(e => [e.producto, e.bultos, e.kilos]), [[241, 1, 4], [244, 2, 35]]);
  assert.deepEqual(envios[1].indicesBultos, [2, 3]);
});

test('los bordes de cada rango caen en el rango de abajo', () => {
  assert.equal(ln.rangoParaKilos(5).producto, 241);
  assert.equal(ln.rangoParaKilos(5.001).producto, 242);
  assert.equal(ln.rangoParaKilos(10).producto, 242);
  assert.equal(ln.rangoParaKilos(15).producto, 243);
  assert.equal(ln.rangoParaKilos(20).producto, 244);
  assert.equal(ln.rangoParaKilos(25).producto, 280);
});

test('un bulto de más de 25 kg frena el envío por La Nave', () => {
  const r = ln.agruparBultos([10, 25.5]);
  assert.ok(r.error);
  assert.equal(r.envios, undefined);
  assert.deepEqual(r.bultosExcedidos, [{ bulto: 2, kilos: 25.5 }]);
  assert.match(r.error, /bulto 2 \(25,5 kg\)/);
});

test('los kilos del envío van enteros, redondeando hacia arriba', () => {
  assert.equal(ln.agruparBultos([4.2]).envios[0].kilos, 5);
  assert.equal(ln.agruparBultos([0.3]).envios[0].kilos, 1);
  // 4.4 + 5.6 en coma flotante es 10.000000000000002: no tiene que dar 11
  assert.equal(ln.agruparBultos([4.4, 5.6]).envios.reduce((s, e) => s + e.kilos, 0), 5 + 6);
  assert.equal(ln.agruparBultos([6.4, 3.6 + 3]).envios[0].kilos, 13);
  const { envios } = ln.agruparBultos([9.4, 9.6]);
  assert.equal(envios[0].kilos, 19);
});

test('la clave externa es corta, solo letras/números/guiones, y lleva el intento', () => {
  assert.equal(ln.claveExterna('12773', 244, 1), '12773-244-1');
  assert.equal(ln.claveExterna('O/V 0001-12773', 241, 2), 'OV000112773-241-2');
  assert.ok(ln.claveExterna('x'.repeat(200), 280, 99).length <= 30);
  assert.equal(ln.claveExterna('', 241, 1), 'SN-241-1');
});

test('la localidad de las Observaciones manda sobre la Ciudad', () => {
  const r = ln.localidadDelCliente({
    Ciudad: 'Ciudad de la Costa',
    Observaciones: 'Entregas: L a V 08:00 a 17:00 hs.\nLocalidad: Shangrilá'
  });
  assert.deepEqual(r, { localidad: 'SHANGRILA', origen: 'observaciones' });
});

test('la línea de localidad se lee aunque esté escrita distinto', () => {
  assert.equal(ln.localidadDeObservaciones('LOCALIDAD:neptunia'), 'NEPTUNIA');
  assert.equal(ln.localidadDeObservaciones('  localidad - Santa Ana ; tocar timbre'), 'SANTA ANA');
  assert.equal(ln.localidadDeObservaciones('Entregas: L a V | Localidad: La Tuna'), 'LA TUNA');
  assert.equal(ln.localidadDeObservaciones('Entregas: L a V 8 a 12'), null);
  // En Contabilium el campo es casi siempre UN renglón: la etiqueta puede
  // venir después de los horarios, o antes de otro dato.
  assert.equal(ln.localidadDeObservaciones('Entregas: L a V 08:00 a 17:00 hs. Localidad: SHANGRILA'), 'SHANGRILA');
  assert.equal(ln.localidadDeObservaciones('Entregas: L a V 8 a 17 hs.\t\t\nLocalidad: Colinas de Solymar.'), 'COLINAS DE SOLYMAR');
  assert.equal(ln.localidadDeObservaciones('Localidad: El Pinar Entregas: L a V 8 a 12'), 'EL PINAR');
  assert.equal(ln.localidadDeObservaciones('Localidad: Sta. Lucía / tocar timbre'), 'STA. LUCIA');
  // "Sublocalidad:" no es la etiqueta
  assert.equal(ln.localidadDeObservaciones('Sublocalidad: X'), null);
});

test('sin línea en Observaciones se usa la Ciudad, sin tildes y en mayúsculas', () => {
  assert.deepEqual(ln.localidadDelCliente({ Ciudad: 'Paysandú', Observaciones: '' }), { localidad: 'PAYSANDU', origen: 'ciudad' });
});

test('Montevideo nunca va por La Nave', () => {
  assert.equal(ln.localidadDelCliente({ Ciudad: 'Montevideo' }).montevideo, true);
  assert.equal(ln.localidadDelCliente({ Ciudad: 'SIN IDENTIFICAR', Provincia: 'Montevideo' }).montevideo, true);
  assert.equal(ln.localidadDelCliente({ Ciudad: 'Pando', Observaciones: 'Localidad: Montevideo' }).montevideo, true);
});

test('ciudad sin cargar: se explica qué corregir', () => {
  const r = ln.localidadDelCliente({ Ciudad: 'SIN IDENTIFICAR', Provincia: 'SIN IDENTIFICAR' });
  assert.ok(r.problema);
  assert.equal(r.localidad, undefined);
  assert.match(r.problema, /Observaciones/);
  assert.ok(ln.localidadDelCliente({ Ciudad: '' }).problema);
  assert.ok(ln.localidadDelCliente(null).problema);
});

test('el destinatario sale sin comillas dobles ni saltos de línea', () => {
  const d = ln.destinatarioDelCliente({
    RazonSocial: 'PEÑAROL "EL MANYA" S.A.',
    Domicilio: 'AV. RIVERA 1234\nESQ. "X"',
    Cp: '15800',
    Telefono: '099 123 456'
  });
  assert.equal(d.destinatario, "PEÑAROL 'EL MANYA' S.A.");
  assert.equal(d.direccion, "AV. RIVERA 1234 ESQ. 'X' CP 15800");
  assert.equal(d.telefono, '099 123 456');
  assert.equal(ln.destinatarioDelCliente({ NombreFantasia: 'FERRETERIA X', RazonSocial: 'X SA' }).destinatario, 'FERRETERIA X');
});

// Etiqueta real devuelta por el ambiente de prueba el 24/11/2026.
const EPL_REAL = 'Q302,39,20\n<ENTER>ZT\n<ENTER>q540\n<ENTER>N\n<ENTER>A65,10,0,3,1,1,N,"OPE:111"\n<ENTER>A190,10,0,3,1,1,N,"LaNave-CargoPostal"\n<ENTER>A65,40,0,2,1,1,N,"PRUEBA SUPRABOND NO ENTREGAR"\n<ENTER>A65,65,0,2,1,1,N,""\n<ENTER>A65,90,0,2,1,1,N,"GIANNATTASIO KM 20 - PRUEBA"\n<ENTER>A420,90,0,4,1,1,N,""\n<ENTER>A340,125,0,2,1,1,N,""\n<ENTER>A65,125,0,2,1,1,N,"2497 SUPRABOND"\n<ENTER>B110,145,0,2,2,7,42,N,"0010996053012975"\n<ENTER>X65,198,2,513,228\n<ENTER>A73,205,0,2,1,1,N,"Tracking:"\n<ENTER>A185,201,0,4,1,1,N,"001 0 996053 01 2975"\n<ENTER>A65,230,0,2,1,1,N,"SHANGRILA (ZONA REPARTO)"\n<ENTER>A65,254,0,2,1,1,N,"24/11/26"\n<ENTER>A65,302,0,2,1,1,N,"242 BULTO HASTA 10 KG"\n<ENTER>A65,326,0,2,1,1,N,"PRUEBA-GSU-0001"\n<ENTER>A420,250,0,4,1,1,N,"R10-979"\n<ENTER>A420,275,0,4,1,1,N,"CF2"\n<ENTER>A420,300,0,4,1,1,N,"1/2"\n<ENTER>f130\n<ENTER>P1\n<ENTER>';

test('la etiqueta EPL de La Nave se desarma completa, códigos internos incluidos', () => {
  const e = ln.parsearEtiquetaEPL(EPL_REAL);
  assert.equal(e.ancho, 540);
  assert.equal(e.alto, 302);
  const textos = e.textos.map(t => t.texto);
  for (const esperado of ['OPE:111', 'R10-979', 'CF2', '1/2', '001 0 996053 01 2975', 'SHANGRILA (ZONA REPARTO)', '242 BULTO HASTA 10 KG']) {
    assert.ok(textos.includes(esperado), `falta "${esperado}"`);
  }
  assert.equal(e.barras.length, 1);
  assert.deepEqual(
    { x: e.barras[0].x, y: e.barras[0].y, tipo: e.barras[0].tipo, angosta: e.barras[0].angosta, ancha: e.barras[0].ancha, alto: e.barras[0].alto, dato: e.barras[0].dato },
    { x: 110, y: 145, tipo: '2', angosta: 2, ancha: 7, alto: 42, dato: '0010996053012975' }
  );
  assert.deepEqual(e.recuadros, [{ x1: 65, y1: 198, grosor: 2, x2: 513, y2: 228 }]);
  // Ningún texto arrastra el marcador <ENTER>
  assert.ok(!JSON.stringify(e).includes('ENTER'));
});

test('un texto con comas adentro no se parte', () => {
  const e = ln.parsearEtiquetaEPL('A65,90,0,2,1,1,N,"RUTA 8, KM 20, LOCAL 3"');
  assert.equal(e.textos[0].texto, 'RUTA 8, KM 20, LOCAL 3');
});
