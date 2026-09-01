#!/usr/bin/env node
//
// Regenera products.json (el catálogo EAN/DUN que va adentro del .exe) a
// partir de la planilla maestra.
//
// Usa el MISMO parser que la pantalla "Subir tabla" de la app: levanta el
// servidor, le postea el .xlsx a /api/catalogo y se queda con lo que guardó.
// Parsear la planilla por afuera es la forma de que el catálogo embebido y el
// que sube el depósito terminen distintos — y ya pasó: leyendo el .xlsx con
// otra herramienta, dos celdas con fórmula (`=+F428`) entraron como texto y
// esos productos quedaron con el DUN roto y sin factor. Un DUN sin factor no
// falla: escanear la caja suma 1 unidad en vez de 24, en silencio.
//
// Uso:  node pkg/regenerar-catalogo.js "<ruta del xlsx>"

const fs = require('fs');
const path = require('path');
const http = require('http');
const os = require('os');
const { spawn, execFileSync } = require('child_process');

const RAIZ = path.join(__dirname, '..');
const PUERTO = 3000;
const xlsx = process.argv[2] || path.join(
  RAIZ, '..', 'Gestión de Vendedores - Claude + GSU', 'assets', 'Tabla EAN DUN GSU.xlsx',
);

if (!fs.existsSync(xlsx)) {
  console.error(`No encuentro la planilla: ${xlsx}`);
  process.exit(1);
}

// Carpeta aparte para no pisar el config.json ni el catalogo.json de nadie.
const carpeta = fs.mkdtempSync(path.join(os.tmpdir(), 'picking-catalogo-'));
for (const f of ['server-standalone.js', 'products.json', 'package.json']) {
  fs.copyFileSync(path.join(RAIZ, f), path.join(carpeta, f));
}
fs.cpSync(path.join(RAIZ, 'public'), path.join(carpeta, 'public'), { recursive: true });
fs.symlinkSync(path.join(RAIZ, 'node_modules'), path.join(carpeta, 'node_modules'));

const hijo = spawn(process.execPath, ['server-standalone.js'], {
  cwd: carpeta, stdio: 'ignore', detached: true,
});

function pedir(opciones, cuerpo) {
  return new Promise((resolve, reject) => {
    const req = http.request({ host: '127.0.0.1', port: PUERTO, timeout: 30000, ...opciones },
      res => {
        let body = '';
        res.on('data', c => { body += c; });
        res.on('end', () => resolve({ status: res.statusCode, body }));
      });
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
    req.on('error', reject);
    if (cuerpo) req.write(cuerpo);
    req.end();
  });
}

function limpiar() {
  try { process.kill(-hijo.pid); } catch { /* ya murió */ }
  try { hijo.kill('SIGKILL'); } catch { /* idem */ }
  fs.rmSync(carpeta, { recursive: true, force: true });
}

(async () => {
  let listo = false;
  for (let i = 0; i < 30 && !listo; i++) {
    try { await pedir({ method: 'GET', path: '/api/auth/estado' }); listo = true; }
    catch { await new Promise(r => setTimeout(r, 500)); }
  }
  if (!listo) { limpiar(); console.error('El servidor no arrancó.'); process.exit(1); }

  const cred = JSON.stringify({ usuario: 'catalogo', password: 'catalogo-temporal' });
  const alta = await pedir({
    method: 'POST', path: '/api/auth/registrar-primero',
    headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(cred) },
  }, cred);
  const token = JSON.parse(alta.body).token;

  const buffer = fs.readFileSync(xlsx);
  const subida = await pedir({
    method: 'POST', path: '/api/catalogo',
    headers: {
      'x-session-token': token,
      'x-filename': encodeURIComponent(path.basename(xlsx)),
      'Content-Type': 'application/octet-stream',
      'Content-Length': buffer.length,
    },
  }, buffer);

  if (subida.status !== 200) {
    limpiar();
    console.error('La app rechazó la planilla:', subida.body);
    process.exit(1);
  }

  const guardado = JSON.parse(fs.readFileSync(path.join(carpeta, 'catalogo.json'), 'utf8'));
  limpiar();

  const nuevos = guardado.productos;
  const viejos = JSON.parse(fs.readFileSync(path.join(RAIZ, 'products.json'), 'utf8'));
  const con = (l, k) => l.filter(x => x[k]).length;

  const dunSinFactor = nuevos.filter(p => p.dun && !p.factor);
  if (dunSinFactor.length) {
    console.error('\n⚠ Productos con DUN y sin factor — escanear la caja sumaría 1 unidad:');
    for (const p of dunSinFactor) console.error(`   ${p.sku} — ${p.nombre}`);
    console.error('Arreglá la planilla antes de usar este catálogo.');
    process.exit(1);
  }

  fs.writeFileSync(path.join(RAIZ, 'products.json'), JSON.stringify(nuevos, null, 2) + '\n');
  console.log(`productos : ${viejos.length} -> ${nuevos.length}`);
  console.log(`con EAN   : ${con(viejos, 'ean')} -> ${con(nuevos, 'ean')}`);
  console.log(`con DUN   : ${con(viejos, 'dun')} -> ${con(nuevos, 'dun')}`);
  console.log(`con factor: ${con(viejos, 'factor')} -> ${con(nuevos, 'factor')}`);
})();
