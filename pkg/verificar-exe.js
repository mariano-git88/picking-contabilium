#!/usr/bin/env node
//
// Prueba que el .exe recién compilado SIRVA LA PANTALLA, no solo que exista.
//
// Por qué existe este script: hasta la v1.2.1 el `pkg` config vivía en
// pkg/pkg-package.json y los globs de `assets` se resuelven relativo a la
// carpeta del config, así que "public/**/*" apuntaba a pkg/public/ — que no
// existe. El index.html nunca entró al ejecutable. El .exe arrancaba bien,
// abría el navegador solo, respondía la API... y devolvía 404 en "/". Nadie
// lo vio porque la única verificación era que el archivo estuviera.
//
// Encima `pkg` puede fallar y devolver exit code 0, así que ni el build
// rompe. La única prueba que vale es levantar el binario y pedirle la página.
//
// Uso:  npm run verificar    (después de npm run build)

const fs = require('fs');
const http = require('http');
const path = require('path');
const { spawn, execFileSync } = require('child_process');

const RAIZ = path.join(__dirname, '..');
const EXE = path.join(RAIZ, 'dist', 'PickingContabilium.exe');
const { version } = require(path.join(RAIZ, 'package.json'));

// Textos que tienen que estar en la página servida por el .exe. Si el
// index.html no entró al snapshot, no está ninguno.
const ESPERADOS = [
  'Lista de Picking',
  'El pedido sale incompleto',
  'Volver a leer esta orden',
];

const enWindows = process.platform === 'win32';
// El puerto está fijo en el server (const PORT = 3000), no se puede elegir.
// Si ya hay algo escuchando ahí, el .exe muere con EADDRINUSE y este script
// lo confundiría con "no arranca", así que lo chequeamos antes.
const PUERTO = 3000;

function salir(codigo, mensaje) {
  console[codigo ? 'error' : 'log'](mensaje);
  process.exit(codigo);
}

if (!fs.existsSync(EXE)) {
  salir(1, `No existe ${EXE}. Corré primero: npm run build`);
}

// El .exe es de Windows. Desde WSL se ejecuta por interop; desde Linux puro
// no hay forma, y avisamos en vez de dar un falso OK.
const esWSL = !enWindows && fs.existsSync('/proc/version')
  && /microsoft/i.test(fs.readFileSync('/proc/version', 'utf8'));
if (!enWindows && !esWSL) {
  salir(1, 'Este script necesita Windows o WSL para ejecutar el .exe.');
}

// La carpeta de prueba va adentro de dist/ y no en /tmp: desde WSL, un .exe
// de Windows lanzado con cwd en el sistema de archivos de Linux no arranca.
// dist/ está donde está el repo, que es donde Windows puede llegar.
const carpeta = fs.mkdtempSync(path.join(RAIZ, 'dist', '.verificacion-'));
fs.copyFileSync(EXE, path.join(carpeta, 'PickingContabilium.exe'));

// El .exe escucha del lado de WINDOWS. Desde WSL, 127.0.0.1 es el localhost
// de Linux: el reenvío de puertos de WSL2 va Windows -> WSL, no al revés. Por
// eso, corriendo en WSL, hay que sondear con powershell.exe y no con http.get
// — si no, el .exe anda perfecto y el script informa "nunca respondió".
function pedirHttp(ruta, timeout) {
  return new Promise((resolve, reject) => {
    const req = http.get({ host: '127.0.0.1', port: PUERTO, path: ruta, timeout }, res => {
      let body = '';
      res.on('data', c => { body += c; });
      res.on('end', () => resolve({ status: res.statusCode, body }));
    });
    req.on('timeout', () => { req.destroy(); reject(new Error('timeout')); });
    req.on('error', reject);
  });
}

// Ruta de WSL -> ruta de Windows (/mnt/c/x -> C:\x).
const aWin = p => p
  .replace(/^\/mnt\/([a-z])\//, (_, d) => d.toUpperCase() + ':\\')
  .replace(/\//g, '\\');

const PS1 = path.join(carpeta, 'sondear.ps1');
const RESP_BODY = path.join(carpeta, 'body.txt');
const RESP_STATUS = path.join(carpeta, 'status.txt');

// El script va a un .ps1 en vez de ir en -Command: escapar comillas a través
// de bash + WSL + PowerShell da falsos negativos silenciosos.
function escribirSondaPs1() {
  fs.writeFileSync(PS1, `param([string]$Ruta)
$ErrorActionPreference = 'Stop'
try {
  $r = Invoke-WebRequest ('http://127.0.0.1:${PUERTO}' + $Ruta) -UseBasicParsing -TimeoutSec 5
  [IO.File]::WriteAllText('${aWin(RESP_STATUS)}', [string]$r.StatusCode)
  [IO.File]::WriteAllText('${aWin(RESP_BODY)}', $r.Content, [Text.Encoding]::UTF8)
} catch {
  $codigo = ''
  if ($_.Exception.Response) { $codigo = [int]$_.Exception.Response.StatusCode }
  [IO.File]::WriteAllText('${aWin(RESP_STATUS)}', [string]$codigo)
  [IO.File]::WriteAllText('${aWin(RESP_BODY)}', '')
}
`);
}

function pedirPowershell(ruta) {
  fs.rmSync(RESP_STATUS, { force: true });
  fs.rmSync(RESP_BODY, { force: true });
  execFileSync('powershell.exe', [
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', aWin(PS1), '-Ruta', ruta,
  ], { stdio: 'ignore' });
  // PowerShell escribe UTF-8 CON BOM, y el BOM rompe el JSON.parse.
  const leer = f => (fs.existsSync(f) ? fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, '') : '');
  const status = leer(RESP_STATUS).trim();
  if (!status) throw new Error('sin respuesta');
  return { status: Number(status), body: leer(RESP_BODY) };
}

async function pedir(ruta, timeout = 4000) {
  if (enWindows) return pedirHttp(ruta, timeout);
  return pedirPowershell(ruta);
}

let hijo = null;

function matarlo() {
  try {
    if (enWindows && hijo) {
      execFileSync('taskkill', ['/PID', String(hijo.pid), '/T', '/F'], { stdio: 'ignore' });
    } else if (!enWindows) {
      execFileSync('powershell.exe', ['-NoProfile', '-Command',
        'Get-Process PickingContabilium -ErrorAction SilentlyContinue | Stop-Process -Force',
      ], { stdio: 'ignore' });
    }
  } catch { /* ya estaba muerto */ }
  fs.rmSync(carpeta, { recursive: true, force: true });
}

(async () => {
  // El puerto está fijo en el server, así que si ya hay algo escuchando el
  // .exe muere con EADDRINUSE y lo confundiríamos con "no arranca".
  try {
    await pedir('/', 1500);
    fs.rmSync(carpeta, { recursive: true, force: true });
    salir(1, `Ya hay algo escuchando en el puerto ${PUERTO}. Cerrá la app de picking (o lo que sea) y probá de nuevo.`);
  } catch { /* libre, seguimos */ }

    escribirSondaPs1();
  console.log(`Levantando el .exe v${version} (puerto ${PUERTO})...`);
  hijo = spawn(path.join(carpeta, 'PickingContabilium.exe'), [], {
    cwd: carpeta, stdio: 'ignore', detached: true,
  });
  hijo.unref();

  let pagina = null;
  for (let intento = 0; intento < 30; intento++) {
    try { pagina = await pedir('/'); break; }
    catch { await new Promise(r => setTimeout(r, 1000)); }
  }

  if (!pagina) { matarlo(); salir(1, 'El .exe nunca respondió. ¿Arranca?'); }

  const problemas = [];
  if (pagina.status !== 200) {
    problemas.push(`GET / devolvió ${pagina.status}, no 200. El index.html no entró al ejecutable (revisá "pkg.assets" en package.json).`);
  }
  for (const texto of ESPERADOS) {
    if (!pagina.body.includes(texto)) problemas.push(`La página no contiene "${texto}".`);
  }

  try {
    const estado = await pedir('/api/auth/estado');
    const datos = JSON.parse(estado.body);
    if (datos.version !== version) {
      problemas.push(`El .exe dice ser la ${datos.version} y package.json la ${version}.`);
    }
    if (!(datos.motivosFaltante || []).length) {
      problemas.push('El .exe no expone los motivos de faltante.');
    }
  } catch (err) {
    problemas.push('No se pudo leer /api/auth/estado: ' + err.message);
  }

  matarlo();

  if (problemas.length) salir(1, 'NO PASA:\n  - ' + problemas.join('\n  - '));
  salir(0, `OK: el .exe v${version} sirve la pantalla (${pagina.body.length} bytes) y la API responde.`);
})();
