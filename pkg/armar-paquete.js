#!/usr/bin/env node
//
// Arma el ZIP que se le manda al depósito: el .exe recién compilado, el logo
// de las etiquetas, la credencial del buzón, la dirección de la planilla y un
// LEEME con los 3 pasos.
//
// Se hace por script y no a mano porque olvidarse un archivo no se nota hasta
// que la app ya está instalada en la PC del depósito: sin `buzon-sa.json` los
// armados no le llegan a nadie, y sin `assets/` las etiquetas salen sin logo.
// El script verifica que estén todos antes de comprimir.
//
// Uso:  npm run paquete   (después de npm run build)

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const RAIZ = path.join(__dirname, '..');
const { version } = require(path.join(RAIZ, 'package.json'));
const DIST = path.join(RAIZ, 'dist');
const SALIDA = path.join(DIST, 'actualizacion');

// origen -> destino dentro del paquete
const ARCHIVOS = [
  [path.join(DIST, 'PickingContabilium.exe'), 'PickingContabilium.exe'],
  [path.join(RAIZ, 'assets', 'logo_suprabond.png'), path.join('assets', 'logo_suprabond.png')],
  [path.join(RAIZ, 'buzon-sa.json'), 'buzon-sa.json'],
];

const faltan = ARCHIVOS.filter(([origen]) => !fs.existsSync(origen)).map(([o]) => o);
if (faltan.length) {
  console.error('Faltan archivos para armar el paquete:\n  ' + faltan.join('\n  '));
  console.error('\nSi falta el .exe, corré primero: npm run build');
  process.exit(1);
}

fs.rmSync(SALIDA, { recursive: true, force: true });
for (const [origen, destino] of ARCHIVOS) {
  const final = path.join(SALIDA, destino);
  fs.mkdirSync(path.dirname(final), { recursive: true });
  fs.copyFileSync(origen, final);
}

// La dirección de la planilla del buzón viaja en el paquete para que la
// instalación no tenga ningún paso de configuración: la app la lee sola si
// config.json no trae sheetId. Pegarla a mano era el paso más fácil de errar
// y el más difícil de diagnosticar — la app parecía andar pero los pedidos no
// le llegaban a nadie.
const SHEET_ID = '1b7-f7TRNlgGT8fODUPftn2q4ZFFkdW0r6866N8N83rA';
fs.writeFileSync(path.join(SALIDA, 'buzon-sheet.txt'), SHEET_ID + '\n');

fs.writeFileSync(path.join(SALIDA, 'LEEME.txt'), `ACTUALIZACION DE LA APP DE PICKING  -  version ${version}
=====================================================

Que trae: ahora un pedido puede cerrarse aunque falte mercaderia, indicando
por que falto y con la autorizacion de un supervisor.

Son 3 pasos. No hay que configurar nada.


PASO 1 - CERRAR LA APP
----------------------
Cerrar la ventana negra de la app (la X de arriba a la derecha).
Si queda abierta, Windows no deja reemplazar el programa.


PASO 2 - COPIAR LOS ARCHIVOS
----------------------------
NO ABRAS EL PROGRAMA DESDE ADENTRO DEL ZIP. Si haces doble click en
PickingContabilium.exe sin extraer, Windows lo copia solo a una carpeta
temporal, sin los demas archivos, y el programa arranca a medias: no
encuentra la credencial y los pedidos armados no le llegan a nadie.
Primero extraer, despues abrir.

Copiar TODO lo que esta en esta carpeta dentro de la carpeta donde ya vive
el programa, y aceptar cuando pregunte si reemplaza.

Son 4 cosas y tienen que quedar las 4 SUELTAS, al lado del programa:

    PickingContabilium.exe
    buzon-sa.json
    buzon-sheet.txt
    assets  (una carpeta)

OJO CON ESTO: cuando Windows pregunta donde extraer, propone crear una
carpeta nueva con el nombre del ZIP. Si se acepta esa propuesta, los
archivos quedan adentro de esa carpeta nueva y el programa no los
encuentra. Hay que borrar el nombre que propone y elegir la carpeta donde
ya esta el programa.

IMPORTANTE: no borrar la carpeta vieja antes de copiar. Adentro hay tres
archivos (config.json, armados.json, catalogo.json) con las contrasenas, el
historial de todos los pedidos armados y el catalogo de codigos de barra.
Esos archivos no estan en esta carpeta, o sea que copiando encima no se
tocan. Se pierden solo si alguien borra la carpeta.


PASO 3 - ABRIR Y REVISAR
------------------------
Doble click en PickingContabilium.exe y entrar con el usuario de siempre.

Dos cosas tienen que estar bien:

1. La pantalla se ve normal, con el listado de pedidos.
2. Arriba a la derecha, donde dice "Envio al facturador", dice "al dia"
   en verde.


SI ALGO NO COINCIDE
-------------------
Si la pantalla aparece en blanco o dice "No encontrado", si en vez de
"al dia" aparece un cartel rojo, o si dice que falta un archivo:
NO seguir. Avisarle a Mariano y dejar la app cerrada.
`);

const zip = path.join(DIST, `PickingContabilium-v${version}.zip`);
fs.rmSync(zip, { force: true });
// Compress-Archive de PowerShell: no hace falta instalar nada en Windows.
const aWin = p => p.replace(/^\/mnt\/([a-z])\//, (_, d) => d.toUpperCase() + ':\\').replace(/\//g, '\\');
execFileSync('powershell.exe', [
  '-NoProfile', '-Command',
  `Compress-Archive -Path '${aWin(SALIDA)}\\*' -DestinationPath '${aWin(zip)}' -Force`,
], { stdio: 'inherit' });

const mb = (fs.statSync(zip).size / 1024 / 1024).toFixed(1);
console.log(`\nPaquete listo: dist/PickingContabilium-v${version}.zip (${mb} MB)`);
console.log('Tiene una credencial adentro: va por pendrive o acceso remoto, no por mail ni WhatsApp.');
