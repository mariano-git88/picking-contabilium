# Picking Contabilium

App de picking y armado de pedidos, integrada con la API de Contabilium.
Incluye control por código de barra (EAN/DUN), manejo de combos, impresión
directa de etiquetas de despacho vía ZPL en impresoras Zebra, consulta de
pedidos armados con reporte en Excel, y sistema de usuarios.

## Requisitos
- Node.js 18 o superior
- Windows (la impresión de etiquetas usa PowerShell y APIs nativas de Windows)

## Instalación
```
npm install
```

## Uso en desarrollo
```
npm start
```
Abrí `http://localhost:3000` en el navegador.

Por defecto la app escucha **solo en la propia PC**. Para usarla desde otra
máquina de la red, agregar `"host": "0.0.0.0"` en `config.json` — teniendo en
cuenta que ahí cualquiera de la red llega a la pantalla de login.

## Empaquetado como ejecutable (.exe)

```
npm run build       # compila dist/PickingContabilium.exe
npm run verificar   # lo levanta y comprueba que sirva la pantalla
npm run paquete     # verifica y arma el ZIP para el depósito
```

**`npm run build` no alcanza como verificación, y no es un detalle.** Hasta la
v1.2.1 la config de `pkg` vivía en `pkg/pkg-package.json`, y los globs de
`assets` se resuelven **relativo a la carpeta del archivo de config**: así
`"public/**/*"` apuntaba a `pkg/public/`, que no existe. El `index.html` nunca
entró al ejecutable. El `.exe` arrancaba, imprimía su banner, abría el
navegador solo y respondía la API — pero devolvía **404 en `/`**. Encima `pkg`
puede fallar y devolver exit code 0, así que ni el build se quejaba.

Por eso la config de `pkg` ahora vive en `package.json` (donde el glob se
resuelve desde la raíz) y `npm run paquete` **no arma el ZIP si
`npm run verificar` no pasa**. La única prueba que vale es levantar el binario
y pedirle la página.

El ZIP queda en `dist/PickingContabilium-v<version>.zip` con el `.exe`,
`assets/logo_suprabond.png`, `buzon-sa.json`, `buzon-sheet.txt` y un
`LEEME.txt`. **Tiene una credencial adentro**: va por pendrive o acceso
remoto, no por mail ni por WhatsApp.

## Estructura
- `server-standalone.js` — backend (Node, sin dependencias de framework)
- `public/index.html` — frontend (una sola página)
- `products.json` — catálogo EAN/DUN por defecto (se puede actualizar desde la app)
- `assets/logo_suprabond.png` — logo usado en las etiquetas

## La instalación no pide escribir nada

Todo lo que la app necesita viaja en el paquete, al lado del `.exe`:

| Archivo | Qué es |
|---|---|
| `contabilium.json` | email + API Key de Contabilium |
| `buzon-sa.json` | Service Account del buzón |
| `buzon-sheet.txt` | dirección de la planilla del buzón |
| `usuario-inicial.json` | supervisor de fábrica, ya hasheado |

Ninguno está en git. Los genera `npm run paquete` a partir de archivos locales
(`contabilium-credenciales.json`, `supervisor-inicial.json`), y falla si
alguno falta.

**Por qué no se piden por pantalla:** ya se erró. En el depósito alguien puso
el usuario de la app (`GABI`) en el campo "Email de tu cuenta Contabilium", la
app lo guardó sin chequear, y recién fallaba al buscar órdenes con un error que
parecía de otra cosa. Un secreto que hay que copiar a mano es el paso más fácil
de errar y el más difícil de diagnosticar.

**Se arregla solo.** Al arrancar, si las credenciales guardadas **no
funcionan**, la app adopta las del paquete. Una instalación que quedó con
basura adentro se recupera con solo actualizar.

**Y si igual alguien las escribe a mano**, se prueban contra Contabilium
**antes** de guardarlas: si no andan, no se guarda nada y el mensaje dice por
qué. Guardar algo que no funciona y enterarse después es lo que hay que evitar.

## Usuarios y roles

Cada usuario es **operario** o **supervisor**. Armar, escanear e imprimir lo
hace cualquiera; **autorizar un pedido que sale incompleto y administrar
usuarios, solo un supervisor**. Lo pidió Gabriel Parodi (2026-08-10): cerrar un
pedido con menos es lo que después dispara la factura por menos.

**El paquete trae un supervisor ya creado.** Así ninguna instalación queda sin
alguien que pueda autorizar un faltante. Viaja en `usuario-inicial.json`, al
lado del `.exe`, y **ya viene hasheado** (scrypt, mismo formato que
`hashPassword()`): la contraseña en claro no está ni en el repo —que es
**público**— ni en el disco de la PC del depósito.

La contraseña en claro vive **solo en la máquina desde la que se arma el
paquete**, en `supervisor-inicial.json` (gitignoreado):

```json
{ "usuario": "GABI", "password": "...", "rol": "supervisor" }
```

`npm run paquete` lo lee, calcula el hash y falla si el archivo no está.
Para cambiar la contraseña: editar ese archivo y volver a armar el paquete.

**Solo crea lo que falta.** Si el usuario ya existe, no lo toca: una
actualización nunca le pisa la contraseña a alguien que ya la cambió. Y si la
instalación ya tenía otros usuarios, el aviso de "faltan definir los roles"
sigue apareciendo — que exista GABI no decide quién es operario.

**La app no adivina quién es supervisor.** Al actualizar desde una versión sin
roles, **todos quedan supervisores** —exactamente lo que podían hacer antes— y
`config.json` marca `rolesDefinidos: false`. La pantalla de Usuarios avisa que
faltan definir, y el aviso desaparece cuando una persona cambia un rol o crea
un usuario eligiéndolo.

La 1.3.0 intentó adivinar con "el primero de la lista es quien instaló la app".
En el depósito el primer usuario es **Jesica**, que es la que usa la app todos
los días, así que quien administra quedó de operario y **sin forma de
arreglarlo desde la app**. De un control cerrado sobre la persona equivocada no
se sale sin editar `config.json` a mano en la PC del depósito; de uno abierto y
avisado, sí. Por eso ahora falla abierto. Las instalaciones que quedaron mal
con la 1.3.0/1.3.1 se reparan solas al arrancar la 1.3.2.

No se puede quedar sin ningún supervisor: la app rechaza borrar o degradar al
último.

## Productos sin código de barra

El escaneo es la única forma de contar una línea: si el producto no tiene EAN
ni DUN en el catálogo, esa línea quedaba clavada en cero y el pedido no se
podía cerrar. Desde la v1.4 esas líneas —**y solo esas**— tienen un casillero
para anotar la cantidad a mano.

**No cuentan como faltante.** Si se anotó todo lo pedido, el armado cierra como
completo y no pide supervisor: la mercadería salió entera, lo que falta es el
código en el catálogo. Si se anotó de menos, sí es un faltante real y pide
motivo y firma como cualquier otro.

`productoSinCodigo()` valida esto **en el servidor**: anotar a mano un producto
que sí tiene código se rechaza. Si no, el casillero sería la puerta para
saltearse el escaneo escribiendo un número.

La pantalla lista arriba del pedido los SKU que no se pueden escanear, para
pasárselos a quien mantiene el catálogo, y el armado viaja al buzón con
`contadoAMano` por item.

**Cómo se actualiza el catálogo embebido:**

```
npm run catalogo    # regenera products.json desde la planilla maestra
```

No parsea la planilla por su cuenta: levanta el servidor y le postea el
`.xlsx` a `/api/catalogo`, o sea **el mismo parser que usa el botón "Subir
tabla"** del depósito. Parsear por afuera es la forma de que el catálogo
embebido y el que sube el depósito terminen distintos, y ya pasó: leyendo el
`.xlsx` con otra herramienta, dos celdas con fórmula (`=+F428`) entraron como
texto y esos dos productos quedaron con el DUN roto y sin factor. **Un DUN sin
factor no falla: escanear la caja cerrada suma 1 unidad en vez de 24, en
silencio.** Por eso el script aborta si encuentra alguno.

**Por qué el catálogo es una planilla y no la API** (medido el 1/9/2026):
Contabilium expone `CodigoBarras` en `/api/conceptos/search`, pero lo tiene
cargado en **7 de 598 productos** — los 463 EAN que usa el depósito viven solo
en la planilla. Y **no hay ningún campo para el DUN ni para el factor**, que
son los que permiten escanear una caja cerrada y sumar 24 unidades de una. O
sea que la API no puede reemplazar la planilla ni aunque alguien cargara todos
los códigos. Ver `assets/Tabla EAN DUN GSU.xlsx` en el repo de Gestión de
Vendedores.

## Pedidos incompletos

Cuando falta escanear mercadería hay dos caminos, y el orden importa:

1. **Volver a leer esta orden** — el camino normal. Los dos casos reales que
   planteó Gabriel (no había stock y se ajustó la nota de pedido; el vendedor
   agregó productos) se resuelven **editando la orden en la web de
   Contabilium** y releyéndola acá. La app no edita órdenes: por API no se
   puede, `POST /api/ordenesventa/<lo que sea>?id=` es un handler genérico que
   **cancela** la orden. Al releer, el conteo se reinicia — los renglones
   pueden haber cambiado de posición o de cantidad.
2. **El pedido sale incompleto** — la excepción. Pide un motivo por línea
   (faltante de stock · no apto para despachar · producto vencido · otro) y la
   contraseña de un supervisor. Son botones y no un campo de texto a propósito.

**Escanear de más sigue bloqueado**, sin excepción.

El motivo viaja al buzón **adentro de `items_json`**, no en una columna nueva:
el buzón es un log append-only con las dos puntas desplegadas por separado (el
depósito se actualiza a mano, la nube no), así que migrar el encabezado sería
un problema. Quién autorizó va en la columna `observacion`, que en los eventos
`armado` iba vacía.

**Un pedido incompleto no lo factura el depósito.** Llega marcado a la pantalla
de administración con el detalle y el motivo, y ahí está bloqueado a propósito:
el body de la factura se arma con las cantidades de la **orden**, no con lo
preparado, así que emitirlo desde ahí facturaría de más.

## Tutorial y novedades
La app tiene una pantalla propia de **Tutorial y novedades**, dentro de
`public/index.html` (`#pantalla-ayuda`). El tutorial está escrito para la gente
del depósito, sin tecnicismos.

Al publicar una versión nueva:

1. Subir `version` en `package.json` — el servidor la expone en
   `/api/auth/estado` y el front la usa para el punto verde de "novedades sin
   leer" (se guarda en el `localStorage` del navegador de cada persona).
2. Agregar la entrada correspondiente en la solapa **Novedades**, contando el
   cambio desde el punto de vista de quien usa la app, no del código.

## Archivos que NO se versionan (ver .gitignore)
La app genera estos archivos junto al ejecutable en tiempo de ejecución,
y contienen datos sensibles (credenciales, usuarios, datos de clientes):
`config.json`, `armados.json`, `catalogo.json`, `clientes_cache.json`.

Cada uno se escribe de forma atómica (primero un `.tmp`, después se reemplaza)
y conserva la versión anterior en un `.bak`. Si al arrancar un archivo no se
puede leer porque quedó dañado, la app **no lo pisa**: lo guarda aparte como
`.dañado-<fecha>` y recupera los datos desde el `.bak`.

## Envío al facturador (buzón)
La app corre local en el depósito y el facturador corre en la nube: no se ven
entre sí. Se comunican por una tab `armados` en una planilla de Google
dedicada, a la que esta app le escribe y de la que el facturador lee.

**Configuración (una sola vez):**

1. Copiar el archivo del Service Account como **`buzon-sa.json`**, al lado del
   ejecutable. La app avisa en pantalla si no lo encuentra.
2. Pegar la dirección de la planilla en el recuadro **Envío al facturador** y
   guardar. La pestaña `armados` y su encabezado se crean solos.

**El Service Account tiene que ser exclusivo del buzón y tener acceso a esa
planilla solamente.** Los permisos de Google son por archivo, no por pestaña:
un service account con acceso a la planilla de comisiones o a los logs de
facturación pondría todo eso al alcance de la PC del depósito.

El token de Google se firma con el módulo `crypto` de Node, sin librerías de
Google, para no engordar el `.exe`.

El envío nunca bloquea al depósito: el armado se guarda en disco primero y, si
el envío falla, queda pendiente y se reintenta al arrancar, cada 5 minutos y
con cada armado nuevo. Reenviar de más es inofensivo (el buzón es un log de
eventos y el facturador se queda con uno por orden); perder un envío no lo es,
así que la pantalla muestra cuántos quedan sin enviar.

## Respaldo
Lo que hay que respaldar es **`armados.json`** (el historial de pedidos
armados). El reporte en Excel de la pantalla "Consultar armados" sirve como
respaldo legible.

**`config.json` no va al respaldo en la nube**: adentro está la API Key de
Contabilium, que permite emitir y modificar documentos.
