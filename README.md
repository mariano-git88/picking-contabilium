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
Este proyecto usa `pkg` para generar un ejecutable standalone para Windows:
```
npm run build
```

## Estructura
- `server-standalone.js` — backend (Node, sin dependencias de framework)
- `public/index.html` — frontend (una sola página)
- `products.json` — catálogo EAN/DUN por defecto (se puede actualizar desde la app)
- `assets/logo_suprabond.png` — logo usado en las etiquetas

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

## Respaldo
Lo que hay que respaldar es **`armados.json`** (el historial de pedidos
armados). El reporte en Excel de la pantalla "Consultar armados" sirve como
respaldo legible.

**`config.json` no va al respaldo en la nube**: adentro está la API Key de
Contabilium, que permite emitir y modificar documentos.
