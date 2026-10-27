# SENATI — Compañero 2 — Backend REST

Este directorio contiene exclusivamente el backend Node.js + Express en CommonJS.

## Instalación (Node.js >=24)

```powershell
cd backend
npm install
Copy-Item .env.example .env
npm test
npm start
```

API: http://localhost:3000/api

## Rutas
- GET /api/health
- GET /api/parametros
- GET /api/registros (filtros: expediente_siged, accion, empresa_supervisora, provincia, mes_presenta)
- GET /api/registros/:id
- POST /api/registros
- PUT /api/registros/:id
- DELETE /api/registros/:id
- GET /api/exportar-excel

JSON de éxito: `{"data": ...}`. Errores: `{"error":"..."}`. DELETE devuelve 204 sin cuerpo.
POST/PUT reciben solo los campos de `contrato-api.json`. `id` y `contrato` no son editables.
`numero_supervisores` es entero positivo; `expediente_siged` y `carta_linea` son texto.

## Integración pendiente del compañero 3
- `../database/repositorio.js` exporta métodos async: listar(filtros), obtener(id), crear(registro), actualizar(id,registro), eliminar(id).
- `./services/catalogos.js` exporta obtenerParametros() que retorna `{catalogos:{},relaciones:{}}`, usando datos reales.
- `./services/exportacion.js` exporta exportarExcel(registros,plantillaPath); el adaptador aquí espera un `Buffer` o `Uint8Array` con el XLSM terminado.
- Colocar el libro original en la ruta de plantilla documentada o establecer `PLANTILLA_XLSM`. El libro NO debe versionarse.

Si faltan servicios, las rutas afectadas responden 503, nunca datos ficticios.
No se ha verificado el Excel real ni los catálogos. No subir secretos, base de datos ni Excel.

## Pruebas
`npm test` ejecuta pruebas de rutas con repositorio simulado **solo en tests**.
Verificación manual una vez integrados servicios: `Invoke-RestMethod http://localhost:3000/api/health`.

## Git
Subir únicamente archivos propios a `feature/backend` y abrir PR contra `main` sin fusionar.
