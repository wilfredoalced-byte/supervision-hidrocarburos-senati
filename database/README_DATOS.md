# Persistencia y datos de Excel

## SQLite

- Requiere Node.js 22.13 o posterior por el módulo integrado `node:sqlite` (sin dependencia npm adicional).
- Inicialización: `node -e "require('./database/init').inicializar().close()"`.
- La ruta predeterminada es `database/data.sqlite`; puede cambiarse con `DATABASE_PATH`.
- El archivo de base de datos no se versiona. `database/schema.sql` declara el identificador entero autogenerado, los 16 campos del contrato, el control positivo de `numero_supervisores` y los índices de filtros.
- `database/repositorio.js` exporta `listar(filtros)`, `obtener(id)`, `crear(registro)`, `actualizar(id, registro)` y `eliminar(id)`, todas como funciones async. Devuelve arrays, registro o `null`, registro creado/actualizado o `null`, y booleano respectivamente. `contrato` no es una columna: es un campo calculado fuera de esta persistencia. Los valores de usuario se enlazan mediante parámetros SQLite.
- `expediente_siged` y `carta_linea` permanecen como texto; `numero_supervisores` debe ser entero positivo. Las consultas filtradas admiten `expediente_siged`, `accion`, `empresa_supervisora`, `provincia` y `mes_presenta`.

## Libro fuente, catálogos y exportación

El libro `02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm` no estaba presente en `plantilla/` ni en otra ruta del workspace durante esta implementación. No se inspeccionaron hojas, encabezados, fórmulas, rangos con nombre, validaciones, relaciones, estilos o macros; no se afirma que se hayan verificado.

`backend/services/catalogos.js` responde con error `statusCode: 503` cuando falta el libro. Incluso si se configura `PLANTILLA_EXCEL_PATH`, el servicio se mantiene no disponible hasta inspeccionar y documentar el contenido real, sin responder con catálogos vacíos o simulados. Al recibir el libro se deben documentar las claves y relaciones verificadas antes de habilitar `{ catalogos, relaciones }`.

`backend/services/exportacion.js` declara `exportarExcel(registros, plantillaPath)`. Si se habilita con un libro real, retorna `{ path, mimeType, filename }` para que backend entregue el archivo binario con `Content-Disposition`. En Windows utiliza Excel COM sobre una copia temporal, abre la hoja `Registro`, exige encabezados exactos del contrato y comprueba la presencia de `xl/vbaProject.bin` antes y después de guardar. Una estructura incompatible o una pérdida de macros produce error `503`; nunca modifica el original. No se ejecutó esta exportación porque falta la plantilla, por lo que la preservación de hojas, fórmulas, estilos y validaciones permanece sin verificar.

## Pruebas

Ejecutar `node --test database/test/*.test.js`. La prueba usa una base SQLite temporal, inserta 15 registros sintéticos solo en esa base, edita 3, elimina 2, consulta filtros, comprueba los tipos de texto y la restricción del entero positivo, e inspecciona el 503 de los servicios bloqueados. No son datos reales ni prueban el libro Excel.
