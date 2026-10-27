# Persistencia y datos de Excel

## SQLite

- Requiere Node.js 22.13 o posterior por el módulo integrado `node:sqlite` (sin dependencia npm adicional).
- Inicialización: `node -e "require('./database/init').inicializar().close()"`.
- La ruta predeterminada es `database/data.sqlite`; puede cambiarse con `DATABASE_PATH`.
- El archivo de base de datos no se versiona. `database/schema.sql` declara el identificador entero autogenerado, los 16 campos del contrato, el control positivo de `numero_supervisores` y los índices de filtros.
- `database/repositorio.js` exporta `listar(filtros)`, `obtener(id)`, `crear(registro)`, `actualizar(id, registro)` y `eliminar(id)`, todas como funciones async. Devuelve arrays, registro o `null`, registro creado/actualizado o `null`, y booleano respectivamente. `crear` y `actualizar` aceptan los campos parciales del contrato y texto `null`, como la API; un objeto vacío no es válido. `contrato` no es una columna: es un campo calculado fuera de esta persistencia. Los valores de usuario se enlazan mediante parámetros SQLite.
- `expediente_siged` y `carta_linea` permanecen como texto; `numero_supervisores`, cuando se proporciona, debe ser entero positivo. Las consultas filtradas admiten `expediente_siged`, `accion`, `empresa_supervisora`, `provincia` y `mes_presenta`.

## Libro fuente, catálogos y exportación

Se inspeccionó el libro real en la ruta del usuario: `C:\Users\BORUTO\02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm`. La evidencia comprobada incluye estas hojas: `Registro`, `Parametros`, `Accion`, `AccionEspecifica`, `SubAccion`, `ActividadAgente`, `POA`, `Lista de Empresas Supervisoras` y `Base`.

`Registro` contiene la tabla de captura con encabezados reales y el campo calculado `CONTRATO` (fila 21). Los encabezados observados son: `Acción`, `Acción Especifica`, `Sub Acción Específica`, `Tipo de Agente`, `Actividad`, `Zona (Provincia)`, `Zona (Distrito)`, `Tipo de Transporte Empleado`, `N° de Supervisores participantes de la Sup.`, `Empresa Supervisora`, `Calidad del entregable (informe mensual)`, `N°  de Expediente SIGED Involucrado`, `Carta Linea`, `Observaciones (Para futuras Mejoras)` y `CONTRATO`.

`Accion`, `AccionEspecifica`, `SubAccion` y `ActividadAgente` son las fuentes reales de los catálogos y relaciones; `Lista de Empresas Supervisoras` aporta la lista maestra con `OFICINA REGIONAL`, `CONTRATO`, `RAZON SOCIAL`, `SUB SECTOR` y `LISTA DE EMPRESAS`.

`backend/services/catalogos.js` lee el libro real con `openpyxl` y devuelve `{ catalogos, relaciones }` con las claves verificadas del libro. Los errores de disponibilidad de ambos servicios incluyen `status = 503` para el middleware de la API. `backend/services/exportacion.js` usa la misma resolución de ruta y garantiza que, si hay un libro real, no modifique el original; copia a una salida temporal y valida la presencia de `xl/vbaProject.bin` y la estructura de `Registro` antes de guardar. La función devuelve un `Buffer` XLSM; la ruta HTTP establece `Content-Type` y `Content-Disposition` y envía el binario. La validación admite campos de texto nulos de los registros SQLite.

## Pruebas

Ejecutar `node --test database/test/*.test.js`. Las pruebas usan una base SQLite temporal, insertan 15 registros sintéticos solo en esa base, editan 3, eliminan 2, consultan filtros y comprueban los tipos de texto, campos parciales y la restricción del entero positivo. Si el libro real está disponible, también leen catálogos y exportan un registro sintético a una copia XLSM; verifican la respuesta binaria y que el hash del libro fuente no cambie. Esta prueba de exportación requiere Windows con Excel COM y `openpyxl` disponible para la lectura de catálogos. Sin el libro, se comprueban las respuestas 503; no se usan datos simulados en producción.
