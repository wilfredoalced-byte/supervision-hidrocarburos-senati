# Persistencia y datos de Excel

## SQLite

- Requiere Node.js 22.13 o posterior por el módulo integrado `node:sqlite` (sin dependencia npm adicional).
- Inicialización: `node -e "require('./database/init').inicializar().close()"`.
- La ruta predeterminada es `database/data.sqlite`; puede cambiarse con `DATABASE_PATH`.
- El archivo de base de datos no se versiona. `database/schema.sql` declara el identificador entero autogenerado, los 16 campos del contrato, el control positivo de `numero_supervisores` y los índices de filtros.
- `database/repositorio.js` exporta `listar(filtros)`, `obtener(id)`, `crear(registro)`, `actualizar(id, registro)` y `eliminar(id)`, todas como funciones async. Devuelve arrays, registro o `null`, registro creado/actualizado o `null`, y booleano respectivamente. `contrato` no es una columna: es un campo calculado fuera de esta persistencia. Los valores de usuario se enlazan mediante parámetros SQLite.
- `expediente_siged` y `carta_linea` permanecen como texto; `numero_supervisores` debe ser entero positivo. Las consultas filtradas admiten `expediente_siged`, `accion`, `empresa_supervisora`, `provincia` y `mes_presenta`.

## Libro fuente, catálogos y exportación

Se inspeccionó el libro real en la ruta del usuario: `C:\Users\BORUTO\02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm`. La evidencia comprobada incluye estas hojas: `Registro`, `Parametros`, `Accion`, `AccionEspecifica`, `SubAccion`, `ActividadAgente`, `POA`, `Lista de Empresas Supervisoras` y `Base`.

`Registro` contiene la tabla de captura con encabezados reales y el campo calculado `CONTRATO` (fila 21). Los encabezados observados son: `Acción`, `Acción Especifica`, `Sub Acción Específica`, `Tipo de Agente`, `Actividad`, `Zona (Provincia)`, `Zona (Distrito)`, `Tipo de Transporte Empleado`, `N° de Supervisores participantes de la Sup.`, `Empresa Supervisora`, `Calidad del entregable (informe mensual)`, `N°  de Expediente SIGED Involucrado`, `Carta Linea`, `Observaciones (Para futuras Mejoras)` y `CONTRATO`.

`Accion`, `AccionEspecifica`, `SubAccion` y `ActividadAgente` son las fuentes reales de los catálogos y relaciones; `Lista de Empresas Supervisoras` aporta la lista maestra con `OFICINA REGIONAL`, `CONTRATO`, `RAZON SOCIAL`, `SUB SECTOR` y `LISTA DE EMPRESAS`.

`backend/services/catalogos.js` lee el libro real con `openpyxl` y devuelve `{ catalogos, relaciones }` con las claves verificadas del libro. `backend/services/exportacion.js` usa la misma resolución de ruta y garantiza que, si hay un libro real, no modifique el original; copia a una salida temporal y valida la presencia de `xl/vbaProject.bin` y la estructura de `Registro` antes de guardar.

## Pruebas

Ejecutar `node --test database/test/*.test.js`. La prueba usa una base SQLite temporal, inserta 15 registros sintéticos solo en esa base, edita 3, elimina 2, consulta filtros, comprueba los tipos de texto y la restricción del entero positivo, e inspecciona el 503 de los servicios bloqueados. No son datos reales ni prueban el libro Excel.
