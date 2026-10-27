'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const directorio = fs.mkdtempSync(path.join(os.tmpdir(), 'supervision-db-test-'));
process.env.DATABASE_PATH = path.join(directorio, 'pruebas.sqlite');

const repositorio = require('../repositorio');
const { inicializar } = require('../init');
const { obtenerParametros } = require('../../backend/services/catalogos');
const { exportarExcel } = require('../../backend/services/exportacion');
const plantillaReal = (() => {
  const candidates = [
    process.env.PLANTILLA_EXCEL_PATH,
    path.join(process.env.HOME || os.homedir(), '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm'),
    path.join(os.homedir(), 'Downloads', '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm'),
    path.join(__dirname, '..', '..', 'plantilla', '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm'),
  ];
  return candidates.find((candidate) => candidate && fs.existsSync(candidate)) || null;
})();

function registro(indice) {
  return {
    mes_presenta: `2024-${String((indice % 12) + 1).padStart(2, '0')}`,
    oficina_regional: `Oficina ${indice % 3}`,
    accion: `Accion ${indice % 2}`,
    accion_especifica: `Especifica ${indice}`,
    sub_accion_especifica: `Subaccion ${indice}`,
    tipo_agente: `Agente ${indice % 2}`,
    actividad: `Actividad ${indice}`,
    provincia: `Provincia ${indice % 3}`,
    distrito: `Distrito ${indice}`,
    tipo_transporte: `Transporte ${indice % 2}`,
    numero_supervisores: (indice % 4) + 1,
    empresa_supervisora: `Empresa ${indice % 2}`,
    calidad_entregable: `Calidad ${indice % 2}`,
    expediente_siged: `SIGED-00${indice}`,
    carta_linea: `CARTA-00${indice}`,
    observaciones: `Observacion ${indice}`,
  };
}

test('SQLite inicializa el esquema, mantiene datos y ejecuta el ciclo CRUD solicitado', async (t) => {
  t.after(() => {
    fs.rmSync(directorio, { recursive: true, force: true });
  });

  const db = inicializar();
  const columnas = db.prepare('PRAGMA table_info(registros)').all().map((columna) => columna.name);
  assert.deepEqual(columnas, ['id', ...Object.keys(registro(0))]);
  assert.equal(db.prepare("SELECT COUNT(*) AS cantidad FROM sqlite_master WHERE type = 'index' AND tbl_name = 'registros'").get().cantidad, 5);
  assert.throws(
    () => db.prepare('INSERT INTO registros (numero_supervisores) VALUES (0)').run(),
    /CHECK constraint failed/
  );
  db.close();

  const creados = [];
  for (let indice = 1; indice <= 15; indice++) {
    creados.push(await repositorio.crear(registro(indice)));
  }
  assert.equal(creados.length, 15);
  assert.equal(creados[0].id, 1);
  assert.equal(creados[0].expediente_siged, 'SIGED-001');
  assert.equal(typeof creados[0].carta_linea, 'string');
  assert.equal((await repositorio.listar()).length, 15);

  for (const id of [1, 2, 3]) {
    const actualizado = await repositorio.actualizar(id, {
      ...registro(id),
      observaciones: `Editado ${id}`,
      expediente_siged: `SIGED-TEXTO-${id}`,
    });
    assert.equal(actualizado.observaciones, `Editado ${id}`);
    assert.equal(actualizado.expediente_siged, `SIGED-TEXTO-${id}`);
  }

  assert.equal((await repositorio.listar({ accion: 'Accion 1' })).length, 8);
  assert.equal((await repositorio.listar({ provincia: 'Provincia 1', empresa_supervisora: 'Empresa 1' })).length, 3);
  assert.equal((await repositorio.listar({ expediente_siged: 'SIGED-TEXTO-1' })).length, 1);

  assert.equal(await repositorio.eliminar(14), true);
  assert.equal(await repositorio.eliminar(15), true);
  assert.equal(await repositorio.eliminar(15), false);
  assert.equal((await repositorio.listar()).length, 13);
  assert.equal((await repositorio.obtener(999)), null);
  await assert.rejects(repositorio.crear({ ...registro(16), numero_supervisores: 0 }), /entero positivo/);
});

test('servicios Excel leen el libro real cuando existe y bloquean faltantes', async () => {
  const sinLibro = path.join(directorio, 'libro-ausente.xlsm');
  if (plantillaReal) {
    process.env.PLANTILLA_EXCEL_PATH = plantillaReal;
    const params = await obtenerParametros();
    assert.ok(params && params.catalogos && params.relaciones);
    assert.ok(Array.isArray(params.catalogos.accion));
    assert.ok(Array.isArray(params.catalogos.empresa_supervisora));
    await assert.rejects(exportarExcel([], plantillaReal), (error) =>
      error.statusCode === 503 && /No se pudo verificar y exportar|no se encontró/.test(error.message)
    );
  } else {
    await assert.rejects(obtenerParametros(), (error) => error.statusCode === 503 && /no se encontró el libro fuente/.test(error.message));
    await assert.rejects(exportarExcel([], sinLibro), (error) =>
      error.statusCode === 503 && /no se encontró el libro fuente/.test(error.message)
    );
  }
});
