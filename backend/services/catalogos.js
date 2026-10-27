'use strict';

const fs = require('node:fs');
const path = require('node:path');

const NOMBRE_PLANTILLA = '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm';

function crearErrorServicio(mensaje, causa) {
  const error = new Error(mensaje, causa ? { cause } : undefined);
  error.statusCode = 503;
  return error;
}

async function obtenerParametros() {
  const plantillaPath = process.env.PLANTILLA_EXCEL_PATH || path.join(__dirname, '..', '..', 'plantilla', NOMBRE_PLANTILLA);
  if (!fs.existsSync(plantillaPath)) {
    throw crearErrorServicio(`Catálogos no disponibles: no se encontró el libro fuente "${NOMBRE_PLANTILLA}".`);
  }

  throw crearErrorServicio(
    'Catálogos no disponibles: el libro fuente requiere inspección para identificar sus hojas, rangos y relaciones sin asumir datos.'
  );
}

module.exports = { obtenerParametros };
