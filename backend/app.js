require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('node:path');
const { errorHandler } = require('./middleware/errores');
const { crearRouter: rutasRegistros } = require('./routes/registros');
const { crearRouter: rutasParametros } = require('./routes/parametros');
const { crearRouter: rutasExcel } = require('./routes/exportacion');
function cargar(archivo) {
  try { return require(archivo); }
  catch (e) { if (e.code !== 'MODULE_NOT_FOUND' || !String(e.message).includes(archivo)) throw e; return null; }
}
function noDisponible(nombre, metodos) {
  return Object.fromEntries(metodos.map(m => [m, async () => {
    const err = new Error(nombre + ' todavía no está disponible'); err.status = 503; throw err;
  }]));
}
function crearApp(dependencias = {}) {
  const app = express();
  app.disable('x-powered-by');
  app.use(express.json({limit:'1mb'}));
  const origen = process.env.CORS_ORIGIN;
  if (origen) app.use(cors({origin:origen.split(',').map(s=>s.trim())}));
  else if (process.env.NODE_ENV !== 'production') app.use(cors());
  const repo = dependencias.repo || cargar('../database/repositorio.js') ||
    noDisponible('Repositorio SQLite', ['listar','obtener','crear','actualizar','eliminar']);
  const catalogos = dependencias.catalogos || cargar('./services/catalogos.js') ||
    noDisponible('Servicio de catálogos', ['obtenerParametros']);
  const exportacion = dependencias.exportacion || cargar('./services/exportacion.js') ||
    noDisponible('Servicio de exportación', ['exportarExcel']);
  const plantillaPath = dependencias.plantillaPath ||
    process.env.PLANTILLA_XLSM ||
    path.resolve(__dirname,'../plantilla/02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm');
  app.get('/api/health', (req,res) => res.json({data:{status:'ok'}}));
  app.use('/api/registros', rutasRegistros(repo));
  app.use('/api/parametros', rutasParametros(catalogos));
  app.use('/api/exportar-excel', rutasExcel(repo, exportacion, plantillaPath));
  app.use((req,res) => res.status(404).json({error:'Ruta no encontrada'}));
  app.use(errorHandler);
  return app;
}
const app = crearApp();
module.exports = { app, crearApp };
