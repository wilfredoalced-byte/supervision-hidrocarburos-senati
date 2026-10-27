const express = require('express');
const fs = require('node:fs');
const path = require('node:path');
function crearRouter(repo, servicio, plantillaPath) {
  const router = express.Router();
  router.get('/', async (req,res,next) => {
    try {
      if (!fs.existsSync(plantillaPath)) return res.status(503).json({error:'Plantilla Excel original no disponible'});
      const registros = await repo.listar({});
      const resultado = await servicio.exportarExcel(registros, plantillaPath);
      if (!Buffer.isBuffer(resultado) && !(resultado instanceof Uint8Array)) {
        return res.status(503).json({error:'El servicio de exportación debe devolver Buffer o Uint8Array'});
      }
      return res.status(200)
        .set('Content-Type','application/vnd.ms-excel.sheet.macroEnabled.12')
        .set('Content-Disposition','attachment; filename="supervision-hidrocarburos.xlsm"')
        .send(Buffer.from(resultado));
    } catch (error) { return next(error); }
  });
  return router;
}
module.exports = { crearRouter };
