const express = require('express');
function crearRouter(catalogos) {
  const router = express.Router();
  router.get('/', async (req,res,next) => {
    try {
      const data = await catalogos.obtenerParametros();
      if (!data || typeof data !== 'object' || !data.catalogos || !data.relaciones) {
        return res.status(503).json({error:'Los catálogos reales todavía no están disponibles'});
      }
      return res.json({data:{catalogos:data.catalogos, relaciones:data.relaciones}});
    } catch (error) { return next(error); }
  });
  return router;
}
module.exports = { crearRouter };
