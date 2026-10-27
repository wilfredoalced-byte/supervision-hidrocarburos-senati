const express = require('express');
const { validarId, validarRegistro, validarFiltros } = require('../controllers/registrosController');
function crearRouter(repo) {
  const r = express.Router();
  const wrap = fn => (req,res,next) => Promise.resolve().then(() => fn(req,res)).catch(next);
  r.get('/', wrap(async (req,res) => res.json({data: await repo.listar(validarFiltros(req.query))})));
  r.get('/:id', wrap(async (req,res) => {
    const registro = await repo.obtener(validarId(req.params.id));
    if (!registro) return res.status(404).json({error:'Registro no encontrado'});
    return res.json({data:registro});
  }));
  r.post('/', wrap(async (req,res) => {
    const data = await repo.crear(validarRegistro(req.body));
    return res.status(201).json({data});
  }));
  r.put('/:id', wrap(async (req,res) => {
    const id = validarId(req.params.id);
    if (!await repo.obtener(id)) return res.status(404).json({error:'Registro no encontrado'});
    const data = await repo.actualizar(id,validarRegistro(req.body));
    if (!data) return res.status(404).json({error:'Registro no encontrado'});
    return res.json({data});
  }));
  r.delete('/:id', wrap(async (req,res) => {
    const id = validarId(req.params.id);
    if (!await repo.obtener(id)) return res.status(404).json({error:'Registro no encontrado'});
    await repo.eliminar(id);
    return res.status(204).end();
  }));
  return r;
}
module.exports = { crearRouter };
