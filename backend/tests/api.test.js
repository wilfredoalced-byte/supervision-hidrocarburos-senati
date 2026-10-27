const { test } = require('node:test');
const assert = require('node:assert/strict');
const request = require('supertest');
const { crearApp } = require('../app');
function preparar() {
  const data = new Map(); let nextId = 1;
  const repo = {
    async listar(filtros) { return [...data.values()].filter(item => Object.entries(filtros).every(([k,v])=>item[k]===v)); },
    async obtener(id) { return data.get(id) || null; },
    async crear(r) { const item={id:nextId++,...r};data.set(item.id,item);return item; },
    async actualizar(id,r) {const item={...data.get(id),...r};data.set(id,item);return item;},
    async eliminar(id) {data.delete(id);}
  };
  return crearApp({repo,catalogos:{async obtenerParametros(){return {catalogos:{},relaciones:{}};}},exportacion:{async exportarExcel(){return Buffer.from('test');}},plantillaPath:__filename});
}
test('health responde ok', async () => {
  const res=await request(preparar()).get('/api/health');
  assert.equal(res.status,200);assert.equal(res.body.data.status,'ok');
});
test('CRUD, filtros y respuesta 204', async () => {
  const app=preparar();
  const c=await request(app).post('/api/registros').send({expediente_siged:'000123',numero_supervisores:2,provincia:'Lima'});
  assert.equal(c.status,201);assert.equal(c.body.data.id,1);
  assert.equal((await request(app).get('/api/registros?provincia=Lima')).body.data.length,1);
  assert.equal((await request(app).get('/api/registros?provincia=Tacna')).body.data.length,0);
  assert.equal((await request(app).get('/api/registros/1')).body.data.expediente_siged,'000123');
  assert.equal((await request(app).put('/api/registros/1').send({carta_linea:'001'})).body.data.carta_linea,'001');
  assert.equal((await request(app).delete('/api/registros/1')).status,204);
  assert.equal((await request(app).get('/api/registros/1')).status,404);
});
test('validación bloquea id, contrato y tipos incorrectos', async () => {
  const app=preparar();
  assert.equal((await request(app).post('/api/registros').send({id:1})).status,400);
  assert.equal((await request(app).post('/api/registros').send({contrato:'x'})).status,400);
  assert.equal((await request(app).post('/api/registros').send({numero_supervisores:0})).status,400);
  assert.equal((await request(app).post('/api/registros').send({expediente_siged:123})).status,400);
  assert.equal((await request(app).get('/api/registros/nope')).status,400);
  assert.equal((await request(app).get('/api/registros?foo=x')).status,400);
});
test('parámetros devuelve envoltorio estable', async () => {
  const res=await request(preparar()).get('/api/parametros');
  assert.equal(res.status,200);assert.deepEqual(res.body.data,{catalogos:{},relaciones:{}});
});
test('dependencias no disponibles devuelven 503', async () => {
  const app=crearApp({repo:{...Object.fromEntries(['listar','obtener','crear','actualizar','eliminar'].map(k=>[k,async()=>{throw Object.assign(new Error('No disponible'),{status:503});}]))},catalogos:{async obtenerParametros(){throw Object.assign(new Error('No disponible'),{status:503});}},exportacion:{async exportarExcel(){return Buffer.from('x');}},plantillaPath:'/archivo/que/no/existe.xlsm'});
  assert.equal((await request(app).get('/api/registros')).status,503);
  assert.equal((await request(app).get('/api/parametros')).status,503);
  assert.equal((await request(app).get('/api/exportar-excel')).status,503);
});
test('exportación envía attachment con bytes', async () => {
  const res=await request(preparar()).get('/api/exportar-excel');
  assert.equal(res.status,200);assert.match(res.headers['content-disposition'],/attachment/);
  assert.match(res.headers['content-type'],/macroenabled/);
});
