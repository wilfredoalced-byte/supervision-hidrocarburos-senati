const CAMPOS = ['mes_presenta','oficina_regional','accion','accion_especifica','sub_accion_especifica','tipo_agente','actividad','provincia','distrito','tipo_transporte','numero_supervisores','empresa_supervisora','calidad_entregable','expediente_siged','carta_linea','observaciones'];
const FILTROS = ['expediente_siged','accion','empresa_supervisora','provincia','mes_presenta'];
function validarId(raw) {
  if (!/^[1-9]\\d*$/.test(String(raw)) || !Number.isSafeInteger(Number(raw))) {
    const error = new Error('id inválido'); error.status = 400; throw error;
  }
  return Number(raw);
}
function validarRegistro(body) {
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    const error = new Error('Se requiere un objeto JSON'); error.status = 400; throw error;
  }
  const desconocidos = Object.keys(body).filter(k => !CAMPOS.includes(k));
  if (desconocidos.length) {
    const error = new Error('Campos no permitidos: ' + desconocidos.join(', ')); error.status = 400; throw error;
  }
  for (const [campo, valor] of Object.entries(body)) {
    if (campo === 'numero_supervisores') {
      if (!Number.isSafeInteger(valor) || valor <= 0) {
        const error = new Error('numero_supervisores debe ser entero positivo'); error.status = 400; throw error;
      }
    } else if (valor !== null && typeof valor !== 'string') {
      const error = new Error(campo + ' debe ser texto o null'); error.status = 400; throw error;
    }
  }
  return Object.fromEntries(Object.entries(body).filter(([k]) => CAMPOS.includes(k)));
}
function validarFiltros(query) {
  const invalidos = Object.keys(query).filter(k => !FILTROS.includes(k));
  if (invalidos.length) {
    const error = new Error('Filtros no permitidos: ' + invalidos.join(', ')); error.status = 400; throw error;
  }
  for (const value of Object.values(query)) {
    if (typeof value !== 'string') {
      const error = new Error('Filtros deben ser texto'); error.status = 400; throw error;
    }
  }
  return Object.fromEntries(FILTROS.filter(k => query[k] !== undefined).map(k => [k, query[k]]));
}
module.exports = { CAMPOS, FILTROS, validarId, validarRegistro, validarFiltros };
