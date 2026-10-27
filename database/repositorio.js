'use strict';

const { inicializar } = require('./init');

const CAMPOS = [
  'mes_presenta',
  'oficina_regional',
  'accion',
  'accion_especifica',
  'sub_accion_especifica',
  'tipo_agente',
  'actividad',
  'provincia',
  'distrito',
  'tipo_transporte',
  'numero_supervisores',
  'empresa_supervisora',
  'calidad_entregable',
  'expediente_siged',
  'carta_linea',
  'observaciones',
];

const FILTROS = [
  'expediente_siged',
  'accion',
  'empresa_supervisora',
  'provincia',
  'mes_presenta',
];

function validarRegistro(registro) {
  if (!registro || typeof registro !== 'object' || Array.isArray(registro)) {
    throw new TypeError('El registro debe ser un objeto.');
  }

  const camposDesconocidos = Object.keys(registro).filter((campo) => !CAMPOS.includes(campo));
  if (camposDesconocidos.length) {
    throw new TypeError(`Campo(s) no permitido(s): ${camposDesconocidos.join(', ')}`);
  }

  if (Object.keys(registro).length === 0) {
    throw new TypeError('Debe proporcionarse al menos un campo del registro.');
  }

  for (const [campo, valor] of Object.entries(registro)) {
    if (campo === 'numero_supervisores') {
      if (valor !== null && (!Number.isInteger(valor) || valor <= 0)) {
        throw new TypeError('numero_supervisores debe ser un entero positivo o null.');
      }
    } else if (valor !== null && typeof valor !== 'string') {
      throw new TypeError(`${campo} debe ser texto o null.`);
    }
  }
}

function withDatabase(callback) {
  const database = inicializar();
  try {
    return callback(database);
  } finally {
    database.close();
  }
}

async function listar(filtros = {}) {
  if (!filtros || typeof filtros !== 'object' || Array.isArray(filtros)) {
    throw new TypeError('Los filtros deben ser un objeto.');
  }

  const condiciones = [];
  const valores = [];
  for (const campo of FILTROS) {
    const valor = filtros[campo];
    if (valor !== undefined && valor !== null && valor !== '') {
      if (typeof valor !== 'string') {
        throw new TypeError(`${campo} debe ser texto.`);
      }
      condiciones.push(`${campo} = ?`);
      valores.push(valor);
    }
  }

  const sql = `SELECT id, ${CAMPOS.join(', ')} FROM registros${
    condiciones.length ? ` WHERE ${condiciones.join(' AND ')}` : ''
  } ORDER BY id`;

  return withDatabase((database) => database.prepare(sql).all(...valores));
}

async function obtener(id) {
  if (!Number.isInteger(id) || id <= 0) {
    return null;
  }

  return withDatabase((database) =>
    database.prepare(`SELECT id, ${CAMPOS.join(', ')} FROM registros WHERE id = ?`).get(id) || null
  );
}

async function crear(registro) {
  validarRegistro(registro);
  const campos = Object.keys(registro);
  const placeholders = campos.map(() => '?').join(', ');
  const sql = `INSERT INTO registros (${campos.join(', ')}) VALUES (${placeholders}) RETURNING id, ${CAMPOS.join(', ')}`;
  const valores = campos.map((campo) => registro[campo]);
  return withDatabase((database) => database.prepare(sql).get(...valores));
}

async function actualizar(id, registro) {
  if (!Number.isInteger(id) || id <= 0) {
    return null;
  }
  validarRegistro(registro);

  const campos = Object.keys(registro);
  const asignaciones = campos.map((campo) => `${campo} = ?`).join(', ');
  const sql = `UPDATE registros SET ${asignaciones} WHERE id = ? RETURNING id, ${CAMPOS.join(', ')}`;
  const valores = [...campos.map((campo) => registro[campo]), id];
  return withDatabase((database) => database.prepare(sql).get(...valores) || null);
}

async function eliminar(id) {
  if (!Number.isInteger(id) || id <= 0) {
    return false;
  }

  return withDatabase((database) =>
    database.prepare('DELETE FROM registros WHERE id = ?').run(id).changes > 0
  );
}

module.exports = { listar, obtener, crear, actualizar, eliminar };
