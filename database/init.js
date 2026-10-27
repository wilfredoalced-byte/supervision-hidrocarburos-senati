'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { DatabaseSync } = require('node:sqlite');

const DEFAULT_DATABASE_PATH = path.join(__dirname, 'data.sqlite');
const SCHEMA_PATH = path.join(__dirname, 'schema.sql');

function getDatabasePath(databasePath = process.env.DATABASE_PATH) {
  return databasePath || DEFAULT_DATABASE_PATH;
}

function inicializar(databasePath) {
  const resolvedPath = getDatabasePath(databasePath);
  if (resolvedPath !== ':memory:') {
    fs.mkdirSync(path.dirname(path.resolve(resolvedPath)), { recursive: true });
  }

  const database = new DatabaseSync(resolvedPath);
  try {
    database.exec(fs.readFileSync(SCHEMA_PATH, 'utf8'));
    return database;
  } catch (error) {
    database.close();
    throw error;
  }
}

module.exports = { inicializar, getDatabasePath };
