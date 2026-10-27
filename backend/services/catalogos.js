'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const NOMBRE_PLANTILLA = '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm';

function crearErrorServicio(mensaje, causa) {
  const error = new Error(mensaje, causa ? { cause: causa } : undefined);
  error.status = 503;
  error.statusCode = 503;
  return error;
}

function resolverPlantillaPath(plantillaPath) {
  const candidates = [];
  if (plantillaPath) candidates.push(plantillaPath);
  if (process.env.PLANTILLA_EXCEL_PATH) candidates.push(process.env.PLANTILLA_EXCEL_PATH);
  candidates.push(path.join(__dirname, '..', '..', 'plantilla', NOMBRE_PLANTILLA));
  candidates.push(path.join(os.homedir(), NOMBRE_PLANTILLA));
  candidates.push(path.join(os.homedir(), 'Downloads', NOMBRE_PLANTILLA));

  const seen = new Set();
  for (const candidate of candidates) {
    if (!candidate) continue;
    const resolved = path.resolve(candidate);
    if (!seen.has(resolved)) {
      seen.add(resolved);
      if (fs.existsSync(resolved)) {
        return resolved;
      }
    }
  }

  return path.resolve(candidates[0] || path.join(__dirname, '..', '..', 'plantilla', NOMBRE_PLANTILLA));
}

function ejecutarPython(script, env) {
  const resultado = spawnSync('python', ['-c', script], {
    encoding: 'utf8',
    env: { ...process.env, ...env },
    windowsHide: true,
  });

  if (resultado.status !== 0) {
    const detalle = (resultado.stderr || resultado.stdout || '').trim();
    throw new Error(detalle || 'La inspección del Excel falló.');
  }

  return JSON.parse((resultado.stdout || '').trim() || '{}');
}

async function obtenerParametros() {
  const plantillaPath = resolverPlantillaPath();
  if (!fs.existsSync(plantillaPath)) {
    throw crearErrorServicio(`Catálogos no disponibles: no se encontró el libro fuente "${NOMBRE_PLANTILLA}".`);
  }

  const script = `
import json, os
from openpyxl import load_workbook
path = os.environ['PLANTILLA_EXCEL_PATH']
wb = load_workbook(path, read_only=True, data_only=True)

def clean(value):
    if value is None:
        return None
    text = str(value).strip()
    # El libro original contiene algunas cadenas UTF-8 interpretadas como Latin-1.
    if 'Ã' in text or 'Â' in text or '�' in text:
        try:
            reparado = text.encode('latin-1').decode('utf-8')
            if reparado.count('Ã') + reparado.count('�') < text.count('Ã') + text.count('�'):
                text = reparado
        except (UnicodeError, ValueError):
            pass
    return text if text else None

accion = []
for row in wb['Accion'].iter_rows(min_row=2, values_only=True):
    if row is None or all(v is None for v in row):
        continue
    item = {
        'nombre_abreviado': clean(row[0]),
        'accion': clean(row[1]),
        'codigo_accion': clean(row[2]),
        'codigo_plan_accion': row[3],
    }
    if item['accion'] or item['codigo_accion'] or item['nombre_abreviado']:
        accion.append({k: v for k, v in item.items() if v is not None})

accion_especifica = []
for row in wb['AccionEspecifica'].iter_rows(min_row=2, values_only=True):
    if row is None or all(v is None for v in row):
        continue
    item = {
        'nombre_abreviado': clean(row[0]),
        'accion_especifica': clean(row[1]),
        'codigo_accion_especifica': clean(row[2]),
        'codigo_accion': clean(row[3]),
    }
    if item['accion_especifica'] or item['codigo_accion_especifica']:
        accion_especifica.append({k: v for k, v in item.items() if v is not None})

sub_accion = []
for row in wb['SubAccion'].iter_rows(min_row=2, values_only=True):
    if row is None or all(v is None for v in row):
        continue
    item = {
        'nombre_abreviado': clean(row[0]),
        'sub_accion_especifica': clean(row[1]),
        'concat': clean(row[2]),
        'codigo_sub_accion_especifica': clean(row[3]),
        'codigo_accion_especifica': clean(row[4]),
    }
    if item['sub_accion_especifica'] or item['codigo_sub_accion_especifica']:
        sub_accion.append({k: v for k, v in item.items() if v is not None})

tipo_agente = []
actividades = []
for row in wb['ActividadAgente'].iter_rows(min_row=2, values_only=True):
    if row is None or all(v is None for v in row):
        continue
    codigo = clean(row[0]) or clean(row[5])
    nombre = clean(row[1]) or clean(row[6])
    if codigo or nombre:
        tipo_agente.append({'codigo': codigo, 'nombre': nombre})
    actividad = {
        'codigo_tipo_agente': codigo,
        'tipo_agente': nombre,
        'codigo_actividad': clean(row[2]),
        'actividad': clean(row[3]),
    }
    if actividad['codigo_actividad'] or actividad['actividad']:
        actividades.append({k: v for k, v in actividad.items() if v is not None})

empresas = []
for row in wb['Lista de Empresas Supervisoras'].iter_rows(min_row=7, values_only=True):
    if row is None or all(v is None for v in row):
        continue
    item = {
        'oficina_regional': clean(row[1]),
        'contrato': clean(row[2]),
        'razon_social': clean(row[3]),
        'sub_sector': clean(row[4]),
        'lista_empresas': clean(row[5]),
    }
    if item['lista_empresas'] or item['contrato'] or item['razon_social']:
        empresas.append({k: v for k, v in item.items() if v is not None})

relaciones = {
    'accion_especifica_por_accion': {},
    'sub_accion_por_accion_especifica': {},
    'actividad_por_tipo_agente': {},
    'empresas_supervisora_por_oficina_regional': {},
}
for item in accion_especifica:
    relaciones['accion_especifica_por_accion'][item.get('codigo_accion')] = relaciones['accion_especifica_por_accion'].get(item.get('codigo_accion'), []) + [item.get('codigo_accion_especifica')]
for item in sub_accion:
    relaciones['sub_accion_por_accion_especifica'][item.get('codigo_accion_especifica')] = relaciones['sub_accion_por_accion_especifica'].get(item.get('codigo_accion_especifica'), []) + [item.get('codigo_sub_accion_especifica')]
for item in actividades:
    codigo_ta = item.get('codigo_tipo_agente')
    if codigo_ta:
        relaciones['actividad_por_tipo_agente'].setdefault(codigo_ta, []).append(item.get('codigo_actividad'))
for item in empresas:
    oficina = item.get('oficina_regional')
    if oficina:
        relaciones['empresas_supervisora_por_oficina_regional'].setdefault(oficina, []).append(item.get('lista_empresas'))

# Catálogos complementarios tomados del libro fuente, nunca de datos inventados.
registro = wb['Registro']
def valores_registro(col):
    return sorted({v for fila in registro.iter_rows(min_row=22, max_col=col, values_only=True)
                   if (v := clean(fila[col-1])) and not v.startswith('#')})
meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
mes_actual = clean(registro['A3'].value)
if mes_actual and mes_actual not in meses: meses.append(mes_actual)
oficinas = sorted({v for item in empresas if (v := item.get('oficina_regional'))} | ({clean(registro['A6'].value)} if clean(registro['A6'].value) else set()))
# La lista de empresas tiene celdas de oficina combinadas: heredar la oficina previa.
oficina_anterior = None
for item in empresas:
    if item.get('oficina_regional'): oficina_anterior = item['oficina_regional']
    else: item['oficina_regional'] = oficina_anterior
oficinas = sorted({v for item in empresas if (v := item.get('oficina_regional'))})
relaciones['empresas_supervisora_por_oficina_regional'] = {}
for item in empresas:
    oficina = item.get('oficina_regional')
    if oficina and item.get('lista_empresas'):
        relaciones['empresas_supervisora_por_oficina_regional'].setdefault(oficina, []).append(item['lista_empresas'])
# Evitar entradas duplicadas de tipos de agente.
tipo_agente = list({item['codigo']: item for item in tipo_agente if item.get('codigo')}.values())
resultado = {
    'catalogos': {
        'mes_presenta': meses,
        'oficina_regional': oficinas,
        'provincia': valores_registro(6),
        'distrito': valores_registro(7),
        'tipo_transporte': valores_registro(8),
        'calidad_entregable': valores_registro(11),
        'accion': accion,
        'accion_especifica': accion_especifica,
        'sub_accion_especifica': sub_accion,
        'tipo_agente': tipo_agente,
        'actividad': actividades,
        'empresa_supervisora': empresas,
    },
    'relaciones': relaciones,
}
print(json.dumps(resultado, ensure_ascii=False))
`;

  try {
    return ejecutarPython(script, { PLANTILLA_EXCEL_PATH: plantillaPath });
  } catch (error) {
    throw crearErrorServicio('Catálogos no disponibles: no se pudo leer el libro fuente.', error);
  }
}

module.exports = { obtenerParametros };
