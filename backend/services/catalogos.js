'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const NOMBRE_PLANTILLA = '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm';

let cache = null;

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
    env: { ...process.env, ...env, PYTHONIOENCODING: 'utf-8' },
    maxBuffer: 64 * 1024 * 1024,
    windowsHide: true,
  });

  if (resultado.error) {
    throw resultado.error;
  }
  if (resultado.status !== 0) {
    const detalle = (resultado.stderr || resultado.stdout || '').trim();
    throw new Error(detalle || 'La inspección del Excel falló.');
  }

  return JSON.parse((resultado.stdout || '').trim() || '{}');
}

// Cada opción incluye `valor` y `etiqueta` (texto tal como lo guarda el libro en la hoja Registro)
// más los campos originales de la hoja de la que proviene. Las listas dependientes se
// resuelven con `relaciones`, indexadas por el `valor` del campo padre.
const SCRIPT = String.raw`
import json, os, unicodedata
from openpyxl import load_workbook

wb = load_workbook(os.environ['PLANTILLA_EXCEL_PATH'], read_only=True, data_only=True)


def clean(value):
    if value is None:
        return None
    text = str(value).strip()
    if 'Ã' in text or 'Â' in text or '�' in text:
        try:
            reparado = text.encode('latin-1').decode('utf-8')
            if reparado.count('Ã') + reparado.count('�') < text.count('Ã') + text.count('�'):
                text = reparado
        except (UnicodeError, ValueError):
            pass
    return text if text else None


def filas(hoja, min_row, max_row=None, min_col=1, max_col=None):
    return list(wb[hoja].iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col, values_only=True))


def exigir(condicion, mensaje):
    if not condicion:
        raise ValueError(mensaje)


def opcion(valor, **extra):
    resultado = {'valor': valor, 'etiqueta': valor}
    resultado.update({k: v for k, v in extra.items() if v is not None})
    return resultado


def unicos(items, clave):
    vistos = set()
    salida = []
    for item in items:
        k = clave(item)
        if k in vistos:
            continue
        vistos.add(k)
        salida.append(item)
    return salida


def agregar(relacion, padre, hijo):
    if padre is None or hijo is None:
        return
    lista = relacion.setdefault(padre, [])
    if hijo not in lista:
        lista.append(hijo)


# --- Listas de la hoja Base: Base!AV3:BG34 (columna AV = 48) ---
base = filas('Base', 3, 34, 48, 59)


def celda_base(fila, columna):
    return base[fila - 3][columna - 48]


exigir(clean(celda_base(4, 50)) == 'OFICINAS REGIONALES', 'Base!AX4 no contiene el encabezado OFICINAS REGIONALES.')
exigir(clean(celda_base(11, 52)) == 'Meses', 'Base!AZ11 no contiene el encabezado Meses.')
exigir(clean(celda_base(3, 54)) == 'Tipo de Transporte', 'Base!BB3 no contiene el encabezado Tipo de Transporte.')
exigir(clean(celda_base(3, 57)) == 'Estado del Informe', 'Base!BE3 no contiene el encabezado Estado del Informe.')

oficinas = [opcion(clean(celda_base(r, 50))) for r in range(5, 30) if clean(celda_base(r, 50))]
meses = unicos([opcion(clean(celda_base(r, 52))) for r in range(12, 35) if clean(celda_base(r, 52))], lambda i: i['valor'])
tipo_transporte = [
    opcion(clean(celda_base(r, 54)), horas=celda_base(r, 55))
    for r in range(4, 10) if clean(celda_base(r, 54))
]
calidad = [
    opcion(clean(celda_base(r, 57)), criterio=clean(celda_base(r, 58)), puntaje=celda_base(r, 59))
    for r in range(4, 7) if clean(celda_base(r, 57))
]

# --- Tabla2 (Base!AD6:AS1937): oficina regional, provincia y distrito ---
tabla2 = filas('Base', 6, 1937, 30, 36)
cabecera = [clean(v) for v in tabla2[0]]
exigir(cabecera[0] == 'Oficina Regional' and cabecera[2] == 'Provincia'
       and cabecera[4] == 'Concatenado' and cabecera[6] == 'Distrito',
       'Los encabezados de Base!AD6:AJ6 no coinciden con lo esperado.')

provincia = []
distrito = []
provincia_por_oficina = {}
distrito_por_oficina_provincia = {}
for fila in tabla2[1:]:
    oficina, prov, concatenado, dist = clean(fila[0]), clean(fila[2]), clean(fila[4]), clean(fila[6])
    if not (oficina and prov and concatenado):
        continue
    provincia.append(opcion(prov, oficina_regional=oficina, concatenado=concatenado))
    agregar(provincia_por_oficina, oficina, prov)
    if dist:
        distrito.append(opcion(dist, oficina_regional=oficina, provincia=prov, concatenado=concatenado))
        agregar(distrito_por_oficina_provincia, concatenado, dist)
provincia = unicos(provincia, lambda i: i['concatenado'])
distrito = unicos(distrito, lambda i: (i['concatenado'], i['valor']))

# --- Acción, acción específica y subacción (valor = Nombre_Abreviado, como en Registro) ---
accion = []
codigo_accion_a_valor = {}
for fila in filas('Accion', 2, None, 1, 4):
    nombre = clean(fila[0])
    if not nombre:
        continue
    accion.append(opcion(nombre, accion=clean(fila[1]), codigo_accion=clean(fila[2]), codigo_plan_accion=clean(fila[3])))
    codigo_accion_a_valor[clean(fila[2])] = nombre

accion_especifica = []
codigo_ae_a_valor = {}
accion_especifica_por_accion = {}
for fila in filas('AccionEspecifica', 2, None, 1, 4):
    nombre = clean(fila[0])
    if not nombre:
        continue
    accion_especifica.append(opcion(nombre, accion_especifica=clean(fila[1]),
                                    codigo_accion_especifica=clean(fila[2]), codigo_accion=clean(fila[3])))
    codigo_ae_a_valor[clean(fila[2])] = nombre
    agregar(accion_especifica_por_accion, codigo_accion_a_valor.get(clean(fila[3])), nombre)

sub_accion = []
sub_accion_por_ae = {}
for fila in filas('SubAccion', 2, None, 1, 5):
    nombre = clean(fila[0])
    if not nombre:
        continue
    sub_accion.append(opcion(nombre, sub_accion_especifica=clean(fila[1]), concat=clean(fila[2]),
                             codigo_sub_accion_especifica=clean(fila[3]), codigo_accion_especifica=clean(fila[4])))
    agregar(sub_accion_por_ae, codigo_ae_a_valor.get(clean(fila[4])), nombre)
sub_accion = unicos(sub_accion, lambda i: i['concat'] or i['valor'])

# --- Tipo de agente (TablaAgente, ActividadAgente!F:G) y actividades (A:D) ---
tipo_agente = []
codigo_ta_a_valor = {}
for fila in filas('ActividadAgente', 2, None, 6, 7):
    nombre = clean(fila[1])
    if not nombre:
        continue
    tipo_agente.append(opcion(nombre, codigo=clean(fila[0])))
    codigo_ta_a_valor[clean(fila[0])] = nombre
tipo_agente = unicos(tipo_agente, lambda i: i['valor'])

actividad = []
actividad_por_tipo_agente = {}
for fila in filas('ActividadAgente', 2, None, 1, 4):
    nombre = clean(fila[3])
    if not nombre:
        continue
    padre = clean(fila[1]) or codigo_ta_a_valor.get(clean(fila[0]))
    actividad.append(opcion(nombre, codigo_actividad=clean(fila[2]), codigo_tipo_agente=clean(fila[0]), tipo_agente=padre))
    agregar(actividad_por_tipo_agente, padre, nombre)
actividad = unicos(actividad, lambda i: (i.get('tipo_agente'), i['valor']))

# --- Empresas supervisoras ---
# La hoja escribe la oficina solo en la primera fila de cada bloque; las filas siguientes
# pertenecen a la misma oficina. Las filas con "-" son relleno y no se listan.
def normalizar(texto):
    sin_tildes = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    return sin_tildes.replace('_', ' ').casefold()


oficinas_base = {}
for item in oficinas:
    oficinas_base.setdefault(normalizar(item['valor']), []).append(item['valor'])

empresas = []
empresas_por_oficina = {}
oficina_hoja = None
for fila in filas('Lista de Empresas Supervisoras', 7, None, 1, 6):
    if clean(fila[1]):
        oficina_hoja = clean(fila[1])
    lista = clean(fila[5])
    if not lista or lista == '-':
        continue
    coincidencias = oficinas_base.get(normalizar(oficina_hoja), []) if oficina_hoja else []
    oficina = coincidencias[0] if len(coincidencias) == 1 else None
    empresas.append(opcion(lista, oficina_regional=oficina, oficina_regional_hoja=oficina_hoja,
                           contrato=clean(fila[2]), razon_social=clean(fila[3]), sub_sector=clean(fila[4])))
    agregar(empresas_por_oficina, oficina, lista)

resultado = {
    'catalogos': {
        'mes_presenta': meses,
        'oficina_regional': oficinas,
        'accion': accion,
        'accion_especifica': accion_especifica,
        'sub_accion_especifica': sub_accion,
        'tipo_agente': tipo_agente,
        'actividad': actividad,
        'provincia': provincia,
        'distrito': distrito,
        'tipo_transporte': tipo_transporte,
        'empresa_supervisora': empresas,
        'calidad_entregable': calidad,
    },
    'relaciones': {
        'accion_especifica_por_accion': accion_especifica_por_accion,
        'sub_accion_por_accion_especifica': sub_accion_por_ae,
        'actividad_por_tipo_agente': actividad_por_tipo_agente,
        'provincia_por_oficina_regional': provincia_por_oficina,
        'distrito_por_oficina_provincia': distrito_por_oficina_provincia,
        'empresas_supervisora_por_oficina_regional': empresas_por_oficina,
    },
}
print(json.dumps(resultado))
`;

async function obtenerParametros() {
  const plantillaPath = resolverPlantillaPath();
  if (!fs.existsSync(plantillaPath)) {
    throw crearErrorServicio(`Catálogos no disponibles: no se encontró el libro fuente "${NOMBRE_PLANTILLA}".`);
  }

  const { mtimeMs } = fs.statSync(plantillaPath);
  if (cache && cache.path === plantillaPath && cache.mtimeMs === mtimeMs) {
    return cache.datos;
  }

  try {
    const datos = ejecutarPython(SCRIPT, { PLANTILLA_EXCEL_PATH: plantillaPath });
    cache = { path: plantillaPath, mtimeMs, datos };
    return datos;
  } catch (error) {
    throw crearErrorServicio(`Catálogos no disponibles: no se pudo leer el libro fuente. ${error.message}`, error);
  }
}

module.exports = { obtenerParametros };
