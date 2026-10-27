'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');

const NOMBRE_PLANTILLA = '02. Sustento del Detalle de las Acciones de Supervisión de Hidrocarburos 2024 nov.xlsm';
const MIME_XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12';
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

function crearErrorServicio(mensaje, causa) {
  const error = new Error(mensaje, causa ? { cause } : undefined);
  error.statusCode = 503;
  return error;
}

function ejecutarPowerShell(scriptPath, argumentos) {
  return new Promise((resolve, reject) => {
    const proceso = spawn('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scriptPath,
      ...argumentos,
    ], { windowsHide: true });

    let salida = '';
    let error = '';
    proceso.stdout.setEncoding('utf8').on('data', (chunk) => { salida += chunk; });
    proceso.stderr.setEncoding('utf8').on('data', (chunk) => { error += chunk; });
    proceso.once('error', reject);
    proceso.once('close', (codigo) => {
      if (codigo === 0) {
        resolve(salida.trim());
      } else {
        reject(new Error(error.trim() || salida.trim() || `PowerShell terminó con código ${codigo}.`));
      }
    });
  });
}

async function exportarExcel(registros, plantillaPath) {
  if (!Array.isArray(registros)) {
    throw new TypeError('registros debe ser un array.');
  }

  const fuente = plantillaPath || process.env.PLANTILLA_EXCEL_PATH || path.join(__dirname, '..', '..', 'plantilla', NOMBRE_PLANTILLA);
  if (!fs.existsSync(fuente)) {
    throw crearErrorServicio(`Exportación no disponible: no se encontró el libro fuente "${NOMBRE_PLANTILLA}".`);
  }
  if (process.platform !== 'win32') {
    throw crearErrorServicio('Exportación no disponible: la preservación del libro requiere Excel COM en Windows.');
  }

  const datos = registros.map((registro) => {
    if (!registro || typeof registro !== 'object') {
      throw new TypeError('Cada elemento de registros debe ser un objeto.');
    }
    const fila = {};
    for (const campo of CAMPOS) {
      if (!Object.hasOwn(registro, campo)) {
        throw new TypeError(`Falta el campo ${campo} en un registro.`);
      }
      if (campo === 'numero_supervisores') {
        if (!Number.isInteger(registro[campo]) || registro[campo] <= 0) {
          throw new TypeError('numero_supervisores debe ser un entero positivo.');
        }
      } else if (typeof registro[campo] !== 'string') {
        throw new TypeError(`${campo} debe ser texto.`);
      }
      fila[campo] = registro[campo];
    }
    return fila;
  });

  const idTemporal = `${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const directorioTemporal = fs.mkdtempSync(path.join(os.tmpdir(), 'supervision-excel-'));
  const salidaPath = path.join(directorioTemporal, `reporte-${idTemporal}.xlsm`);
  const datosPath = path.join(directorioTemporal, 'registros.json');
  const scriptPath = path.join(directorioTemporal, 'exportar.ps1');
  const script = String.raw`param(
  [Parameter(Mandatory = $true)][string]$SourcePath,
  [Parameter(Mandatory = $true)][string]$OutputPath,
  [Parameter(Mandatory = $true)][string]$DataPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$required = @('mes_presenta','oficina_regional','accion','accion_especifica','sub_accion_especifica','tipo_agente','actividad','provincia','distrito','tipo_transporte','numero_supervisores','empresa_supervisora','calidad_entregable','expediente_siged','carta_linea','observaciones')
$excel = $null
$book = $null
function Get-WorkbookSignature($targetBook) {
  $names = @()
  $formulaCount = [long]0
  $validationCount = [long]0
  foreach ($targetSheet in $targetBook.Worksheets) {
    $names += $targetSheet.Name
    $range = $targetSheet.UsedRange
    try { $formulaCount += [long]$range.SpecialCells(-4123).CountLarge } catch {}
    try { $validationCount += [long]$range.SpecialCells(-4174).CountLarge } catch {}
  }
  $styleNames = @()
  for ($i = 1; $i -le $targetBook.Styles.Count; $i++) {
    $styleNames += $targetBook.Styles.Item($i).Name
  }
  return @{
    sheets = $names -join [char]31
    formulas = $formulaCount
    validations = $validationCount
    styles = $styleNames -join [char]31
  }
}
try {
  $sourceZip = [System.IO.Compression.ZipFile]::OpenRead($SourcePath)
  try {
    if (-not ($sourceZip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' })) {
      throw 'El libro fuente no contiene xl/vbaProject.bin; no se puede garantizar la preservación de macros.'
    }
  } finally { $sourceZip.Dispose() }

  Copy-Item -LiteralPath $SourcePath -Destination $OutputPath
  $rows = Get-Content -LiteralPath $DataPath -Raw | ConvertFrom-Json
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Open($OutputPath, 0, $false)
  $before = Get-WorkbookSignature $book
  $sheet = $book.Worksheets.Item('Registro')
  $used = $sheet.UsedRange
  $values = $used.Value2
  $headerRow = 0
  $columns = @{}
  for ($r = 1; $r -le $used.Rows.Count -and $r -le 50; $r++) {
    $candidate = @{}
    for ($c = 1; $c -le $used.Columns.Count; $c++) {
      $value = if ($values -is [Array]) { $values.GetValue($r, $c) } else { $values }
      if ($value -is [string] -and $required -contains $value.Trim()) {
        $candidate[$value.Trim()] = $used.Column + $c - 1
      }
    }
    if (@($required | Where-Object { -not $candidate.ContainsKey($_) }).Count -eq 0) {
      $headerRow = $used.Row + $r - 1
      $columns = $candidate
      break
    }
  }
  if ($headerRow -eq 0) {
    throw 'No se encontró en la hoja Registro una fila con todos los encabezados exactos del contrato; se evitó modificar el libro.'
  }

  $lastRow = $headerRow
  for ($rowIndex = $headerRow + 1; $rowIndex -le ($used.Row + $used.Rows.Count - 1); $rowIndex++) {
    $hasData = $false
    foreach ($field in $required) {
      if ($sheet.Cells.Item($rowIndex, $columns[$field]).Value2 -ne $null) { $hasData = $true; break }
    }
    if ($hasData) { $lastRow = $rowIndex }
  }

  $startRow = $lastRow + 1
  for ($i = 0; $i -lt $rows.Count; $i++) {
    $targetRow = $startRow + $i
    foreach ($field in $required) {
      $cell = $sheet.Cells.Item($targetRow, $columns[$field])
      if ($cell.HasFormula -or $cell.Value2 -ne $null) {
        throw "La fila destino $targetRow contiene datos o fórmulas; se evitó sobrescribir el libro."
      }
      $cell.Value2 = $rows[$i].$field
    }
  }

  $book.Save()
  $book.Close($true)
  $book = $null
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  $excel = $null

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AutomationSecurity = 3
  $book = $excel.Workbooks.Open($OutputPath, 0, $true)
  $after = Get-WorkbookSignature $book
  $book.Close($false)
  $book = $null
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  $excel = $null
  foreach ($key in @('sheets','formulas','validations','styles')) {
    if ($before[$key] -cne $after[$key]) {
      throw "La verificación de $key no coincide tras guardar el libro."
    }
  }

  $outputZip = [System.IO.Compression.ZipFile]::OpenRead($OutputPath)
  try {
    if (-not ($outputZip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' })) {
      throw 'Excel eliminó xl/vbaProject.bin al guardar; se rechaza la exportación.'
    }
  } finally { $outputZip.Dispose() }
  Write-Output 'OK'
} catch {
  if ($book -ne $null) { $book.Close($false) }
  if ($excel -ne $null) { $excel.Quit() }
  throw
}`;

  try {
    fs.writeFileSync(datosPath, JSON.stringify(datos), 'utf8');
    fs.writeFileSync(scriptPath, script, 'utf8');
    await ejecutarPowerShell(scriptPath, [fuente, salidaPath, datosPath]);
    return {
      path: salidaPath,
      mimeType: MIME_XLSM,
      filename: path.basename(salidaPath),
    };
  } catch (error) {
    fs.rmSync(directorioTemporal, { recursive: true, force: true });
    throw crearErrorServicio(`No se pudo verificar y exportar el libro Excel sin alterar sus macros: ${error.message}`, error);
  } finally {
    for (const archivo of [datosPath, scriptPath]) {
      if (fs.existsSync(archivo)) {
        fs.rmSync(archivo);
      }
    }
  }
}

module.exports = { exportarExcel };
