
"use strict";

const API_BASE = "http://localhost:3000/api";

const CAMPOS = [
  "mes_presenta",
  "oficina_regional",
  "accion",
  "accion_especifica",
  "sub_accion_especifica",
  "tipo_agente",
  "actividad",
  "provincia",
  "distrito",
  "tipo_transporte",
  "numero_supervisores",
  "empresa_supervisora",
  "calidad_entregable",
  "expediente_siged",
  "carta_linea",
  "observaciones"
];

const SELECTS = CAMPOS.filter(
  campo => document.getElementById(campo)?.tagName === "SELECT"
);

const form = document.getElementById("formRegistro");
const tabla = document.getElementById("tablaRegistros");
const mensaje = document.getElementById("mensaje");
const buscar = document.getElementById("buscarRegistro");
const btnGuardar = document.getElementById("btnGuardar");
const btnExportar = document.getElementById("btnExportar");

let registros = [];
let catalogos = {};
let relaciones = {};

function mostrarMensaje(texto, error = false) {
  mensaje.textContent = texto;
  mensaje.style.color = error ? "#b42318" : "#087e8b";
}

async function peticion(ruta, opciones = {}) {
  let respuesta;

  try {
    respuesta = await fetch(`${API_BASE}${ruta}`, opciones);
  } catch {
    throw new Error("No se puede conectar con el servidor.");
  }

  if (respuesta.status === 204) return null;

  const contenido = await respuesta.json().catch(() => null);

  if (!respuesta.ok) {
    throw new Error(
      contenido?.error || `Error del servidor: ${respuesta.status}`
    );
  }

  return contenido?.data;
}

function llenarSelect(campo, opciones, seleccionado = "") {
  const select = document.getElementById(campo);
  if (!select) return;
  select.replaceChildren(new Option("Seleccione una opción", ""));
  if (!Array.isArray(opciones)) { select.disabled = true; return; }
  const vistos = new Set();
  for (const item of opciones) {
    let valor, etiqueta;
    if (typeof item === "string" || typeof item === "number") {
      valor = etiqueta = String(item);
    } else if (item && typeof item === "object") {
      // Registro utiliza nombres abreviados; los códigos solo sirven para relaciones.
      if (campo === "accion") valor = item.nombre_abreviado ?? item.accion;
      else if (campo === "accion_especifica") valor = item.nombre_abreviado ?? item.accion_especifica;
      else if (campo === "sub_accion_especifica") valor = item.nombre_abreviado ?? item.sub_accion_especifica;
      else if (campo === "tipo_agente") valor = item.nombre;
      else if (campo === "actividad") valor = item.actividad;
      else if (campo === "empresa_supervisora") valor = item.lista_empresas;
      else valor = item.valor ?? item.nombre ?? item.codigo;
      etiqueta = valor;
    }
    if (valor == null || vistos.has(String(valor))) continue;
    vistos.add(String(valor));
    select.add(new Option(String(etiqueta), String(valor)));
  }
  select.value = String(seleccionado ?? "");
  select.disabled = select.options.length < 2;
}

// Relaciones basadas en los códigos reales del Excel; no en el texto mostrado.
const DEPENDENCIAS = {
  accion: { destino: "accion_especifica", codigoPadre: "codigo_accion", codigoHijo: "codigo_accion" },
  accion_especifica: { destino: "sub_accion_especifica", codigoPadre: "codigo_accion_especifica", codigoHijo: "codigo_accion_especifica" },
  tipo_agente: { destino: "actividad", codigoPadre: "codigo", codigoHijo: "codigo_tipo_agente" },
  oficina_regional: { destino: "empresa_supervisora", codigoPadre: null, codigoHijo: "oficina_regional" }
};

function valorCatalogo(campo, item) {
  if (typeof item !== "object" || item === null) return String(item ?? "");
  const claves = {
    accion: ["nombre_abreviado", "accion"],
    accion_especifica: ["nombre_abreviado", "accion_especifica"],
    sub_accion_especifica: ["nombre_abreviado", "sub_accion_especifica"],
    tipo_agente: ["nombre"], actividad: ["actividad"],
    empresa_supervisora: ["lista_empresas"]
  };
  const clave = (claves[campo] ?? ["valor", "nombre", "codigo"]).find(k => item[k] != null);
  return String(clave ? item[clave] : "");
}

function dependientes(campo, seleccionado = "") {
  const cfg = DEPENDENCIAS[campo];
  if (!cfg) return;
  const elegido = document.getElementById(campo)?.value ?? "";
  const padre = (catalogos[campo] ?? []).find(item => valorCatalogo(campo, item) === elegido);
  const codigo = cfg.codigoPadre ? padre?.[cfg.codigoPadre] : elegido;
  const opciones = codigo == null || codigo === "" ? [] :
    (catalogos[cfg.destino] ?? []).filter(item =>
      String(item[cfg.codigoHijo] ?? "") === String(codigo));
  llenarSelect(cfg.destino, opciones, seleccionado);
  if (!seleccionado) dependientes(cfg.destino);
}

for (const campo of Object.keys(DEPENDENCIAS)) {
  document.getElementById(campo)?.addEventListener("change", () => dependientes(campo));
}

async function cargarCatalogos() {
  try {
    const datos = await peticion("/parametros");

    catalogos = datos?.catalogos ?? {};
    relaciones = datos?.relaciones ?? {};

    for (const campo of SELECTS) {
      llenarSelect(campo, catalogos[campo]);
    }

    for (const campo of Object.keys(DEPENDENCIAS)) dependientes(campo);
    mostrarMensaje("Catálogos cargados correctamente.");
  } catch (error) {
    for (const campo of SELECTS) {
      llenarSelect(campo, []);
    }

    mostrarMensaje(`Catálogos no disponibles: ${error.message}`, true);
  }
}

function obtenerFormulario() {
  const datos = {};

  for (const campo of CAMPOS) {
    const elemento = document.getElementById(campo);
    datos[campo] = elemento.value.trim();
  }

  if (datos.numero_supervisores !== "") {
    datos.numero_supervisores = Number(datos.numero_supervisores);
  } else {
    datos.numero_supervisores = null;
  }

  return datos;
}

function validarFormulario(datos) {
  if (
    !Number.isInteger(datos.numero_supervisores) ||
    datos.numero_supervisores < 1
  ) {
    throw new Error("Ingrese un número válido de supervisores.");
  }
}

function crearCelda(fila, texto) {
  const celda = document.createElement("td");
  celda.textContent = String(texto ?? "");
  fila.appendChild(celda);
  return celda;
}

function dibujarTabla(datos) {
  tabla.replaceChildren();

  if (!datos.length) {
    const fila = document.createElement("tr");
    const celda = crearCelda(fila, "No hay registros para mostrar.");
    celda.colSpan = 8;
    tabla.appendChild(fila);
    return;
  }

  for (const registro of datos) {
    const fila = document.createElement("tr");

    [
      "id",
      "mes_presenta",
      "oficina_regional",
      "accion",
      "provincia",
      "empresa_supervisora",
      "expediente_siged"
    ].forEach(campo => crearCelda(fila, registro[campo]));

    const operaciones = document.createElement("td");

    const editar = document.createElement("button");
    editar.type = "button";
    editar.textContent = "Editar";
    editar.className = "secundario";
    editar.addEventListener("click", () => editarRegistro(registro.id));

    const eliminar = document.createElement("button");
    eliminar.type = "button";
    eliminar.textContent = "Eliminar";
    eliminar.addEventListener("click", () => eliminarRegistro(registro.id));

    operaciones.append(editar, document.createTextNode(" "), eliminar);
    fila.appendChild(operaciones);
    tabla.appendChild(fila);
  }
}

async function cargarRegistros() {
  try {
    const datos = await peticion("/registros");
    if (!Array.isArray(datos)) {
      throw new Error("El servidor devolvió un listado inválido.");
    }

    registros = datos;
    filtrarRegistros();
  } catch (error) {
    registros = [];
    tabla.replaceChildren();

    const fila = document.createElement("tr");
    const celda = crearCelda(fila, error.message);
    celda.colSpan = 8;
    tabla.appendChild(fila);

    mostrarMensaje(error.message, true);
  }
}

function filtrarRegistros() {
  const texto = buscar.value.trim().toLocaleLowerCase("es");

  const filtrados = registros.filter(registro =>
    [
      registro.expediente_siged,
      registro.accion,
      registro.empresa_supervisora,
      registro.provincia
    ].some(valor =>
      String(valor ?? "").toLocaleLowerCase("es").includes(texto)
    )
  );

  dibujarTabla(filtrados);
}

function limpiarFormulario() {
  document.getElementById("registroId").value = "";
  btnGuardar.textContent = "Guardar registro";
  for (const campo of Object.keys(DEPENDENCIAS)) dependientes(campo);
  mostrarMensaje("");
}

async function editarRegistro(id) {
  try {
    const registro = await peticion(`/registros/${encodeURIComponent(id)}`);

    for (const campo of CAMPOS) {
      const elemento = document.getElementById(campo);
      if (!elemento) continue;

      if (elemento.tagName === "SELECT") {
        llenarSelect(campo, catalogos[campo], registro[campo]);
      } else {
        elemento.value = registro[campo] ?? "";
      }
    }

    for (const campo of Object.keys(DEPENDENCIAS)) {
      dependientes(campo, registro[DEPENDENCIAS[campo].destino]);
    }
    document.getElementById("registroId").value = registro.id;
    btnGuardar.textContent = "Actualizar registro";

    window.scrollTo({ top: 0, behavior: "smooth" });
    mostrarMensaje(`Editando registro ${registro.id}.`);
  } catch (error) {
    mostrarMensaje(error.message, true);
  }
}

async function eliminarRegistro(id) {
  if (!confirm(`¿Desea eliminar el registro ${id}?`)) return;

  try {
    await peticion(`/registros/${encodeURIComponent(id)}`, {
      method: "DELETE"
    });

    await cargarRegistros();
    mostrarMensaje("Registro eliminado correctamente.");
  } catch (error) {
    mostrarMensaje(error.message, true);
  }
}

form.addEventListener("submit", async evento => {
  evento.preventDefault();

  try {
    const datos = obtenerFormulario();
    validarFormulario(datos);

    const id = document.getElementById("registroId").value;
    const editando = id !== "";

    btnGuardar.disabled = true;

    await peticion(
      editando ? `/registros/${encodeURIComponent(id)}` : "/registros",
      {
        method: editando ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(datos)
      }
    );

    form.reset();
    limpiarFormulario();
    await cargarRegistros();

    mostrarMensaje(
      editando
        ? "Registro actualizado correctamente."
        : "Registro guardado correctamente."
    );
  } catch (error) {
    mostrarMensaje(error.message, true);
  } finally {
    btnGuardar.disabled = false;
  }
});

form.addEventListener("reset", () => {
  limpiarFormulario();
});

buscar.addEventListener("input", filtrarRegistros);

btnExportar.addEventListener("click", async () => {
  btnExportar.disabled = true;

  try {
    const respuesta = await fetch(`${API_BASE}/exportar-excel`);

    if (!respuesta.ok) {
      const error = await respuesta.json().catch(() => null);
      throw new Error(error?.error || "No se pudo exportar el Excel.");
    }

    const archivo = await respuesta.blob();
    const url = URL.createObjectURL(archivo);

    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download = "supervision-hidrocarburos.xlsm";
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();

    URL.revokeObjectURL(url);
    mostrarMensaje("Exportación descargada correctamente.");
  } catch (error) {
    mostrarMensaje(error.message, true);
  } finally {
    btnExportar.disabled = false;
  }
});

async function iniciar() {
  await cargarCatalogos();
  await cargarRegistros();
}

iniciar();
