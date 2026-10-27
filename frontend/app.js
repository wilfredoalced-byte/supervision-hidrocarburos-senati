
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

  if (!Array.isArray(opciones)) {
    select.disabled = true;
    return;
  }

  for (const item of opciones) {
    let valor;
    let etiqueta;

    if (typeof item === "string" || typeof item === "number") {
      valor = String(item);
      etiqueta = String(item);
    } else if (item && typeof item === "object") {
      valor = item.valor ?? item.value ?? item.codigo ?? item.id ?? item.codigo_accion ?? item.codigo_accion_especifica ?? item.codigo_sub_accion_especifica ?? item.codigo_actividad ?? item.lista_empresas ?? item.oficina_regional ?? item.nombre_abreviado ?? item.accion ?? item.accion_especifica ?? item.sub_accion_especifica ?? item.actividad;
      etiqueta = item.etiqueta ?? item.label ?? item.nombre ?? item.accion ?? item.accion_especifica ?? item.sub_accion_especifica ?? item.actividad ?? item.lista_empresas ?? item.razon_social ?? item.nombre_abreviado ?? valor;
    }

    if (valor === undefined || valor === null) continue;

    select.add(new Option(String(etiqueta), String(valor)));
  }

  select.value = String(seleccionado ?? "");
  select.disabled = opciones.length === 0;
}

async function cargarCatalogos() {
  try {
    const datos = await peticion("/parametros");

    catalogos = datos?.catalogos ?? {};
    relaciones = datos?.relaciones ?? {};

    for (const campo of SELECTS) {
      llenarSelect(campo, catalogos[campo]);
    }

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
