/* ==========================================================================
   Sistema de Gestión Contractual y Financiera — lógica de interfaz
   ========================================================================== */

const API_BASE = (typeof window.BACKEND_URL === "string" ? window.BACKEND_URL : "").replace(/\/$/, "");

const MODULOS = {
  extraccion:    { eyebrow: "Módulo 01", titulo: "Extracción de causaciones" },
  plantilla:     { eyebrow: "Módulo 02", titulo: "Plantilla de pagos" },
  pac:           { eyebrow: "Módulo 03", titulo: "Cuadre PAC" },
  estadocuenta:  { eyebrow: "Módulo 04", titulo: "Estado de cuenta" },
};

/* ---------------------------- Navegación ---------------------------- */

function cambiarVista(idVista) {
  document.querySelectorAll(".vista").forEach(v => v.hidden = true);
  document.getElementById(`vista-${idVista}`).hidden = false;

  document.querySelectorAll(".nav-item").forEach(b => b.classList.remove("activo"));
  document.querySelector(`.nav-item[data-vista="${idVista}"]`).classList.add("activo");

  document.getElementById("eyebrow-modulo").textContent = MODULOS[idVista].eyebrow;
  document.getElementById("titulo-modulo").textContent = MODULOS[idVista].titulo;
}

document.querySelectorAll(".nav-item").forEach(btn => {
  btn.addEventListener("click", () => cambiarVista(btn.dataset.vista));
});

/* ---------------------------- Dropzones ---------------------------- */

function formatearTamano(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function configurarDropzoneUnico(dropId, inputId, listaId) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const lista = document.getElementById(listaId);

  function render() {
    lista.innerHTML = "";
    if (input.files.length) {
      const f = input.files[0];
      const chip = document.createElement("div");
      chip.className = "archivo-chip";
      chip.innerHTML = `<span>${f.name} · ${formatearTamano(f.size)}</span>`;
      lista.appendChild(chip);
    }
  }

  input.addEventListener("change", render);

  ["dragover", "dragenter"].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("arrastrando"); })
  );
  ["dragleave", "drop"].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("arrastrando"); })
  );
  drop.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) {
      input.files = e.dataTransfer.files;
      render();
    }
  });
}

function configurarDropzoneMultiple(dropId, inputId, listaId) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const lista = document.getElementById(listaId);
  let archivos = [];

  function sincronizarInput() {
    const dt = new DataTransfer();
    archivos.forEach(f => dt.items.add(f));
    input.files = dt.files;
  }

  function render() {
    lista.innerHTML = "";
    archivos.forEach((f, idx) => {
      const chip = document.createElement("div");
      chip.className = "archivo-chip";
      chip.innerHTML = `<span>${f.name} · ${formatearTamano(f.size)}</span>`;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "quitar";
      btn.textContent = "×";
      btn.addEventListener("click", () => {
        archivos.splice(idx, 1);
        sincronizarInput();
        render();
      });
      chip.appendChild(btn);
      lista.appendChild(chip);
    });
  }

  function agregarArchivos(nuevos) {
    for (const f of nuevos) {
      if (!archivos.some(a => a.name === f.name && a.size === f.size)) {
        archivos.push(f);
      }
    }
    sincronizarInput();
    render();
  }

  input.addEventListener("change", () => agregarArchivos(input.files));

  ["dragover", "dragenter"].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("arrastrando"); })
  );
  ["dragleave", "drop"].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove("arrastrando"); })
  );
  drop.addEventListener("drop", e => {
    if (e.dataTransfer.files.length) agregarArchivos(e.dataTransfer.files);
  });
}

configurarDropzoneUnico("drop-basegen", "input-basegen", "lista-basegen");
configurarDropzoneMultiple("drop-pdfs", "input-pdfs", "lista-pdfs");
configurarDropzoneUnico("drop-extraccion-in", "input-extraccion-in", "lista-extraccion-in");
configurarDropzoneUnico("drop-plantilla-pac", "input-plantilla-pac", "lista-plantilla-pac");
configurarDropzoneUnico("drop-reporte-pac", "input-reporte-pac", "lista-reporte-pac");
configurarDropzoneUnico("drop-plantilla-ec", "input-plantilla-ec", "lista-plantilla-ec");
configurarDropzoneUnico("drop-historico-ec", "input-historico-ec", "lista-historico-ec");

/* ---------------------------- Estado del backend ---------------------------- */
/* Servicios gratuitos como Render "duermen" el backend tras inactividad y la
   primera petición puede tardar 30-50s en responder mientras despierta. Esta
   verificación evita que la primera persona que use el formulario piense que
   está roto: muestra un aviso visible mientras el servidor arranca. */

async function verificarBackend() {
  const banner = document.getElementById("banner-backend");
  if (!API_BASE) { banner.hidden = true; return; }

  banner.hidden = false;
  banner.textContent = "Conectando con el servidor… puede tardar hasta 1 minuto si está inactivo.";
  banner.className = "banner-backend banner-cargando";

  try {
    const resp = await fetch(`${API_BASE}/api/salud`, { method: "GET" });
    if (resp.ok) {
      banner.className = "banner-backend banner-listo";
      banner.textContent = "Servidor conectado. Todo listo para procesar.";
      setTimeout(() => { banner.hidden = true; }, 2500);
    } else {
      throw new Error("respuesta no ok");
    }
  } catch (err) {
    banner.className = "banner-backend banner-error";
    banner.textContent = "No se pudo conectar con el servidor. Intenta de nuevo en unos segundos o recarga la página.";
  }
}

verificarBackend();

/* ---------------------------- Utilidades UI ---------------------------- */

function mostrarEstado(id, mostrar) {
  document.getElementById(id).classList.toggle("visible", mostrar);
}

function mostrarAlerta(id, mensaje) {
  const el = document.getElementById(id);
  if (!mensaje) { el.classList.remove("visible"); el.textContent = ""; return; }
  el.textContent = mensaje;
  el.classList.add("visible");
}

function crearMetrica(valor, etiqueta, acento) {
  const div = document.createElement("div");
  div.className = `metrica-card${acento ? " acento-" + acento : ""}`;
  div.innerHTML = `<div class="metrica-valor">${valor}</div><div class="metrica-etiqueta">${etiqueta}</div>`;
  return div;
}

function formatearMoneda(valor) {
  const n = Number(valor) || 0;
  return n.toLocaleString("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 });
}

/* ---------------------------- Módulo 01: Extracción ---------------------------- */

document.getElementById("form-extraccion").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-extraccion", "");
  document.getElementById("resultados-extraccion").classList.remove("visible");

  const basegen = document.getElementById("input-basegen").files[0];
  const pdfs = document.getElementById("input-pdfs").files;

  if (!basegen || !pdfs.length) {
    mostrarAlerta("alerta-extraccion", "Debes adjuntar el archivo BASEGEN y al menos un PDF.");
    return;
  }

  const fd = new FormData();
  fd.append("basegen", basegen);
  for (const f of pdfs) fd.append("pdfs", f);

  const btn = document.getElementById("btn-extraccion");
  btn.disabled = true;
  mostrarEstado("estado-extraccion", true);

  try {
    const resp = await fetch(`${API_BASE}/api/extraccion/procesar`, { method: "POST", body: fd });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-extraccion", data.mensaje || "No se pudo completar la extracción.");
      return;
    }

    const cont = document.getElementById("metricas-extraccion");
    cont.innerHTML = "";
    cont.appendChild(crearMetrica(data.total_pdfs, "PDFs recibidos"));
    cont.appendChild(crearMetrica(data.procesados, "Procesados con éxito", "verde"));
    cont.appendChild(crearMetrica(data.errores.length, "Con errores", data.errores.length ? "rojo" : ""));
    cont.appendChild(crearMetrica(data.registros, "Registros finales", "ocre"));

    if (data.errores && data.errores.length) {
      mostrarAlerta("alerta-extraccion", `Algunos archivos no se pudieron procesar: ${data.errores.slice(0, 5).join("; ")}${data.errores.length > 5 ? "…" : ""}`);
    }

    document.getElementById("link-descarga-extraccion").href = API_BASE + data.descarga;
    document.getElementById("resultados-extraccion").classList.add("visible");

  } catch (err) {
    mostrarAlerta("alerta-extraccion", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-extraccion", false);
  }
});

/* ---------------------------- Módulo 02: Plantilla de pagos ---------------------------- */

document.getElementById("form-plantilla").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-plantilla", "");
  document.getElementById("resultados-plantilla").classList.remove("visible");

  const archivo = document.getElementById("input-extraccion-in").files[0];
  if (!archivo) {
    mostrarAlerta("alerta-plantilla", "Debes adjuntar el archivo de extracción.");
    return;
  }

  const fd = new FormData();
  fd.append("extraccion", archivo);

  const btn = document.getElementById("btn-plantilla");
  btn.disabled = true;
  mostrarEstado("estado-plantilla", true);

  try {
    const resp = await fetch(`${API_BASE}/api/plantilla-pagos/procesar`, { method: "POST", body: fd });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-plantilla", data.mensaje || "No se pudo generar la plantilla.");
      return;
    }

    const cont = document.getElementById("metricas-plantilla");
    cont.innerHTML = "";
    cont.appendChild(crearMetrica(data.registros_procesados, "Registros procesados"));
    cont.appendChild(crearMetrica(data.filas_con_codigo_44, "Con honorarios (cód. 44)", "ocre"));
    cont.appendChild(crearMetrica(data.tiene_columnas_honorarios ? "Sí" : "No", "Columnas de honorarios detectadas"));

    document.getElementById("link-descarga-plantilla").href = API_BASE + data.descarga;
    document.getElementById("resultados-plantilla").classList.add("visible");

  } catch (err) {
    mostrarAlerta("alerta-plantilla", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-plantilla", false);
  }
});

/* ---------------------------- Módulo 03: Cuadre PAC ---------------------------- */

document.getElementById("form-pac").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-pac", "");
  mostrarAlerta("aviso-pac", "");
  document.getElementById("resultados-pac").classList.remove("visible");

  const plantilla = document.getElementById("input-plantilla-pac").files[0];
  const reporte = document.getElementById("input-reporte-pac").files[0];
  const mes = document.getElementById("input-mes-pac").value.trim();

  if (!plantilla || !reporte || !mes) {
    mostrarAlerta("alerta-pac", "Debes adjuntar ambos archivos e indicar el mes a evaluar.");
    return;
  }

  const fd = new FormData();
  fd.append("plantilla", plantilla);
  fd.append("pac", reporte);
  fd.append("mes", mes);

  const btn = document.getElementById("btn-pac");
  btn.disabled = true;
  mostrarEstado("estado-pac", true);

  try {
    const resp = await fetch(`${API_BASE}/api/cuadre-pac/procesar`, { method: "POST", body: fd });
    const data = await resp.json();

    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-pac", data.mensaje || "No se pudo realizar el cuadre PAC.");
      return;
    }

    const cont = document.getElementById("metricas-pac");
    cont.innerHTML = "";
    cont.appendChild(crearMetrica(data.combinaciones, "Combinaciones Rubro+Fondos"));
    cont.appendChild(crearMetrica(data.cubre, "Cubren en " + data.mes, "verde"));
    cont.appendChild(crearMetrica(data.no_cubre, "No cubren", data.no_cubre ? "rojo" : ""));
    cont.appendChild(crearMetrica(data.sin_disponibilidad, "Sin disponibilidad", data.sin_disponibilidad ? "ocre" : ""));

    if (data.sin_match_en_pac > 0) {
      mostrarAlerta("aviso-pac", `${data.sin_match_en_pac} rubro(s) de la plantilla no tienen coincidencia en el reporte PAC del mes indicado.`);
    }

    const tbody = document.getElementById("tabla-pac-body");
    tbody.innerHTML = "";
    data.tabla.forEach(fila => {
      const tr = document.createElement("tr");
      let pillClase = "pill-sin-disp";
      if (fila.Cobertura.startsWith("CUBRE")) pillClase = "pill-cubre";
      else if (fila.Cobertura.startsWith("NO CUBRE")) pillClase = "pill-no-cubre";

      tr.innerHTML = `
        <td>${fila.Rubro ?? ""}</td>
        <td>${fila.Fondos ?? ""}</td>
        <td>${formatearMoneda(fila.Importe_Total)}</td>
        <td>${formatearMoneda(fila.Disponibilidad_PAC)}</td>
        <td>${formatearMoneda(fila.Diferencia)}</td>
        <td><span class="pill-cobertura ${pillClase}">${fila.Cobertura}</span></td>
      `;
      tbody.appendChild(tr);
    });

    document.getElementById("link-descarga-pac").href = API_BASE + data.descarga;
    document.getElementById("resultados-pac").classList.add("visible");

  } catch (err) {
    mostrarAlerta("alerta-pac", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-pac", false);
  }
});

