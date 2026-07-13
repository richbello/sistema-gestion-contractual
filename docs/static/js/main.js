/* ==========================================================================
   Sistema de Gestión Contractual y Financiera — lógica de interfaz
   ========================================================================== */
const API_BASE = (typeof window.BACKEND_URL === "string" ? window.BACKEND_URL : "").replace(/\/$/, "");

/* Descarga la plantilla saneada (base64 -> .xlsx) */
function descargarSaneada(b64) {
  const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Plantilla_Pagos_CORREGIDA.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}



const MODULOS = {
  extraccion:    { eyebrow: "Módulo 01", titulo: "Extracción de causaciones" },
  plantilla:     { eyebrow: "Módulo 02", titulo: "Plantilla de pagos" },
  pac:           { eyebrow: "Módulo 03", titulo: "Cuadre PAC" },
  estadocuenta:  { eyebrow: "Módulo 04", titulo: "Estado de cuenta" },
  historico:     { eyebrow: "Módulo 05", titulo: "Histórico de Pagos" },
  presupuestal:  { eyebrow: "Módulo 06", titulo: "Análisis Presupuestal" },
  contratacion:  { eyebrow: "Módulo 07", titulo: "Contratación" },
  planeacion:    { eyebrow: "Módulo 08", titulo: "Planeación PDL" },
  ciberseguridad:{ eyebrow: "Módulo 09", titulo: "Ciberseguridad" },
  licencias:     { eyebrow: "Módulo 10", titulo: "Licencias" },
  validacion:    { eyebrow: "Módulo 11", titulo: "Validación de plantilla" },
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

configurarDropzoneUnico("drop-crp-ec", "input-crp-ec", "lista-crp-ec");
configurarDropzoneUnico("drop-consolidado-ec", "input-consolidado-ec", "lista-consolidado-ec");



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
/* Con entregas grandes, el free tier de Render se queda sin memoria si recibe
   todos los PDFs en una sola petición. Por eso, cuando hay muchos PDFs, se
   trocean en lotes pequeños (LOTE_PDFS) que se procesan uno por uno; los
   registros se acumulan y el Excel final se genera de una sola vez al final. */

const LOTE_PDFS = 8;          // tamaño de lote seguro para el free tier
const UMBRAL_TROCEO = 10;     // hasta este número, se usa el flujo clásico

function _actualizarProgresoExtraccion(hechos, total) {
  const el = document.getElementById("estado-extraccion");
  if (!el) return;
  const txt = el.querySelector(".estado-texto") || el;
  txt.textContent = `Procesando lote ${hechos} de ${total}…`;
}

async function _extraerPorLotes(basegen, pdfsArray) {
  const totalPdfs = pdfsArray.length;
  const lotes = [];
  for (let i = 0; i < totalPdfs; i += LOTE_PDFS) {
    lotes.push(pdfsArray.slice(i, i + LOTE_PDFS));
  }

  let registros = [];
  let errores = [];

  for (let idx = 0; idx < lotes.length; idx++) {
    _actualizarProgresoExtraccion(idx + 1, lotes.length);
    const fd = new FormData();
    for (const f of lotes[idx]) fd.append("pdfs", f);

    const resp = await fetch(`${API_BASE}/api/extraccion/extraer-lote`, {
      method: "POST",
      body: fd,
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      throw new Error(data.mensaje || `Falló el lote ${idx + 1} de ${lotes.length}.`);
    }
    registros = registros.concat(data.registros || []);
    if (data.errores && data.errores.length) errores = errores.concat(data.errores);
  }

  // Consolidación final: cruce con BASEGEN + orden + Excel.
  _actualizarProgresoExtraccion(lotes.length, lotes.length);
  const fdFinal = new FormData();
  fdFinal.append("basegen", basegen);
  fdFinal.append("registros", JSON.stringify(registros));
  fdFinal.append("total_pdfs", String(totalPdfs));
  fdFinal.append("errores", JSON.stringify(errores));

  const respFinal = await fetch(`${API_BASE}/api/extraccion/consolidar`, {
    method: "POST",
    body: fdFinal,
  });
  const dataFinal = await respFinal.json();
  if (!respFinal.ok || !dataFinal.ok) {
    throw new Error(dataFinal.mensaje || "No se pudo consolidar el Excel final.");
  }
  return dataFinal;
}

async function _extraerDirecto(basegen, pdfs) {
  const fd = new FormData();
  fd.append("basegen", basegen);
  for (const f of pdfs) fd.append("pdfs", f);
  const resp = await fetch(`${API_BASE}/api/extraccion/procesar`, { method: "POST", body: fd });
  const data = await resp.json();
  if (!resp.ok || !data.ok) {
    throw new Error(data.mensaje || "No se pudo completar la extracción.");
  }
  return data;
}

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
  const btn = document.getElementById("btn-extraccion");
  btn.disabled = true;
  mostrarEstado("estado-extraccion", true);
  try {
    const pdfsArray = Array.from(pdfs);
    const data = pdfsArray.length > UMBRAL_TROCEO
      ? await _extraerPorLotes(basegen, pdfsArray)
      : await _extraerDirecto(basegen, pdfs);

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
    mostrarAlerta("alerta-extraccion", err.message || "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
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
    cont.appendChild(crearMetrica(data.cuentas_desde_registro ?? "—", "Cuentas verificadas contra el registro"));
    const sinCruzar = data.cuentas_sin_cruzar || [];
    cont.appendChild(crearMetrica(sinCruzar.length, "Cuentas por revisar", sinCruzar.length ? "ocre" : ""));
    // Aviso detallado cuando hay cedulas que no cruzaron con el registro
    const idAviso = "aviso-cuentas-plantilla";
    let aviso = document.getElementById(idAviso);
    if (aviso) aviso.remove();
    if (sinCruzar.length) {
      aviso = document.createElement("div");
      aviso.id = idAviso;
      aviso.className = "alerta alerta-aviso";
      aviso.style.display = "block";
      const lista = sinCruzar.map(c => `${c.nombre} (${c.cedula})`).join(", ");
      aviso.innerHTML = `<strong>Revisar cuentas:</strong> ${sinCruzar.length} contratista(s) no están en el registro depurado, por lo que su cuenta se tomó del archivo de extracción. Verifica antes de pagar: ${lista}.`;
      cont.parentNode.insertBefore(aviso, cont.nextSibling);
    }
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

/* ---------------------------- Módulo 11: Validación de plantilla ---------------------------- */
configurarDropzoneUnico("drop-validacion", "input-validacion", "lista-validacion");

document.getElementById("form-validacion").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-validacion", "");
  document.getElementById("resultados-validacion").classList.remove("visible");
  const archivo = document.getElementById("input-validacion").files[0];
  if (!archivo) {
    mostrarAlerta("alerta-validacion", "Debes adjuntar la plantilla a validar.");
    return;
  }
  const fd = new FormData();
  fd.append("plantilla", archivo);
  const btn = document.getElementById("btn-validacion");
  btn.disabled = true;
  mostrarEstado("estado-validacion", true);
  try {
    const resp = await fetch(`${API_BASE}/api/validacion-plantilla/procesar`, { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-validacion", data.mensaje || "No se pudo validar la plantilla.");
      return;
    }
    const rs = data.resumen || {};
    const cont = document.getElementById("metricas-validacion");
    cont.innerHTML = "";
    cont.appendChild(crearMetrica(rs.bloques_total ?? "—", "Bloques (contratistas)"));
    cont.appendChild(crearMetrica(rs.bloques_ok ?? "—", "Bloques correctos", "ok"));
    cont.appendChild(crearMetrica(rs.bloques_con_error ?? 0, "Bloques con error", (rs.bloques_con_error ? "ocre" : "")));
    cont.appendChild(crearMetrica(rs.celdas_con_caracteres ?? 0, "Celdas con caracteres", (rs.celdas_con_caracteres ? "ocre" : "")));

    cont.appendChild(crearMetrica(data.total_cambios ?? 0, "Correcciones aplicadas", (data.total_cambios ? "ok" : "")));

    // Detalle de estructura
    const estr = data.estructura || [];
    const boxE = document.getElementById("detalle-estructura");
    if (estr.length) {
      let html = '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;"><strong>Errores de estructura:</strong> ' + estr.length + ' bloque(s) a revisar.</div>';
      html += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr><th>Fila</th><th>Contratista</th><th>Problemas</th></tr></thead><tbody>';
      estr.forEach(x => {
        html += `<tr><td>${x.fila}</td><td>${x.contratista || ""}</td><td>${x.problemas.join("; ")}</td></tr>`;
      });
      html += '</tbody></table></div>';
      boxE.innerHTML = html;
    } else {
      boxE.innerHTML = '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">Estructura de bloques C/P40/P31 correcta.</div>';
    }

    // Detalle de caracteres
    const chars = data.caracteres || [];
    const boxC = document.getElementById("detalle-caracteres");
    if (chars.length) {
      let html = '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;"><strong>Caracteres problemáticos:</strong> ' + chars.length + ' celda(s)' + (data.caracteres_truncado ? " (mostrando las primeras 200)" : "") + '.</div>';
      html += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr><th>Celda</th><th>Valor</th><th>Problemas</th></tr></thead><tbody>';
      chars.forEach(x => {
        html += `<tr><td>${x.celda}</td><td>${(x.valor||"").replace(/</g,"&lt;")}</td><td>${x.problemas.join("; ")}</td></tr>`;
      });
      html += '</tbody></table></div>';
      boxC.innerHTML = html;
    } else {
      boxC.innerHTML = '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">Sin caracteres problemáticos para SAP.</div>';
    }

    // Plantilla saneada + descarga
    const cambios = data.cambios_saneado || [];
    const boxS = document.getElementById("detalle-saneado");
    if (data.plantilla_saneada_b64) {
      let html = "";
      if (cambios.length) {
        html += '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;"><strong>Correcciones aplicadas:</strong> ' + cambios.length + ' celda(s) saneada(s) para SAP.</div>';
        html += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr><th>Celda</th><th>Antes</th><th>Después</th></tr></thead><tbody>';
        cambios.forEach(x => {
          html += `<tr><td>${x.celda}</td><td>${(x.antes||"").replace(/</g,"&lt;")}</td><td>${(x.despues||"").replace(/</g,"&lt;")}</td></tr>`;
        });
        html += '</tbody></table></div>';
      } else {
        html += '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">La plantilla ya estaba limpia; no hubo correcciones.</div>';
      }
      html += '<button type="button" class="btn-primario" id="btn-descargar-saneada" style="margin-top:10px;">Descargar plantilla corregida</button>';
      boxS.innerHTML = html;
      document.getElementById("btn-descargar-saneada").onclick = () => descargarSaneada(data.plantilla_saneada_b64);
    } else {
      boxS.innerHTML = "";
    }

// --- V2: BP creado ---
    const bp = data.bp || [];
    const sinBp = bp.filter(x => x.estado !== "OK");
    const boxBp = document.getElementById("detalle-bp");
    if (bp.length && sinBp.length) {
      let h = '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;"><strong>BP no creado:</strong> ' + sinBp.length + ' contratista(s) sin BP.</div>';
      h += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr><th>Contrato</th><th>Documento</th><th>Contratista</th></tr></thead><tbody>';
      sinBp.forEach(x => { h += `<tr><td>${x.contrato}</td><td>${x.doc}</td><td>${x.contratista||""}</td></tr>`; });
      h += '</tbody></table></div>';
      boxBp.innerHTML = h;
    } else if (bp.length) {
      boxBp.innerHTML = '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">Todos los contratistas tienen BP creado.</div>';
    } else { boxBp.innerHTML = ""; }

    // --- V3: recursos CRP ---
    const crp = data.crp || [];
    const probCrp = crp.filter(x => x.estado !== "OK");
    const boxCrp = document.getElementById("detalle-crp");
    if (crp.length && probCrp.length) {
      let h = '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;"><strong>Recursos CRP:</strong> ' + probCrp.length + ' pago(s) a revisar.</div>';
      h += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr><th>Contrato</th><th>CRP</th><th>Detalle</th></tr></thead><tbody>';
      probCrp.forEach(x => { h += `<tr><td>${x.contrato}</td><td>${x.crp||""}</td><td>${x.mensaje}</td></tr>`; });
      h += '</tbody></table></div>';
      boxCrp.innerHTML = h;
    } else if (crp.length) {
      boxCrp.innerHTML = '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">Todos los pagos tienen recursos CRP suficientes.</div>';
    } else { boxCrp.innerHTML = ""; }



    document.getElementById("resultados-validacion").classList.add("visible");
  } catch (err) {
    mostrarAlerta("alerta-validacion", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-validacion", false);
  }
});
/* ---------------------------- Módulo 04: Estado de cuenta ---------------------------- */
document.getElementById("form-estadocuenta").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-estadocuenta", "");
  document.getElementById("resultados-estadocuenta").classList.remove("visible");

  const plantilla   = document.getElementById("input-plantilla-ec").files[0];
  const crp         = document.getElementById("input-crp-ec").files[0];
  const consolidado = document.getElementById("input-consolidado-ec").files[0];
  const historico   = document.getElementById("input-historico-ec").files[0];
  const contrato    = document.getElementById("input-contrato-ec").value.trim();

  if (!plantilla || !crp || !contrato) {
    mostrarAlerta("alerta-estadocuenta", "Debes adjuntar la plantilla, el Reporte CRP e indicar el número de contrato.");
    return;
  }

  const fd = new FormData();
  fd.append("plantilla", plantilla);
  fd.append("reporte_crp", crp);
  if (consolidado) fd.append("consolidado", consolidado);
  if (historico)   fd.append("historico", historico);
  fd.append("contrato", contrato);

  const btn = document.getElementById("btn-estadocuenta");
  btn.disabled = true;
  mostrarEstado("estado-estadocuenta", true);

  try {
    const resp = await fetch(`${API_BASE}/api/estado-cuenta/procesar`, { method: "POST", body: fd });

    if (!resp.ok) {
      let msg = "No se pudo generar el estado de cuenta.";
      try { const err = await resp.json(); if (err.mensaje) msg = err.mensaje; } catch (_) {}
      mostrarAlerta("alerta-estadocuenta", msg);
      return;
    }

    const blob = await resp.blob();
    const url  = URL.createObjectURL(blob);
    const nombre = `Estado_de_Cuenta_${contrato}.xlsx`;

    const link = document.getElementById("link-descarga-estadocuenta");
    link.href = url;
    link.download = nombre;
    document.getElementById("nombre-descarga-ec").textContent = nombre;
    document.getElementById("resultados-estadocuenta").classList.add("visible");
  } catch (err) {
    mostrarAlerta("alerta-estadocuenta", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-estadocuenta", false);
  }
});