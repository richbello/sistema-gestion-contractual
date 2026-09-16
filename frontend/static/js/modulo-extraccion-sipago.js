/* ==========================================================================
   Módulo — Extracción SIPAGO
   Cruza el Reporte del Periodo con dos bases guardadas en el servidor:
     · Base CRP       -> CRP y SALDO CRP   (cruce por cédula, gana el de mayor saldo)
     · Base bancaria  -> Código Bco, Tipo Cta, No Cuenta (gana la cuenta más reciente)
   ========================================================================== */

const API_SIPAGO = `${API_BASE}/api/extraccion-sipago`;

/* ---------- Dropzones ---------- */
configurarDropzoneUnico("drop-sipago-reporte", "input-sipago-reporte", "lista-sipago-reporte");
configurarDropzoneUnico("drop-sipago-crp", "input-sipago-crp", "lista-sipago-crp");
configurarDropzoneMultiple("drop-sipago-historicos", "input-sipago-historicos", "lista-sipago-historicos");

/* ---------- Estado de las bases guardadas ---------- */
async function sipagoRefrescarBases() {
  const boxCrp = document.getElementById("sipago-estado-crp");
  const boxBan = document.getElementById("sipago-estado-bancaria");
  if (!boxCrp || !boxBan) return;

  try {
    const resp = await fetch(`${API_SIPAGO}/estado-bases`);
    const data = await resp.json();
    if (!data.ok) throw new Error();

    const pintar = (el, info, etiqueta) => {
      if (info.existe) {
        el.className = "alerta alerta-aviso visible";
        el.style.background = "#e3f2e9";
        el.style.color = "#1f7a4d";
        el.textContent =
          `${etiqueta}: ${info.registros} contratistas · actualizada el ${info.actualizada}`;
      } else {
        el.className = "alerta alerta-aviso visible";
        el.style.background = "";
        el.style.color = "";
        el.textContent = `${etiqueta}: sin cargar. Debes subirla antes de consolidar.`;
      }
    };
    pintar(boxCrp, data.crp, "Base CRP");
    pintar(boxBan, data.bancaria, "Base bancaria");
  } catch (err) {
    boxCrp.textContent = "No se pudo consultar el estado de las bases.";
    boxCrp.className = "alerta alerta-error visible";
  }
}
sipagoRefrescarBases();

/* ---------- Reemplazar base CRP ---------- */
document.getElementById("form-sipago-crp").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-sipago-bases", "");

  const archivo = document.getElementById("input-sipago-crp").files[0];
  if (!archivo) {
    mostrarAlerta("alerta-sipago-bases", "Debes adjuntar el Reporte CRP.");
    return;
  }

  const fd = new FormData();
  fd.append("reporte_crp", archivo);

  const btn = document.getElementById("btn-sipago-crp");
  btn.disabled = true;
  mostrarEstado("estado-sipago-crp", true);
  try {
    const resp = await fetch(`${API_SIPAGO}/cargar-base-crp`, { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-sipago-bases", data.mensaje || "No se pudo cargar la base CRP.");
      return;
    }
    await sipagoRefrescarBases();
  } catch (err) {
    mostrarAlerta("alerta-sipago-bases", "Error de conexión con el servidor.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-sipago-crp", false);
  }
});

/* ---------- Reemplazar base bancaria ---------- */
document.getElementById("form-sipago-bancaria").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-sipago-bases", "");

  const archivos = document.getElementById("input-sipago-historicos").files;
  if (!archivos.length) {
    mostrarAlerta("alerta-sipago-bases", "Debes adjuntar al menos un histórico de pagos.");
    return;
  }

  const fd = new FormData();
  for (const f of archivos) fd.append("historicos", f);

  const btn = document.getElementById("btn-sipago-bancaria");
  btn.disabled = true;
  mostrarEstado("estado-sipago-bancaria", true);
  try {
    const resp = await fetch(`${API_SIPAGO}/cargar-base-bancaria`, { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-sipago-bases", data.mensaje || "No se pudo cargar la base bancaria.");
      return;
    }
    if (data.omitidos && data.omitidos.length) {
      mostrarAlerta("alerta-sipago-bases",
        `Archivos omitidos: ${data.omitidos.join("; ")}`);
    }
    await sipagoRefrescarBases();
  } catch (err) {
    mostrarAlerta("alerta-sipago-bases", "Error de conexión con el servidor.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-sipago-bancaria", false);
  }
});

/* ---------- Consolidar el Reporte del Periodo ---------- */
document.getElementById("form-sipago-procesar").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-sipago", "");
  document.getElementById("resultados-sipago").classList.remove("visible");

  const reporte = document.getElementById("input-sipago-reporte").files[0];
  if (!reporte) {
    mostrarAlerta("alerta-sipago", "Debes adjuntar el Reporte del Periodo.");
    return;
  }

  const fd = new FormData();
  fd.append("reporte", reporte);

  const btn = document.getElementById("btn-sipago-procesar");
  btn.disabled = true;
  mostrarEstado("estado-sipago", true);

  try {
    const resp = await fetch(`${API_SIPAGO}/procesar`, { method: "POST", body: fd });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-sipago", data.mensaje || "No se pudo consolidar el reporte.");
      return;
    }

    const cont = document.getElementById("metricas-sipago");
    cont.innerHTML = "";
    cont.appendChild(crearMetrica(data.registros, "Registros del reporte"));
    cont.appendChild(crearMetrica(data.completos, "Completos", "verde"));
    cont.appendChild(crearMetrica(data.sin_crp, "Sin CRP", data.sin_crp ? "ocre" : ""));
    cont.appendChild(crearMetrica(data.sin_cuenta, "Sin cuenta bancaria", data.sin_cuenta ? "ocre" : ""));

    // Detalle de los registros que quedaron incompletos
    const box = document.getElementById("detalle-sipago-pendientes");
    const pend = data.pendientes || [];
    if (pend.length) {
      let html = '<div class="alerta alerta-aviso" style="display:block;margin-bottom:10px;">'
               + `<strong>Revisar antes de pagar:</strong> ${pend.length} registro(s) sin CRP o sin cuenta bancaria.</div>`;
      html += '<div class="tabla-wrap"><table class="tabla-datos"><thead><tr>'
            + '<th>Contrato</th><th>Contratista</th><th>Cédula</th><th>Falta</th>'
            + '</tr></thead><tbody>';
      pend.forEach(p => {
        html += `<tr><td>${p["Contrato"] || ""}</td><td>${p["Contratista"] || ""}</td>`
              + `<td>${p["NIT_CC"] || ""}</td><td>${p["falta"] || ""}</td></tr>`;
      });
      html += "</tbody></table></div>";
      box.innerHTML = html;
    } else {
      box.innerHTML = '<div class="alerta alerta-aviso" style="display:block;background:#e3f2e9;color:#1f7a4d;">'
                    + "Todos los registros quedaron con CRP y cuenta bancaria.</div>";
    }

    document.getElementById("link-descarga-sipago").href = API_BASE + data.descarga;
    document.getElementById("resultados-sipago").classList.add("visible");
  } catch (err) {
    mostrarAlerta("alerta-sipago", "Error de conexión con el servidor.");
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-sipago", false);
  }
});