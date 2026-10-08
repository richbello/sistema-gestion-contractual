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

/* ==========================================================================
   Base bancaria — la extracción ocurre AQUÍ, en el navegador.

   Los históricos de pago pesan decenas de MB y al abrirlos en el servidor
   ocupaban cientos de MB de RAM, lo que mataba al worker de Render (512 MB).
   El navegador tiene memoria de sobra, así que lee los Excel con SheetJS y
   solo manda al backend la lista final de cédula + cuenta: unos cientos de KB.
   ========================================================================== */

const SIPAGO_RE_CUENTA = /^\s*(\S+)\s+(AHORROS|CORRIENTE)\s+(.+)$/i;
const SIPAGO_TIPO_CTA  = { AHORROS: "02", CORRIENTE: "01" };

// Orden de prioridad para ubicar la cédula. 'documento' va de último a
// propósito: el histórico trae 'Nº documento' (consecutivo contable de SAP)
// mucho antes que 'Nº identificación', y buscar columna por columna hacía que
// ganara el consecutivo y la base quedara con llaves equivocadas.
const SIPAGO_ID_CLAVES = ["identificación", "identificacion", "nit", "cedula", "cédula", "documento"];
const SIPAGO_EXCLUIR   = ["clase", "tipo doc", "tipo de doc"];

function sipagoDigitos(valor) {
  if (valor === null || valor === undefined) return "";
  return String(valor).trim().replace(/\.0$/, "").replace(/[^\d]/g, "");
}

function sipagoIdxPorPrioridad(encabezados, claves) {
  for (const clave of claves) {
    for (let i = 0; i < encabezados.length; i++) {
      const h = encabezados[i];
      if (SIPAGO_EXCLUIR.some((x) => h.includes(x))) continue;
      if (h.includes(clave)) return i;
    }
  }
  return -1;
}

/* La columna de cuenta se detecta por CONTENIDO, no por nombre: el histórico
   trae varias columnas llamadas 'Cuenta bancaria' y solo una viene con el
   formato '<numero> AHORROS|CORRIENTE <banco>'. */
function sipagoDetectarColCuenta(filas, muestra = 300) {
  const conteo = new Map();
  const tope = Math.min(filas.length, muestra);
  for (let f = 0; f < tope; f++) {
    const fila = filas[f] || [];
    for (let i = 0; i < fila.length; i++) {
      const v = fila[i];
      if (v === null || v === undefined) continue;
      if (SIPAGO_RE_CUENTA.test(String(v))) {
        conteo.set(i, (conteo.get(i) || 0) + 1);
      }
    }
  }
  let mejor = -1, max = 0;
  conteo.forEach((n, i) => { if (n > max) { max = n; mejor = i; } });
  return mejor;
}

/* Extrae las cuentas de un archivo ya leído. Devuelve {registros, motivo}. */
function sipagoExtraerCuentas(arrayBuffer, orden) {
  const libro = XLSX.read(arrayBuffer, {
    type: "array",
    dense: true,       // menos objetos en memoria para hojas grandes
    cellDates: false,
    cellStyles: false,
    cellNF: false,
  });
  const hoja = libro.Sheets[libro.SheetNames[0]];
  if (!hoja) return { registros: [], motivo: "el archivo está vacío" };

  const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, raw: true, blankrows: false });
  if (!filas.length) return { registros: [], motivo: "el archivo está vacío" };

  const encabezados = (filas[0] || []).map((h) =>
    h === null || h === undefined ? "" : String(h).trim().toLowerCase());

  const iId = sipagoIdxPorPrioridad(encabezados, SIPAGO_ID_CLAVES);
  if (iId < 0) return { registros: [], motivo: "sin columna de identificación" };

  const iCod = sipagoIdxPorPrioridad(encabezados, ["codigo de banco", "código de banco"]);

  const cuerpo = filas.slice(1);
  const iCta = sipagoDetectarColCuenta(cuerpo);

  // Formato A: una sola columna "<numero> AHORROS|CORRIENTE <banco>"
  if (iCta >= 0) {
    const registros = [];
    for (const fila of cuerpo) {
      const bruto = fila[iCta];
      if (bruto === null || bruto === undefined) continue;
      const m = SIPAGO_RE_CUENTA.exec(String(bruto));
      if (!m) continue;

      const cedula = sipagoDigitos(fila[iId]);
      const cuenta = m[1].replace(/[^\d]/g, "");
      if (!cedula || !cuenta) continue;

      registros.push({
        cedula,
        No_Cuenta: cuenta,
        Tipo_Cta: SIPAGO_TIPO_CTA[m[2].toUpperCase()] || "",
        banco_txt: m[3].trim(),
        cod_banco: iCod >= 0 ? fila[iCod] : null,
        _orden: orden,
      });
    }
    return { registros, motivo: null };
  }

  // Formato B: la cuenta viene en columnas separadas (número, tipo, código, banco)
  const iNum = sipagoIdxPorPrioridad(encabezados, SIPAGO_NUM_CTA_CLAVES);
  if (iNum < 0) return { registros: [], motivo: "sin columna de cuenta bancaria" };
  const iTipo = sipagoIdxPorPrioridad(encabezados, SIPAGO_TIPO_CTA_CLAVES);
  const iCodB = sipagoIdxPorPrioridad(encabezados, SIPAGO_COD_BANCO_CLAVES);
  const iBco = sipagoIdxPorPrioridad(encabezados, SIPAGO_NOMBRE_BANCO_CLAVES);

  const registros = [];
  for (const fila of cuerpo) {
    const cuenta = sipagoSoloDigitosNum(fila[iNum]);
    const cedula = sipagoDigitos(fila[iId]);
    if (!cedula || !cuenta) continue;
    const banco = iBco >= 0 ? fila[iBco] : null;
    registros.push({
      cedula,
      No_Cuenta: cuenta,
      Tipo_Cta: iTipo >= 0 ? sipagoTipoCuentaB(fila[iTipo]) : "02",
      banco_txt: banco === null || banco === undefined ? "" : String(banco).trim(),
      cod_banco: iCodB >= 0 ? fila[iCodB] : null,
      _orden: orden,
    });
  }
  return { registros, motivo: null };
}

/* Formato B: claves y normalizadores (espejo de la versión del servidor). */
const SIPAGO_NUM_CTA_CLAVES = ["no cuenta", "nº cuenta", "n° cuenta", "numero de cuenta",
  "número de cuenta", "numero cuenta"];
const SIPAGO_TIPO_CTA_CLAVES = ["tipo cta", "tipo de cuenta", "tipo cuenta"];
const SIPAGO_COD_BANCO_CLAVES = ["codigo de banco", "código de banco", "codigo banco", "codigo", "código"];
const SIPAGO_NOMBRE_BANCO_CLAVES = ["nombre del banco", "nombre banco", "banco"];

function sipagoSoloDigitosNum(valor) {
  if (valor === null || valor === undefined) return "";
  return String(valor).trim().replace(/\.0$/, "").replace(/\D/g, "");
}

function sipagoTipoCuentaB(valor) {
  if (valor === null || valor === undefined) return "";
  const s = String(valor).trim().toUpperCase().replace(/\.0$/, "");
  if (s === "01" || s === "02") return s;
  if (SIPAGO_TIPO_CTA[s]) return SIPAGO_TIPO_CTA[s];
  return /^\d+$/.test(s) ? s.padStart(2, "0") : "";
}

/* Deduplica dejando la cuenta MÁS RECIENTE (la del archivo de mayor orden;
   dentro de un mismo archivo gana la primera aparición) y arma el mapa
   banco -> código con todas las filas, antes de perderlas en la deduplicación. */
function sipagoConsolidar(todos) {
  const votos = new Map();
  for (const r of todos) {
    if (r.cod_banco === null || r.cod_banco === undefined || String(r.cod_banco).trim() === "") continue;
    const cod = parseInt(Number(r.cod_banco), 10);
    if (!Number.isFinite(cod)) continue;
    if (!r.banco_txt) continue; // sin nombre de banco no hay de qué aprender el código
    if (!votos.has(r.banco_txt)) votos.set(r.banco_txt, new Map());
    const m = votos.get(r.banco_txt);
    m.set(cod, (m.get(cod) || 0) + 1);
  }
  const mapa_bancos = {};
  votos.forEach((m, banco) => {
    let mejor = null, max = -1;
    [...m.keys()].sort((a, b) => a - b).forEach((cod) => {
      if (m.get(cod) > max) { max = m.get(cod); mejor = cod; }
    });
    if (mejor !== null) mapa_bancos[banco] = mejor;
  });

  const elegidos = new Map();
  for (const r of todos) {
    const previo = elegidos.get(r.cedula);
    if (!previo || r._orden > previo._orden) elegidos.set(r.cedula, r);
  }

  const registros = [...elegidos.values()].map((r) => ({
    cedula: r.cedula,
    "No Cuenta": r.No_Cuenta,
    "Tipo Cta": r.Tipo_Cta,
    banco_txt: r.banco_txt,
    cod_banco: r.cod_banco,
    _orden: r._orden,
  }));

  return { registros, mapa_bancos };
}

/* ---------- Reemplazar base bancaria ---------- */
document.getElementById("form-sipago-bancaria").addEventListener("submit", async (e) => {
  e.preventDefault();
  mostrarAlerta("alerta-sipago-bases", "");

  const archivos = [...document.getElementById("input-sipago-historicos").files];
  if (!archivos.length) {
    mostrarAlerta("alerta-sipago-bases", "Debes adjuntar al menos un histórico de pagos.");
    return;
  }

  const btn = document.getElementById("btn-sipago-bancaria");
  const estado = document.getElementById("estado-sipago-bancaria");
  const textoEstado = (txt) => {
    if (!estado) return;
    const nodo = [...estado.childNodes].find((n) => n.nodeType === Node.TEXT_NODE);
    if (nodo) nodo.nodeValue = " " + txt;
  };

  btn.disabled = true;
  mostrarEstado("estado-sipago-bancaria", true);

  try {
    const todos = [];
    const omitidos = [];

    for (let i = 0; i < archivos.length; i++) {
      const archivo = archivos[i];
      textoEstado(`Leyendo ${archivo.name} (${i + 1} de ${archivos.length})…`);
      // Cede el hilo para que el navegador alcance a repintar el estado.
      await new Promise((r) => setTimeout(r, 0));

      try {
        const buffer = await archivo.arrayBuffer();
        const { registros, motivo } = sipagoExtraerCuentas(buffer, i);
        if (motivo) omitidos.push(`${archivo.name}: ${motivo}`);
        else todos.push(...registros);
      } catch (err) {
        omitidos.push(`${archivo.name}: no se pudo leer (${err.message})`);
      }
    }

    if (!todos.length) {
      mostrarAlerta("alerta-sipago-bases",
        "No se encontraron cuentas bancarias en los archivos."
        + (omitidos.length ? ` ${omitidos.join("; ")}` : ""));
      return;
    }

    const { registros, mapa_bancos } = sipagoConsolidar(todos);

    textoEstado(`Guardando ${registros.length} cuentas…`);
    const resp = await fetch(`${API_SIPAGO}/cargar-base-bancaria-json`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ registros, mapa_bancos, omitidos }),
    });
    const data = await resp.json();
    if (!resp.ok || !data.ok) {
      mostrarAlerta("alerta-sipago-bases", data.mensaje || "No se pudo cargar la base bancaria.");
      return;
    }
    if (data.omitidos && data.omitidos.length) {
      mostrarAlerta("alerta-sipago-bases", `Archivos omitidos: ${data.omitidos.join("; ")}`);
    }
    await sipagoRefrescarBases();
  } catch (err) {
    mostrarAlerta("alerta-sipago-bases", `No se pudo procesar el histórico: ${err.message}`);
  } finally {
    btn.disabled = false;
    mostrarEstado("estado-sipago-bancaria", false);
    textoEstado("Procesando…");
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