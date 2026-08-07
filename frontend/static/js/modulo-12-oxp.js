/* ==========================================================================
 * Módulo 12 · CDP-CRP de OXP  (submenú: CRP OXP)
 * --------------------------------------------------------------------------
 * Lee el reporte CRP en el NAVEGADOR con SheetJS (window.XLSX), filtra OXP
 * (Com.Sin.Aut.Giro > 0), mapea a claves canónicas (incluye N° Interno CRP y
 * N° Posición CRP para el agrupamiento) y envía solo las filas necesarias al
 * backend, que agrupa por CRP y devuelve el .xlsx diligenciado.
 *
 * Reutiliza helpers globales de main.js:
 *   API_BASE, configurarDropzoneUnico, mostrarAlerta, mostrarEstado, crearMetrica
 * ==========================================================================*/
(function () {
  "use strict";

  // Encabezado del reporte CRP -> clave canónica (insensible a may/tildes/espacios/punt.)
  var MAPEO_FUENTE = {
    importe:         ["Com.Sin.Aut.Giro"],
    objeto:          ["Objeto"],
    tipo_compromiso: ["Tipo de compromiso"],
    no_compromiso:   ["No. Compromiso"],
    modo_seleccion:  ["Modalidad de selección"],
    tipo_doc_benef:  ["Tipo Doc. BP Beneficiario"],
    id_benef:        ["Número Doc. BP Beneficiario"],
    id_solicitante:  ["ID Solicitante"],
    id_responsable:  ["ID Responsable"],
    interno_crp:     ["N° Interno CRP"],     // clave de agrupamiento
    pos_crp:         ["N° Posición CRP"],     // -> Posición (col B)
    num_crp:         ["Número de CRP"]        // -> Num. Ext. Entidad (col V)
  };

  function norm(t) {
    if (t === null || t === undefined) return "";
    return String(t).normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().replace(/[^a-z0-9]/g, "");
  }
  function aNumero(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    var n = parseFloat(String(v).replace(/[.,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  function procesarReporte(arrayBuffer) {
    var wb = window.XLSX.read(arrayBuffer, { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    if (!aoa.length) throw new Error("El reporte está vacío.");

    var idx = {};
    aoa[0].forEach(function (h, i) { if (idx[norm(h)] === undefined) idx[norm(h)] = i; });

    var colDe = {};
    Object.keys(MAPEO_FUENTE).forEach(function (clave) {
      for (var k = 0; k < MAPEO_FUENTE[clave].length; k++) {
        var i = idx[norm(MAPEO_FUENTE[clave][k])];
        if (i !== undefined) { colDe[clave] = i; break; }
      }
    });
    if (colDe.importe === undefined)
      throw new Error("No se encontró la columna 'Com.Sin.Aut.Giro' en el reporte.");

    var filas = [], total = 0, gruposSet = {};
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      if (!row || row.every(function (c) { return c === null || c === ""; })) continue;
      total++;
      if (aNumero(row[colDe.importe]) <= 0) continue;   // filtro OXP
      var obj = {};
      Object.keys(colDe).forEach(function (clave) {
        var v = row[colDe[clave]];
        if (v !== null && v !== undefined && v !== "") obj[clave] = v;
      });
      if (obj.interno_crp !== undefined) gruposSet[obj.interno_crp] = 1;
      filas.push(obj);
    }
    var grupos = Object.keys(gruposSet).length || filas.length;
    return { filas: filas, total: total, grupos: grupos };
  }

  function descargarBlob(blob, nombre) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
  }

  // --- Inicialización (los <script> cargan al final: el DOM ya existe) ----
  var input = document.getElementById("input-oxp");
  if (!input) return;

  // Pestañas del submenú (CRP / CDP)
  Array.prototype.forEach.call(document.querySelectorAll(".oxp-tab"), function (tab) {
    tab.addEventListener("click", function () {
      if (tab.disabled) return;
      document.querySelectorAll(".oxp-tab").forEach(function (t) { t.classList.remove("activo"); });
      document.querySelectorAll(".oxp-panel").forEach(function (p) { p.classList.remove("activo"); });
      tab.classList.add("activo");
      document.getElementById(tab.dataset.panel).classList.add("activo");
    });
  });

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-oxp", "input-oxp", "lista-oxp");

  var estado = { filas: [], total: 0, grupos: 0 };
  var resumen = document.getElementById("resumen-oxp");
  var btnCrp  = document.getElementById("btn-oxp-crp");
  var metricas = document.getElementById("metricas-oxp");

  input.addEventListener("change", function (e) {
    var file = e.target.files[0];
    if (!file) return;
    mostrarAlerta("alerta-oxp", "");
    document.getElementById("resultados-oxp").classList.remove("visible");
    btnCrp.disabled = true;
    resumen.textContent = "Leyendo reporte…";
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        estado = procesarReporte(ev.target.result);
        resumen.innerHTML =
          "Filas en reporte: <b>" + estado.total + "</b> · " +
          "A constituir (saldo &gt; 0): <b>" + estado.filas.length + "</b> · " +
          "CRP únicos: <b>" + estado.grupos + "</b>";
        var hay = estado.filas.length > 0;
        btnCrp.disabled = !hay;
        if (!hay) mostrarAlerta("alerta-oxp", "No hay filas con saldo pendiente (Com.Sin.Aut.Giro > 0).");
      } catch (err) {
        resumen.textContent = "";
        mostrarAlerta("alerta-oxp", err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  btnCrp.addEventListener("click", function () {
    mostrarAlerta("alerta-oxp", "");
    btnCrp.disabled = true;
    mostrarEstado("estado-oxp", true);
    fetch(API_BASE + "/api/oxp/crp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ filas: estado.filas })
    }).then(function (resp) {
      if (!resp.ok) return resp.text().then(function (t) { throw new Error(t || ("HTTP " + resp.status)); });
      var n = resp.headers.get("X-OXP-Registros") || estado.filas.length;
      return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
    }).then(function (res) {
      descargarBlob(res.blob, "CRP_de_OXP_diligenciado.xlsx");
      metricas.innerHTML = "";
      metricas.appendChild(crearMetrica(estado.total, "Filas en reporte"));
      metricas.appendChild(crearMetrica(estado.grupos, "CRP únicos", "verde"));
      metricas.appendChild(crearMetrica(res.n, "Filas escritas", "ocre"));
      document.getElementById("resultados-oxp").classList.add("visible");
    }).catch(function (err) {
      mostrarAlerta("alerta-oxp", "Error: " + err.message);
    }).finally(function () {
      btnCrp.disabled = false;
      mostrarEstado("estado-oxp", false);
    });
  });

  // Despierta el backend (Render duerme tras ~15 min)
  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
