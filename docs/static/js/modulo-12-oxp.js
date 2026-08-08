/* ==========================================================================
 * Módulo 12 · CDP-CRP de OXP  (CRP + CDP con conversión de rubros)
 * --------------------------------------------------------------------------
 * Pestañas CRP y CDP: lee el reporte CRP en el NAVEGADOR con SheetJS,
 * filtra OXP (Com.Sin.Aut.Giro > 0), mapea a claves canónicas (incluye
 * N° Interno CRP, N° Posición CRP y Rubro para conversión). El backend
 * agrupa por CRP y aplica la conversión de rubros en CDP.
 *
 * Reutiliza helpers globales de main.js:
 *   API_BASE, configurarDropzoneUnico, mostrarAlerta, mostrarEstado, crearMetrica
 * ==========================================================================*/
(function () {
  "use strict";

  // Encabezado del reporte CRP -> clave canónica
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
    interno_crp:     ["N° Interno CRP"],
    pos_crp:         ["N° Posición CRP"],
    num_cdp:         ["Número de CDP"],      // -> col A del CDP y Objeto
    num_crp:         ["Número de CRP"],      // -> Objeto
    rubro:           ["Rubro"],              // -> Posición Presupuestal (fallback)
    nuevo_rubro:     ["Nuevo Rubro"],        // -> Posición Presupuestal (si lleno)
    elemento_pep:    ["Elemento PEP"],       // -> determina Posición Presupuestal
    fondos:          ["Fondos", "Fondo"]
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
      if (aNumero(row[colDe.importe]) <= 0) continue;
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

  // ========== PESTAÑAS ==========
  Array.prototype.forEach.call(document.querySelectorAll(".oxp-tab"), function (tab) {
    tab.addEventListener("click", function () {
      if (tab.disabled) return;
      document.querySelectorAll(".oxp-tab").forEach(function (t) { t.classList.remove("activo"); });
      document.querySelectorAll(".oxp-panel").forEach(function (p) { p.classList.remove("activo"); });
      tab.classList.add("activo");
      document.getElementById(tab.dataset.panel).classList.add("activo");
    });
  });

  // ========== CRP OXP ==========
  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-oxp", "input-oxp", "lista-oxp");

  var estado_crp = { filas: [], total: 0, grupos: 0 };
  var input_crp = document.getElementById("input-oxp");
  var resumen_crp = document.getElementById("resumen-oxp");
  var btn_crp = document.getElementById("btn-oxp-crp");
  var metricas_crp = document.getElementById("metricas-oxp");

  if (input_crp) {
    input_crp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      mostrarAlerta("alerta-oxp", "");
      document.getElementById("resultados-oxp").classList.remove("visible");
      btn_crp.disabled = true;
      resumen_crp.textContent = "Leyendo reporte…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          estado_crp = procesarReporte(ev.target.result);
          resumen_crp.innerHTML =
            "Filas en reporte: <b>" + estado_crp.total + "</b> · " +
            "A constituir (saldo &gt; 0): <b>" + estado_crp.filas.length + "</b> · " +
            "CRP únicos: <b>" + estado_crp.grupos + "</b>";
          var hay = estado_crp.filas.length > 0;
          btn_crp.disabled = !hay;
          if (!hay) mostrarAlerta("alerta-oxp", "No hay filas con saldo pendiente.");
        } catch (err) {
          resumen_crp.textContent = "";
          mostrarAlerta("alerta-oxp", err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    });

    btn_crp.addEventListener("click", function () {
      mostrarAlerta("alerta-oxp", "");
      btn_crp.disabled = true;
      mostrarEstado("estado-oxp", true);
      fetch(API_BASE + "/api/oxp/crp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: estado_crp.filas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t || ("HTTP " + resp.status)); });
        var n = resp.headers.get("X-OXP-Registros") || estado_crp.filas.length;
        return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
      }).then(function (res) {
        descargarBlob(res.blob, "CRP_de_OXP_diligenciado.xlsx");
        metricas_crp.innerHTML = "";
        metricas_crp.appendChild(crearMetrica(estado_crp.total, "Filas en reporte"));
        metricas_crp.appendChild(crearMetrica(estado_crp.grupos, "CRP únicos", "verde"));
        metricas_crp.appendChild(crearMetrica(res.n, "Filas escritas", "ocre"));
        document.getElementById("resultados-oxp").classList.add("visible");
      }).catch(function (err) {
        mostrarAlerta("alerta-oxp", "Error: " + err.message);
      }).finally(function () {
        btn_crp.disabled = false;
        mostrarEstado("estado-oxp", false);
      });
    });
  }

  // ========== CDP OXP ==========
  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-oxp-cdp", "input-oxp-cdp", "lista-oxp-cdp");

  var estado_cdp = { filas: [], total: 0, grupos: 0 };
  var input_cdp = document.getElementById("input-oxp-cdp");
  var resumen_cdp = document.getElementById("resumen-oxp-cdp");
  var btn_cdp = document.getElementById("btn-oxp-cdp");
  var metricas_cdp = document.getElementById("metricas-oxp-cdp");

  if (input_cdp) {
    input_cdp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      mostrarAlerta("alerta-oxp-cdp", "");
      document.getElementById("resultados-oxp-cdp").classList.remove("visible");
      btn_cdp.disabled = true;
      resumen_cdp.textContent = "Leyendo reporte…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          estado_cdp = procesarReporte(ev.target.result);
          resumen_cdp.innerHTML =
            "Filas en reporte: <b>" + estado_cdp.total + "</b> · " +
            "A constituir (saldo &gt; 0): <b>" + estado_cdp.filas.length + "</b> · " +
            "CRP únicos: <b>" + estado_cdp.grupos + "</b><br><em style='color:#7c3a3a;'>Los rubros se convertirán según la tabla (ej. O230689 → O230690).</em>";
          var hay = estado_cdp.filas.length > 0;
          btn_cdp.disabled = !hay;
          if (!hay) mostrarAlerta("alerta-oxp-cdp", "No hay filas con saldo pendiente.");
        } catch (err) {
          resumen_cdp.textContent = "";
          mostrarAlerta("alerta-oxp-cdp", err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    });

    btn_cdp.addEventListener("click", function () {
      mostrarAlerta("alerta-oxp-cdp", "");
      btn_cdp.disabled = true;
      mostrarEstado("estado-oxp-cdp", true);
      fetch(API_BASE + "/api/oxp/cdp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: estado_cdp.filas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t || ("HTTP " + resp.status)); });
        var n = resp.headers.get("X-OXP-Registros") || estado_cdp.filas.length;
        return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
      }).then(function (res) {
        descargarBlob(res.blob, "CDP_de_OXP_diligenciado.xlsx");
        metricas_cdp.innerHTML = "";
        metricas_cdp.appendChild(crearMetrica(estado_cdp.total, "Filas en reporte"));
        metricas_cdp.appendChild(crearMetrica(estado_cdp.grupos, "CRP únicos", "verde"));
        metricas_cdp.appendChild(crearMetrica(res.n, "Filas escritas", "ocre"));
        document.getElementById("resultados-oxp-cdp").classList.add("visible");
      }).catch(function (err) {
        mostrarAlerta("alerta-oxp-cdp", "Error: " + err.message);
      }).finally(function () {
        btn_cdp.disabled = false;
        mostrarEstado("estado-oxp-cdp", false);
      });
    });
  }

  // Despierta el backend
  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
