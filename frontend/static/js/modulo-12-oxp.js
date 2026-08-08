(function () {
  "use strict";

  var MAPEO_CRP = {
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
    num_cdp:         ["Número de CDP"],
    num_crp:         ["Número de CRP"],
    rubro:           ["Rubro"],
    nuevo_rubro:     ["Nuevo Rubro"],
    elemento_pep:    ["Elemento PEP"],
    fondos:          ["Fondos", "Fondo"],
    no_interno_cdp:  ["N° Interno CDP"],
    no_posicion_cdp: ["N° Posición CDP"]
  };

  function norm(t) {
    if (!t) return "";
    return String(t).normalize("NFKD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function aNumero(v) {
    if (!v) return 0;
    var n = parseFloat(String(v).replace(/[.,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  function mapearColumnas(aoa, mapeo) {
    var idx = {};
    aoa[0].forEach(function (h, i) {
      if (idx[norm(h)] === undefined) idx[norm(h)] = i;
    });
    var colDe = {};
    Object.keys(mapeo).forEach(function (clave) {
      for (var k = 0; k < mapeo[clave].length; k++) {
        var i = idx[norm(mapeo[clave][k])];
        if (i !== undefined) { colDe[clave] = i; break; }
      }
    });
    return colDe;
  }

  function procesarReporte(arrayBuffer) {
    var wb = window.XLSX.read(arrayBuffer, { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
    if (!aoa.length) throw new Error("El reporte está vacío.");

    var colDe = mapearColumnas(aoa, MAPEO_CRP);
    if (colDe.importe === undefined)
      throw new Error("No se encontró 'Com.Sin.Aut.Giro'.");

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

  Array.prototype.forEach.call(document.querySelectorAll(".oxp-tab"), function (tab) {
    tab.addEventListener("click", function () {
      if (tab.disabled) return;
      document.querySelectorAll(".oxp-tab").forEach(function (t) { t.classList.remove("activo"); });
      document.querySelectorAll(".oxp-panel").forEach(function (p) { p.classList.remove("activo"); });
      tab.classList.add("activo");
      var panel = document.getElementById(tab.dataset.panel);
      if (panel) panel.classList.add("activo");
    });
  });

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
      btn_crp.disabled = true;
      resumen_crp.textContent = "Leyendo reporte CRP…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          estado_crp = procesarReporte(ev.target.result);
          resumen_crp.innerHTML =
            "Filas: <b>" + estado_crp.total + "</b> · A constituir: <b>" + estado_crp.filas.length + "</b> · CRP únicos: <b>" + estado_crp.grupos + "</b>";
          btn_crp.disabled = estado_crp.filas.length === 0;
          if (estado_crp.filas.length === 0) mostrarAlerta("alerta-oxp", "No hay filas con saldo pendiente.");
        } catch (err) {
          resumen_crp.textContent = "";
          mostrarAlerta("alerta-oxp", err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  if (btn_crp) {
    btn_crp.addEventListener("click", function () {
      mostrarAlerta("alerta-oxp", "");
      btn_crp.disabled = true;
      mostrarEstado("estado-oxp", true);
      fetch(API_BASE + "/api/oxp/crp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: estado_crp.filas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t); });
        var n = resp.headers.get("X-OXP-Registros") || estado_crp.filas.length;
        return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
      }).then(function (res) {
        descargarBlob(res.blob, "CRP_de_OXP_diligenciado.xlsx");
        if (metricas_crp) {
          metricas_crp.innerHTML = "";
          if (typeof crearMetrica === "function") {
            metricas_crp.appendChild(crearMetrica(estado_crp.total, "Filas en reporte"));
            metricas_crp.appendChild(crearMetrica(estado_crp.grupos, "CRP únicos", "verde"));
            metricas_crp.appendChild(crearMetrica(res.n, "Filas escritas", "ocre"));
          }
        }
        var resDiv = document.getElementById("resultados-oxp");
        if (resDiv) resDiv.classList.add("visible");
      }).catch(function (err) {
        mostrarAlerta("alerta-oxp", "Error: " + err.message);
      }).finally(function () {
        btn_crp.disabled = false;
        mostrarEstado("estado-oxp", false);
      });
    });
  }

  // ===== CDP OXP =====
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
      btn_cdp.disabled = true;
      resumen_cdp.textContent = "Leyendo reporte…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          estado_cdp = procesarReporte(ev.target.result);
          resumen_cdp.innerHTML =
            "Filas: <b>" + estado_cdp.total + "</b> · A constituir: <b>" + estado_cdp.filas.length + "</b>";
          btn_cdp.disabled = estado_cdp.filas.length === 0;
          if (estado_cdp.filas.length === 0) mostrarAlerta("alerta-oxp-cdp", "No hay filas con saldo pendiente.");
        } catch (err) {
          resumen_cdp.textContent = "";
          mostrarAlerta("alerta-oxp-cdp", err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  if (btn_cdp) {
    btn_cdp.addEventListener("click", function () {
      mostrarAlerta("alerta-oxp-cdp", "");
      btn_cdp.disabled = true;
      mostrarEstado("estado-oxp-cdp", true);
      fetch(API_BASE + "/api/oxp/cdp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: estado_cdp.filas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t); });
        var n = resp.headers.get("X-OXP-Registros") || estado_cdp.filas.length;
        return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
      }).then(function (res) {
        descargarBlob(res.blob, "CDP_de_OXP_diligenciado.xlsx");
        if (metricas_cdp) {
          metricas_cdp.innerHTML = "";
          if (typeof crearMetrica === "function") {
            metricas_cdp.appendChild(crearMetrica(estado_cdp.total, "Filas en reporte"));
            metricas_cdp.appendChild(crearMetrica(estado_cdp.filas.length, "A constituir", "verde"));
            metricas_cdp.appendChild(crearMetrica(res.n, "Filas escritas", "ocre"));
          }
        }
        var resDiv = document.getElementById("resultados-oxp-cdp");
        if (resDiv) resDiv.classList.add("visible");
      }).catch(function (err) {
        mostrarAlerta("alerta-oxp-cdp", "Error: " + err.message);
      }).finally(function () {
        btn_cdp.disabled = false;
        mostrarEstado("estado-oxp-cdp", false);
      });
    });
  }

  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
