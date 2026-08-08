/* ==========================================================================
 * Módulo 12 · CDP-CRP de OXP  (CRP + CDP con cruce por Número de CDP)
 * --------------------------------------------------------------------------
 * Pestañas CRP y CDP: lee reportes CRP y CDP en el NAVEGADOR con SheetJS,
 * filtra OXP (Com.Sin.Aut.Giro > 0), mapea a claves canónicas, y cruza
 * automáticamente por Número de CDP para extraer No.Interno CDP y No.Posición CDP.
 * ==========================================================================*/
(function () {
  "use strict";

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
    num_cdp:         ["Número de CDP"],
    num_crp:         ["Número de CRP"],
    rubro:           ["Rubro"],
    nuevo_rubro:     ["Nuevo Rubro"],
    elemento_pep:    ["Elemento PEP"],
    fondos:          ["Fondos", "Fondo"]
  };

  var MAPEO_CDP = {
    no_cdp:          ["No. CDP"],
    no_interno_cdp:  ["No.Interno CDP"],
    no_posicion_cdp: ["No.Posición CDP"]
  };

  var datosReporteCDP = {};

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

  function procesarReporteCDP(arrayBuffer) {
    var wb = window.XLSX.read(arrayBuffer, { type: "array" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    if (!aoa.length) throw new Error("El reporte CDP está vacío.");

    var idx = {};
    aoa[0].forEach(function (h, i) { if (idx[norm(h)] === undefined) idx[norm(h)] = i; });

    var colDe = {};
    Object.keys(MAPEO_CDP).forEach(function (clave) {
      for (var k = 0; k < MAPEO_CDP[clave].length; k++) {
        var i = idx[norm(MAPEO_CDP[clave][k])];
        if (i !== undefined) { colDe[clave] = i; break; }
      }
    });

    var mapa = {};
    for (var r = 1; r < aoa.length; r++) {
      var row = aoa[r];
      if (!row || row.every(function (c) { return c === null || c === ""; })) continue;
      var no_cdp = row[colDe.no_cdp];
      if (!no_cdp) continue;
      try {
        var num = parseInt(parseFloat(String(no_cdp).trim()));
        var clave = String(num).padStart(10, '0');
        mapa[clave] = {
          no_interno_cdp: row[colDe.no_interno_cdp] || "",
          no_posicion_cdp: row[colDe.no_posicion_cdp] || ""
        };
      } catch (e) {}
    }
    return mapa;
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
      document.getElementById(tab.dataset.panel).classList.add("activo");
    });
  });

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-oxp", "input-oxp", "lista-oxp");

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-reporte-cdp", "input-reporte-cdp", "lista-reporte-cdp");

  var estado_crp = { filas: [], total: 0, grupos: 0 };
  var input_crp = document.getElementById("input-oxp");
  var input_cdp = document.getElementById("input-reporte-cdp");
  var resumen_crp = document.getElementById("resumen-oxp");
  var resumen_cdp = document.getElementById("resumen-reporte-cdp");
  var btn_crp = document.getElementById("btn-oxp-crp");
  var metricas_crp = document.getElementById("metricas-oxp");

  if (input_crp) {
    input_crp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      mostrarAlerta("alerta-oxp", "");
      document.getElementById("resultados-oxp").classList.remove("visible");
      btn_crp.disabled = true;
      resumen_crp.textContent = "Leyendo reporte CRP…";
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
  }

  if (input_cdp) {
    input_cdp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      mostrarAlerta("alerta-reporte-cdp", "");
      resumen_cdp.textContent = "Leyendo reporte CDP…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          datosReporteCDP = procesarReporteCDP(ev.target.result);
          var numRegistros = Object.keys(datosReporteCDP).length;
          resumen_cdp.innerHTML =
            "Reporte CDP cargado: <b>" + numRegistros + "</b> registros. Se cruzarán por Número de CDP.";
        } catch (err) {
          resumen_cdp.textContent = "";
          mostrarAlerta("alerta-reporte-cdp", err.message);
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

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-oxp-cdp", "input-oxp-cdp", "lista-oxp-cdp");

  var estado_cdp = { filas: [], total: 0, grupos: 0 };
  var input_cdp_final = document.getElementById("input-oxp-cdp");
  var resumen_cdp_final = document.getElementById("resumen-oxp-cdp");
  var btn_cdp = document.getElementById("btn-oxp-cdp");
  var metricas_cdp = document.getElementById("metricas-oxp-cdp");

  if (input_cdp_final) {
    input_cdp_final.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      mostrarAlerta("alerta-oxp-cdp", "");
      document.getElementById("resultados-oxp-cdp").classList.remove("visible");
      btn_cdp.disabled = true;
      resumen_cdp_final.textContent = "Leyendo reporte…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          estado_cdp = procesarReporte(ev.target.result);
          resumen_cdp_final.innerHTML =
            "Filas en reporte: <b>" + estado_cdp.total + "</b> · " +
            "A constituir (saldo &gt; 0): <b>" + estado_cdp.filas.length + "</b>";
          var hay = estado_cdp.filas.length > 0;
          btn_cdp.disabled = !hay;
          if (!hay) mostrarAlerta("alerta-oxp-cdp", "No hay filas con saldo pendiente.");
        } catch (err) {
          resumen_cdp_final.textContent = "";
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

      var filasEnriquecidas = estado_cdp.filas.map(function (fila) {
        var filaComp = Object.assign({}, fila);
        var num_cdp_str = String(fila.num_cdp || "").trim();
        try {
          var num = parseInt(parseFloat(num_cdp_str));
          var clave = String(num).padStart(10, '0');
          if (datosReporteCDP[clave]) {
            var datos = datosReporteCDP[clave];
            filaComp.no_interno_cdp = datos.no_interno_cdp;
            filaComp.no_posicion_cdp = datos.no_posicion_cdp;
          }
        } catch (e) {}
        return filaComp;
      });

      fetch(API_BASE + "/api/oxp/cdp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: filasEnriquecidas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t || ("HTTP " + resp.status)); });
        var n = resp.headers.get("X-OXP-Registros") || estado_cdp.filas.length;
        return resp.blob().then(function (blob) { return { blob: blob, n: n }; });
      }).then(function (res) {
        descargarBlob(res.blob, "CDP_de_OXP_diligenciado.xlsx");
        metricas_cdp.innerHTML = "";
        metricas_cdp.appendChild(crearMetrica(estado_cdp.total, "Filas en reporte"));
        metricas_cdp.appendChild(crearMetrica(estado_cdp.filas.length, "Filas a constituir", "verde"));
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

  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
