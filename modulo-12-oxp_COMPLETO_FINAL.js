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
    fondos:          ["Fondos", "Fondo"]
  };

  var MAPEO_CDP = {
    no_cdp:          ["No. CDP"],
    no_interno_cdp:  ["No.Interno CDP"],
    no_posicion_cdp: ["No.Posición CDP"]
  };

  var datosReporteCDP = {};

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

  function mapearFilas(aoa, mapeo) {
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

  if (typeof configurarDropzoneUnico === "function") {
    configurarDropzoneUnico("drop-oxp", "input-oxp", "lista-oxp");
    configurarDropzoneUnico("drop-reporte-cdp", "input-reporte-cdp", "lista-reporte-cdp");
  }

  var input_crp = document.getElementById("input-oxp");
  var input_cdp = document.getElementById("input-reporte-cdp");
  var btn_crp = document.getElementById("btn-oxp-crp");
  var resumen_crp = document.getElementById("resumen-oxp");
  var resumen_cdp = document.getElementById("resumen-reporte-cdp");
  var metricas_crp = document.getElementById("metricas-oxp");

  var estado_crp = { filas: [], total: 0 };

  if (input_crp) {
    input_crp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var wb = window.XLSX.read(ev.target.result, { type: "array" });
          var ws = wb.Sheets[wb.SheetNames[0]];
          var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
          var colDe = mapearFilas(aoa, MAPEO_CRP);

          var filas = [];
          for (var r = 1; r < aoa.length; r++) {
            var row = aoa[r];
            if (!row || aNumero(row[colDe.importe]) <= 0) continue;
            var obj = {};
            Object.keys(colDe).forEach(function (clave) {
              var v = row[colDe[clave]];
              if (v !== null && v !== undefined && v !== "") obj[clave] = v;
            });
            filas.push(obj);
          }

          estado_crp = { filas: filas, total: aoa.length - 1 };
          resumen_crp.innerHTML = "Filas: <b>" + estado_crp.total + "</b> · A constituir: <b>" + filas.length + "</b>";
          btn_crp.disabled = filas.length === 0;
        } catch (err) {
          resumen_crp.textContent = "Error: " + err.message;
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  if (input_cdp) {
    input_cdp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var wb = window.XLSX.read(ev.target.result, { type: "array" });
          var ws = wb.Sheets[wb.SheetNames[0]];
          var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
          var colDe = mapearFilas(aoa, MAPEO_CDP);

          datosReporteCDP = {};
          for (var r = 1; r < aoa.length; r++) {
            var row = aoa[r];
            if (!row) continue;
            var no_cdp = row[colDe.no_cdp];
            if (!no_cdp) continue;
            try {
              var num = parseInt(parseFloat(String(no_cdp).trim()));
              var clave = String(num).padStart(10, '0');
              datosReporteCDP[clave] = {
                no_interno_cdp: row[colDe.no_interno_cdp] || "",
                no_posicion_cdp: row[colDe.no_posicion_cdp] || ""
              };
            } catch (e) {}
          }

          resumen_cdp.innerHTML = "CDP cargado: <b>" + Object.keys(datosReporteCDP).length + "</b> registros";
        } catch (err) {
          resumen_cdp.textContent = "Error: " + err.message;
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  if (btn_crp) {
    btn_crp.addEventListener("click", function () {
      btn_crp.disabled = true;

      var filasEnriquecidas = estado_crp.filas.map(function (fila) {
        var filaComp = Object.assign({}, fila);
        var num_cdp_str = String(fila.num_cdp || "").trim();
        try {
          var num = parseInt(parseFloat(num_cdp_str));
          var clave = String(num).padStart(10, '0');
          if (datosReporteCDP[clave]) {
            filaComp.no_interno_cdp = datosReporteCDP[clave].no_interno_cdp;
            filaComp.no_posicion_cdp = datosReporteCDP[clave].no_posicion_cdp;
          }
        } catch (e) {}
        return filaComp;
      });

      fetch(API_BASE + "/api/oxp/crp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: filasEnriquecidas })
      }).then(function (resp) {
        if (!resp.ok) return resp.text().then(function (t) { throw new Error(t); });
        return resp.blob();
      }).then(function (blob) {
        var a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "CRP_de_OXP_diligenciado.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(a.href);
        if (metricas_crp) {
          metricas_crp.innerHTML = "";
          if (typeof crearMetrica === "function") {
            metricas_crp.appendChild(crearMetrica(estado_crp.total, "Filas en reporte"));
            metricas_crp.appendChild(crearMetrica(estado_crp.filas.length, "A constituir", "verde"));
          }
        }
        var res = document.getElementById("resultados-oxp");
        if (res) res.classList.add("visible");
        btn_crp.disabled = false;
      }).catch(function (err) {
        alert("Error: " + err.message);
        btn_crp.disabled = false;
      });
    });
  }

  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
