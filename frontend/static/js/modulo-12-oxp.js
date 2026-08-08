(function () {
  "use strict";
  
  if (typeof configurarDropzoneUnico === "function") {
    configurarDropzoneUnico("drop-oxp", "input-oxp", "lista-oxp");
    configurarDropzoneUnico("drop-reporte-cdp", "input-reporte-cdp", "lista-reporte-cdp");
  }

  var estado_crp = { filas: [], total: 0, grupos: 0 };
  var datosReporteCDP = {};
  
  var input_crp = document.getElementById("input-oxp");
  var input_cdp = document.getElementById("input-reporte-cdp");
  var btn_crp = document.getElementById("btn-oxp-crp");
  var resumen_crp = document.getElementById("resumen-oxp");
  var resumen_cdp = document.getElementById("resumen-reporte-cdp");

  function aNumero(v) {
    if (!v) return 0;
    var n = parseFloat(String(v).replace(/[.,\s]/g, ""));
    return isNaN(n) ? 0 : n;
  }

  if (input_crp) {
    input_crp.addEventListener("change", function (e) {
      var file = e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var wb = XLSX.read(ev.target.result, { type: "array" });
          var ws = wb.Sheets[wb.SheetNames[0]];
          var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
          
          var filas = [];
          for (var r = 1; r < aoa.length; r++) {
            var row = aoa[r];
            if (!row || aNumero(row[40]) <= 0) continue; // AO = columna 41
            filas.push(row);
          }
          
          estado_crp = { filas: filas, total: aoa.length - 1, grupos: 0 };
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
          var wb = XLSX.read(ev.target.result, { type: "array" });
          var ws = wb.Sheets[wb.SheetNames[0]];
          var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
          
          datosReporteCDP = {};
          for (var r = 1; r < aoa.length; r++) {
            var row = aoa[r];
            if (!row) continue;
            var no_cdp = row[7]; // Columna H (No. CDP)
            if (!no_cdp) continue;
            try {
              var num = parseInt(parseFloat(String(no_cdp).trim()));
              var clave = String(num).padStart(10, '0');
              datosReporteCDP[clave] = {
                no_interno: row[20], // Columna U
                no_posicion: row[21]  // Columna V
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
      
      var filasEnriquecidas = estado_crp.filas.map(function (row) {
        var num_cdp_str = String(row[13] || "").trim(); // Columna N
        try {
          var num = parseInt(parseFloat(num_cdp_str));
          var clave = String(num).padStart(10, '0');
          if (datosReporteCDP[clave]) {
            row[40] = datosReporteCDP[clave].no_interno;  // Guardar en algún lugar temporal
            row[41] = datosReporteCDP[clave].no_posicion;
          }
        } catch (e) {}
        return row;
      });
      
      fetch(API_BASE + "/api/oxp/crp", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filas: filasEnriquecidas })
      }).then(function (r) { return r.blob(); })
        .then(function (blob) {
          var a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = "CRP_de_OXP.xlsx";
          a.click();
          btn_crp.disabled = false;
        })
        .catch(function (err) {
          alert("Error: " + err.message);
          btn_crp.disabled = false;
        });
    });
  }

  if (API_BASE) fetch(API_BASE + "/api/salud").catch(function () {});
})();
