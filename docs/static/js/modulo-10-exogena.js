/* ==========================================================================
 * Módulo 10 · Reportes Exógena  (Distrital Bogotá / DIAN Nacional)
 * --------------------------------------------------------------------------
 * Andamiaje: pestañas + carga de insumo con SheetJS en el navegador.
 * El generador se conecta cuando se defina la ficha técnica del formato.
 * ==========================================================================*/
(function () {
  "use strict";

  var input = document.getElementById("input-exo-d");
  if (!input) return;   // la vista no está en esta página

  // Pestañas Distrital / Nacional
  Array.prototype.forEach.call(document.querySelectorAll(".exo-tab"), function (tab) {
    tab.addEventListener("click", function () {
      document.querySelectorAll(".exo-tab").forEach(function (t) { t.classList.remove("activo"); });
      document.querySelectorAll(".exo-panel").forEach(function (p) { p.classList.remove("activo"); });
      tab.classList.add("activo");
      document.getElementById(tab.dataset.panel).classList.add("activo");
    });
  });

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-exo-d", "input-exo-d", "lista-exo-d");

  var resumen = document.getElementById("resumen-exo");
  var btn = document.getElementById("btn-exo-d");

  input.addEventListener("change", function (e) {
    var file = e.target.files[0];
    if (!file) return;
    resumen.textContent = "Leyendo insumo…";
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var wb = window.XLSX.read(ev.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var filas = aoa.length ? aoa.length - 1 : 0;
        var cols = aoa.length ? (aoa[0] || []).filter(function (c) { return c !== null && c !== ""; }).length : 0;
        resumen.innerHTML = "Insumo leído: <b>" + filas + "</b> filas · <b>" + cols + "</b> columnas. " +
          "(Generador pendiente de ficha técnica.)";
        // btn.disabled = false;  // se habilita al conectar el generador
      } catch (err) {
        resumen.textContent = "No se pudo leer el archivo: " + err.message;
      }
    };
    reader.readAsArrayBuffer(file);
  });
})();
