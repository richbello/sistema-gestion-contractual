/* ============================================================
   MÓDULO 06 — Análisis Presupuestal CRP
   Lee el Reporte CRP (SAP Distrital, hoja "Data") y muestra
   KPIs, gráficas y tabla con datos REALES.
   ============================================================ */
(function () {
  "use strict";

  let DATOS = [];
  const GRAF = {};

  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();

  function init() {
    const input = document.getElementById("carga-06");
    if (!input || input.dataset.ligado === "1") return;
    input.dataset.ligado = "1";
    input.addEventListener("change", (e) => leerExcel(e.target.files[0]));
  }

  /* ---------- lectura del archivo ---------- */
  function leerExcel(archivo) {
    if (!archivo) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { cellDates: true });
        // usa la hoja "Data" si existe, si no la primera
        const nombreHoja = wb.SheetNames.includes("Data") ? "Data" : wb.SheetNames[0];
        const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { defval: "" });
        DATOS = filas.filter((r) => r["Ejercicio"] || r["Valor CRP"]);
        if (DATOS.length === 0) {
          alert("El archivo no contiene datos reconocibles de CRP.");
          return;
        }
        render(archivo.name);
      } catch (err) {
        console.error(err);
        alert("No se pudo leer el archivo: " + err.message);
      }
    };
    reader.readAsArrayBuffer(archivo);
  }

  /* ---------- utilidades ---------- */
  const num = (x) => {
    const n = parseFloat(String(x).replace(/[^\d.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };
  const money = (v) =>
    new Intl.NumberFormat("es-CO", {
      style: "currency",
      currency: "COP",
      maximumFractionDigits: 0,
    }).format(v);
  const moneyCorto = (v) => {
    if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM";
    if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M";
    if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K";
    return "$" + v.toFixed(0);
  };

  /* ---------- render principal ---------- */
  function render(nombreArchivo) {
    const vista = document.getElementById("vista-presupuestal");
    if (!vista) return;

    // ocultar la pantalla de carga decorativa si sigue visible
    const cargaDecor = document.getElementById("pantalla-carga-crp");
    if (cargaDecor) cargaDecor.style.display = "none";

    // eliminar dashboard previo si el usuario recarga otro archivo
    const previo = document.getElementById("dash-06");
    if (previo) previo.remove();

    // totales
    let crp = 0, giro = 0, pend = 0, anul = 0, rein = 0;
    DATOS.forEach((r) => {
      crp += num(r["Valor CRP"]);
      giro += num(r["Autorizacion giro"]);
      pend += num(r["Com.Sin.Aut.Giro"]);
      anul += num(r["Anulaciones"]);
      rein += num(r["Reintegros"]);
    });
    const ejec = crp ? (giro / crp) * 100 : 0;

    // construir dashboard
    const dash = document.createElement("div");
    dash.id = "dash-06";
    dash.style.cssText = "padding:0 4px;";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:12px;padding:18px 22px;margin:18px 0;color:#fff;">
        <div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8fb8e8;">Reporte CRP · ${nombreArchivo}</div>
        <div style="font-size:22px;font-weight:700;margin-top:4px;">${DATOS.length.toLocaleString("es-CO")} registros analizados</div>
      </div>

      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:18px;">
        ${kpi("Valor CRP", money(crp), "#667eea")}
        ${kpi("Girado", money(giro), "#28a745")}
        ${kpi("Pendiente de giro", money(pend), "#ffc107")}
        ${kpi("Anulaciones", money(anul), "#dc3545")}
        ${kpi("% Ejecución", ejec.toFixed(1) + "%", "#17a2b8")}
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">CRP vs Girado por período</div>
          <canvas id="c06-periodo"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Ejecución del presupuesto</div>
          <canvas id="c06-ejec"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Distribución por modalidad</div>
          <canvas id="c06-modalidad"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Top 10 beneficiarios (Valor CRP)</div>
          <canvas id="c06-benef"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);grid-column:1/-1;">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Valor CRP por rubro (Top 10)</div>
          <canvas id="c06-rubro"></canvas>
        </div>
      </div>

      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <input id="c06-busca" placeholder="🔎 Buscar beneficiario, CRP, CDP o compromiso…"
          style="flex:1;padding:11px;border:1.5px solid #e8edf2;border-radius:8px;font-size:13px;">
        <button id="c06-btn-tabla" style="padding:11px 18px;background:#667eea;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;">Ver tabla</button>
      </div>
      <div id="c06-tabla-wrap" style="display:none;background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 10px rgba(0,0,0,.06);overflow-x:auto;"></div>
    `;

    // insertar justo después del panel de carga
    vista.appendChild(dash);

    // gráficas
    graficaPeriodo();
    graficaEjecucion(giro, pend, anul);
    graficaModalidad();
    graficaBeneficiarios();
    graficaRubro();

    // interacción tabla / búsqueda
    document.getElementById("c06-btn-tabla").addEventListener("click", () => {
      const w = document.getElementById("c06-tabla-wrap");
      w.style.display = w.style.display === "none" ? "block" : "none";
      if (w.style.display === "block") pintarTabla(DATOS.slice(0, 200));
    });
    document.getElementById("c06-busca").addEventListener("input", (e) => {
      const q = e.target.value.toLowerCase().trim();
      const w = document.getElementById("c06-tabla-wrap");
      w.style.display = "block";
      if (!q) return pintarTabla(DATOS.slice(0, 200));
      const filtr = DATOS.filter((r) =>
        [r["Nombre BP Beneficiario"], r["Número de CRP"], r["Número de CDP"], r["No. Compromiso"]]
          .some((v) => String(v).toLowerCase().includes(q))
      );
      pintarTabla(filtr.slice(0, 200));
    });

    dash.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function kpi(label, valor, color) {
    return `<div style="background:${color};color:#fff;padding:16px;border-radius:10px;text-align:center;">
      <div style="font-size:11px;opacity:.85;text-transform:uppercase;letter-spacing:.04em;">${label}</div>
      <div style="font-size:20px;font-weight:700;margin-top:4px;line-height:1.1;">${valor}</div>
    </div>`;
  }

  /* ---------- gráficas ---------- */
  function agrupar(campo, valor) {
    const m = {};
    DATOS.forEach((r) => {
      const k = String(r[campo] || "—");
      m[k] = (m[k] || 0) + (valor ? num(r[valor]) : 1);
    });
    return m;
  }

  function graficaPeriodo() {
    const crpP = agrupar("Período", "Valor CRP");
    const giroP = agrupar("Período", "Autorizacion giro");
    const periodos = Object.keys(crpP).sort((a, b) => num(a) - num(b));
    dibujar("c06-periodo", "bar", {
      labels: periodos,
      datasets: [
        { label: "CRP", data: periodos.map((p) => crpP[p]), backgroundColor: "#667eea" },
        { label: "Girado", data: periodos.map((p) => giroP[p] || 0), backgroundColor: "#28a745" },
      ],
    });
  }

  function graficaEjecucion(giro, pend, anul) {
    dibujar("c06-ejec", "doughnut", {
      labels: ["Girado", "Pendiente de giro", "Anulado"],
      datasets: [{ data: [giro, pend, anul], backgroundColor: ["#28a745", "#ffc107", "#dc3545"] }],
    });
  }

  function graficaModalidad() {
    const m = agrupar("Descripcion Mod. Selec", "Valor CRP");
    const orden = Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 6);
    dibujar("c06-modalidad", "pie", {
      labels: orden.map((x) => x[0]),
      datasets: [{ data: orden.map((x) => x[1]), backgroundColor: ["#667eea", "#28a745", "#ffc107", "#dc3545", "#17a2b8", "#8e44ad"] }],
    });
  }

  function graficaBeneficiarios() {
    const m = agrupar("Nombre BP Beneficiario", "Valor CRP");
    const top = Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10);
    dibujar("c06-benef", "bar", {
      labels: top.map((x) => (x[0].length > 28 ? x[0].slice(0, 28) + "…" : x[0])),
      datasets: [{ label: "Valor CRP", data: top.map((x) => x[1]), backgroundColor: "#667eea" }],
    }, { indexAxis: "y" });
  }

  function graficaRubro() {
    const m = agrupar("Descripción del Rubro", "Valor CRP");
    const top = Object.entries(m).sort((a, b) => b[1] - a[1]).slice(0, 10);
    dibujar("c06-rubro", "bar", {
      labels: top.map((x) => (x[0].length > 30 ? x[0].slice(0, 30) + "…" : x[0])),
      datasets: [{ label: "Valor CRP", data: top.map((x) => x[1]), backgroundColor: "#1e3a6e" }],
    });
  }

  function dibujar(id, tipo, data, extra) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    if (GRAF[id]) GRAF[id].destroy();
    GRAF[id] = new Chart(el, {
      type: tipo,
      data,
      options: Object.assign(
        {
          responsive: true,
          maintainAspectRatio: true,
          plugins: {
            legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
            tooltip: {
              callbacks: {
                label: (ctx) => {
                  const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed;
                  return ctx.dataset.label
                    ? `${ctx.dataset.label}: ${moneyCorto(v)}`
                    : moneyCorto(v);
                },
              },
            },
          },
          scales:
            tipo === "bar"
              ? { y: { ticks: { callback: (v) => moneyCorto(v) } } }
              : {},
        },
        extra || {}
      ),
    });
  }

  /* ---------- tabla ---------- */
  function pintarTabla(filas) {
    const w = document.getElementById("c06-tabla-wrap");
    if (!w) return;
    const cols = [
      ["Período", "Período"],
      ["Número de CRP", "CRP"],
      ["Número de CDP", "CDP"],
      ["Nombre BP Beneficiario", "Beneficiario"],
      ["Descripcion Mod. Selec", "Modalidad"],
      ["Valor CRP", "Valor CRP"],
      ["Autorizacion giro", "Girado"],
      ["Com.Sin.Aut.Giro", "Pendiente"],
    ];
    let html =
      '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr>';
    cols.forEach(
      (c) =>
        (html += `<th style="text-align:left;padding:8px;background:#1a2742;color:#fff;font-size:11px;position:sticky;top:0;">${c[1]}</th>`)
    );
    html += "</tr></thead><tbody>";
    filas.forEach((r, i) => {
      html += `<tr style="background:${i % 2 ? "#f7f9fc" : "#fff"};">`;
      cols.forEach(([campo]) => {
        let val = r[campo];
        if (["Valor CRP", "Autorizacion giro", "Com.Sin.Aut.Giro"].includes(campo))
          val = money(num(val));
        html += `<td style="padding:7px 8px;border-bottom:1px solid #eef2f7;">${val ?? ""}</td>`;
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    if (filas.length === 0) html = '<div style="padding:20px;color:#888;">Sin resultados.</div>';
    w.innerHTML = html;
  }

  console.log("✅ Módulo 06 (CRP real) listo");
})();