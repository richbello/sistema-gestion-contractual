(function () {
  "use strict";
  let SECTORES = [], PROYECTOS = [], TOTALS = {}, fileName = "";
  const GRAF = {};
  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();
  function init() {
    const input = document.getElementById("carga-08");
    if (!input || input.dataset.ligado === "1") return;
    input.dataset.ligado = "1";
    input.addEventListener("change", (e) => leerExcel(e.target.files[0]));
  }
  function leerExcel(archivo) {
    if (!archivo) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const names = wb.SheetNames;
        const ejN = names.find(n => n.toLowerCase().includes("ejec")) || names[0];
        const ejRaw = XLSX.utils.sheet_to_json(wb.Sheets[ejN], { header: 1 });
        let ejRows = [];
        if (ejRaw.length > 1) {
          const heads = ejRaw[0].map(h => String(h || "").trim());
          for (let i = 1; i < ejRaw.length; i++) {
            const row = ejRaw[i];
            if (!row.some(c => c != null && c !== "")) continue;
            const o = {}; heads.forEach((h, j) => o[h] = row[j] ?? ""); ejRows.push(o);
          }
        }
        procesar(ejRows);
        fileName = archivo.name;
        if (PROYECTOS.length === 0) { alert("Sin datos de proyectos"); return; }
        render(archivo.name);
      } catch (err) { console.error(err); alert("Error: " + err.message); }
    };
    reader.readAsArrayBuffer(archivo);
  }
  const num = (x) => { const n = parseFloat(String(x).replace(/[^\d.-]/g, "")); return isNaN(n) ? 0 : n; };
  const money = (v) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);
  const moneyCorto = (v) => { if (Math.abs(v) >= 1e12) return "$" + (v / 1e12).toFixed(1) + "B"; if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM"; if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M"; if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K"; return "$" + v.toFixed(0); };
  const gcP = (row, keys) => { for (const k of keys) if (row[k] != null && row[k] !== "") return row[k]; return ""; };
  const pct = (v) => (v * 100).toFixed(1) + "%";
  function procesar(ejec) {
    const proyMap = {};
    ejec.forEach(row => {
      const cod = num(gcP(row, ["Código_Proyecto", "Codigo_Proyecto", "Cód. Proyecto"]));
      if (!cod) return;
      const poai = num(gcP(row, ["POAI 2026", "Recursos POAI 2026", "POAI"]));
      const aprop = num(gcP(row, ["Apropiaciones", "apropiaciones"]));
      const comp = num(gcP(row, ["Compromisos proyecto", "Compromisos", "compromisos"]));
      const giros = num(gcP(row, ["Giros proyecto", "Giros", "giros"]));
      if (poai === 0) return;
      proyMap[cod] = { codigo: cod, poai, apropiaciones: aprop, compromisos: comp, giros: giros, pctComp: aprop > 0 ? comp / aprop : 0, sector: "SIN SECTOR" };
    });
    PROYECTOS = Object.values(proyMap).filter(p => p.poai > 0).sort((a, b) => b.poai - a.poai);
    const secMap = { "SIN SECTOR": { nombre: "SIN SECTOR", poai: 0, apropiaciones: 0, compromisos: 0, giros: 0, proyCount: 0 } };
    PROYECTOS.forEach(p => {
      if (!secMap[p.sector]) secMap[p.sector] = { nombre: p.sector, poai: 0, apropiaciones: 0, compromisos: 0, giros: 0, proyCount: 0 };
      secMap[p.sector].poai += p.poai;
      secMap[p.sector].apropiaciones += p.apropiaciones;
      secMap[p.sector].compromisos += p.compromisos;
      secMap[p.sector].giros += p.giros;
      secMap[p.sector].proyCount += 1;
    });
    Object.values(secMap).forEach(s => { s.pctComp = s.apropiaciones > 0 ? s.compromisos / s.apropiaciones : 0; });
    SECTORES = Object.values(secMap).sort((a, b) => b.poai - a.poai);
    TOTALS = { poai: PROYECTOS.reduce((s, p) => s + p.poai, 0), apropiaciones: PROYECTOS.reduce((s, p) => s + p.apropiaciones, 0), compromisos: PROYECTOS.reduce((s, p) => s + p.compromisos, 0), giros: PROYECTOS.reduce((s, p) => s + p.giros, 0) };
  }
  function render(nombreArchivo) {
    const vista = document.getElementById("vista-planeacion");
    if (!vista) return;
    const contenedor = vista.querySelector("div[style*='max-width']") || vista;
    const cargaDecor = document.getElementById("pantalla-carga-08");
    if (cargaDecor) cargaDecor.style.display = "none";
    const previo = document.getElementById("dash-08");
    if (previo) previo.remove();
    const pctComp = TOTALS.apropiaciones ? (TOTALS.compromisos / TOTALS.apropiaciones) * 100 : 0;
    const dash = document.createElement("div");
    dash.id = "dash-08";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:12px;padding:18px 22px;margin:18px 0;color:#fff;">
        <div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8fb8e8;">Plan de Desarrollo Local · ${nombreArchivo}</div>
        <div style="font-size:22px;font-weight:700;margin-top:4px;">${PROYECTOS.length.toLocaleString("es-CO")} proyectos</div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:20px;">
        ${kpi("POAI 2026", moneyCorto(TOTALS.poai), "#667eea", "💰")}
        ${kpi("Apropiaciones", moneyCorto(TOTALS.apropiaciones), "#17a2b8", "📌")}
        ${kpi("Compromisos", moneyCorto(TOTALS.compromisos), "#28a745", "✅")}
        ${kpi("Giros", moneyCorto(TOTALS.giros), "#fd7e14", "💸")}
        ${kpi("% Ejecución", pctComp.toFixed(1) + "%", "#8e44ad", "📊")}
      </div>
      <div style="background:#fff;padding:20px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:16px;">
        <div style="font-size:14px;font-weight:700;margin-bottom:12px;color:#1a2742;">🎯 POAI 2026</div>
        <canvas id="c08-poai" style="max-height:350px;"></canvas>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Cascada</div><canvas id="c08-cascada"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Ejecución</div><canvas id="c08-ejec"></canvas></div>
      </div>
      <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:16px;">
        <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Top 15 Proyectos</div>
        <canvas id="c08-top" style="max-height:300px;"></canvas>
      </div>
      <button id="c08-btn-tabla" style="padding:11px 18px;background:#667eea;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;margin-bottom:15px;">Ver tabla completa</button>
      <div id="c08-tabla-wrap" style="display:none;background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 10px rgba(0,0,0,.06);overflow-x:auto;"></div>
    `;
    contenedor.appendChild(dash);
    setTimeout(() => {
      graficaPOAI(); graficaCascada(); graficaEjecucion(); graficaTop();
    }, 100);
    document.getElementById("c08-btn-tabla").addEventListener("click", () => { const w = document.getElementById("c08-tabla-wrap"); w.style.display = w.style.display === "none" ? "block" : "none"; if (w.style.display === "block") pintarTabla(PROYECTOS); });
    dash.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function kpi(label, valor, color, icon) { return `<div style="background:${color};color:#fff;padding:16px;border-radius:10px;text-align:center;"><div style="font-size:20px;margin-bottom:4px;">${icon}</div><div style="font-size:11px;opacity:.85;text-transform:uppercase;letter-spacing:.04em;">${label}</div><div style="font-size:18px;font-weight:700;margin-top:4px;">${valor}</div></div>`; }
  function graficaPOAI() { dibujar("c08-poai", "bar", { labels: PROYECTOS.slice(0, 15).map(p => `${p.codigo}`), datasets: [{ label: "POAI 2026", data: PROYECTOS.slice(0, 15).map(p => p.poai), backgroundColor: "#667eea" }] }); }
  function graficaCascada() { dibujar("c08-cascada", "bar", { labels: ["POAI", "Aprop", "Comp", "Giros"], datasets: [{ label: "Valor", data: [TOTALS.poai, TOTALS.apropiaciones, TOTALS.compromisos, TOTALS.giros], backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14"] }] }, { plugins: { legend: { display: false } } }); }
  function graficaEjecucion() { const noComp = Math.max(0, TOTALS.apropiaciones - TOTALS.compromisos); const compSinGiro = Math.max(0, TOTALS.compromisos - TOTALS.giros); dibujar("c08-ejec", "doughnut", { labels: ["Girado", "Comprometido sin girar", "Sin comprometer"], datasets: [{ data: [TOTALS.giros, compSinGiro, noComp], backgroundColor: ["#28a745", "#ffc107", "#dc3545"] }] }); }
  function graficaTop() { const top = PROYECTOS.slice(0, 15); dibujar("c08-top", "bar", { labels: top.map(p => `Proy ${p.codigo}`), datasets: [{ label: "POAI", data: top.map(p => p.poai), backgroundColor: "#667eea" }, { label: "Comp", data: top.map(p => p.compromisos), backgroundColor: "#28a745" }, { label: "Giros", data: top.map(p => p.giros), backgroundColor: "#fd7e14" }] }); }
  function dibujar(id, tipo, data, extra) { const el = document.getElementById(id); if (!el || typeof Chart === "undefined") return; if (GRAF[id]) GRAF[id].destroy(); GRAF[id] = new Chart(el, { type: tipo, data, options: Object.assign({ responsive: true, maintainAspectRatio: true, plugins: { legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } }, tooltip: { callbacks: { label: (ctx) => { const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed; return ctx.dataset.label ? `${ctx.dataset.label}: ${moneyCorto(v)}` : moneyCorto(v); } } } }, scales: tipo === "bar" ? { y: { ticks: { callback: (v) => moneyCorto(v) } } } : {} }, extra || {}) }); }
  function pintarTabla(filas) { const w = document.getElementById("c08-tabla-wrap"); if (!w) return; const cols = [["codigo", "Cód"], ["poai", "POAI 2026"], ["apropiaciones", "Aprop"], ["compromisos", "Comp"], ["giros", "Giros"], ["pctComp", "% Ej"]]; let html = '<table style="width:100%;border-collapse:collapse;font-size:11px;"><thead><tr>'; cols.forEach(c => html += `<th style="text-align:left;padding:8px;background:#1a2742;color:#fff;font-size:10px;position:sticky;top:0;font-weight:600;">${c[1]}</th>`); html += "</tr></thead><tbody>"; filas.forEach((r, i) => { html += `<tr style="background:${i % 2 ? "#f7f9fc" : "#fff"};">`; cols.forEach(([campo]) => { let val = r[campo]; if (["poai", "apropiaciones", "compromisos", "giros"].includes(campo)) val = money(num(val)); else if (campo === "pctComp") val = pct(r.pctComp); html += `<td style="padding:6px 8px;border-bottom:1px solid #eef2f7;font-size:11px;">${val ?? ""}</td>`; }); html += "</tr>"; }); html += "</tbody></table>"; w.innerHTML = html; }
  console.log("✅ Módulo 08 — POAI 2026 capturado correctamente");
})();
