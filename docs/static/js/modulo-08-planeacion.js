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
        const pdluN = names.find(n => n.toLowerCase().includes("pdlu") || n.toLowerCase().includes("pdl")) || names[0];
        const ejN = names.find(n => n.toLowerCase().includes("ejec")) || names[1];
        const pdlu = XLSX.utils.sheet_to_json(wb.Sheets[pdluN], { defval: "" });
        const ejRaw = XLSX.utils.sheet_to_json(wb.Sheets[ejN], { header: 1 });
        let ejRows = [];
        if (ejRaw.length > 1) {
          let hIdx = 0;
          for (let i = 0; i < Math.min(5, ejRaw.length); i++) {
            if (ejRaw[i].some(c => String(c || "").toLowerCase().includes("proyecto") || String(c || "").toLowerCase().includes("aprop"))) { hIdx = i; break; }
          }
          const heads = ejRaw[hIdx].map(h => String(h || "").trim());
          for (let i = hIdx + 1; i < ejRaw.length; i++) {
            const row = ejRaw[i];
            if (!row.some(c => c != null && c !== "")) continue;
            const o = {}; heads.forEach((h, j) => o[h] = row[j] ?? ""); ejRows.push(o);
          }
        }
        procesar(pdlu, ejRows);
        fileName = archivo.name;
        if (SECTORES.length === 0) { alert("El archivo no contiene datos reconocibles de PDL."); return; }
        render(archivo.name);
      } catch (err) { console.error(err); alert("No se pudo leer: " + err.message); }
    };
    reader.readAsArrayBuffer(archivo);
  }
  const num = (x) => { const n = parseFloat(String(x).replace(/[^\d.-]/g, "")); return isNaN(n) ? 0 : n; };
  const money = (v) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);
  const moneyCorto = (v) => { if (Math.abs(v) >= 1e12) return "$" + (v / 1e12).toFixed(1) + "B"; if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM"; if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M"; if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K"; return "$" + v.toFixed(0); };
  const gcP = (row, keys) => { for (const k of keys) if (row[k] != null && row[k] !== "") return row[k]; return ""; };
  const pct = (v) => (v * 100).toFixed(1) + "%";
  function procesar(pdlu, ejec) {
    const ejMap = {};
    ejec.forEach(row => {
      const cod = num(gcP(row, ["Código_Proyecto", "codigo_proyecto", "Cód. Proyecto"]));
      if (!cod) return;
      ejMap[cod] = { aprop: num(gcP(row, ["Apropiaciones", "apropiaciones"])), comp: num(gcP(row, ["Compromisos proyecto", "compromisos"])), giros: num(gcP(row, ["Giros proyecto", "giros"])) };
    });
    const proyMap = {};
    pdlu.forEach(row => {
      const cod = num(gcP(row, ["Cód. Proyecto de Inversión SEGPLAN", "Cod Proyecto SEGPLAN"]));
      if (!cod) return;
      const nombre = String(gcP(row, ["Nombre del Proyecto", "nombre_proyecto"]) || "");
      const sector = String(gcP(row, ["Sector", "sector"]) || "SIN SECTOR");
      const poai = num(gcP(row, ["Recursos POAI DEFINITIVO 2026", "recursos_poai"]));
      const ej = ejMap[cod] || { aprop: 0, comp: 0, giros: 0 };
      if (!proyMap[cod]) proyMap[cod] = { codigo: cod, nombre, sector, poai: 0, apropiaciones: ej.aprop, compromisos: ej.comp, giros: ej.giros, pctComp: ej.aprop > 0 ? ej.comp / ej.aprop : 0, pctGiros: ej.aprop > 0 ? ej.giros / ej.aprop : 0 };
      proyMap[cod].poai += poai;
    });
    PROYECTOS = Object.values(proyMap).sort((a, b) => b.poai - a.poai);
    const secMap = {};
    PROYECTOS.forEach(p => {
      if (!secMap[p.sector]) secMap[p.sector] = { nombre: p.sector, poai: 0, apropiaciones: 0, compromisos: 0, giros: 0, proyCount: 0 };
      secMap[p.sector].poai += p.poai; secMap[p.sector].apropiaciones += p.apropiaciones; secMap[p.sector].compromisos += p.compromisos; secMap[p.sector].giros += p.giros; secMap[p.sector].proyCount += 1;
    });
    Object.values(secMap).forEach(s => { s.pctComp = s.apropiaciones > 0 ? s.compromisos / s.apropiaciones : 0; s.pctGiros = s.apropiaciones > 0 ? s.giros / s.apropiaciones : 0; });
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
    const sinAprop = PROYECTOS.filter(p => p.apropiaciones === 0).length;
    const dash = document.createElement("div");
    dash.id = "dash-08";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:12px;padding:18px 22px;margin:18px 0;color:#fff;">
        <div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8fb8e8;">Plan de Desarrollo Local · ${nombreArchivo}</div>
        <div style="font-size:22px;font-weight:700;margin-top:4px;">${PROYECTOS.length.toLocaleString("es-CO")} proyectos · ${SECTORES.length} sectores</div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:18px;">
        ${kpi("POAI 2026", moneyCorto(TOTALS.poai), "#667eea")}
        ${kpi("Apropiaciones", moneyCorto(TOTALS.apropiaciones), "#17a2b8")}
        ${kpi("Compromisos", moneyCorto(TOTALS.compromisos), "#28a745")}
        ${kpi("Giros", moneyCorto(TOTALS.giros), "#fd7e14")}
        ${kpi("% Ejecución", pctComp.toFixed(1) + "%", "#8e44ad")}
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Cascada presupuestal (POAI → Giros)</div><canvas id="c08-cascada"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Ejecución del presupuesto</div><canvas id="c08-ejec"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Distribución POAI por sector</div><canvas id="c08-sector-pie"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Top 10 proyectos (POAI)</div><canvas id="c08-top-proy"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);grid-column:1/-1;"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">POAI · Apropiaciones · Compromisos · Giros por sector</div><canvas id="c08-sector-multi"></canvas></div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);grid-column:1/-1;"><div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">% Ejecución por sector — semáforo de gestión</div><canvas id="c08-semaforo"></canvas></div>
      </div>
      ${sinAprop > 0 ? `<div style="background:#fff3cd;border-left:4px solid #ffc107;border-radius:8px;padding:12px 16px;margin-bottom:15px;font-size:13px;color:#856404;">⚠️ <strong>${sinAprop} proyectos</strong> del PDL no registran apropiaciones. Verificar si están pendientes de presupuestar.</div>` : ""}
      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <input id="c08-busca" placeholder="🔎 Buscar proyecto, código o sector…" style="flex:1;padding:11px;border:1.5px solid #e8edf2;border-radius:8px;font-size:13px;">
        <select id="c08-filtro-sector" style="padding:11px;border:1.5px solid #e8edf2;border-radius:8px;font-size:13px;"><option value="TODOS">🏛️ Todos los sectores</option>${[...new Set(PROYECTOS.map(p => p.sector))].map(s => `<option value="${s}">${s.split(",")[0]}</option>`).join("")}</select>
        <button id="c08-btn-tabla" style="padding:11px 18px;background:#667eea;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;">Ver tabla</button>
      </div>
      <div id="c08-tabla-wrap" style="display:none;background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 10px rgba(0,0,0,.06);overflow-x:auto;"></div>
    `;
    contenedor.appendChild(dash);
    graficaCascada(); graficaEjecucion(); graficaSectorPie(); graficaTopProyectos(); graficaSectorMulti(); graficaSemaforo();
    document.getElementById("c08-btn-tabla").addEventListener("click", () => { const w = document.getElementById("c08-tabla-wrap"); w.style.display = w.style.display === "none" ? "block" : "none"; if (w.style.display === "block") pintarTabla(filtrar()); });
    document.getElementById("c08-busca").addEventListener("input", () => { document.getElementById("c08-tabla-wrap").style.display = "block"; pintarTabla(filtrar()); });
    document.getElementById("c08-filtro-sector").addEventListener("change", () => { document.getElementById("c08-tabla-wrap").style.display = "block"; pintarTabla(filtrar()); });
    dash.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function kpi(label, valor, color) { return `<div style="background:${color};color:#fff;padding:16px;border-radius:10px;text-align:center;"><div style="font-size:11px;opacity:.85;text-transform:uppercase;letter-spacing:.04em;">${label}</div><div style="font-size:20px;font-weight:700;margin-top:4px;line-height:1.1;">${valor}</div></div>`; }
  function filtrar() { const q = (document.getElementById("c08-busca").value || "").toLowerCase().trim(); const sec = document.getElementById("c08-filtro-sector").value; let res = PROYECTOS; if (sec !== "TODOS") res = res.filter(p => p.sector === sec); if (q) res = res.filter(p => [p.nombre, String(p.codigo), p.sector].some(v => String(v).toLowerCase().includes(q))); return res.slice(0, 200); }
  function graficaCascada() { dibujar("c08-cascada", "bar", { labels: ["POAI 2026", "Apropiaciones", "Compromisos", "Giros"], datasets: [{ label: "Valor", data: [TOTALS.poai, TOTALS.apropiaciones, TOTALS.compromisos, TOTALS.giros], backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14"] }] }, { plugins: { legend: { display: false } } }); }
  function graficaEjecucion() { const noComp = Math.max(0, TOTALS.apropiaciones - TOTALS.compromisos); const compSinGiro = Math.max(0, TOTALS.compromisos - TOTALS.giros); dibujar("c08-ejec", "doughnut", { labels: ["Girado", "Comprometido sin girar", "Sin comprometer"], datasets: [{ data: [TOTALS.giros, compSinGiro, noComp], backgroundColor: ["#28a745", "#ffc107", "#dc3545"] }] }); }
  function graficaSectorPie() { const top = SECTORES.slice(0, 8); dibujar("c08-sector-pie", "pie", { labels: top.map(s => s.nombre.split(",")[0].length > 22 ? s.nombre.split(",")[0].slice(0, 22) + "…" : s.nombre.split(",")[0]), datasets: [{ data: top.map(s => s.poai), backgroundColor: ["#667eea", "#28a745", "#ffc107", "#dc3545", "#17a2b8", "#8e44ad", "#fd7e14", "#20c997"] }] }); }
  function graficaTopProyectos() { const top = PROYECTOS.slice(0, 10); dibujar("c08-top-proy", "bar", { labels: top.map(p => (p.nombre.length > 28 ? p.nombre.slice(0, 28) + "…" : p.nombre)), datasets: [{ label: "POAI", data: top.map(p => p.poai), backgroundColor: "#667eea" }] }, { indexAxis: "y" }); }
  function graficaSectorMulti() { const sec = SECTORES.slice(0, 10); dibujar("c08-sector-multi", "bar", { labels: sec.map(s => s.nombre.split(",")[0].length > 16 ? s.nombre.split(",")[0].slice(0, 16) + "…" : s.nombre.split(",")[0]), datasets: [{ label: "POAI", data: sec.map(s => s.poai), backgroundColor: "#667eea" }, { label: "Apropiaciones", data: sec.map(s => s.apropiaciones), backgroundColor: "#17a2b8" }, { label: "Compromisos", data: sec.map(s => s.compromisos), backgroundColor: "#28a745" }, { label: "Giros", data: sec.map(s => s.giros), backgroundColor: "#fd7e14" }] }); }
  function graficaSemaforo() { const sec = SECTORES.filter(s => s.apropiaciones > 0).sort((a, b) => b.pctComp - a.pctComp); const colores = sec.map(s => s.pctComp >= 0.6 ? "#28a745" : s.pctComp >= 0.4 ? "#17a2b8" : s.pctComp >= 0.2 ? "#ffc107" : "#dc3545"); const el = document.getElementById("c08-semaforo"); if (!el || typeof Chart === "undefined") return; if (GRAF["c08-semaforo"]) GRAF["c08-semaforo"].destroy(); GRAF["c08-semaforo"] = new Chart(el, { type: "bar", data: { labels: sec.map(s => s.nombre.split(",")[0].length > 20 ? s.nombre.split(",")[0].slice(0, 20) + "…" : s.nombre.split(",")[0]), datasets: [{ label: "% Ejecución", data: sec.map(s => +(s.pctComp * 100).toFixed(1)), backgroundColor: colores }] }, options: { indexAxis: "y", responsive: true, maintainAspectRatio: true, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx) => `Ejecución: ${ctx.parsed.x}%` } } }, scales: { x: { max: 100, ticks: { callback: (v) => v + "%" } } } } }); }
  function dibujar(id, tipo, data, extra) { const el = document.getElementById(id); if (!el || typeof Chart === "undefined") return; if (GRAF[id]) GRAF[id].destroy(); GRAF[id] = new Chart(el, { type: tipo, data, options: Object.assign({ responsive: true, maintainAspectRatio: true, plugins: { legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } }, tooltip: { callbacks: { label: (ctx) => { const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed; return ctx.dataset.label ? `${ctx.dataset.label}: ${moneyCorto(v)}` : moneyCorto(v); } } } }, scales: tipo === "bar" ? { [extra && extra.indexAxis === "y" ? "x" : "y"]: { ticks: { callback: (v) => moneyCorto(v) } } } : {} }, extra || {}) }); }
  function pintarTabla(filas) { const w = document.getElementById("c08-tabla-wrap"); if (!w) return; const cols = [["codigo", "Código"], ["nombre", "Proyecto"], ["sector", "Sector"], ["poai", "POAI"], ["apropiaciones", "Apropiaciones"], ["compromisos", "Compromisos"], ["giros", "Giros"], ["pctComp", "% Ejec."]]; let html = '<table style="width:100%;border-collapse:collapse;font-size:12px;"><thead><tr>'; cols.forEach(c => html += `<th style="text-align:left;padding:8px;background:#1a2742;color:#fff;font-size:11px;position:sticky;top:0;">${c[1]}</th>`); html += "</tr></thead><tbody>"; filas.forEach((r, i) => { html += `<tr style="background:${i % 2 ? "#f7f9fc" : "#fff"};">`; cols.forEach(([campo]) => { let val = r[campo]; if (["poai", "apropiaciones", "compromisos", "giros"].includes(campo)) val = money(num(val)); else if (campo === "pctComp") val = pct(r.pctComp); else if (campo === "nombre") val = String(val).length > 45 ? String(val).slice(0, 45) + "…" : val; html += `<td style="padding:7px 8px;border-bottom:1px solid #eef2f7;">${val ?? ""}</td>`; }); html += "</tr>"; }); html += "</tbody></table>"; if (filas.length === 0) html = '<div style="padding:20px;color:#888;">Sin resultados.</div>'; w.innerHTML = html; }
  console.log("✅ Módulo 08 (Planeación PDL) listo");
})();
