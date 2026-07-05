/* ============================================================
   MÓDULO 08 — Planeación PDL (Plan de Desarrollo Local)
   Convertido de React a Vanilla JS + Chart.js
   Cruza hojas PDLU + EJECUCIONES. 5 tabs completos.
   ============================================================ */
(function () {
  "use strict";
  const C = { dark: "#0B1221", darkMid: "#111929", darkCard: "#1A2744", navy: "#1a2742", cyan: "#00C8E0", verde: "#10B981", amar: "#F59E0B", rojo: "#EF4444", border: "#243058", texto: "#E2E8F0", secondary: "#CBD5E1", muted2: "#94A3B8" };
  const PAL = ["#00C8E0","#10B981","#818CF8","#F59E0B","#34D399","#A78BFA","#F472B6","#FB7185","#FBBF24","#38BDF8"];
  let SECTORES = [], PROYECTOS = [], TOTALS = {}, TAB = "resumen", FILTRO = "TODOS";
  let fileName = "";
  const GRAF = {};
  const toN = (v) => { if (v == null || v === "") return 0; const n = parseFloat(String(v).replace(/[^0-9.-]/g, "")); return isNaN(n) ? 0 : n; };
  const fmtCOP = (v) => { if (!v) return "–"; if (v >= 1e12) return "$" + (v/1e12).toFixed(1) + "T"; if (v >= 1e9) return "$" + (v/1e9).toFixed(1) + "B"; if (v >= 1e6) return "$" + (v/1e6).toFixed(0) + "M"; return "$" + v.toLocaleString("es-CO"); };
  const pct = (v) => (v * 100).toFixed(1) + "%";
  const gcP = (row, keys) => { for (const k of keys) if (row[k] != null && row[k] !== "") return row[k]; return ""; };
  const riskColor = (p) => p >= 0.6 ? C.verde : p >= 0.4 ? C.cyan : p >= 0.2 ? C.amar : C.rojo;
  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();
  function init() {
    const input = document.getElementById("carga-08");
    if (!input || input.dataset.ligado === "1") return;
    input.dataset.ligado = "1";
    input.addEventListener("change", (e) => leer(e.target.files[0]));
  }
  function leer(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const names = wb.SheetNames;
        const pdluN = names.find(n => n.toLowerCase().includes("pdlu") || n.toLowerCase().includes("pdl")) || names[0];
        const ejN = names.find(n => n.toLowerCase().includes("ejec")) || names[1];
        const pdlu = XLSX.utils.sheet_to_json(wb.Sheets[pdluN]);
        const ejRaw = XLSX.utils.sheet_to_json(wb.Sheets[ejN], { header: 1 });
        let ejRows = [];
        if (ejRaw.length > 1) {
          let hIdx = 0;
          for (let i = 0; i < Math.min(5, ejRaw.length); i++) {
            if (ejRaw[i].some(c => String(c||"").toLowerCase().includes("proyecto") || String(c||"").toLowerCase().includes("aprop"))) { hIdx = i; break; }
          }
          const heads = ejRaw[hIdx].map(h => String(h||"").trim());
          for (let i = hIdx+1; i < ejRaw.length; i++) {
            const row = ejRaw[i];
            if (!row.some(c => c != null && c !== "")) continue;
            const o = {}; heads.forEach((h, j) => o[h] = row[j] ?? ""); ejRows.push(o);
          }
        }
        procesar(pdlu, ejRows);
        fileName = file.name;
        if (SECTORES.length === 0) { alert("No datos encontrados"); return; }
        render();
      } catch (e) { console.error(e); alert("Error: " + e.message); }
    };
    reader.readAsArrayBuffer(file);
  }
  function procesar(pdlu, ejec) {
    const ejMap = {};
    ejec.forEach(row => {
      const cod = toN(gcP(row, ["Código_Proyecto","Codigo_Proyecto","codigo_proyecto","Cód. Proyecto"]));
      if (!cod) return;
      ejMap[cod] = { aprop: toN(gcP(row, ["Apropiaciones","apropiaciones"])), comp: toN(gcP(row, ["Compromisos proyecto","compromisos"])), giros: toN(gcP(row, ["Giros proyecto","giros"])) };
    });
    const proyMap = {};
    pdlu.forEach(row => {
      const cod = toN(gcP(row, ["Cód. Proyecto de Inversión SEGPLAN","Cod Proyecto SEGPLAN"]));
      if (!cod) return;
      const nombre = String(gcP(row, ["Nombre del Proyecto","nombre_proyecto"]) || "");
      const sector = String(gcP(row, ["Sector","sector"]) || "SIN SECTOR");
      const poai = toN(gcP(row, ["Recursos POAI DEFINITIVO 2026","recursos_poai"]));
      const ej = ejMap[cod] || { aprop: 0, comp: 0, giros: 0 };
      if (!proyMap[cod]) proyMap[cod] = { codigo: cod, nombre, sector, poai: 0, apropiaciones: ej.aprop, compromisos: ej.comp, giros: ej.giros, pctComp: ej.aprop > 0 ? ej.comp/ej.aprop : 0, indicadores: [] };
      proyMap[cod].poai += poai;
    });
    PROYECTOS = Object.values(proyMap).sort((a, b) => b.poai - a.poai);
    const secMap = {};
    PROYECTOS.forEach(p => {
      if (!secMap[p.sector]) secMap[p.sector] = { nombre: p.sector, poai: 0, apropiaciones: 0, compromisos: 0, giros: 0 };
      secMap[p.sector].poai += p.poai; secMap[p.sector].apropiaciones += p.apropiaciones; secMap[p.sector].compromisos += p.compromisos; secMap[p.sector].giros += p.giros;
    });
    Object.values(secMap).forEach(s => { s.pctComp = s.apropiaciones > 0 ? s.compromisos/s.apropiaciones : 0; });
    SECTORES = Object.values(secMap).sort((a, b) => b.poai - a.poai);
    TOTALS = { poai: PROYECTOS.reduce((s,p)=>s+p.poai,0), apropiaciones: PROYECTOS.reduce((s,p)=>s+p.apropiaciones,0), compromisos: PROYECTOS.reduce((s,p)=>s+p.compromisos,0), giros: PROYECTOS.reduce((s,p)=>s+p.giros,0) };
  }
  function render() {
    const vista = document.getElementById("vista-planeacion");
    if (!vista) return;
    const root = document.createElement("div");
    root.id = "plan-root";
    root.style.cssText = "background:#0B1221;color:#E2E8F0;padding:0;";
    const header = document.createElement("div");
    header.style.cssText = `background:linear-gradient(135deg,#0B1221,#162040);border-bottom:1px solid ${C.border};padding:14px;`;
    header.innerHTML = `<div style="font-size:14px;font-weight:800;">📐 Planeación PDL 2026</div><div style="font-size:10px;color:${C.cyan};margin-top:4px;">${fileName} · ${PROYECTOS.length} proyectos</div>`;
    root.appendChild(header);
    const tabs = document.createElement("div");
    tabs.style.cssText = `background:#111929;border-bottom:1px solid ${C.border};padding:0 12px;display:flex;gap:2px;`;
    [{ key: 'resumen', label: '📊 Resumen' },{ key: 'sectores', label: '🏛️ Sectores' },{ key: 'proyectos', label: '🎯 Proyectos' },{ key: 'alertas', label: '🚨 Alertas' }].forEach(t => {
      const btn = document.createElement("button");
      btn.style.cssText = `padding:10px 12px;font-size:11px;border:none;background:transparent;cursor:pointer;color:${TAB===t.key?C.cyan:C.muted2};border-bottom:${TAB===t.key?`2px solid ${C.cyan}`:'2px solid transparent'};white-space:nowrap;`;
      btn.textContent = t.label;
      btn.onclick = () => { TAB = t.key; renderTab(root); };
      tabs.appendChild(btn);
    });
    root.appendChild(tabs);
    const content = document.createElement("div");
    content.id = "plan-content";
    content.style.cssText = "padding:16px;background:#0B1221;";
    root.appendChild(content);
    vista.appendChild(root);
    renderTab(root);
  }
  function renderTab(root) {
    const content = root.querySelector("#plan-content");
    content.innerHTML = "";
    if (TAB === "resumen") tabResumen(content);
    else if (TAB === "sectores") tabSectores(content);
    else if (TAB === "proyectos") tabProyectos(content);
    else tabAlertas(content);
  }
  function tabResumen(c) {
    const pctComp = TOTALS.apropiaciones > 0 ? TOTALS.compromisos/TOTALS.apropiaciones : 0;
    c.innerHTML = `<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;">
      <div style="background:${C.darkCard};border-top:3px solid ${C.cyan};border-radius:10px;padding:14px;"><div style="font-size:10px;color:${C.muted2};">POAI 2026</div><div style="font-size:18px;font-weight:800;color:${C.cyan};">${fmtCOP(TOTALS.poai)}</div></div>
      <div style="background:${C.darkCard};border-top:3px solid ${C.amar};border-radius:10px;padding:14px;"><div style="font-size:10px;color:${C.muted2};">Apropiaciones</div><div style="font-size:18px;font-weight:800;color:${C.amar};">${fmtCOP(TOTALS.apropiaciones)}</div></div>
      <div style="background:${C.darkCard};border-top:3px solid ${C.verde};border-radius:10px;padding:14px;"><div style="font-size:10px;color:${C.muted2};">Compromisos</div><div style="font-size:18px;font-weight:800;color:${C.verde};">${fmtCOP(TOTALS.compromisos)}</div></div>
      <div style="background:${C.darkCard};border-top:3px solid ${C.rojo};border-radius:10px;padding:14px;"><div style="font-size:10px;color:${C.muted2};">Giros</div><div style="font-size:18px;font-weight:800;color:${C.rojo};">${fmtCOP(TOTALS.giros)}</div></div>
    </div>
    <div style="background:${C.darkCard};border:1px solid ${C.border};border-radius:10px;padding:14px;">
      <div style="font-size:13px;font-weight:600;margin-bottom:8px;">Sectores</div>
      <canvas id="pl-donut" style="max-height:250px;"></canvas>
    </div>`;
    setTimeout(() => {
      const top = SECTORES.slice(0, 8);
      dibujar("pl-donut", "doughnut", { labels: top.map(s => s.nombre.split(",")[0].slice(0,15)), datasets: [{ data: top.map(s => +(s.poai/1e9).toFixed(2)), backgroundColor: top.map((_,i)=>PAL[i%PAL.length]) }] });
    }, 100);
  }
  function tabSectores(c) {
    c.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:${C.navy};"><th style="padding:8px;color:${C.cyan};">Sector</th><th style="padding:8px;color:${C.cyan};">POAI</th><th style="padding:8px;color:${C.cyan};">Compromisos</th><th style="padding:8px;color:${C.cyan};">% Comp</th></tr>${SECTORES.map((s,i)=>`<tr style="background:${i%2?C.darkCard:"transparent"};"><td style="padding:8px;">${s.nombre}</td><td>${fmtCOP(s.poai)}</td><td>${fmtCOP(s.compromisos)}</td><td style="color:${riskColor(s.pctComp)};font-weight:700;">${pct(s.pctComp)}</td></tr>`).join("")}</table>`;
  }
  function tabProyectos(c) {
    c.innerHTML = `<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:${C.navy};"><th style="padding:8px;color:${C.cyan};">Código</th><th style="padding:8px;color:${C.cyan};">Proyecto</th><th style="padding:8px;color:${C.cyan};">POAI</th><th style="padding:8px;color:${C.cyan};">% Comp</th></tr>${PROYECTOS.slice(0,50).map((p,i)=>`<tr style="background:${i%2?C.darkCard:"transparent"};"><td style="padding:8px;color:${C.cyan};font-weight:700;">${p.codigo}</td><td style="padding:8px;">${p.nombre.substring(0,40)}</td><td style="padding:8px;">${fmtCOP(p.poai)}</td><td style="color:${riskColor(p.pctComp)};font-weight:700;">${pct(p.pctComp)}</td></tr>`).join("")}</table></div>`;
  }
  function tabAlertas(c) {
    const conAprop = PROYECTOS.filter(p => p.apropiaciones > 0);
    const criticos = conAprop.filter(p => p.pctComp < 0.1);
    c.innerHTML = `<div style="background:${C.rojo}15;border-left:4px solid ${C.rojo};padding:12px;border-radius:8px;margin-bottom:16px;"><strong>🚨 CRÍTICOS:</strong> ${criticos.length} proyectos con ejecución < 10%</div>`;
    if (criticos.length) c.innerHTML += `<table style="width:100%;border-collapse:collapse;font-size:11px;"><tr style="background:${C.navy};"><th style="padding:8px;color:${C.cyan};">Código</th><th style="padding:8px;color:${C.cyan};">Proyecto</th><th style="padding:8px;color:${C.cyan};">% Comp</th></tr>${criticos.map((p,i)=>`<tr style="background:${i%2?C.darkCard:"transparent"};"><td style="padding:8px;color:${C.rojo};font-weight:700;">${p.codigo}</td><td style="padding:8px;">${p.nombre.substring(0,50)}</td><td style="color:${C.rojo};font-weight:700;">${pct(p.pctComp)}</td></tr>`).join("")}</table>`;
  }
  function dibujar(id, tipo, data) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    if (GRAF[id]) GRAF[id].destroy();
    GRAF[id] = new Chart(el, { type: tipo, data, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: C.muted2 } } } } });
  }
  console.log("✅ Módulo 08 Planeación PDL listo");
})();
