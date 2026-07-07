/* ============================================================
   MÓDULO 08 — Planeación PDL (Plan de Desarrollo Local)
   Vanilla JS + Chart.js. Cruza hojas PDLU + EJECUCIONES.
   Colores del módulo 07.
   ============================================================ */
(function () {
  "use strict";

  const C = {
    navy: "#1a2742", navy2: "#1e3a6e", card: "#ffffff", bg: "#f0f2f5",
    cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12", rojo: "#e74c3c",
    violet: "#8e44ad", muted: "#7f8c8d", border: "#e8edf2", texto: "#2c3e50",
  };
  const PAL = ["#5dade2","#27ae60","#8e44ad","#f39c12","#1abc9c","#e74c3c","#e67e22","#3498db","#9b59b6","#16a085","#d35400","#7f8c8d"];

  let SECTORES = [], PROYECTOS = [], TOTALS = {}, TAB = "resumen", FILTRO = "TODOS";
  const GRAF = {};

  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();

  function init() {
    const input = document.getElementById("carga-08");
    if (!input || input.dataset.ligado === "1") return;
    input.dataset.ligado = "1";
    input.addEventListener("change", (e) => leer(e.target.files[0]));
  }

  const toN = (v) => { if (v == null || v === "") return 0; const n = parseFloat(String(v).replace(/[^0-9.-]/g, "")); return isNaN(n) ? 0 : n; };
  const fmtCOP = (v) => { if (!v) return "–"; if (v >= 1e12) return "$" + (v/1e12).toFixed(1) + "B"; if (v >= 1e9) return "$" + (v/1e9).toFixed(1) + "MM"; if (v >= 1e6) return "$" + (v/1e6).toFixed(0) + "M"; return "$" + v.toLocaleString("es-CO"); };
  const pct = (v) => (v * 100).toFixed(1) + "%";
  const gcP = (row, keys) => { for (const k of keys) if (row[k] != null && row[k] !== "") return row[k]; return ""; };
  const semColor = (p) => p >= 0.6 ? C.verde : p >= 0.4 ? C.cyan : p >= 0.2 ? C.amar : p > 0 ? C.rojo : C.muted;
  const semTxt = (p) => !p ? "Sin datos" : p >= 0.6 ? "✅ En meta" : p >= 0.4 ? "🔵 Bueno" : p >= 0.2 ? "🟡 Seguim." : "🔴 Bajo";

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

        // ejecuciones: detectar header
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
        if (SECTORES.length === 0) { alert("No se encontraron datos. Verifica que el Excel tenga las hojas PDLU y EJECUCIONES."); return; }
        render(file.name);
      } catch (e) { console.error(e); alert("Error leyendo archivo: " + e.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function procesar(pdlu, ejec) {
    const ejMap = {};
    ejec.forEach(row => {
      const cod = toN(gcP(row, ["Código_Proyecto","Codigo_Proyecto","codigo_proyecto","CODIGO_PROYECTO","Cód. Proyecto","Cod. Proyecto"]));
      if (!cod) return;
      ejMap[cod] = {
        aprop: toN(gcP(row, ["Apropiaciones","apropiaciones","APROPIACIONES"])),
        comp: toN(gcP(row, ["Compromisos proyecto","Compromisos_proyecto","compromisos","Compromisos","COMPROMISOS"])),
        giros: toN(gcP(row, ["Giros proyecto","Giros_proyecto","giros","Giros","GIROS"])),
      };
    });

    const proyMap = {};
    pdlu.forEach(row => {
      const cod = toN(gcP(row, ["Cód. Proyecto de Inversión SEGPLAN","Cod. Proyecto de Inversión SEGPLAN","Código Proyecto SEGPLAN","Codigo Proyecto SEGPLAN","Cód. Proyecto","Cod Proyecto"]));
      if (!cod) return;
      const nombre = String(gcP(row, ["Nombre del Proyecto","Nombre Proyecto","nombre_proyecto"]) || "");
      const sector = String(gcP(row, ["Sector","sector","SECTOR"]) || "SIN SECTOR");
      const poai = toN(gcP(row, ["Recursos POAI DEFINITIVO 2026\n","Recursos POAI DEFINITIVO 2026","Recursos POAI","recursos_poai","POAI"]));
      const meta = toN(gcP(row, ["Magnitud Meta anualizada DEFINITIVA 2026","Magnitud Meta anualizada","Meta anualizada","meta_anualizada"]));
      const ind = String(gcP(row, ["Indicador de producto","Indicador_Producto","indicador_producto"]) || "");
      const prog = String(gcP(row, ["Programa","programa","PROGRAMA"]) || "");
      const ej = ejMap[cod] || { aprop: 0, comp: 0, giros: 0 };
      if (!proyMap[cod]) proyMap[cod] = { codigo: cod, nombre, sector, poai: 0, apropiaciones: ej.aprop, compromisos: ej.comp, giros: ej.giros, pctComp: ej.aprop > 0 ? ej.comp/ej.aprop : 0, pctGiros: ej.aprop > 0 ? ej.giros/ej.aprop : 0, indicadores: [] };
      proyMap[cod].poai += poai;
      if (ind) proyMap[cod].indicadores.push({ descripcion: ind, programa: prog, meta });
    });

    PROYECTOS = Object.values(proyMap).sort((a, b) => b.poai - a.poai);
    const secMap = {};
    PROYECTOS.forEach(p => {
      if (!secMap[p.sector]) secMap[p.sector] = { nombre: p.sector, poai: 0, apropiaciones: 0, compromisos: 0, giros: 0 };
      secMap[p.sector].poai += p.poai; secMap[p.sector].apropiaciones += p.apropiaciones;
      secMap[p.sector].compromisos += p.compromisos; secMap[p.sector].giros += p.giros;
    });
    Object.values(secMap).forEach(s => { s.pctComp = s.apropiaciones > 0 ? s.compromisos/s.apropiaciones : 0; s.pctGiros = s.apropiaciones > 0 ? s.giros/s.apropiaciones : 0; });
    SECTORES = Object.values(secMap).sort((a, b) => b.poai - a.poai);
    TOTALS = {
      poai: PROYECTOS.reduce((s,p)=>s+p.poai,0),
      apropiaciones: PROYECTOS.reduce((s,p)=>s+p.apropiaciones,0),
      compromisos: PROYECTOS.reduce((s,p)=>s+p.compromisos,0),
      giros: PROYECTOS.reduce((s,p)=>s+p.giros,0),
    };
  }

  function render(nombre) {
    const vista = document.getElementById("vista-planeacion");
    if (!vista) return;
    const decor = document.getElementById("pantalla-carga-08");
    if (decor) decor.style.display = "none";
    const prev = document.getElementById("plan-root");
    if (prev) prev.remove();

    const pctComp = TOTALS.apropiaciones > 0 ? TOTALS.compromisos/TOTALS.apropiaciones : 0;
    const pctGiros = TOTALS.apropiaciones > 0 ? TOTALS.giros/TOTALS.apropiaciones : 0;
    const sinAprop = PROYECTOS.filter(p => p.apropiaciones === 0).length;

    const root = document.createElement("div");
    root.id = "plan-root";
    root.innerHTML = `
      <div style="background:linear-gradient(135deg,${C.navy},${C.navy2});border-radius:12px;padding:18px 24px;margin:18px 0 16px;color:#fff;">
        <div style="font-size:10px;color:${C.cyan};letter-spacing:3px;font-weight:700;text-transform:uppercase;">🎯 Planeación · Plan de Desarrollo Local</div>
        <h2 style="margin:6px 0 4px;font-size:20px;font-weight:900;">Ejecución PDL — ${nombre}</h2>
        <div style="font-size:11px;color:rgba(255,255,255,.7);">${PROYECTOS.length} proyectos · ${SECTORES.length} sectores</div>
      </div>

      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px;">
        ${kpi("POAI 2026", fmtCOP(TOTALS.poai), C.cyan, "Presupuesto programado")}
        ${kpi("Apropiaciones", fmtCOP(TOTALS.apropiaciones), C.amar, "Disponible")}
        ${kpi("Compromisos", fmtCOP(TOTALS.compromisos), C.verde, pct(pctComp) + " ejecutado")}
        ${kpi("Giros", fmtCOP(TOTALS.giros), C.rojo, pct(pctGiros) + " girado")}
        ${kpi("Sin apropiación", sinAprop + " proy.", C.violet, "Requieren gestión")}
      </div>

      <div style="background:#fff;border-radius:10px 10px 0 0;border-bottom:2px solid ${C.border};display:flex;padding:0 12px;overflow-x:auto;">
        ${tabBtn("resumen","📊 Resumen")}
        ${tabBtn("sectores","🏛️ Sectores")}
        ${tabBtn("proyectos","🎯 Proyectos")}
        ${tabBtn("alertas","🚨 Alertas")}
      </div>
      <div id="plan-tab" style="background:#fff;border-radius:0 0 10px 10px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);"></div>
    `;
    vista.appendChild(root);
    root.querySelectorAll("[data-tab]").forEach(b => b.addEventListener("click", () => { TAB = b.dataset.tab; pintarTab(); }));
    pintarTab();
    root.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function kpi(label, valor, color, sub) {
    return `<div style="background:#fff;border-top:3px solid ${color};border-radius:10px;padding:14px 16px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
      <div style="font-size:10px;color:${C.muted};text-transform:uppercase;letter-spacing:.5px;">${label}</div>
      <div style="font-size:20px;font-weight:800;color:${color};margin:4px 0 2px;line-height:1;">${valor}</div>
      <div style="font-size:10px;color:${C.muted};">${sub}</div>
    </div>`;
  }
  function tabBtn(id, label) {
    const a = TAB === id;
    return `<button data-tab="${id}" style="padding:12px 18px;background:transparent;border:none;border-bottom:${a?`2px solid ${C.cyan}`:"2px solid transparent"};color:${a?C.cyan:C.muted};cursor:pointer;font-size:12px;font-weight:${a?700:500};white-space:nowrap;">${label}</button>`;
  }

  function pintarTab() {
    const root = document.getElementById("plan-root");
    root.querySelectorAll("[data-tab]").forEach(b => {
      const a = b.dataset.tab === TAB;
      b.style.borderBottom = a ? `2px solid ${C.cyan}` : "2px solid transparent";
      b.style.color = a ? C.cyan : C.muted; b.style.fontWeight = a ? 700 : 500;
    });
    const c = document.getElementById("plan-tab");
    if (TAB === "resumen") tabResumen(c);
    else if (TAB === "sectores") tabSectores(c);
    else if (TAB === "proyectos") tabProyectos(c);
    else tabAlertas(c);
  }

  function tabResumen(c) {
    c.innerHTML = `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;">
        <div><div style="font-size:13px;font-weight:600;color:${C.texto};margin-bottom:8px;">Distribución POAI por sector</div><canvas id="pl-donut"></canvas></div>
        <div><div style="font-size:13px;font-weight:600;color:${C.texto};margin-bottom:8px;">Cascada presupuestal</div><canvas id="pl-cascada"></canvas></div>
        <div style="grid-column:1/-1;"><div style="font-size:13px;font-weight:600;color:${C.texto};margin-bottom:8px;">% Compromisos vs Giros por sector</div><canvas id="pl-compgiros"></canvas></div>
      </div>`;
    const top = SECTORES.slice(0, 10);
    dibujar("pl-donut", "doughnut", {
      labels: top.map(s => s.nombre.split(",")[0].slice(0, 18)),
      datasets: [{ data: top.map(s => s.poai), backgroundColor: top.map((_,i)=>PAL[i%PAL.length]) }],
    });
    dibujar("pl-cascada", "bar", {
      labels: ["POAI","Apropiaciones","Compromisos","Giros"],
      datasets: [{ data: [TOTALS.poai,TOTALS.apropiaciones,TOTALS.compromisos,TOTALS.giros], backgroundColor: [C.cyan,C.amar,C.verde,C.rojo] }],
    }, { sinLeyenda: true });
    const sec = SECTORES.filter(s => s.apropiaciones > 0);
    dibujar("pl-compgiros", "bar", {
      labels: sec.map(s => s.nombre.split(",")[0].slice(0, 14)),
      datasets: [
        { label: "% Comp", data: sec.map(s => +(s.pctComp*100).toFixed(1)), backgroundColor: C.verde },
        { label: "% Giros", data: sec.map(s => +(s.pctGiros*100).toFixed(1)), backgroundColor: C.cyan },
      ],
    }, { pct: true });
  }

  function tabSectores(c) {
    c.innerHTML = `<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:12px;">
      <thead><tr style="background:${C.navy};">${["Sector","POAI","Apropiac.","Compromisos","Giros","% Comp","Estado"].map(h=>`<th style="padding:9px 10px;text-align:left;color:${C.cyan};font-size:10px;white-space:nowrap;">${h}</th>`).join("")}</tr></thead>
      <tbody>${SECTORES.map((s,i)=>`<tr style="background:${i%2?C.bg:"#fff"};border-bottom:1px solid ${C.border};">
        <td style="padding:8px 10px;font-weight:700;color:${C.texto};max-width:180px;">${s.nombre}</td>
        <td style="padding:8px 10px;">${fmtCOP(s.poai)}</td>
        <td style="padding:8px 10px;">${fmtCOP(s.apropiaciones)}</td>
        <td style="padding:8px 10px;color:${C.verde};font-weight:600;">${fmtCOP(s.compromisos)}</td>
        <td style="padding:8px 10px;color:${C.cyan};">${fmtCOP(s.giros)}</td>
        <td style="padding:8px 10px;font-weight:700;color:${semColor(s.pctComp)};">${pct(s.pctComp)}</td>
        <td style="padding:8px 10px;"><span style="background:${semColor(s.pctComp)}22;color:${semColor(s.pctComp)};border-radius:12px;padding:2px 8px;font-size:10px;font-weight:700;white-space:nowrap;">${semTxt(s.pctComp)}</span></td>
      </tr>`).join("")}</tbody></table></div>`;
  }

  function tabProyectos(c) {
    const sectores = ["TODOS", ...new Set(PROYECTOS.map(p => p.sector))];
    const filt = FILTRO === "TODOS" ? PROYECTOS : PROYECTOS.filter(p => p.sector === FILTRO);
    c.innerHTML = `
      <div style="margin-bottom:12px;">
        <select id="pl-filtro" style="padding:8px 12px;border:1.5px solid ${C.border};border-radius:8px;font-size:12px;">
          ${sectores.map(s => `<option value="${s}" ${s===FILTRO?"selected":""}>${s === "TODOS" ? "🏛️ Todos los sectores" : s}</option>`).join("")}
        </select>
        <span style="margin-left:10px;font-size:11px;color:${C.muted};">${filt.length} proyectos</span>
      </div>
      <div style="overflow-x:auto;max-height:500px;overflow-y:auto;"><table style="width:100%;border-collapse:collapse;font-size:11px;">
        <thead><tr style="background:${C.navy};position:sticky;top:0;">${["Cód","Proyecto","Sector","POAI","Compromisos","% Comp","Estado"].map(h=>`<th style="padding:9px 8px;text-align:left;color:${C.cyan};font-size:10px;white-space:nowrap;">${h}</th>`).join("")}</tr></thead>
        <tbody>${filt.slice(0,150).map((p,i)=>`<tr style="background:${i%2?C.bg:"#fff"};border-bottom:1px solid ${C.border};">
          <td style="padding:7px 8px;font-weight:700;color:${C.navy2};">${p.codigo}</td>
          <td style="padding:7px 8px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${p.nombre}">${p.nombre}</td>
          <td style="padding:7px 8px;color:${C.muted};max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${p.sector.split(",")[0]}</td>
          <td style="padding:7px 8px;">${fmtCOP(p.poai)}</td>
          <td style="padding:7px 8px;color:${C.verde};">${fmtCOP(p.compromisos)}</td>
          <td style="padding:7px 8px;font-weight:700;color:${semColor(p.pctComp)};">${pct(p.pctComp)}</td>
          <td style="padding:7px 8px;"><span style="background:${semColor(p.pctComp)}22;color:${semColor(p.pctComp)};border-radius:12px;padding:2px 7px;font-size:9px;font-weight:700;white-space:nowrap;">${semTxt(p.pctComp)}</span></td>
        </tr>`).join("")}</tbody></table></div>`;
    c.querySelector("#pl-filtro").addEventListener("change", (e) => { FILTRO = e.target.value; tabProyectos(c); });
  }

  function tabAlertas(c) {
    const criticos = PROYECTOS.filter(p => p.apropiaciones > 0 && p.pctComp < 0.1);
    const sinAprop = PROYECTOS.filter(p => p.apropiaciones === 0);
    const brecha = TOTALS.compromisos > 0 ? (TOTALS.compromisos - TOTALS.giros) / TOTALS.compromisos : 0;
    c.innerHTML = `
      <div style="display:flex;flex-direction:column;gap:12px;">
        ${alerta("🚨", "Proyectos en ejecución crítica (<10%)", `${criticos.length} proyectos con apropiación pero compromiso menor al 10%.`, C.rojo)}
        ${alerta("⚠️", "Proyectos sin apropiación", `${sinAprop.length} proyectos aún no tienen recursos apropiados.`, C.amar)}
        ${alerta("💧", "Brecha compromisos vs giros", `${pct(brecha)} de lo comprometido aún no se ha girado.`, C.cyan)}
      </div>
      ${criticos.length ? `<div style="margin-top:16px;font-size:12px;font-weight:700;color:${C.texto};margin-bottom:8px;">Detalle proyectos críticos</div>
      <div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:11px;">
        <thead><tr style="background:${C.navy};">${["Cód","Proyecto","Apropiac.","% Comp"].map(h=>`<th style="padding:8px;text-align:left;color:${C.cyan};font-size:10px;">${h}</th>`).join("")}</tr></thead>
        <tbody>${criticos.slice(0,30).map((p,i)=>`<tr style="background:${i%2?C.bg:"#fff"};">
          <td style="padding:7px 8px;font-weight:700;color:${C.navy2};">${p.codigo}</td>
          <td style="padding:7px 8px;max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${p.nombre}</td>
          <td style="padding:7px 8px;">${fmtCOP(p.apropiaciones)}</td>
          <td style="padding:7px 8px;color:${C.rojo};font-weight:700;">${pct(p.pctComp)}</td>
        </tr>`).join("")}</tbody></table></div>` : ""}`;
  }
  function alerta(icon, titulo, texto, color) {
    return `<div style="background:${color}12;border:1px solid ${color}44;border-left:4px solid ${color};border-radius:8px;padding:14px 16px;display:flex;gap:12px;align-items:center;">
      <span style="font-size:22px;">${icon}</span>
      <div><div style="font-weight:700;font-size:13px;color:${C.texto};">${titulo}</div><div style="font-size:12px;color:${C.muted};margin-top:2px;">${texto}</div></div>
    </div>`;
  }

  function dibujar(id, tipo, data, extra) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    if (GRAF[id]) GRAF[id].destroy();
    const esPct = extra && extra.pct;
    GRAF[id] = new Chart(el, {
      type: tipo, data,
      options: {
        responsive: true, maintainAspectRatio: true,
        plugins: {
          legend: { display: !(extra && extra.sinLeyenda), position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
          tooltip: { callbacks: { label: (ctx) => {
            const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed;
            if (esPct) return `${ctx.dataset.label}: ${v}%`;
            return ctx.dataset.label ? `${ctx.dataset.label}: ${fmtCOP(v)}` : fmtCOP(v);
          }}},
        },
        scales: tipo === "bar" ? { y: { ticks: { callback: (v) => esPct ? v + "%" : fmtCOP(v) } } } : {},
      },
    });
  }

  console.log("✅ Módulo 08 (Planeación PDL) listo");
})();
