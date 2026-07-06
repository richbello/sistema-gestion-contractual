(function() {
  "use strict";
  let PROYECTOS = [], TOTALS = {}, fileName = "";
  const GRAF = {};
  
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
  
  function init() {
    const input = document.getElementById("carga-08");
    if (input && input.dataset.ligado !== "1") {
      input.dataset.ligado = "1";
      input.addEventListener("change", (e) => leerExcel(e.target.files[0]));
    }
  }
  
  function leerExcel(archivo) {
    if (!archivo) return;
    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const ws = wb.Sheets["EJECUCIONES"];
        const ejRaw = XLSX.utils.sheet_to_json(ws, { header: 1 });
        
        PROYECTOS = [];
        let poaiTotal = 0;
        
        for (let i = 2; i < ejRaw.length; i++) {
          const row = ejRaw[i];
          if (!row || row.length < 8) continue;
          
          const codigo = String(row[2] || "").trim();
          const poai = parseFloat(String(row[1] || 0).replace(/[^0-9.-]/g, "")) || 0;
          const aprop = parseFloat(String(row[3] || 0).replace(/[^0-9.-]/g, "")) || 0;
          const comp = parseFloat(String(row[4] || 0).replace(/[^0-9.-]/g, "")) || 0;
          const giros = parseFloat(String(row[6] || 0).replace(/[^0-9.-]/g, "")) || 0;
          
          if (poai > poaiTotal) poaiTotal = poai;
          
          if (codigo && aprop > 0) {
            PROYECTOS.push({
              codigo: codigo,
              nombre: "Proyecto " + codigo,
              sector: "Sector Usme",
              poai: 0,
              apropiaciones: aprop,
              compromisos: comp,
              giros: giros,
              pctComp: aprop > 0 ? comp / aprop : 0,
              pctGiros: comp > 0 ? giros / comp : 0
            });
          }
        }
        
        const totalAprop = PROYECTOS.reduce((s, p) => s + p.apropiaciones, 0);
        if (poaiTotal > 0 && totalAprop > 0) {
          PROYECTOS.forEach(p => { p.poai = (p.apropiaciones / totalAprop) * poaiTotal; });
        } else {
          PROYECTOS.forEach(p => { p.poai = p.apropiaciones; });
        }
        
        PROYECTOS.sort((a, b) => b.apropiaciones - a.apropiaciones);
        
        TOTALS = {
          poai: poaiTotal || totalAprop,
          apropiaciones: PROYECTOS.reduce((s, p) => s + p.apropiaciones, 0),
          compromisos: PROYECTOS.reduce((s, p) => s + p.compromisos, 0),
          giros: PROYECTOS.reduce((s, p) => s + p.giros, 0)
        };
        
        console.log("✅ " + PROYECTOS.length + " proyectos cargados");
        
        fileName = archivo.name;
        if (PROYECTOS.length === 0) { alert("Sin datos"); return; }
        render(archivo.name);
      } catch (err) { console.error(err); alert("Error: " + err.message); }
    };
    reader.readAsArrayBuffer(archivo);
  }
  
  const money = (v) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);
  const moneyCorto = (v) => {
    if (Math.abs(v) >= 1e12) return "$" + (v / 1e12).toFixed(1) + "B";
    if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM";
    if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M";
    if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K";
    return "$" + v.toFixed(0);
  };
  
  function render(nombreArchivo) {
    const vista = document.getElementById("vista-planeacion");
    if (!vista) return;
    const contenedor = vista.querySelector("div[style*=\'max-width\']") || vista;
    const cargaDecor = document.getElementById("pantalla-carga-08");
    if (cargaDecor) cargaDecor.style.display = "none";
    const previo = document.getElementById("dash-08");
    if (previo) previo.remove();
    
    const pctComp = TOTALS.apropiaciones ? (TOTALS.compromisos / TOTALS.apropiaciones) * 100 : 0;
    const pctGiros = TOTALS.compromisos ? (TOTALS.giros / TOTALS.compromisos) * 100 : 0;
    
    const dash = document.createElement("div");
    dash.id = "dash-08";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:16px;padding:28px;margin:20px 0;color:#fff;box-shadow:0 8px 32px rgba(26,39,66,.2);">
        <div style="display:flex;justify-content:space-between;">
          <div><div style="font-size:13px;letter-spacing:.12em;text-transform:uppercase;color:#8fb8e8;font-weight:600;">Planeación PDL 2026</div><div style="font-size:32px;font-weight:800;margin-top:8px;">${PROYECTOS.length}</div><div style="font-size:12px;color:#8fb8e8;">proyectos activos</div></div>
          <div style="text-align:right;"><div style="font-size:24px;font-weight:700;color:#5dade2;">${moneyCorto(TOTALS.poai)}</div><div style="font-size:11px;color:#8fb8e8;">POAI Total</div></div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px;">
        <div style="background:#667eea;color:#fff;padding:18px;border-radius:12px;"><div style="font-size:11px;text-transform:uppercase;opacity:.85;">Apropiaciones</div><div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.apropiaciones)}</div><div style="font-size:10px;opacity:.7;">100%</div></div>
        <div style="background:#17a2b8;color:#fff;padding:18px;border-radius:12px;"><div style="font-size:11px;text-transform:uppercase;opacity:.85;">Compromisos</div><div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.compromisos)}</div><div style="font-size:10px;opacity:.7;">${pctComp.toFixed(1)}%</div></div>
        <div style="background:#28a745;color:#fff;padding:18px;border-radius:12px;"><div style="font-size:11px;text-transform:uppercase;opacity:.85;">Giros</div><div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.giros)}</div><div style="font-size:10px;opacity:.7;">${pctGiros.toFixed(1)}%</div></div>
        <div style="background:#fd7e14;color:#fff;padding:18px;border-radius:12px;"><div style="font-size:11px;text-transform:uppercase;opacity:.85;">Pendiente</div><div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.compromisos - TOTALS.giros)}</div><div style="font-size:10px;opacity:.7;">${(100 - pctGiros).toFixed(1)}%</div></div>
      </div>
      <div style="background:#fff;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px;"><label style="display:block;font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">📋 Seleccionar Proyecto (${PROYECTOS.length} disponibles)</label><select id="c08-select" style="width:100%;padding:14px;border:2px solid #667eea;border-radius:8px;font-size:14px;color:#1a2742;font-weight:600;"><option value="">-- Selecciona un proyecto --</option></select></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;"><div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);"><div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">📊 Apropiaciones</div><canvas id="c08-poai"></canvas></div><div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);"><div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">📈 Cascada</div><canvas id="c08-cascada"></canvas></div></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;"><div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);"><div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">💰 Distribución</div><canvas id="c08-pie"></canvas></div><div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);"><div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">🎯 Semáforo</div><canvas id="c08-semaforo"></canvas></div></div>
      <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:15px;"><div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;">📋 Análisis Multidimensional</div><canvas id="c08-multi" style="max-height:300px;"></canvas></div>
      <div id="c08-detalle-wrap" style="background:#fff;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);"></div>
    `;
    contenedor.appendChild(dash);
    
    const select = document.getElementById("c08-select");
    PROYECTOS.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p.codigo;
      opt.textContent = p.codigo + " - " + moneyCorto(p.apropiaciones);
      select.appendChild(opt);
    });
    
    select.addEventListener("change", function() {
      if (this.value) {
        const proy = PROYECTOS.find(p => p.codigo == this.value);
        if (proy) mostrarDetalle(proy);
      } else {
        document.getElementById("c08-detalle-wrap").innerHTML = "";
      }
    });
    
    setTimeout(() => { graficaPOAI(); graficaCascada(); graficaPie(); graficaSemaforo(); graficaMulti(); }, 300);
  }
  
  function mostrarDetalle(p) {
    const wrap = document.getElementById("c08-detalle-wrap");
    const noGirado = Math.max(0, p.compromisos - p.giros);
    wrap.innerHTML = `<div style="padding:24px;"><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:12px;margin-bottom:20px;"><div style="background:#f0f6ff;padding:16px;border-radius:10px;border-left:4px solid #667eea;"><div style="font-size:9px;color:#667eea;font-weight:700;text-transform:uppercase;">Código</div><div style="font-size:18px;font-weight:800;color:#667eea;">${p.codigo}</div></div><div style="background:#f0f6ff;padding:16px;border-radius:10px;border-left:4px solid #667eea;"><div style="font-size:9px;color:#667eea;font-weight:700;text-transform:uppercase;">POAI 2026</div><div style="font-size:16px;font-weight:800;color:#667eea;">${money(p.poai)}</div></div><div style="background:#f0f8f5;padding:16px;border-radius:10px;border-left:4px solid #17a2b8;"><div style="font-size:9px;color:#17a2b8;font-weight:700;text-transform:uppercase;">Apropiaciones</div><div style="font-size:16px;font-weight:800;color:#17a2b8;">${money(p.apropiaciones)}</div></div><div style="background:#f0f8f5;padding:16px;border-radius:10px;border-left:4px solid #28a745;"><div style="font-size:9px;color:#28a745;font-weight:700;text-transform:uppercase;">Compromisos</div><div style="font-size:16px;font-weight:800;color:#28a745;">${money(p.compromisos)}</div></div><div style="background:#fff8f0;padding:16px;border-radius:10px;border-left:4px solid #fd7e14;"><div style="font-size:9px;color:#fd7e14;font-weight:700;text-transform:uppercase;">Giros</div><div style="font-size:16px;font-weight:800;color:#fd7e14;">${money(p.giros)}</div></div><div style="background:#fff5f5;padding:16px;border-radius:10px;border-left:4px solid #dc3545;"><div style="font-size:9px;color:#dc3545;font-weight:700;text-transform:uppercase;">Pendiente</div><div style="font-size:16px;font-weight:800;color:#dc3545;">${money(noGirado)}</div></div></div><div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;"><div style="background:linear-gradient(135deg,#667eea,#5569d8);color:#fff;padding:14px;border-radius:10px;text-align:center;"><div style="font-size:9px;opacity:.9;text-transform:uppercase;">% Comprometido</div><div style="font-size:20px;font-weight:800;">${(p.pctComp * 100).toFixed(1)}%</div></div><div style="background:linear-gradient(135deg,#28a745,#218838);color:#fff;padding:14px;border-radius:10px;text-align:center;"><div style="font-size:9px;opacity:.9;text-transform:uppercase;">% Girado</div><div style="font-size:20px;font-weight:800;">${(p.pctGiros * 100).toFixed(1)}%</div></div><div style="background:linear-gradient(135deg,#fd7e14,#e76905);color:#fff;padding:14px;border-radius:10px;text-align:center;"><div style="font-size:9px;opacity:.9;text-transform:uppercase;">% Pendiente</div><div style="font-size:20px;font-weight:800;">${(100 - p.pctGiros * 100).toFixed(1)}%</div></div></div></div>`;
  }
  
  function graficaPOAI() { const top = PROYECTOS.slice(0, 15); dibujar("c08-poai", "bar", { labels: top.map(p => p.codigo), datasets: [{ label: "Apropiaciones", data: top.map(p => p.apropiaciones), backgroundColor: "#667eea" }] }); }
  function graficaCascada() { dibujar("c08-cascada", "bar", { labels: ["POAI", "Aprop", "Comp", "Giros"], datasets: [{ label: "Valor", data: [TOTALS.poai, TOTALS.apropiaciones, TOTALS.compromisos, TOTALS.giros], backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14"] }] }, { plugins: { legend: { display: false } } }); }
  function graficaPie() { const top = PROYECTOS.slice(0, 8); dibujar("c08-pie", "doughnut", { labels: top.map(p => p.codigo), datasets: [{ data: top.map(p => p.apropiaciones), backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14", "#8e44ad", "#20c997", "#ffc107", "#dc3545"] }] }); }
  function graficaSemaforo() { const sorted = PROYECTOS.filter(p => p.apropiaciones > 0).sort((a, b) => b.pctComp - a.pctComp).slice(0, 10); const colores = sorted.map(p => p.pctComp >= 0.6 ? "#28a745" : p.pctComp >= 0.4 ? "#17a2b8" : "#dc3545"); const el = document.getElementById("c08-semaforo"); if (el && typeof Chart !== "undefined") { if (GRAF["c08-semaforo"]) GRAF["c08-semaforo"].destroy(); GRAF["c08-semaforo"] = new Chart(el, { type: "bar", data: { labels: sorted.map(p => p.codigo), datasets: [{ label: "% Ej", data: sorted.map(p => +(p.pctComp * 100).toFixed(1)), backgroundColor: colores }] }, options: { indexAxis: "y", responsive: true, plugins: { legend: { display: false } }, scales: { x: { max: 100 } } } }); } }
  function graficaMulti() { const top = PROYECTOS.slice(0, 10); dibujar("c08-multi", "bar", { labels: top.map(p => p.codigo), datasets: [{ label: "Aprop", data: top.map(p => p.apropiaciones), backgroundColor: "#667eea" }, { label: "Comp", data: top.map(p => p.compromisos), backgroundColor: "#28a745" }, { label: "Giros", data: top.map(p => p.giros), backgroundColor: "#fd7e14" }] }); }
  
  function dibujar(id, tipo, data, extra) { const el = document.getElementById(id); if (!el || typeof Chart === "undefined") return; if (GRAF[id]) GRAF[id].destroy(); const opts = { responsive: true, maintainAspectRatio: true, plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 }, padding: 12 } }, tooltip: { backgroundColor: "rgba(26, 39, 66, 0.9)", padding: 12, borderRadius: 8, callbacks: { label: function(ctx) { const v = ctx.parsed.y !== undefined ? ctx.parsed.y : ctx.parsed.x; return ctx.dataset.label + ": " + moneyCorto(v); } } } }, scales: tipo === "bar" ? { y: { ticks: { callback: function(v) { return moneyCorto(v); } } } } : {} }; if (extra) Object.assign(opts, extra); GRAF[id] = new Chart(el, { type: tipo, data: data, options: opts }); }
})();
