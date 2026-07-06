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
        const names = wb.SheetNames;
        const ejN = names.find(n => n.toLowerCase().includes("ejec")) || names[0];
        const ejRaw = XLSX.utils.sheet_to_json(wb.Sheets[ejN], { header: 1 });
        
        let ejRows = [];
        if (ejRaw.length > 1) {
          const heads = ejRaw[0].map(h => String(h || "").trim());
          for (let i = 1; i < ejRaw.length; i++) {
            const row = ejRaw[i];
            if (!row.some(c => c != null && c !== "")) continue;
            const o = {};
            heads.forEach((h, j) => o[h] = row[j] || "");
            ejRows.push(o);
          }
        }
        
        procesar(ejRows);
        fileName = archivo.name;
        if (PROYECTOS.length === 0) { alert("Sin datos"); return; }
        render(archivo.name);
      } catch (err) { console.error(err); alert("Error: " + err.message); }
    };
    reader.readAsArrayBuffer(archivo);
  }
  
  const num = (x) => { const n = parseFloat(String(x).replace(/[^\d.-]/g, "")); return isNaN(n) ? 0 : n; };
  const money = (v) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);
  const moneyCorto = (v) => {
    if (Math.abs(v) >= 1e12) return "$" + (v / 1e12).toFixed(1) + "B";
    if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM";
    if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M";
    if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K";
    return "$" + v.toFixed(0);
  };
  const gcP = (row, keys) => { for (const k of keys) { if (row[k] != null && row[k] !== "") return row[k]; } return ""; };
  
  function procesar(ejec) {
    const proyMap = {};
    ejec.forEach(row => {
      const cod = num(gcP(row, ["Código_Proyecto", "Codigo_Proyecto"]));
      if (!cod) return;
      const poai = num(gcP(row, ["POAI 2026", "Recursos POAI 2026"]));
      const aprop = num(gcP(row, ["Apropiaciones"]));
      const comp = num(gcP(row, ["Compromisos proyecto", "Compromisos"]));
      const giros = num(gcP(row, ["Giros proyecto", "Giros"]));
      if (poai === 0) return;
      proyMap[cod] = { codigo: cod, poai, apropiaciones: aprop, compromisos: comp, giros, pctComp: aprop > 0 ? comp / aprop : 0 };
    });
    PROYECTOS = Object.values(proyMap).filter(p => p.poai > 0).sort((a, b) => b.poai - a.poai);
    TOTALS = {
      poai: PROYECTOS.reduce((s, p) => s + p.poai, 0),
      apropiaciones: PROYECTOS.reduce((s, p) => s + p.apropiaciones, 0),
      compromisos: PROYECTOS.reduce((s, p) => s + p.compromisos, 0),
      giros: PROYECTOS.reduce((s, p) => s + p.giros, 0)
    };
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
    const pctGiros = TOTALS.compromisos ? (TOTALS.giros / TOTALS.compromisos) * 100 : 0;
    
    const dash = document.createElement("div");
    dash.id = "dash-08";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742 0%,#1e3a6e 100%);border-radius:16px;padding:28px;margin:20px 0;color:#fff;box-shadow:0 8px 32px rgba(26,39,66,.2);">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div>
            <div style="font-size:13px;letter-spacing:.12em;text-transform:uppercase;color:#8fb8e8;font-weight:600;">Planeación PDL 2026</div>
            <div style="font-size:32px;font-weight:800;margin-top:8px;line-height:1;">${PROYECTOS.length}</div>
            <div style="font-size:12px;color:#8fb8e8;margin-top:4px;">proyectos activos</div>
          </div>
          <div style="text-align:right;">
            <div style="font-size:24px;font-weight:700;color:#5dade2;">${moneyCorto(TOTALS.poai)}</div>
            <div style="font-size:11px;color:#8fb8e8;margin-top:4px;">POAI Total</div>
          </div>
        </div>
      </div>
      
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px;">
        <div style="background:#667eea;color:#fff;padding:18px;border-radius:12px;box-shadow:0 4px 12px rgba(102,126,234,.15);">
          <div style="font-size:11px;text-transform:uppercase;opacity:.85;letter-spacing:.05em;">Apropiaciones</div>
          <div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.apropiaciones)}</div>
          <div style="font-size:10px;opacity:.7;margin-top:4px;">100%</div>
        </div>
        <div style="background:#17a2b8;color:#fff;padding:18px;border-radius:12px;box-shadow:0 4px 12px rgba(23,162,184,.15);">
          <div style="font-size:11px;text-transform:uppercase;opacity:.85;letter-spacing:.05em;">Compromisos</div>
          <div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.compromisos)}</div>
          <div style="font-size:10px;opacity:.7;margin-top:4px;">${pctComp.toFixed(1)}%</div>
        </div>
        <div style="background:#28a745;color:#fff;padding:18px;border-radius:12px;box-shadow:0 4px 12px rgba(40,167,69,.15);">
          <div style="font-size:11px;text-transform:uppercase;opacity:.85;letter-spacing:.05em;">Giros</div>
          <div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.giros)}</div>
          <div style="font-size:10px;opacity:.7;margin-top:4px;">${pctGiros.toFixed(1)}%</div>
        </div>
        <div style="background:#fd7e14;color:#fff;padding:18px;border-radius:12px;box-shadow:0 4px 12px rgba(253,126,20,.15);">
          <div style="font-size:11px;text-transform:uppercase;opacity:.85;letter-spacing:.05em;">Pendiente de Giro</div>
          <div style="font-size:18px;font-weight:700;margin-top:8px;">${moneyCorto(TOTALS.compromisos - TOTALS.giros)}</div>
          <div style="font-size:10px;opacity:.7;margin-top:4px;">${(100 - pctGiros).toFixed(1)}%</div>
        </div>
      </div>
      
      <div style="background:#fff;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px;">
        <label style="display:block;font-size:12px;font-weight:600;text-transform:uppercase;color:#1a2742;margin-bottom:8px;letter-spacing:.05em;">Seleccionar proyecto</label>
        <select id="c08-select" style="width:100%;padding:12px;border:2px solid #e8edf2;border-radius:8px;font-size:14px;background:#fff;cursor:pointer;color:#1a2742;font-weight:500;">
          <option value="">-- Todos los proyectos --</option>
        </select>
      </div>
      
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
          <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;letter-spacing:.05em;">📊 POAI por Proyecto</div>
          <canvas id="c08-poai" style="max-height:280px;"></canvas>
        </div>
        <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
          <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;letter-spacing:.05em;">📈 Cascada Presupuestal</div>
          <canvas id="c08-cascada" style="max-height:280px;"></canvas>
        </div>
      </div>
      
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
          <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;letter-spacing:.05em;">💰 Distribución POAI</div>
          <canvas id="c08-pie" style="max-height:280px;"></canvas>
        </div>
        <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
          <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;letter-spacing:.05em;">🎯 % Ejecución (Semáforo)</div>
          <canvas id="c08-semaforo" style="max-height:280px;"></canvas>
        </div>
      </div>
      
      <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:15px;">
        <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;margin-bottom:12px;letter-spacing:.05em;">📋 Análisis Multidimensional (Top 10)</div>
        <canvas id="c08-multi" style="max-height:300px;"></canvas>
      </div>
      
      <div style="background:#fff;padding:18px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:15px;">
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
          <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:#1a2742;letter-spacing:.05em;">🔎 Búsqueda Avanzada</div>
        </div>
        <input id="c08-busca" placeholder="Buscar por código, apropiaciones, compromisos..." style="width:100%;padding:12px;border:2px solid #e8edf2;border-radius:8px;font-size:13px;box-sizing:border-box;">
      </div>
      
      <div id="c08-tabla-wrap" style="background:#fff;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);overflow:hidden;"></div>
    `;
    
    contenedor.appendChild(dash);
    
    // Llenar select
    const select = document.getElementById("c08-select");
    PROYECTOS.slice(0, 20).forEach(p => {
      const opt = document.createElement("option");
      opt.value = p.codigo;
      opt.textContent = `Cód ${p.codigo} - ${moneyCorto(p.poai)}`;
      select.appendChild(opt);
    });
    
    select.addEventListener("change", function() {
      if (this.value) {
        const proy = PROYECTOS.find(p => p.codigo == this.value);
        if (proy) mostrarDetalleProyecto(proy);
      } else {
        document.getElementById("c08-tabla-wrap").innerHTML = "";
      }
    });
    
    document.getElementById("c08-busca").addEventListener("input", function(e) {
      const q = e.target.value.toLowerCase();
      const filtrados = PROYECTOS.filter(p => String(p.codigo).includes(q));
      pintarTablaBusqueda(filtrados);
    });
    
    setTimeout(function() {
      graficaPOAI();
      graficaCascada();
      graficaPie();
      graficaSemaforo();
      graficaMulti();
    }, 300);
  }
  
  function mostrarDetalleProyecto(proy) {
    const wrap = document.getElementById("c08-tabla-wrap");
    const pctComp = proy.apropiaciones > 0 ? (proy.compromisos / proy.apropiaciones) * 100 : 0;
    const pctGiros = proy.compromisos > 0 ? (proy.giros / proy.compromisos) * 100 : 0;
    const noGirado = Math.max(0, proy.compromisos - proy.giros);
    
    wrap.innerHTML = `
      <div style="padding:20px;border-bottom:1px solid #e8edf2;">
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-bottom:16px;">
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">Código</div>
            <div style="font-size:16px;font-weight:700;color:#1a2742;margin-top:4px;">${proy.codigo}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">POAI 2026</div>
            <div style="font-size:16px;font-weight:700;color:#667eea;margin-top:4px;">${money(proy.poai)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">Apropiaciones</div>
            <div style="font-size:16px;font-weight:700;color:#17a2b8;margin-top:4px;">${money(proy.apropiaciones)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">Compromisos</div>
            <div style="font-size:16px;font-weight:700;color:#28a745;margin-top:4px;">${money(proy.compromisos)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">Giros</div>
            <div style="font-size:16px;font-weight:700;color:#fd7e14;margin-top:4px;">${money(proy.giros)}</div>
          </div>
          <div>
            <div style="font-size:11px;color:#888;text-transform:uppercase;">Pendiente</div>
            <div style="font-size:16px;font-weight:700;color:#dc3545;margin-top:4px;">${money(noGirado)}</div>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:16px;">
          <div style="background:#f0f6ff;padding:12px;border-radius:8px;border-left:4px solid #667eea;">
            <div style="font-size:10px;color:#667eea;font-weight:600;">% Comprometido</div>
            <div style="font-size:18px;font-weight:700;color:#667eea;margin-top:4px;">${pctComp.toFixed(1)}%</div>
          </div>
          <div style="background:#f0f8f5;padding:12px;border-radius:8px;border-left:4px solid #28a745;">
            <div style="font-size:10px;color:#28a745;font-weight:600;">% Girado</div>
            <div style="font-size:18px;font-weight:700;color:#28a745;margin-top:4px;">${pctGiros.toFixed(1)}%</div>
          </div>
          <div style="background:#fff5f5;padding:12px;border-radius:8px;border-left:4px solid #dc3545;">
            <div style="font-size:10px;color:#dc3545;font-weight:600;">% Pendiente</div>
            <div style="font-size:18px;font-weight:700;color:#dc3545;margin-top:4px;">${(100 - pctGiros).toFixed(1)}%</div>
          </div>
        </div>
      </div>
    `;
  }
  
  function pintarTablaBusqueda(filas) {
    const wrap = document.getElementById("c08-tabla-wrap");
    if (!filas.length) {
      wrap.innerHTML = '<div style="padding:20px;color:#888;text-align:center;">Sin resultados</div>';
      return;
    }
    
    let html = '<table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#f8f9fa;"><th style="padding:12px;text-align:left;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">Código</th><th style="padding:12px;text-align:left;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">POAI 2026</th><th style="padding:12px;text-align:left;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">Apropiaciones</th><th style="padding:12px;text-align:left;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">Compromisos</th><th style="padding:12px;text-align:left;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">Giros</th><th style="padding:12px;text-align:center;font-size:11px;font-weight:700;color:#1a2742;text-transform:uppercase;border-bottom:2px solid #e8edf2;">% Ej</th></tr></thead><tbody>';
    
    filas.forEach((p, i) => {
      const pctComp = p.apropiaciones > 0 ? (p.compromisos / p.apropiaciones) * 100 : 0;
      html += `<tr style="background:${i % 2 ? '#fff' : '#f8f9fa'};border-bottom:1px solid #e8edf2;"><td style="padding:12px;font-size:13px;font-weight:600;color:#667eea;">${p.codigo}</td><td style="padding:12px;font-size:13px;color:#1a2742;">${money(p.poai)}</td><td style="padding:12px;font-size:13px;color:#1a2742;">${money(p.apropiaciones)}</td><td style="padding:12px;font-size:13px;color:#1a2742;">${money(p.compromisos)}</td><td style="padding:12px;font-size:13px;color:#1a2742;">${money(p.giros)}</td><td style="padding:12px;text-align:center;"><span style="background:${pctComp >= 75 ? '#d4edda' : pctComp >= 50 ? '#fff3cd' : '#f8d7da'};color:${pctComp >= 75 ? '#155724' : pctComp >= 50 ? '#856404' : '#721c24'};padding:4px 8px;border-radius:6px;font-weight:600;font-size:12px;">${pctComp.toFixed(1)}%</span></td></tr>`;
    });
    
    html += '</tbody></table>';
    wrap.innerHTML = html;
  }
  
  function graficaPOAI() {
    const top = PROYECTOS.slice(0, 15);
    dibujar("c08-poai", "bar", {
      labels: top.map(p => String(p.codigo)),
      datasets: [{ label: "POAI", data: top.map(p => p.poai), backgroundColor: "#667eea" }]
    });
  }
  
  function graficaCascada() {
    dibujar("c08-cascada", "bar", {
      labels: ["POAI", "Apropiaciones", "Compromisos", "Giros"],
      datasets: [{ label: "Valor", data: [TOTALS.poai, TOTALS.apropiaciones, TOTALS.compromisos, TOTALS.giros], backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14"] }]
    }, { plugins: { legend: { display: false } } });
  }
  
  function graficaPie() {
    const top = PROYECTOS.slice(0, 8);
    dibujar("c08-pie", "doughnut", {
      labels: top.map(p => String(p.codigo)),
      datasets: [{ data: top.map(p => p.poai), backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14", "#8e44ad", "#20c997", "#ffc107", "#dc3545"] }]
    });
  }
  
  function graficaSemaforo() {
    const sorted = PROYECTOS.filter(p => p.apropiaciones > 0).sort((a, b) => b.pctComp - a.pctComp).slice(0, 10);
    const colores = sorted.map(p => p.pctComp >= 0.6 ? "#28a745" : p.pctComp >= 0.4 ? "#17a2b8" : p.pctComp >= 0.2 ? "#ffc107" : "#dc3545");
    const el = document.getElementById("c08-semaforo");
    if (el && typeof Chart !== "undefined") {
      if (GRAF["c08-semaforo"]) GRAF["c08-semaforo"].destroy();
      GRAF["c08-semaforo"] = new Chart(el, {
        type: "bar",
        data: { labels: sorted.map(p => String(p.codigo)), datasets: [{ label: "% Ej", data: sorted.map(p => +(p.pctComp * 100).toFixed(1)), backgroundColor: colores }] },
        options: { indexAxis: "y", responsive: true, plugins: { legend: { display: false } }, scales: { x: { max: 100, ticks: { callback: function(v) { return v + "%"; } } } } }
      });
    }
  }
  
  function graficaMulti() {
    const top = PROYECTOS.slice(0, 10);
    dibujar("c08-multi", "bar", {
      labels: top.map(p => String(p.codigo)),
      datasets: [
        { label: "POAI", data: top.map(p => p.poai), backgroundColor: "#667eea" },
        { label: "Compromisos", data: top.map(p => p.compromisos), backgroundColor: "#28a745" },
        { label: "Giros", data: top.map(p => p.giros), backgroundColor: "#fd7e14" }
      ]
    });
  }
  
  function dibujar(id, tipo, data, extra) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    if (GRAF[id]) GRAF[id].destroy();
    
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 }, padding: 12 } },
        tooltip: { backgroundColor: "rgba(26, 39, 66, 0.9)", padding: 12, borderRadius: 8, titleFont: { size: 12, weight: "bold" }, bodyFont: { size: 12 }, callbacks: { label: function(ctx) { const v = ctx.parsed.y !== undefined ? ctx.parsed.y : ctx.parsed.x; return ctx.dataset.label + ": " + moneyCorto(v); } } }
      },
      scales: tipo === "bar" ? { y: { ticks: { callback: function(v) { return moneyCorto(v); } } } } : {}
    };
    
    if (extra) Object.assign(opts, extra);
    GRAF[id] = new Chart(el, { type: tipo, data: data, options: opts });
  }
  
  console.log("✅ Módulo 08 PROFESIONAL activado");
})();
