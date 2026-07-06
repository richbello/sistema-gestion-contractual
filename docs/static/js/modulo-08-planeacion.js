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
        if (PROYECTOS.length === 0) {
          alert("Sin datos de proyectos");
          return;
        }
        render(archivo.name);
      } catch (err) {
        console.error(err);
        alert("Error: " + err.message);
      }
    };
    reader.readAsArrayBuffer(archivo);
  }
  
  const num = (x) => {
    const n = parseFloat(String(x).replace(/[^\d.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };
  
  const money = (v) => new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(v);
  
  const moneyCorto = (v) => {
    if (Math.abs(v) >= 1e12) return "$" + (v / 1e12).toFixed(1) + "B";
    if (Math.abs(v) >= 1e9) return "$" + (v / 1e9).toFixed(1) + "MM";
    if (Math.abs(v) >= 1e6) return "$" + (v / 1e6).toFixed(1) + "M";
    if (Math.abs(v) >= 1e3) return "$" + (v / 1e3).toFixed(0) + "K";
    return "$" + v.toFixed(0);
  };
  
  const gcP = (row, keys) => {
    for (const k of keys) {
      if (row[k] != null && row[k] !== "") return row[k];
    }
    return "";
  };
  
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
      proyMap[cod] = {
        codigo: cod,
        poai: poai,
        apropiaciones: aprop,
        compromisos: comp,
        giros: giros,
        pctComp: aprop > 0 ? comp / aprop : 0
      };
    });
    
    PROYECTOS = Object.values(proyMap)
      .filter(p => p.poai > 0)
      .sort((a, b) => b.poai - a.poai);
    
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
    
    const dash = document.createElement("div");
    dash.id = "dash-08";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:12px;padding:20px;margin:20px 0;color:#fff;">
        <div style="font-size:12px;letter-spacing:.1em;text-transform:uppercase;color:#8fb8e8;">Planeación PDL</div>
        <div style="font-size:20px;font-weight:700;margin-top:8px;">${PROYECTOS.length} proyectos | ${moneyCorto(TOTALS.poai)}</div>
      </div>
      
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px;">
        <div style="background:#667eea;color:#fff;padding:16px;border-radius:10px;text-align:center;">
          <div style="font-size:10px;opacity:.9;">POAI 2026</div>
          <div style="font-size:16px;font-weight:700;margin-top:4px;">${moneyCorto(TOTALS.poai)}</div>
        </div>
        <div style="background:#17a2b8;color:#fff;padding:16px;border-radius:10px;text-align:center;">
          <div style="font-size:10px;opacity:.9;">Apropiaciones</div>
          <div style="font-size:16px;font-weight:700;margin-top:4px;">${moneyCorto(TOTALS.apropiaciones)}</div>
        </div>
        <div style="background:#28a745;color:#fff;padding:16px;border-radius:10px;text-align:center;">
          <div style="font-size:10px;opacity:.9;">Compromisos</div>
          <div style="font-size:16px;font-weight:700;margin-top:4px;">${moneyCorto(TOTALS.compromisos)}</div>
        </div>
        <div style="background:#fd7e14;color:#fff;padding:16px;border-radius:10px;text-align:center;">
          <div style="font-size:10px;opacity:.9;">Giros</div>
          <div style="font-size:16px;font-weight:700;margin-top:4px;">${moneyCorto(TOTALS.giros)}</div>
        </div>
        <div style="background:#8e44ad;color:#fff;padding:16px;border-radius:10px;text-align:center;">
          <div style="font-size:10px;opacity:.9;">% Ejecución</div>
          <div style="font-size:16px;font-weight:700;margin-top:4px;">${pctComp.toFixed(1)}%</div>
        </div>
      </div>
      
      <div style="background:#fff;padding:20px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:16px;">
        <div style="font-size:13px;font-weight:600;margin-bottom:12px;color:#1a2742;">📊 POAI 2026</div>
        <canvas id="c08-poai"></canvas>
      </div>
      
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Cascada</div>
          <canvas id="c08-cascada"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Ejecución</div>
          <canvas id="c08-ejec"></canvas>
        </div>
      </div>
      
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Distribución</div>
          <canvas id="c08-pie"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Semáforo</div>
          <canvas id="c08-semaforo"></canvas>
        </div>
      </div>
      
      <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:16px;">
        <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Top 10</div>
        <canvas id="c08-multi"></canvas>
      </div>
      
      <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:16px;">
        <div style="font-size:13px;font-weight:600;margin-bottom:8px;color:#1a2742;">Proyectos POAI</div>
        <canvas id="c08-top"></canvas>
      </div>
      
      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <input id="c08-busca" placeholder="🔎 Buscar por código..." style="flex:1;padding:11px;border:1.5px solid #e8edf2;border-radius:8px;font-size:13px;" />
        <button id="c08-btn-tabla" style="padding:11px 18px;background:#667eea;color:#fff;border:none;border-radius:8px;cursor:pointer;font-weight:600;">📋 Tabla</button>
      </div>
      <div id="c08-tabla-wrap" style="display:none;background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 10px rgba(0,0,0,.06);overflow-x:auto;"></div>
    `;
    
    contenedor.appendChild(dash);
    
    setTimeout(function() {
      graficaPOAI();
      graficaCascada();
      graficaEjecucion();
      graficaPie();
      graficaMulti();
      graficaSemaforo();
      graficaTop();
    }, 500);
    
    const btnTabla = document.getElementById("c08-btn-tabla");
    if (btnTabla) {
      btnTabla.addEventListener("click", function() {
        const w = document.getElementById("c08-tabla-wrap");
        w.style.display = w.style.display === "none" ? "block" : "none";
        if (w.style.display === "block") pintarTabla(PROYECTOS);
      });
    }
    
    const busca = document.getElementById("c08-busca");
    if (busca) {
      busca.addEventListener("input", function(e) {
        const q = e.target.value.toLowerCase();
        const filtrados = PROYECTOS.filter(p => String(p.codigo).includes(q));
        document.getElementById("c08-tabla-wrap").style.display = "block";
        pintarTabla(filtrados);
      });
    }
  }
  
  function graficaPOAI() {
    const top = PROYECTOS.slice(0, 15);
    dibujar("c08-poai", "bar", {
      labels: top.map(p => String(p.codigo)),
      datasets: [{
        label: "POAI",
        data: top.map(p => p.poai),
        backgroundColor: "#667eea"
      }]
    });
  }
  
  function graficaCascada() {
    dibujar("c08-cascada", "bar", {
      labels: ["POAI", "Aprop", "Comp", "Giros"],
      datasets: [{
        label: "Valor",
        data: [TOTALS.poai, TOTALS.apropiaciones, TOTALS.compromisos, TOTALS.giros],
        backgroundColor: ["#667eea", "#17a2b8", "#28a745", "#fd7e14"]
      }]
    }, { plugins: { legend: { display: false } } });
  }
  
  function graficaEjecucion() {
    const noComp = Math.max(0, TOTALS.apropiaciones - TOTALS.compromisos);
    const compSinGiro = Math.max(0, TOTALS.compromisos - TOTALS.giros);
    dibujar("c08-ejec", "doughnut", {
      labels: ["Girado", "Pendiente", "Sin comp"],
      datasets: [{
        data: [TOTALS.giros, compSinGiro, noComp],
        backgroundColor: ["#28a745", "#ffc107", "#dc3545"]
      }]
    });
  }
  
  function graficaPie() {
    const top = PROYECTOS.slice(0, 8);
    dibujar("c08-pie", "pie", {
      labels: top.map(p => String(p.codigo)),
      datasets: [{
        data: top.map(p => p.poai),
        backgroundColor: ["#667eea", "#28a745", "#ffc107", "#dc3545", "#17a2b8", "#8e44ad", "#fd7e14", "#20c997"]
      }]
    });
  }
  
  function graficaMulti() {
    const top = PROYECTOS.slice(0, 10);
    dibujar("c08-multi", "bar", {
      labels: top.map(p => String(p.codigo)),
      datasets: [
        { label: "POAI", data: top.map(p => p.poai), backgroundColor: "#667eea" },
        { label: "Aprop", data: top.map(p => p.apropiaciones), backgroundColor: "#17a2b8" },
        { label: "Comp", data: top.map(p => p.compromisos), backgroundColor: "#28a745" }
      ]
    });
  }
  
  function graficaSemaforo() {
    const sorted = PROYECTOS.filter(p => p.apropiaciones > 0).sort((a, b) => b.pctComp - a.pctComp).slice(0, 10);
    const colores = sorted.map(p => p.pctComp >= 0.6 ? "#28a745" : p.pctComp >= 0.4 ? "#17a2b8" : "#dc3545");
    const el = document.getElementById("c08-semaforo");
    
    if (el && typeof Chart !== "undefined") {
      if (GRAF["c08-semaforo"]) GRAF["c08-semaforo"].destroy();
      GRAF["c08-semaforo"] = new Chart(el, {
        type: "bar",
        data: {
          labels: sorted.map(p => String(p.codigo)),
          datasets: [{
            label: "% Ej",
            data: sorted.map(p => +(p.pctComp * 100).toFixed(1)),
            backgroundColor: colores
          }]
        },
        options: {
          indexAxis: "y",
          responsive: true,
          maintainAspectRatio: true,
          plugins: { legend: { display: false } },
          scales: { x: { max: 100 } }
        }
      });
    }
  }
  
  function graficaTop() {
    const top = PROYECTOS.slice(0, 15);
    dibujar("c08-top", "bar", {
      labels: top.map(p => String(p.codigo)),
      datasets: [{
        label: "POAI",
        data: top.map(p => p.poai),
        backgroundColor: "#667eea"
      }]
    }, { indexAxis: "y" });
  }
  
  function dibujar(id, tipo, data, extra) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    
    if (GRAF[id]) GRAF[id].destroy();
    
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const v = ctx.parsed.y !== undefined ? ctx.parsed.y : ctx.parsed.x !== undefined ? ctx.parsed.x : ctx.parsed;
              return ctx.dataset.label ? ctx.dataset.label + ": " + moneyCorto(v) : moneyCorto(v);
            }
          }
        }
      },
      scales: tipo === "bar" ? {
        y: { ticks: { callback: function(v) { return moneyCorto(v); } } }
      } : {}
    };
    
    if (extra) Object.assign(opts, extra);
    
    GRAF[id] = new Chart(el, {
      type: tipo,
      data: data,
      options: opts
    });
  }
  
  function pintarTabla(filas) {
    const w = document.getElementById("c08-tabla-wrap");
    if (!w) return;
    
    const cols = [["codigo", "Cód"], ["poai", "POAI 2026"], ["apropiaciones", "Aprop"], ["compromisos", "Comp"], ["giros", "Giros"]];
    let html = '<table style="width:100%;border-collapse:collapse;font-size:11px;"><thead><tr>';
    
    cols.forEach(c => html += '<th style="text-align:left;padding:8px;background:#1a2742;color:#fff;">' + c[1] + '</th>');
    html += "</tr></thead><tbody>";
    
    filas.forEach(function(r, i) {
      html += '<tr style="background:' + (i % 2 ? "#f7f9fc" : "#fff") + ';">';
      cols.forEach(function(c) {
        const campo = c[0];
        let val = r[campo];
        if (["poai", "apropiaciones", "compromisos", "giros"].includes(campo)) {
          val = money(num(val));
        }
        html += '<td style="padding:6px 8px;border-bottom:1px solid #eef2f7;">' + (val || "") + '</td>';
      });
      html += "</tr>";
    });
    
    html += "</tbody></table>";
    if (filas.length === 0) html = '<div style="padding:20px;color:#888;">Sin resultados.</div>';
    w.innerHTML = html;
  }
  
  console.log("✅ Módulo 08 iniciado");
})();
