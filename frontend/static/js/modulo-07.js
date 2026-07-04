/* ============================================================
   MÓDULO 07 — Contratación (Análisis Integral)
   Adaptado a vanilla JS + Chart.js desde el componente React.
   Lee el Excel consolidado de contratos y muestra KPIs,
   5 gráficas, filtros combinados y tabla paginada.
   ============================================================ */
(function () {
  "use strict";

  let CONTRATOS = [];
  let RESUMEN = null;
  const GRAF = {};
  const PAL = ["#5dade2","#27ae60","#f39c12","#e74c3c","#8e44ad","#1abc9c","#e67e22","#d63031","#0984e3","#e84393"];
  let pagina = 0;
  const FILAS_PAG = 25;

  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();

  function init() {
    const input = document.getElementById("carga-07");
    if (!input || input.dataset.ligado === "1") return;
    input.dataset.ligado = "1";
    input.addEventListener("change", (e) => leerExcel(e.target.files[0]));
  }

  /* ---------- utilidades ---------- */
  const num = (x) => {
    const n = parseFloat(String(x).replace(/[^\d.-]/g, ""));
    return isNaN(n) ? 0 : n;
  };
  const fmtCOP = (v) =>
    v >= 1e9 ? "$" + (v / 1e9).toFixed(2) + "MM"
    : v >= 1e6 ? "$" + (v / 1e6).toFixed(1) + "M"
    : v >= 1e3 ? "$" + (v / 1e3).toFixed(0) + "K"
    : "$" + v.toFixed(0);
  const fmtNum = (v) => Number(v).toLocaleString("es-CO");

  function gc(headers, terms) {
    return headers.findIndex((h) => terms.some((t) => h.toUpperCase().includes(t.toUpperCase())));
  }

  function parseFecha(val) {
    if (!val || val === "") return "";
    if (val instanceof Date) {
      if (isNaN(val.getTime())) return "";
      return `${String(val.getDate()).padStart(2,"0")}/${String(val.getMonth()+1).padStart(2,"0")}/${val.getFullYear()}`;
    }
    if (typeof val === "number") {
      const dt = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (isNaN(dt.getTime())) return "";
      return `${String(dt.getUTCDate()).padStart(2,"0")}/${String(dt.getUTCMonth()+1).padStart(2,"0")}/${dt.getUTCFullYear()}`;
    }
    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
      const [y, m, d] = s.slice(0, 10).split("-");
      return `${d}/${m}/${y}`;
    }
    return s.slice(0, 10);
  }

  function fechaISO(f) {
    if (!f || !f.includes("/")) return "";
    const [d, m, y] = f.split("/");
    return `${y}-${m}-${d}`;
  }

  const estadoColor = (e) => {
    const map = {
      "LIQUIDADO":"#27ae60","TERMINADO":"#1abc9c","EN EJECUCIÓN":"#5dade2",
      "EN EJECUCION":"#5dade2","SUSPENDIDO":"#f39c12","RESCINDIDO":"#e74c3c","SIN ESTADO":"#7f8c8d",
    };
    return map[e] || "#8e44ad";
  };

  /* ---------- lectura ---------- */
  function leerExcel(archivo) {
    if (!archivo) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        procesar(raw);
        if (CONTRATOS.length === 0) {
          alert("No se encontraron contratos. Verifica que el Excel tenga las columnas esperadas.");
          return;
        }
        render(archivo.name);
      } catch (err) {
        console.error(err);
        alert("Error leyendo archivo: " + err.message);
      }
    };
    reader.readAsArrayBuffer(archivo);
  }

  function procesar(raw) {
    const headers = raw[0].map((h) => String(h ?? "").trim().toUpperCase());
    const rows = raw.slice(1);

    const iCon = gc(headers, ["CONTRATO"]);
    const iNom = gc(headers, ["NOMBRE DEL CONTRATISTA", "NOMBRE DEL", "CONTRATISTA"]);
    const iTip = gc(headers, ["TIPO DE IDENTIFICACION", "TIPO DE ID"]);
    const iFIni = gc(headers, ["FECHA INICIAL DE CONTRATO", "FECHA INICIAL"]);
    const iFTermI = gc(headers, ["FECHA DE TERMINACION INICIAL", "TERMINACION INICIAL"]);
    const iVIni = gc(headers, ["VALOR INICIAL DEL CONTRATO", "VALOR INICIAL"]);
    const iVFin = gc(headers, ["VALOR FINAL DEL CONTRATO", "VALOR FINAL"]);
    const iFTermF = gc(headers, ["FECHA DE TERMINACION FINAL", "TERMINACION FINAL"]);
    const iEst = gc(headers, ["ESTADO ACTUAL DEL CONTRATO", "ESTADO ACTUAL", "ESTADO"]);

    CONTRATOS = [];
    let tIni = 0, tFin = 0, tVar = 0;

    for (const row of rows) {
      const id = String(row[iCon] ?? "").trim();
      if (!id) continue;
      const vIni = num(row[iVIni]);
      const vFin = num(row[iVFin]);
      const variacion = vFin - vIni;
      const fIni = parseFecha(row[iFIni]);
      CONTRATOS.push({
        contrato: id,
        contratista: String(row[iNom] ?? "").trim() || "N/D",
        tipoId: String(row[iTip] ?? "").trim(),
        fechaInicial: fIni,
        fechaTermInicial: parseFecha(row[iFTermI]),
        valorInicial: vIni,
        valorFinal: vFin,
        fechaTermFinal: parseFecha(row[iFTermF]),
        estado: String(row[iEst] ?? "").trim().toUpperCase() || "SIN ESTADO",
        variacion,
        pctVariacion: vIni > 0 ? (variacion / vIni) * 100 : 0,
      });
      tIni += vIni; tFin += vFin; tVar += variacion;
    }

    CONTRATOS.sort((a, b) => b.valorFinal - a.valorFinal);

    RESUMEN = {
      totalContratos: CONTRATOS.length,
      valorInicialTotal: tIni,
      valorFinalTotal: tFin,
      variacionTotal: tVar,
      pctVariacionGlobal: tIni > 0 ? (tVar / tIni) * 100 : 0,
      contratosSobrecosto: CONTRATOS.filter((c) => c.variacion > 0).length,
      estadosUnicos: [...new Set(CONTRATOS.map((c) => c.estado))].sort(),
      tiposIdUnicos: [...new Set(CONTRATOS.map((c) => c.tipoId).filter(Boolean))].sort(),
    };
  }

  /* ---------- agrupaciones sobre lista filtrada ---------- */
  function porEstado(lista) {
    const m = {};
    lista.forEach((c) => {
      if (!m[c.estado]) m[c.estado] = { count: 0, valor: 0 };
      m[c.estado].count++; m[c.estado].valor += c.valorFinal;
    });
    return Object.entries(m).map(([estado, v]) => ({ estado, ...v })).sort((a, b) => b.valor - a.valor);
  }
  function topContratistas(lista) {
    const m = {};
    lista.forEach((c) => {
      if (!m[c.contratista]) m[c.contratista] = { valor: 0, count: 0, variacion: 0 };
      m[c.contratista].valor += c.valorFinal;
      m[c.contratista].count++;
      m[c.contratista].variacion += c.variacion;
    });
    return Object.entries(m).map(([nombre, v]) => ({ nombre, ...v })).sort((a, b) => b.valor - a.valor).slice(0, 10);
  }
  function porMes(lista) {
    const m = {};
    lista.forEach((c) => {
      if (!c.fechaInicial || !c.fechaInicial.includes("/")) return;
      const [, mes, y] = c.fechaInicial.split("/");
      const k = `${y}-${mes}`;
      if (!m[k]) m[k] = { valor: 0, count: 0 };
      m[k].valor += c.valorFinal; m[k].count++;
    });
    return Object.entries(m).sort((a, b) => a[0].localeCompare(b[0])).map(([mes, v]) => ({ mes, ...v })).slice(-24);
  }

  /* ---------- estado de filtros ---------- */
  const F = { busqueda: "", estado: "TODOS", tipoId: "TODOS", desde: "", hasta: "" };

  function filtrar() {
    const q = F.busqueda.toLowerCase().trim();
    return CONTRATOS.filter((c) => {
      const okQ = !q || c.contrato.toLowerCase().includes(q) || c.contratista.toLowerCase().includes(q) || c.tipoId.toLowerCase().includes(q);
      const okE = F.estado === "TODOS" || c.estado === F.estado;
      const okT = F.tipoId === "TODOS" || c.tipoId === F.tipoId;
      const fi = fechaISO(c.fechaInicial);
      const okD = !F.desde || (fi && fi >= F.desde);
      const okH = !F.hasta || (fi && fi <= F.hasta);
      return okQ && okE && okT && okD && okH;
    });
  }

  /* ---------- render ---------- */
  function render(nombreArchivo) {
    const vista = document.getElementById("vista-contratacion");
    if (!vista) return;

    const cargaDecor = document.getElementById("pantalla-carga-07");
    if (cargaDecor) cargaDecor.style.display = "none";

    const previo = document.getElementById("dash-07");
    if (previo) previo.remove();

    const dash = document.createElement("div");
    dash.id = "dash-07";
    dash.innerHTML = `
      <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);border-radius:12px;padding:18px 22px;margin:18px 0;color:#fff;">
        <div style="font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#8fb8e8;">Contratación · ${nombreArchivo}</div>
        <div style="font-size:22px;font-weight:700;margin-top:4px;">${fmtNum(RESUMEN.totalContratos)} contratos cargados</div>
      </div>

      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:18px;">
        ${kpi("Total contratos", fmtNum(RESUMEN.totalContratos), "#5dade2")}
        ${kpi("Valor Final Total", fmtCOP(RESUMEN.valorFinalTotal), "#27ae60")}
        ${kpi("Valor Inicial Total", fmtCOP(RESUMEN.valorInicialTotal), "#f39c12")}
        ${kpi("Con sobrecosto", fmtNum(RESUMEN.contratosSobrecosto), "#e74c3c")}
        ${kpi("Estados distintos", RESUMEN.estadosUnicos.length, "#8e44ad")}
      </div>

      <div style="background:#fff;border-radius:12px;padding:16px 20px;margin-bottom:16px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
        <div style="font-size:13px;font-weight:700;color:#1a2742;margin-bottom:10px;">🔍 Filtros</div>
        <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px;">
          <input id="f07-busca" placeholder="🔎 Contrato, contratista o tipo ID..." style="flex:2;min-width:220px;font-size:12px;padding:8px 12px;border-radius:8px;border:1.5px solid #e8edf2;">
          <select id="f07-estado" style="flex:1;min-width:160px;font-size:12px;padding:8px 12px;border-radius:8px;border:1.5px solid #e8edf2;">
            <option value="TODOS">📋 Todos los estados</option>
            ${RESUMEN.estadosUnicos.map((e) => `<option value="${e}">${e}</option>`).join("")}
          </select>
          <select id="f07-tipo" style="flex:1;min-width:160px;font-size:12px;padding:8px 12px;border-radius:8px;border:1.5px solid #e8edf2;">
            <option value="TODOS">🪪 Todos los tipos ID</option>
            ${RESUMEN.tiposIdUnicos.map((t) => `<option value="${t}">${t}</option>`).join("")}
          </select>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;">
          <span style="font-size:11px;color:#7f8c8d;">📅 Inicio desde</span>
          <input type="date" id="f07-desde" style="font-size:12px;padding:7px 10px;border-radius:8px;border:1.5px solid #e8edf2;">
          <span style="font-size:11px;color:#7f8c8d;">hasta</span>
          <input type="date" id="f07-hasta" style="font-size:12px;padding:7px 10px;border-radius:8px;border:1.5px solid #e8edf2;">
          <button id="f07-limpiar" style="padding:7px 14px;border-radius:8px;border:1px solid #e8edf2;background:#fadbd8;color:#e74c3c;font-weight:700;font-size:12px;cursor:pointer;">✕ Limpiar</button>
          <span id="f07-conteo" style="margin-left:auto;font-size:11px;color:#7f8c8d;"></span>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;color:#1a2742;margin-bottom:8px;">Distribución por Estado</div>
          <canvas id="c07-estado"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;color:#1a2742;margin-bottom:8px;">Top 10 Contratistas por Valor Final</div>
          <canvas id="c07-contratistas"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;color:#1a2742;margin-bottom:8px;">Valor Inicial vs Final (Top 15)</div>
          <canvas id="c07-inifin"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);">
          <div style="font-size:13px;font-weight:600;color:#1a2742;margin-bottom:8px;">Contratos por Estado (cantidad)</div>
          <canvas id="c07-estadocount"></canvas>
        </div>
        <div style="background:#fff;padding:16px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06);grid-column:1/-1;">
          <div style="font-size:13px;font-weight:600;color:#1a2742;margin-bottom:8px;">Contratos por Fecha Inicial</div>
          <canvas id="c07-timeline"></canvas>
        </div>
      </div>

      <div style="background:#fff;border-radius:12px;padding:14px;box-shadow:0 2px 10px rgba(0,0,0,.06);overflow-x:auto;">
        <div id="c07-tabla"></div>
        <div id="c07-paginacion" style="display:flex;justify-content:center;gap:6px;margin-top:14px;"></div>
      </div>
    `;
    vista.appendChild(dash);

    // listeners de filtros
    const reaccionar = () => { pagina = 0; actualizar(); };
    document.getElementById("f07-busca").addEventListener("input", (e) => { F.busqueda = e.target.value; reaccionar(); });
    document.getElementById("f07-estado").addEventListener("change", (e) => { F.estado = e.target.value; reaccionar(); });
    document.getElementById("f07-tipo").addEventListener("change", (e) => { F.tipoId = e.target.value; reaccionar(); });
    document.getElementById("f07-desde").addEventListener("change", (e) => { F.desde = e.target.value; reaccionar(); });
    document.getElementById("f07-hasta").addEventListener("change", (e) => { F.hasta = e.target.value; reaccionar(); });
    document.getElementById("f07-limpiar").addEventListener("click", () => {
      F.busqueda = ""; F.estado = "TODOS"; F.tipoId = "TODOS"; F.desde = ""; F.hasta = "";
      document.getElementById("f07-busca").value = "";
      document.getElementById("f07-estado").value = "TODOS";
      document.getElementById("f07-tipo").value = "TODOS";
      document.getElementById("f07-desde").value = "";
      document.getElementById("f07-hasta").value = "";
      reaccionar();
    });

    actualizar();
    dash.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function kpi(label, valor, color) {
    return `<div style="background:#fff;border-left:3px solid ${color};padding:14px 16px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,.06);">
      <div style="font-size:22px;font-weight:800;color:${color};line-height:1;">${valor}</div>
      <div style="font-size:11px;font-weight:600;color:#2c3e50;margin-top:4px;">${label}</div>
    </div>`;
  }

  /* ---------- actualizar (gráficas + tabla) ---------- */
  function actualizar() {
    const lista = filtrar();
    const total = lista.reduce((s, c) => s + c.valorFinal, 0);
    const conteo = document.getElementById("f07-conteo");
    if (conteo) conteo.innerHTML = `<strong style="color:#1a2742;">${fmtNum(lista.length)}</strong> contratos · <strong style="color:#5dade2;">${fmtCOP(total)}</strong>`;

    // gráficas
    const est = porEstado(lista);
    dibujar("c07-estado", "doughnut", {
      labels: est.map((x) => x.estado.slice(0, 20)),
      datasets: [{ data: est.map((x) => x.valor), backgroundColor: est.map((_, i) => PAL[i % PAL.length]) }],
    });

    const cont = topContratistas(lista);
    dibujar("c07-contratistas", "bar", {
      labels: cont.map((x) => (x.nombre.length > 26 ? x.nombre.slice(0, 26) + "…" : x.nombre)),
      datasets: [{ label: "Valor Final", data: cont.map((x) => x.valor), backgroundColor: "#5dade2" }],
    }, { indexAxis: "y" });

    const top15 = lista.slice(0, 15);
    dibujar("c07-inifin", "bar", {
      labels: top15.map((c) => c.contrato.slice(0, 12)),
      datasets: [
        { label: "Valor Inicial", data: top15.map((c) => c.valorInicial), backgroundColor: "#5dade299" },
        { label: "Valor Final", data: top15.map((c) => c.valorFinal), backgroundColor: "#5dade2" },
      ],
    });

    dibujar("c07-estadocount", "bar", {
      labels: est.map((x) => x.estado.slice(0, 18)),
      datasets: [{ label: "N° Contratos", data: est.map((x) => x.count), backgroundColor: est.map((_, i) => PAL[i % PAL.length] + "bb") }],
    }, { indexAxis: "y", esConteo: true });

    const meses = porMes(lista);
    dibujar("c07-timeline", "line", {
      labels: meses.map((m) => m.mes),
      datasets: [{ label: "Valor Final", data: meses.map((m) => m.valor), borderColor: "#5dade2", backgroundColor: "rgba(93,173,226,.2)", tension: 0.3, fill: true }],
    });

    tabla(lista);
  }

  function dibujar(id, tipo, data, extra) {
    const el = document.getElementById(id);
    if (!el || typeof Chart === "undefined") return;
    if (GRAF[id]) GRAF[id].destroy();
    const esConteo = extra && extra.esConteo;
    const opts = {
      responsive: true,
      maintainAspectRatio: true,
      indexAxis: extra && extra.indexAxis ? extra.indexAxis : "x",
      plugins: {
        legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const v = ctx.parsed.y ?? ctx.parsed.x ?? ctx.parsed;
              if (esConteo) return `${ctx.dataset.label}: ${fmtNum(v)}`;
              return ctx.dataset.label ? `${ctx.dataset.label}: ${fmtCOP(v)}` : fmtCOP(v);
            },
          },
        },
      },
      scales: tipo === "bar"
        ? (extra && extra.indexAxis === "y"
            ? { x: { ticks: { callback: (v) => (esConteo ? fmtNum(v) : fmtCOP(v)) } } }
            : { y: { ticks: { callback: (v) => (esConteo ? fmtNum(v) : fmtCOP(v)) } } })
        : (tipo === "line" ? { y: { ticks: { callback: (v) => fmtCOP(v) } } } : {}),
    };
    GRAF[id] = new Chart(el, { type: tipo, data, options: opts });
  }

  /* ---------- tabla paginada ---------- */
  function tabla(lista) {
    const cont = document.getElementById("c07-tabla");
    const pag = document.getElementById("c07-paginacion");
    if (!cont) return;

    const paginas = Math.ceil(lista.length / FILAS_PAG) || 1;
    if (pagina >= paginas) pagina = paginas - 1;
    const filas = lista.slice(pagina * FILAS_PAG, (pagina + 1) * FILAS_PAG);

    const cols = ["Contrato", "Contratista", "Tipo ID", "Fecha Inicial", "Term. Inicial", "Valor Inicial", "Valor Final", "Term. Final", "Estado"];
    let html = '<table style="width:100%;border-collapse:collapse;font-size:11px;"><thead><tr style="background:#1a2742;color:#fff;">';
    cols.forEach((c) => (html += `<th style="padding:9px 10px;text-align:left;font-size:10px;white-space:nowrap;">${c}</th>`));
    html += "</tr></thead><tbody>";
    filas.forEach((c, i) => {
      const col = estadoColor(c.estado);
      html += `<tr style="border-bottom:1px solid #e8edf2;background:${i % 2 ? "#f0f2f5" : "#fff"};">
        <td style="padding:7px 10px;font-weight:700;color:#1a2742;white-space:nowrap;">${c.contrato}</td>
        <td style="padding:7px 10px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${c.contratista}">${c.contratista}</td>
        <td style="padding:7px 10px;color:#7f8c8d;white-space:nowrap;">${c.tipoId}</td>
        <td style="padding:7px 10px;white-space:nowrap;color:#27ae60;">${c.fechaInicial || "—"}</td>
        <td style="padding:7px 10px;white-space:nowrap;color:#7f8c8d;">${c.fechaTermInicial || "—"}</td>
        <td style="padding:7px 10px;text-align:right;white-space:nowrap;">${fmtCOP(c.valorInicial)}</td>
        <td style="padding:7px 10px;text-align:right;font-weight:700;color:#5dade2;white-space:nowrap;">${fmtCOP(c.valorFinal)}</td>
        <td style="padding:7px 10px;white-space:nowrap;color:${c.variacion > 0 ? "#e74c3c" : "#7f8c8d"};">${c.fechaTermFinal || "—"}</td>
        <td style="padding:7px 10px;"><span style="background:${col}22;color:${col};border:1px solid ${col}44;border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700;white-space:nowrap;">${c.estado.slice(0, 18)}</span></td>
      </tr>`;
    });
    html += "</tbody></table>";
    cont.innerHTML = html;

    // paginación
    if (paginas > 1) {
      let p = "";
      p += botonPag("«", 0, pagina === 0);
      p += botonPag("‹", pagina - 1, pagina === 0);
      const inicio = paginas <= 7 ? 0 : Math.max(0, Math.min(pagina - 3, paginas - 7));
      for (let i = inicio; i < Math.min(inicio + 7, paginas); i++) {
        p += `<button data-pg="${i}" style="padding:5px 10px;border-radius:6px;border:1px solid ${i === pagina ? "#5dade2" : "#e8edf2"};background:${i === pagina ? "#5dade2" : "#fff"};color:${i === pagina ? "#fff" : "#2c3e50"};cursor:pointer;font-weight:${i === pagina ? 700 : 400};font-size:12px;">${i + 1}</button>`;
      }
      p += botonPag("›", pagina + 1, pagina >= paginas - 1);
      p += botonPag("»", paginas - 1, pagina >= paginas - 1);
      pag.innerHTML = p;
      pag.querySelectorAll("button[data-pg]").forEach((b) =>
        b.addEventListener("click", () => { pagina = parseInt(b.dataset.pg); actualizar(); })
      );
    } else {
      pag.innerHTML = "";
    }
  }

  function botonPag(txt, destino, disabled) {
    return `<button data-pg="${destino}" ${disabled ? "disabled" : ""} style="padding:5px 10px;border-radius:6px;border:1px solid #e8edf2;background:#fff;font-size:12px;cursor:${disabled ? "default" : "pointer"};opacity:${disabled ? 0.4 : 1};">${txt}</button>`;
  }

  console.log("✅ Módulo 07 (Contratación) listo");
})();
