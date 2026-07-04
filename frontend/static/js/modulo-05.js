const C = {
  navy: "#1a2742", navyLt: "#1e3a5f", cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12",
  rojo: "#e74c3c", texto: "#2c3e50", muted: "#7f8c8d", bg: "#f0f2f5", border: "#e8edf2",
  violet: "#8e44ad", teal: "#1abc9c", orange: "#e67e22"
};
const PAL = [C.cyan, C.verde, C.amar, C.rojo, C.violet, C.orange, C.teal, "#d63031", "#0984e3", "#e84393"];

const fmtCOP = v => v >= 1e9 ? `$${(v / 1e9).toFixed(2)}MM` : v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : `$${(v / 1e3).toFixed(0)}K`;
const fmtNum = v => v.toLocaleString("es-CO");

let DATOS_GLOBAL = null;
let CHARTS = {};

function gc(headers, terms) {
  return headers.findIndex(h => terms.some(t => h.toLowerCase().includes(t.toLowerCase())));
}

function gcExacto(headers, terms) {
  const exacto = headers.findIndex(h => terms.some(t => h.trim().toLowerCase() === t.toLowerCase()));
  return exacto >= 0 ? exacto : gc(headers, terms);
}

function parseFecha(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    const d = String(val.getDate()).padStart(2, "0");
    const m = String(val.getMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${val.getFullYear()}`;
  }
  if (typeof val === "number") {
    const dt = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (isNaN(dt.getTime())) return "";
    const d = String(dt.getUTCDate()).padStart(2, "0");
    const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${dt.getUTCFullYear()}`;
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const [y, m, d] = s.slice(0, 10).split("-");
    return `${d}/${m}/${y}`;
  }
  return s.slice(0, 10);
}

function procesarExcel(raw) {
  const headers = raw[0].map(c => String(c || "").trim());
  console.log("Encabezados encontrados:", headers);
  
  const rows = raw.slice(1);
  
  // Búsqueda más flexible de índices
  const iNombre = headers.findIndex(h => h.toLowerCase().includes("nombre"));
  const iNroId = headers.findIndex(h => h.toLowerCase().includes("identificación") || h.toLowerCase().includes("nit"));
  const iBruto = headers.findIndex(h => h.toLowerCase().includes("valor bruto") || h.toLowerCase().includes("valor"));
  const iEj = headers.findIndex(h => h.toLowerCase().includes("ejercicio"));
  const iRef = headers.findIndex(h => h.toLowerCase().includes("referencia"));
  const iFecha = headers.findIndex(h => h.toLowerCase().includes("fecha de pago") || h.toLowerCase().includes("fecha pago"));
  const iAsig = headers.findIndex(h => h.toLowerCase().includes("asignación") || h.toLowerCase().includes("asignacion"));

  console.log("Índices:", {iNombre, iNroId, iBruto, iEj, iRef, iFecha, iAsig});

  const filas = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const bruto = Number(row[iBruto]) || 0;
    if (bruto <= 0) continue;
    
    filas.push({
      nombre: String(row[iNombre] || "").trim(),
      nroId: String(row[iNroId] || "").trim(),
      valorBruto: bruto,
      ejercicio: Number(row[iEj]) || 0,
      referencia: String(row[iRef] || "").trim(),
      fechaPago: parseFecha(row[iFecha]),
      asignacion: String(row[iAsig] || "").trim(),
      asignacion3: ""
    });
  }

  console.log("Filas procesadas:", filas.length);

  const ejerciciosSet = [...new Set(filas.map(r => r.ejercicio))].filter(Boolean).sort();
  const totalBruto = filas.reduce((s, r) => s + r.valorBruto, 0);
  const totalRegistros = filas.length;
  const proveedoresUnicos = new Set(filas.map(r => r.nroId)).size;

  console.log({totalBruto, totalRegistros, proveedoresUnicos, ejerciciosSet});

  return { filas, ejerciciosSet, totalBruto, totalRegistros, proveedoresUnicos };
}


function procesarGraficas(filas, ejerciciosSet) {
  // G1: Combo por ejercicio
  const g1 = ejerciciosSet.map(e => ({
    ejercicio: String(e),
    bruto: filas.filter(r => r.ejercicio === e).reduce((s, r) => s + r.valorBruto, 0),
    count: filas.filter(r => r.ejercicio === e).length
  }));

  // G2: Top 10 por nombre
  const mapaNomb = {};
  for (const r of filas) mapaNomb[r.nombre] = (mapaNomb[r.nombre] || 0) + r.valorBruto;
  const g2 = Object.entries(mapaNomb).sort((a, b) => b[1] - a[1]).slice(0, 10)
    .map(([nombre, bruto]) => ({ nombre: nombre.slice(0, 32), bruto }));

  // G3: Top 10 por NIT (Pie)
  const mapaNit = {};
  for (const r of filas) {
    if (!mapaNit[r.nroId]) mapaNit[r.nroId] = { nombre: r.nombre, bruto: 0 };
    mapaNit[r.nroId].bruto += r.valorBruto;
  }
  const g3 = Object.entries(mapaNit).sort((a, b) => b[1].bruto - a[1].bruto).slice(0, 10)
    .map(([nit, v]) => ({ nit, nombre: v.nombre.slice(0, 28), bruto: v.bruto }));

  // G4: Top 10 referencias
  const mapaRef = {};
  for (const r of filas) {
    const k = r.referencia || "SIN REF";
    if (!mapaRef[k]) mapaRef[k] = { bruto: 0, count: 0 };
    mapaRef[k].bruto += r.valorBruto;
    mapaRef[k].count += 1;
  }
  const g4 = Object.entries(mapaRef).sort((a, b) => b[1].bruto - a[1].bruto).slice(0, 10)
    .map(([ref, v]) => ({ ref: ref.slice(0, 14), bruto: v.bruto, count: v.count }));

  // G5: Top 5 proveedores evolución anual
  const top5nombres = Object.entries(mapaNomb).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([n]) => n);
  const g5 = ejerciciosSet.map(e => {
    const row = { ejercicio: String(e) };
    for (const n of top5nombres) {
      row[n.slice(0, 16)] = filas.filter(r => r.ejercicio === e && r.nombre === n).reduce((s, r) => s + r.valorBruto, 0);
    }
    return row;
  });
  const g5nombres = top5nombres.map(n => n.slice(0, 16));

  // G6: NITs únicos por ejercicio
  const g6 = ejerciciosSet.map(e => ({
    ejercicio: String(e),
    nits: new Set(filas.filter(r => r.ejercicio === e).map(r => r.nroId)).size
  }));

  // G7: Scatter referencias
  const g7 = g4.map(r => ({ ref: r.ref, monto: r.bruto, count: r.count }));

  return { g1, g2, g3, g4, g5, g5nombres, g6, g7 };
}

function mostrarMetricas(datos) {
  const { totalBruto, totalRegistros, proveedoresUnicos, ejerciciosSet } = datos;
  const mayorEj = datos.g1.reduce((a, b) => a.bruto > b.bruto ? a : b, datos.g1[0]);

  const html = `
    <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:15px;">
      <div style="background:white; border-radius:10px; padding:15px; border-left:3px solid ${C.cyan};">
        <div style="font-size:11px; color:${C.muted}; margin-bottom:8px;">💰 Valor Bruto Total</div>
        <div style="font-size:22px; font-weight:bold; color:${C.cyan};">${fmtCOP(totalBruto)}</div>
        <div style="font-size:10px; color:${C.muted}; margin-top:5px;">${ejerciciosSet[0]}–${ejerciciosSet[ejerciciosSet.length - 1]}</div>
      </div>
      <div style="background:white; border-radius:10px; padding:15px; border-left:3px solid ${C.verde};">
        <div style="font-size:11px; color:${C.muted}; margin-bottom:8px;">📋 Registros SAP</div>
        <div style="font-size:22px; font-weight:bold; color:${C.verde};">${fmtNum(totalRegistros)}</div>
        <div style="font-size:10px; color:${C.muted}; margin-top:5px;">transacciones</div>
      </div>
      <div style="background:white; border-radius:10px; padding:15px; border-left:3px solid ${C.amar};">
        <div style="font-size:11px; color:${C.muted}; margin-bottom:8px;">🏢 Proveedores únicos</div>
        <div style="font-size:22px; font-weight:bold; color:${C.amar};">${fmtNum(proveedoresUnicos)}</div>
        <div style="font-size:10px; color:${C.muted}; margin-top:5px;">por Nº identificación</div>
      </div>
      <div style="background:white; border-radius:10px; padding:15px; border-left:3px solid ${C.rojo};">
        <div style="font-size:11px; color:${C.muted}; margin-bottom:8px;">📅 Mayor ejercicio</div>
        <div style="font-size:22px; font-weight:bold; color:${C.rojo};">${mayorEj.ejercicio}</div>
        <div style="font-size:10px; color:${C.muted}; margin-top:5px;">${fmtCOP(mayorEj.bruto)}</div>
      </div>
    </div>
  `;

  document.getElementById('metricas-container').innerHTML = html;
  document.getElementById('rango-ejercicios').textContent = `${ejerciciosSet[0]}–${ejerciciosSet[ejerciciosSet.length - 1]}`;
}

function crearGraficas(datos) {
  const { g1, g2, g3, g4, g5, g5nombres, g6, g7 } = datos;

  // Destruir gráficas anteriores
  Object.values(CHARTS).forEach(c => c.destroy?.());
  CHARTS = {};

  // G1: Combo
  CHARTS.g1 = new Chart(document.getElementById('grafico-1'), {
    type: 'bar',
    data: {
      labels: g1.map(d => d.ejercicio),
      datasets: [
        { label: 'Valor Bruto', data: g1.map(d => d.bruto), backgroundColor: C.cyan, yAxisID: 'y' },
        { label: 'Registros', data: g1.map(d => d.count), type: 'line', borderColor: C.amar, backgroundColor: 'rgba(243, 156, 18, 0.1)', borderWidth: 2, yAxisID: 'y1', fill: true, tension: 0.4 }
      ]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '① Valor Bruto por Ejercicio', color: C.navy } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } }, y1: { position: 'right', grid: { drawOnChartArea: false } } } }
  });

  // G2: Top 10 Horizontal
  CHARTS.g2 = new Chart(document.getElementById('grafico-2'), {
    type: 'bar',
    data: {
      labels: g2.map(d => d.nombre),
      datasets: [{ label: 'Valor Bruto', data: g2.map(d => d.bruto), backgroundColor: PAL }]
    },
    options: { indexAxis: 'y', responsive: true, plugins: { title: { display: true, text: '② Top 10 Proveedores' } }, scales: { x: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G3: Pie NITs
  CHARTS.g3 = new Chart(document.getElementById('grafico-3'), {
    type: 'doughnut',
    data: {
      labels: g3.map(d => d.nit),
      datasets: [{ data: g3.map(d => d.bruto), backgroundColor: PAL }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '③ Top 10 por NIT' } } }
  });

  // G4: Referencias
  CHARTS.g4 = new Chart(document.getElementById('grafico-4'), {
    type: 'bar',
    data: {
      labels: g4.map(d => d.ref),
      datasets: [{ label: 'Valor Bruto', data: g4.map(d => d.bruto), backgroundColor: PAL }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '④ Top 10 Referencias' } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G5: Timeline Top 5
  CHARTS.g5 = new Chart(document.getElementById('grafico-5'), {
    type: 'line',
    data: {
      labels: g5.map(d => d.ejercicio),
      datasets: g5nombres.map((n, i) => ({
        label: n,
        data: g5.map(d => d[n] || 0),
        borderColor: PAL[i],
        backgroundColor: PAL[i] + '11',
        borderWidth: 2,
        fill: true,
        tension: 0.4
      }))
    },
    options: { responsive: true, plugins: { title: { display: true, text: '⑤ Evolución Top 5 Proveedores' } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G6: NITs únicos
  CHARTS.g6 = new Chart(document.getElementById('grafico-6'), {
    type: 'bar',
    data: {
      labels: g6.map(d => d.ejercicio),
      datasets: [{ label: 'NITs únicos', data: g6.map(d => d.nits), backgroundColor: C.amar }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '⑥ NITs únicos/Ejercicio' } } }
  });

  // G7: Scatter
  CHARTS.g7 = new Chart(document.getElementById('grafico-7'), {
    type: 'scatter',
    data: {
      datasets: [{
        label: 'Referencias',
        data: g7.map(d => ({ x: d.count, y: d.monto })),
        backgroundColor: PAL.map((c, i) => PAL[i % PAL.length])
      }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '⑦ Monto vs Frecuencia' } }, scales: { x: { title: { display: true, text: 'N° Registros' } }, y: { title: { display: true, text: 'Valor Bruto' }, ticks: { callback: v => fmtCOP(v) } } } }
  });
}

function aplicarFiltros() {
  const q = document.getElementById('filtro-contratista').value.toLowerCase();
  const qRef = document.getElementById('filtro-referencia').value.toLowerCase();
  const fechaDesde = document.getElementById('filtro-fecha-desde').value;
  const fechaHasta = document.getElementById('filtro-fecha-hasta').value;

  const filtradas = DATOS_GLOBAL.filas.filter(r => {
    const campoQ = `${r.nombre} ${r.nroId}`.toLowerCase();
    const okQ = !q || campoQ.includes(q);
    const okRef = !qRef || r.referencia.toLowerCase().includes(qRef);
    return okQ && okRef;
  });

  const datosActualizados = {
    ...DATOS_GLOBAL,
    filas: filtradas,
    totalBruto: filtradas.reduce((s, r) => s + r.valorBruto, 0),
    totalRegistros: filtradas.length,
    ...procesarGraficas(filtradas, DATOS_GLOBAL.ejerciciosSet)
  };

  mostrarMetricas(datosActualizados);
  crearGraficas(datosActualizados);
  mostrarTabla(filtradas);

  // Ficha contratista
  if (q && filtradas.length > 0) {
    const primerPago = filtradas[0];
    document.getElementById('contratista-nombre').textContent = primerPago.nombre;
    document.getElementById('contratista-nit').textContent = `NIT: ${primerPago.nroId}`;
    const totalContratista = filtradas.reduce((s, r) => s + r.valorBruto, 0);
    document.getElementById('contratista-resumen').innerHTML = `
      <div style="background:rgba(255,255,255,0.1); padding:10px; border-radius:6px; text-align:center;">
        <div style="font-size:10px; color:rgba(255,255,255,0.6);">Total pagado</div>
        <div style="font-size:18px; font-weight:bold; color:${C.cyan};">${fmtCOP(totalContratista)}</div>
      </div>
      <div style="background:rgba(255,255,255,0.1); padding:10px; border-radius:6px; text-align:center;">
        <div style="font-size:10px; color:rgba(255,255,255,0.6);">Registros</div>
        <div style="font-size:18px; font-weight:bold; color:${C.amar};">${fmtNum(filtradas.length)}</div>
      </div>
    `;
    document.getElementById('ficha-contratista').style.display = 'block';
  } else {
    document.getElementById('ficha-contratista').style.display = 'none';
  }
}

function mostrarTabla(filas) {
  let html = '<table style="width:100%; border-collapse:collapse; font-size:11px;"><thead><tr style="background:' + C.navy + '; color:white;">';
  ['Nombre', 'NIT', 'Valor Bruto', 'Ejercicio', 'Referencia', 'Fecha Pago'].forEach(h => {
    html += `<th style="padding:10px; text-align:left;">${h}</th>`;
  });
  html += '</tr></thead><tbody>';

  filas.slice(0, 50).forEach((r, i) => {
    html += `<tr style="border-bottom:1px solid ${C.border}; background:${i % 2 === 0 ? 'white' : C.bg};">
      <td style="padding:8px;">${r.nombre}</td>
      <td style="padding:8px; font-family:monospace;">${r.nroId}</td>
      <td style="padding:8px; color:${C.cyan}; font-weight:bold;">${fmtCOP(r.valorBruto)}</td>
      <td style="padding:8px;">${r.ejercicio}</td>
      <td style="padding:8px;">${r.referencia}</td>
      <td style="padding:8px; color:${C.verde};">${r.fechaPago || '—'}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  document.getElementById('contenedor-tabla').innerHTML = html;
}

document.addEventListener('DOMContentLoaded', function() {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input-historico');
  const pantallaCarga = document.getElementById('pantalla-carga');
  const pantallaPrincipal = document.getElementById('pantalla-principal');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.opacity = '0.7'; });
  dropZone.addEventListener('dragleave', () => { dropZone.style.opacity = '1'; });
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) procesarArchivo(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener('change', e => {
    if (e.target.files[0]) procesarArchivo(e.target.files[0]);
  });

  function procesarArchivo(file) {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
      DATOS_GLOBAL = procesarExcel(raw);
      DATOS_GLOBAL = { ...DATOS_GLOBAL, ...procesarGraficas(DATOS_GLOBAL.filas, DATOS_GLOBAL.ejerciciosSet) };

      // Rellenar select vigencias
      const selectVigencia = document.getElementById('filtro-vigencia');
      DATOS_GLOBAL.ejerciciosSet.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        selectVigencia.appendChild(opt);
      });

      pantallaCarga.style.display = 'none';
      pantallaPrincipal.style.display = 'block';

      mostrarMetricas(DATOS_GLOBAL);
      crearGraficas(DATOS_GLOBAL);
      mostrarTabla(DATOS_GLOBAL.filas);
    };
    reader.readAsArrayBuffer(file);
  }

  document.getElementById('btn-buscar').addEventListener('click', aplicarFiltros);
  document.getElementById('btn-limpiar').addEventListener('click', () => {
    document.getElementById('filtro-contratista').value = '';
    document.getElementById('filtro-referencia').value = '';
    document.getElementById('filtro-fecha-desde').value = '';
    document.getElementById('filtro-fecha-hasta').value = '';
    aplicarFiltros();
  });

  document.getElementById('btn-graficas').addEventListener('click', () => {
    document.getElementById('contenedor-graficas').style.display = 'block';
    document.getElementById('contenedor-tabla').style.display = 'none';
  });

  document.getElementById('btn-tabla').addEventListener('click', () => {
    document.getElementById('contenedor-graficas').style.display = 'none';
    document.getElementById('contenedor-tabla').style.display = 'block';
  });

  document.getElementById('btn-descargar').addEventListener('click', () => {
    if (!DATOS_GLOBAL) return;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(DATOS_GLOBAL.filas);
    XLSX.utils.book_append_sheet(wb, ws, 'Historico');
    XLSX.writeFile(wb, `Historico_${new Date().toISOString().split('T')[0]}.xlsx`);
  });
});
