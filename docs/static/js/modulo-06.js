const C = {
  navy: "#1a2742", cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12",
  rojo: "#e74c3c", muted: "#7f8c8d", bg: "#f0f2f5", border: "#e8edf2",
  violet: "#8e44ad", teal: "#1abc9c", orange: "#e67e22"
};
const PAL = [C.cyan, C.verde, C.amar, C.rojo, C.violet, C.teal, C.orange, "#d63031", "#0984e3", "#e84393"];

const fmtCOP = v => v >= 1e9 ? `$${(v / 1e9).toFixed(2)}B` : v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : v >= 1e3 ? `$${(v / 1e3).toFixed(0)}K` : `$${v}`;
const fmtNum = v => v.toLocaleString("es-CO");
const toNum = v => {
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? 0 : n;
};

// Mapeo flexible de columnas
const COLS = {
  ejercicio: ['Ejercicio', 'EJERCICIO'],
  periodo: ['Período', 'Periodo', 'PERÍODO', 'PERIODO'],
  compromiso: ['Compromiso', 'COMPROMISO'],
  noCRP: ['Número de CRP', 'No. CRP', 'NumeroCRP'],
  noCDP: ['Número de CDP', 'No. CDP', 'NumeroCDP'],
  beneficiario: ['Nombre BP Beneficiario', 'Nombre Beneficiario', 'NombreBP'],
  valorCRP: ['Valor CRP', 'VALOR CRP'],
  giro: ['Autorizacion giro', 'Autorización giro', 'AUTORIZACION GIRO'],
  pendGiro: ['Pendiente de Giro', 'Pendiente Giro', 'PENDIENTE DE GIRO'],
  modSelec: ['Mod. Selec', 'Modalidad', 'Mod Selec'],
  tipoDoc: ['Tipo Doc.', 'Tipo de Documento', 'TIPO DOC'],
  anulaciones: ['Anulaciones', 'ANULACIONES'],
  reintegros: ['Reintegros', 'REINTEGROS'],
  valorNeto: ['Valor Neto', 'VALOR NETO']
};

const gc = (row, terms) => {
  for (const t of terms) {
    if (row[t] !== undefined && row[t] !== null && row[t] !== '') return row[t];
  }
  return '';
};

let DATOS_CRP = null;
let CHARTS_CRP = {};

function procesarExcelCRP(raw) {
  const headers = raw[0].map(c => String(c || "").trim());
  console.log("Encabezados CRP:", headers);
  
  const rows = raw.slice(1);
  const filas = [];

  for (const row of rows) {
    const bruto = toNum(gc(row, COLS.valorCRP));
    if (bruto <= 0) continue;

    filas.push({
      ejercicio: String(gc(row, COLS.ejercicio) || "").trim(),
      periodo: String(gc(row, COLS.periodo) || "").trim(),
      compromiso: String(gc(row, COLS.compromiso) || "").trim(),
      noCRP: String(gc(row, COLS.noCRP) || "").trim(),
      noCDP: String(gc(row, COLS.noCDP) || "").trim(),
      beneficiario: String(gc(row, COLS.beneficiario) || "").trim(),
      valorCRP: bruto,
      giro: toNum(gc(row, COLS.giro)),
      pendGiro: toNum(gc(row, COLS.pendGiro)),
      modSelec: String(gc(row, COLS.modSelec) || "").trim(),
      tipoDoc: String(gc(row, COLS.tipoDoc) || "").trim(),
      anulaciones: toNum(gc(row, COLS.anulaciones)),
      reintegros: toNum(gc(row, COLS.reintegros)),
      valorNeto: toNum(gc(row, COLS.valorNeto))
    });
  }

  const ejerciciosSet = [...new Set(filas.map(r => r.ejercicio))].filter(Boolean).sort();
  const periodosSet = [...new Set(filas.map(r => r.periodo))].filter(Boolean).sort();
  const modalidadesSet = [...new Set(filas.map(r => r.modSelec))].filter(Boolean).sort();

  return {
    filas,
    ejerciciosSet,
    periodosSet,
    modalidadesSet,
    totalCRP: filas.reduce((s, r) => s + r.valorCRP, 0),
    totalGiro: filas.reduce((s, r) => s + r.giro, 0),
    totalPend: filas.reduce((s, r) => s + r.pendGiro, 0),
    totalAnulaciones: filas.reduce((s, r) => s + r.anulaciones, 0),
    totalReintegros: filas.reduce((s, r) => s + r.reintegros, 0)
  };
}

function procesarGraficasCRP(filas, ejerciciosSet, periodosSet) {
  // G1: CRP vs Giro vs Pendiente por período
  const g1 = periodosSet.map(p => ({
    periodo: p,
    crp: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.valorCRP, 0),
    giro: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.giro, 0),
    pendiente: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.pendGiro, 0)
  }));

  // G2: Distribución por modalidad
  const g2map = {};
  filas.forEach(r => {
    const m = r.modSelec || 'Sin modalidad';
    g2map[m] = (g2map[m] || 0) + r.valorNeto;
  });
  const g2 = Object.entries(g2map).map(([name, value]) => ({ name: name.slice(0, 24), value })).sort((a, b) => b.value - a.value).slice(0, 8);

  // G3: Anulaciones + Reintegros por período
  const g3 = periodosSet.map(p => ({
    periodo: p,
    anulaciones: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.anulaciones, 0),
    reintegros: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.reintegros, 0)
  }));

  // G4: Top 10 beneficiarios
  const g4map = {};
  filas.forEach(r => {
    const b = r.beneficiario || 'Sin nombre';
    if (!g4map[b]) g4map[b] = { crp: 0, giro: 0 };
    g4map[b].crp += r.valorCRP;
    g4map[b].giro += r.giro;
  });
  const g4 = Object.entries(g4map).map(([nombre, v]) => ({ nombre: nombre.slice(0, 28), ...v })).sort((a, b) => b.crp - a.crp).slice(0, 10);

  // G5: Tipo de documento
  const g5map = {};
  filas.forEach(r => {
    const t = r.tipoDoc || 'Sin tipo';
    g5map[t] = (g5map[t] || 0) + r.valorCRP;
  });
  const g5 = Object.entries(g5map).map(([name, value]) => ({ name: name.slice(0, 20), value })).sort((a, b) => b.value - a.value);

  // G6: N° compromisos por período
  const g6 = periodosSet.map(p => ({
    periodo: p,
    count: filas.filter(r => r.periodo === p).length
  }));

  return { g1, g2, g3, g4, g5, g6 };
}

function mostrarKPIsCRP(datos) {
  const html = `
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.cyan}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">💰 CRP Total</div>
      <div style="font-size:20px; font-weight:bold; color:${C.cyan};">${fmtCOP(datos.totalCRP)}</div>
    </div>
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.verde}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">✅ Giro</div>
      <div style="font-size:20px; font-weight:bold; color:${C.verde};">${fmtCOP(datos.totalGiro)}</div>
    </div>
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.amar}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">⏳ Pendiente</div>
      <div style="font-size:20px; font-weight:bold; color:${C.amar};">${fmtCOP(datos.totalPend)}</div>
    </div>
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.rojo}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">❌ Anulaciones</div>
      <div style="font-size:20px; font-weight:bold; color:${C.rojo};">${fmtCOP(datos.totalAnulaciones)}</div>
    </div>
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.violet}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">↩️ Reintegros</div>
      <div style="font-size:20px; font-weight:bold; color:${C.violet};">${fmtCOP(datos.totalReintegros)}</div>
    </div>
    <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.teal}; flex:1; min-width:140px;">
      <div style="font-size:10px; color:${C.muted}; margin-bottom:6px;">📊 Registros</div>
      <div style="font-size:20px; font-weight:bold; color:${C.teal};">${fmtNum(datos.filas.length)}</div>
    </div>
  `;
  document.getElementById('kpis-crp').innerHTML = html;
}

function crearGraficasCRP(graficas) {
  Object.values(CHARTS_CRP).forEach(c => c.destroy?.());
  CHARTS_CRP = {};

  // G1
  CHARTS_CRP.g1 = new Chart(document.getElementById('grafico-crp-1'), {
    type: 'line',
    data: {
      labels: graficas.g1.map(d => d.periodo),
      datasets: [
        { label: 'CRP', data: graficas.g1.map(d => d.crp), borderColor: C.cyan, backgroundColor: C.cyan + '11', borderWidth: 2, fill: true, tension: 0.4 },
        { label: 'Giro', data: graficas.g1.map(d => d.giro), borderColor: C.verde, backgroundColor: C.verde + '11', borderWidth: 2, fill: true, tension: 0.4 },
        { label: 'Pendiente', data: graficas.g1.map(d => d.pendiente), borderColor: C.amar, backgroundColor: C.amar + '11', borderWidth: 2, fill: true, tension: 0.4 }
      ]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '① CRP vs Giro vs Pendiente' } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G2
  CHARTS_CRP.g2 = new Chart(document.getElementById('grafico-crp-2'), {
    type: 'doughnut',
    data: {
      labels: graficas.g2.map(d => d.name),
      datasets: [{ data: graficas.g2.map(d => d.value), backgroundColor: PAL }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '② Distribución por Modalidad' } } }
  });

  // G3
  CHARTS_CRP.g3 = new Chart(document.getElementById('grafico-crp-3'), {
    type: 'bar',
    data: {
      labels: graficas.g3.map(d => d.periodo),
      datasets: [
        { label: 'Anulaciones', data: graficas.g3.map(d => d.anulaciones), backgroundColor: C.rojo },
        { label: 'Reintegros', data: graficas.g3.map(d => d.reintegros), backgroundColor: C.violet }
      ]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '③ Anulaciones + Reintegros' } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G4
  CHARTS_CRP.g4 = new Chart(document.getElementById('grafico-crp-4'), {
    type: 'bar',
    data: {
      labels: graficas.g4.map(d => d.nombre),
      datasets: [
        { label: 'CRP', data: graficas.g4.map(d => d.crp), backgroundColor: C.cyan },
        { label: 'Giro', data: graficas.g4.map(d => d.giro), backgroundColor: C.verde }
      ]
    },
    options: { indexAxis: 'y', responsive: true, plugins: { title: { display: true, text: '④ Top 10 Beneficiarios' } }, scales: { x: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G5
  CHARTS_CRP.g5 = new Chart(document.getElementById('grafico-crp-5'), {
    type: 'bar',
    data: {
      labels: graficas.g5.map(d => d.name),
      datasets: [{ label: 'Valor CRP', data: graficas.g5.map(d => d.value), backgroundColor: PAL }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '⑤ Tipo de Documento' } }, scales: { y: { ticks: { callback: v => fmtCOP(v) } } } }
  });

  // G6
  CHARTS_CRP.g6 = new Chart(document.getElementById('grafico-crp-6'), {
    type: 'bar',
    data: {
      labels: graficas.g6.map(d => d.periodo),
      datasets: [{ label: 'N° Compromisos', data: graficas.g6.map(d => d.count), backgroundColor: C.amar }]
    },
    options: { responsive: true, plugins: { title: { display: true, text: '⑥ N° Compromisos' } } }
  });
}

function aplicarFiltrosCRP() {
  if (!DATOS_CRP) return;

  const q = document.getElementById('filtro-crp-busca').value.toLowerCase();
  const ej = document.getElementById('filtro-crp-ej').value;
  const per = document.getElementById('filtro-crp-per').value;
  const mod = document.getElementById('filtro-crp-mod').value;

  let filtradas = DATOS_CRP.filas.filter(r => {
    const okQ = !q || `${r.beneficiario} ${r.noCRP} ${r.noCDP} ${r.compromiso}`.toLowerCase().includes(q);
    const okEj = !ej || ej === 'Ejercicio' || r.ejercicio === ej;
    const okPer = !per || per === 'Período' || r.periodo === per;
    const okMod = !mod || mod === 'Modalidad' || r.modSelec === mod;
    return okQ && okEj && okPer && okMod;
  });

  const datosActualizados = {
    ...DATOS_CRP,
    filas: filtradas,
    totalCRP: filtradas.reduce((s, r) => s + r.valorCRP, 0),
    totalGiro: filtradas.reduce((s, r) => s + r.giro, 0),
    totalPend: filtradas.reduce((s, r) => s + r.pendGiro, 0),
    totalAnulaciones: filtradas.reduce((s, r) => s + r.anulaciones, 0),
    totalReintegros: filtradas.reduce((s, r) => s + r.reintegros, 0)
  };

  mostrarKPIsCRP(datosActualizados);
  const graficas = procesarGraficasCRP(filtradas, DATOS_CRP.ejerciciosSet, DATOS_CRP.periodosSet);
  crearGraficasCRP(graficas);
  mostrarTablaCRP(filtradas);
}

function mostrarTablaCRP(filas) {
  let html = '<table style="width:100%; border-collapse:collapse; font-size:10px;"><thead><tr style="background:' + C.navy + '; color:white;">';
  ['Ejercicio', 'Período', 'No. CRP', 'No. CDP', 'Compromiso', 'Beneficiario', 'Modalidad', 'CRP', 'Giro', 'Pendiente', 'Anulaciones'].forEach(h => {
    html += `<th style="padding:8px; text-align:left;">${h}</th>`;
  });
  html += '</tr></thead><tbody>';

  filas.slice(0, 50).forEach((r, i) => {
    html += `<tr style="border-bottom:1px solid ${C.border}; background:${i % 2 === 0 ? 'white' : C.bg};">
      <td style="padding:6px;">${r.ejercicio}</td>
      <td style="padding:6px;">${r.periodo}</td>
      <td style="padding:6px; font-family:monospace;">${r.noCRP}</td>
      <td style="padding:6px; font-family:monospace;">${r.noCDP}</td>
      <td style="padding:6px;">${r.compromiso}</td>
      <td style="padding:6px; max-width:160px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${r.beneficiario}</td>
      <td style="padding:6px;">${r.modSelec}</td>
      <td style="padding:6px; text-align:right; font-weight:bold; color:${C.cyan};">${fmtCOP(r.valorCRP)}</td>
      <td style="padding:6px; text-align:right; color:${C.verde};">${fmtCOP(r.giro)}</td>
      <td style="padding:6px; text-align:right; color:${C.amar};">${fmtCOP(r.pendGiro)}</td>
      <td style="padding:6px; text-align:right; color:${C.rojo};">${fmtCOP(r.anulaciones)}</td>
    </tr>`;
  });
  html += '</tbody></table>';
  document.getElementById('contenedor-tabla-crp').innerHTML = html;
}

document.addEventListener('DOMContentLoaded', function() {
  const dropZone = document.getElementById('drop-zone-crp');
  const fileInput = document.getElementById('file-crp');
  const pantallaC = document.getElementById('pantalla-carga-crp');
  const pantallap = document.getElementById('pantalla-principal-crp');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.opacity = '0.7'; });
  dropZone.addEventListener('dragleave', () => { dropZone.style.opacity = '1'; });
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) procesarArchivoCRP(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener('change', e => {
    if (e.target.files[0]) procesarArchivoCRP(e.target.files[0]);
  });

  function procesarArchivoCRP(file) {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
      DATOS_CRP = procesarExcelCRP(raw);

      document.getElementById('info-crp').textContent = `📁 ${file.name} · ${fmtNum(DATOS_CRP.filas.length)} registros`;

      // Llenar select
      document.getElementById('filtro-crp-ej').innerHTML = '<option>Ejercicio</option>' + DATOS_CRP.ejerciciosSet.map(v => `<option>${v}</option>`).join('');
      document.getElementById('filtro-crp-per').innerHTML = '<option>Período</option>' + DATOS_CRP.periodosSet.map(v => `<option>${v}</option>`).join('');
      document.getElementById('filtro-crp-mod').innerHTML = '<option>Modalidad</option>' + DATOS_CRP.modalidadesSet.map(v => `<option>${v}</option>`).join('');

      pantallaC.style.display = 'none';
      pantallap.style.display = 'block';

      mostrarKPIsCRP(DATOS_CRP);
      const g = procesarGraficasCRP(DATOS_CRP.filas, DATOS_CRP.ejerciciosSet, DATOS_CRP.periodosSet);
      crearGraficasCRP(g);
      mostrarTablaCRP(DATOS_CRP.filas);
    };
    reader.readAsArrayBuffer(file);
  }

  document.getElementById('filtro-crp-busca').addEventListener('input', aplicarFiltrosCRP);
  document.getElementById('filtro-crp-ej').addEventListener('change', aplicarFiltrosCRP);
  document.getElementById('filtro-crp-per').addEventListener('change', aplicarFiltrosCRP);
  document.getElementById('filtro-crp-mod').addEventListener('change', aplicarFiltrosCRP);
  document.getElementById('btn-filtrar-crp').addEventListener('click', aplicarFiltrosCRP);

  document.getElementById('btn-graf-crp').addEventListener('click', () => {
    document.getElementById('contenedor-graficas-crp').style.display = 'block';
    document.getElementById('contenedor-tabla-crp').style.display = 'none';
  });

  document.getElementById('btn-tab-crp').addEventListener('click', () => {
    document.getElementById('contenedor-graficas-crp').style.display = 'none';
    document.getElementById('contenedor-tabla-crp').style.display = 'block';
  });
});
