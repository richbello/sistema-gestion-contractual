(function() {
  const C = {
    navy: "#1a2742", cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12",
    rojo: "#e74c3c", muted: "#7f8c8d", bg: "#f0f2f5", border: "#e8edf2",
    violet: "#8e44ad", teal: "#1abc9c", orange: "#e67e22"
  };
  const PAL = [C.cyan, C.verde, C.amar, C.rojo, C.violet, C.teal, C.orange, "#d63031", "#0984e3", "#e84393"];

  const fmtCOP = v => v >= 1e9 ? `$${(v / 1e9).toFixed(2)}B` : v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : v >= 1e3 ? `$${(v / 1e3).toFixed(0)}K` : `$${v}`;
  const fmtNum = v => v.toLocaleString("es-CO");
  const toNum = v => { if (!v) return 0; const n = parseFloat(String(v).replace(/[^0-9.-]/g, '')); return isNaN(n) ? 0 : n; };

  const COLS = {
    ejercicio: ['Ejercicio', 'EJERCICIO'],
    periodo: ['Período', 'Periodo'],
    noCRP: ['Número de CRP', 'No. CRP'],
    noCDP: ['Número de CDP', 'No. CDP'],
    beneficiario: ['Nombre BP Beneficiario', 'Nombre Beneficiario'],
    valorCRP: ['Valor CRP'],
    giro: ['Autorizacion giro', 'Autorización giro'],
    pendGiro: ['Pendiente de Giro'],
    modSelec: ['Mod. Selec', 'Modalidad'],
    tipoDoc: ['Tipo Doc.'],
    anulaciones: ['Anulaciones'],
    reintegros: ['Reintegros'],
    valorNeto: ['Valor Neto']
  };

  const gc = (row, terms) => {
    for (const t of terms) if (row[t] !== undefined && row[t] !== null && row[t] !== '') return row[t];
    return '';
  };

  let DATOS_CRP = null;
  let CHARTS_CRP = {};

  function procesarExcelCRP(raw) {
    const rows = raw.slice(1);
    const filas = [];
    for (const row of rows) {
      const bruto = toNum(gc(row, COLS.valorCRP));
      if (bruto <= 0) continue;
      filas.push({
        ejercicio: String(gc(row, COLS.ejercicio) || "").trim(),
        periodo: String(gc(row, COLS.periodo) || "").trim(),
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
      filas, ejerciciosSet, periodosSet, modalidadesSet,
      totalCRP: filas.reduce((s, r) => s + r.valorCRP, 0),
      totalGiro: filas.reduce((s, r) => s + r.giro, 0),
      totalPend: filas.reduce((s, r) => s + r.pendGiro, 0),
      totalAnulaciones: filas.reduce((s, r) => s + r.anulaciones, 0),
      totalReintegros: filas.reduce((s, r) => s + r.reintegros, 0)
    };
  }

  function procesarGraficasCRP(filas, periodosSet) {
    const g1 = periodosSet.map(p => ({
      periodo: p,
      crp: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.valorCRP, 0),
      giro: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.giro, 0),
      pendiente: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.pendGiro, 0)
    }));

    const g2map = {};
    filas.forEach(r => { const m = r.modSelec || 'Sin modalidad'; g2map[m] = (g2map[m] || 0) + r.valorNeto; });
    const g2 = Object.entries(g2map).map(([name, value]) => ({ name: name.slice(0, 24), value })).sort((a, b) => b.value - a.value).slice(0, 8);

    const g3 = periodosSet.map(p => ({
      periodo: p,
      anulaciones: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.anulaciones, 0),
      reintegros: filas.filter(r => r.periodo === p).reduce((s, r) => s + r.reintegros, 0)
    }));

    const g4map = {};
    filas.forEach(r => { const b = r.beneficiario || 'Sin nombre'; if (!g4map[b]) g4map[b] = { crp: 0, giro: 0 }; g4map[b].crp += r.valorCRP; g4map[b].giro += r.giro; });
    const g4 = Object.entries(g4map).map(([nombre, v]) => ({ nombre: nombre.slice(0, 28), ...v })).sort((a, b) => b.crp - a.crp).slice(0, 10);

    const g5map = {};
    filas.forEach(r => { const t = r.tipoDoc || 'Sin tipo'; g5map[t] = (g5map[t] || 0) + r.valorCRP; });
    const g5 = Object.entries(g5map).map(([name, value]) => ({ name: name.slice(0, 20), value })).sort((a, b) => b.value - a.value);

    const g6 = periodosSet.map(p => ({ periodo: p, count: filas.filter(r => r.periodo === p).length }));

    return { g1, g2, g3, g4, g5, g6 };
  }

  function mostrarKPIsCRP(datos) {
    const html = `
      <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.cyan}; flex:1; min-width:100px;">
        <div style="font-size:9px; color:${C.muted};">💰 CRP</div>
        <div style="font-size:16px; font-weight:bold; color:${C.cyan};">${fmtCOP(datos.totalCRP)}</div>
      </div>
      <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.verde}; flex:1; min-width:100px;">
        <div style="font-size:9px; color:${C.muted};">✅ Giro</div>
        <div style="font-size:16px; font-weight:bold; color:${C.verde};">${fmtCOP(datos.totalGiro)}</div>
      </div>
      <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.amar}; flex:1; min-width:100px;">
        <div style="font-size:9px; color:${C.muted};">⏳ Pend</div>
        <div style="font-size:16px; font-weight:bold; color:${C.amar};">${fmtCOP(datos.totalPend)}</div>
      </div>
      <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.rojo}; flex:1; min-width:100px;">
        <div style="font-size:9px; color:${C.muted};">❌ Anul</div>
        <div style="font-size:16px; font-weight:bold; color:${C.rojo};">${fmtCOP(datos.totalAnulaciones)}</div>
      </div>
      <div style="background:white; border-radius:10px; padding:12px; border-left:3px solid ${C.violet}; flex:1; min-width:100px;">
        <div style="font-size:9px; color:${C.muted};">↩️ Rein</div>
        <div style="font-size:16px; font-weight:bold; color:${C.violet};">${fmtCOP(datos.totalReintegros)}</div>
      </div>
    `;
    document.getElementById('kpis-crp').innerHTML = html;
  }

  function crearGraficasCRP(graficas) {
    Object.values(CHARTS_CRP).forEach(c => c.destroy?.());
    CHARTS_CRP = {};

    const chartConfig = (id, type, data, title) => {
      CHARTS_CRP[id] = new Chart(document.getElementById(id), {
        type, data, options: { responsive: true, maintainAspectRatio: false, plugins: { title: { display: true, text: title, font: { size: 11 } }, legend: { labels: { font: { size: 10 } } } }, scales: { y: { ticks: { callback: v => fmtCOP(v), font: { size: 9 } } }, x: { ticks: { font: { size: 9 } } } } }
      });
    };

    chartConfig('grafico-crp-1', 'line', {
      labels: graficas.g1.map(d => d.periodo),
      datasets: [
        { label: 'CRP', data: graficas.g1.map(d => d.crp), borderColor: C.cyan, backgroundColor: C.cyan + '11', borderWidth: 2, fill: true, tension: 0.4 },
        { label: 'Giro', data: graficas.g1.map(d => d.giro), borderColor: C.verde, backgroundColor: C.verde + '11', borderWidth: 2, fill: true, tension: 0.4 }
      ]
    }, '① CRP vs Giro');

    chartConfig('grafico-crp-2', 'doughnut', {
      labels: graficas.g2.map(d => d.name),
      datasets: [{ data: graficas.g2.map(d => d.value), backgroundColor: PAL }]
    }, '② Modalidad');

    chartConfig('grafico-crp-3', 'bar', {
      labels: graficas.g3.map(d => d.periodo),
      datasets: [
        { label: 'Anulaciones', data: graficas.g3.map(d => d.anulaciones), backgroundColor: C.rojo },
        { label: 'Reintegros', data: graficas.g3.map(d => d.reintegros), backgroundColor: C.violet }
      ]
    }, '③ Anulaciones');

    chartConfig('grafico-crp-4', 'bar', {
      labels: graficas.g4.map(d => d.nombre),
      datasets: [
        { label: 'CRP', data: graficas.g4.map(d => d.crp), backgroundColor: C.cyan },
        { label: 'Giro', data: graficas.g4.map(d => d.giro), backgroundColor: C.verde }
      ]
    }, '④ Top 10');

    chartConfig('grafico-crp-5', 'bar', {
      labels: graficas.g5.map(d => d.name),
      datasets: [{ label: 'Valor CRP', data: graficas.g5.map(d => d.value), backgroundColor: PAL }]
    }, '⑤ Tipo Doc');

    chartConfig('grafico-crp-6', 'bar', {
      labels: graficas.g6.map(d => d.periodo),
      datasets: [{ label: 'N° Compromisos', data: graficas.g6.map(d => d.count), backgroundColor: C.amar }]
    }, '⑥ Compromisos');
  }

  function aplicarFiltrosCRP() {
    if (!DATOS_CRP) return;
    const q = document.getElementById('filtro-crp-busca').value.toLowerCase();
    const ej = document.getElementById('filtro-crp-ej').value;
    const per = document.getElementById('filtro-crp-per').value;
    const mod = document.getElementById('filtro-crp-mod').value;

    let filtradas = DATOS_CRP.filas.filter(r => {
      const okQ = !q || `${r.beneficiario} ${r.noCRP} ${r.noCDP}`.toLowerCase().includes(q);
      const okEj = !ej || ej === 'Ejercicio' || r.ejercicio === ej;
      const okPer = !per || per === 'Período' || r.periodo === per;
      const okMod = !mod || mod === 'Modalidad' || r.modSelec === mod;
      return okQ && okEj && okPer && okMod;
    });

    const datosAct = { ...DATOS_CRP, filas: filtradas, totalCRP: filtradas.reduce((s, r) => s + r.valorCRP, 0), totalGiro: filtradas.reduce((s, r) => s + r.giro, 0), totalPend: filtradas.reduce((s, r) => s + r.pendGiro, 0), totalAnulaciones: filtradas.reduce((s, r) => s + r.anulaciones, 0), totalReintegros: filtradas.reduce((s, r) => s + r.reintegros, 0) };

    mostrarKPIsCRP(datosAct);
    const g = procesarGraficasCRP(filtradas, DATOS_CRP.periodosSet);
    crearGraficasCRP(g);
  }

  document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('drop-zone-crp');
    const fileInput = document.getElementById('file-crp');
    const pantallaC = document.getElementById('pantalla-carga-crp');
    const pantallap = document.getElementById('pantalla-principal-crp');

    if (!dropZone) return;

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('drop', e => { e.preventDefault(); if (e.dataTransfer.files[0]) procesar(e.dataTransfer.files[0]); });
    fileInput.addEventListener('change', e => { if (e.target.files[0]) procesar(e.target.files[0]); });

    function procesar(file) {
      const reader = new FileReader();
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array' });
          const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
          DATOS_CRP = procesarExcelCRP(raw);
          document.getElementById('info-crp').textContent = `📁 ${file.name} · ${fmtNum(DATOS_CRP.filas.length)} registros`;
          document.getElementById('filtro-crp-ej').innerHTML = '<option>Ejercicio</option>' + DATOS_CRP.ejerciciosSet.map(v => `<option>${v}</option>`).join('');
          document.getElementById('filtro-crp-per').innerHTML = '<option>Período</option>' + DATOS_CRP.periodosSet.map(v => `<option>${v}</option>`).join('');
          document.getElementById('filtro-crp-mod').innerHTML = '<option>Modalidad</option>' + DATOS_CRP.modalidadesSet.map(v => `<option>${v}</option>`).join('');
          pantallaC.style.display = 'none';
          pantallap.style.display = 'block';
          mostrarKPIsCRP(DATOS_CRP);
          const g = procesarGraficasCRP(DATOS_CRP.filas, DATOS_CRP.periodosSet);
          crearGraficasCRP(g);
        } catch (error) { console.error('Error:', error); alert('Error cargando Excel: ' + error.message); }
      };
      reader.readAsArrayBuffer(file);
    }

    document.getElementById('filtro-crp-busca').addEventListener('input', aplicarFiltrosCRP);
    document.getElementById('filtro-crp-ej').addEventListener('change', aplicarFiltrosCRP);
    document.getElementById('filtro-crp-per').addEventListener('change', aplicarFiltrosCRP);
    document.getElementById('filtro-crp-mod').addEventListener('change', aplicarFiltrosCRP);
    document.getElementById('btn-filtrar-crp').addEventListener('click', aplicarFiltrosCRP);
    document.getElementById('btn-graf-crp').addEventListener('click', () => { document.getElementById('contenedor-graficas-crp').style.display = 'block'; document.getElementById('contenedor-tabla-crp').style.display = 'none'; });
    document.getElementById('btn-tab-crp').addEventListener('click', () => { document.getElementById('contenedor-graficas-crp').style.display = 'none'; document.getElementById('contenedor-tabla-crp').style.display = 'block'; });
  });
})();
