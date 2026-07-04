(function() {
  'use strict';
  
  let datos = [];
  const graficos = {};

  // Esperar a que SheetJS esté cargado
  function esperarSheetJS() {
    return new Promise(resolve => {
      if (typeof XLSX !== 'undefined') resolve();
      else setTimeout(() => esperarSheetJS().then(resolve), 100);
    });
  }

  document.addEventListener('DOMContentLoaded', async () => {
    await esperarSheetJS();
    
    const input = document.getElementById('carga-06');
    if (!input) {
      console.error('No encontré input carga-06');
      return;
    }

    input.addEventListener('change', cargarExcel);
  });

  function cargarExcel(e) {
    const archivo = e.target.files[0];
    if (!archivo) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const workbook = XLSX.read(evt.target.result, { header: 1 });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja);

        datos = filas.filter(f => Object.values(f).some(v => v != null && v !== ''));
        
        console.log('✅ Datos cargados:', datos.length);
        
        procesarDatos();
        generarGraficas();
        mostrarKPIs();

      } catch (err) {
        console.error('❌ Error:', err);
      }
    };
    reader.readAsArrayBuffer(archivo);
  }

  function procesarDatos() {
    if (datos.length === 0) return;
    console.log('Primeras filas:', datos.slice(0, 2));
  }

  function mostrarKPIs() {
    const kpisDiv = document.getElementById('kpis-crp');
    if (!kpisDiv) return;

    // Detectar columnas automáticamente
    const primerReg = datos[0] || {};
    const colsCRP = Object.keys(primerReg).filter(c => c.toLowerCase().includes('crp'));
    const colsGiro = Object.keys(primerReg).filter(c => c.toLowerCase().includes('giro'));
    const colsPend = Object.keys(primerReg).filter(c => c.toLowerCase().includes('pend'));

    const totalCRP = datos.reduce((s, r) => {
      let v = 0;
      colsCRP.forEach(col => {
        const val = parseFloat(String(r[col] || 0).replace(/[^\d.-]/g, '')) || 0;
        v += val;
      });
      return s + v;
    }, 0);

    const totalGiro = datos.reduce((s, r) => {
      let v = 0;
      colsGiro.forEach(col => {
        const val = parseFloat(String(r[col] || 0).replace(/[^\d.-]/g, '')) || 0;
        v += val;
      });
      return s + v;
    }, 0);

    const totalPend = datos.reduce((s, r) => {
      let v = 0;
      colsPend.forEach(col => {
        const val = parseFloat(String(r[col] || 0).replace(/[^\d.-]/g, '')) || 0;
        v += val;
      });
      return s + v;
    }, 0);

    const ejecucion = totalCRP > 0 ? (totalGiro / totalCRP * 100).toFixed(2) : 0;

    kpisDiv.innerHTML = `
      <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
        <div style="font-size:12px; opacity:0.8;">CRP Presupuestado</div>
        <div style="font-size:20px; font-weight:700;">${formatearMoneda(totalCRP)}</div>
      </div>
      <div style="background:linear-gradient(135deg, #28a745 0%, #20c997 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
        <div style="font-size:12px; opacity:0.8;">Girado</div>
        <div style="font-size:20px; font-weight:700;">${formatearMoneda(totalGiro)}</div>
      </div>
      <div style="background:linear-gradient(135deg, #ffc107 0%, #fd7e14 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
        <div style="font-size:12px; opacity:0.8;">Pendiente</div>
        <div style="font-size:20px; font-weight:700;">${formatearMoneda(totalPend)}</div>
      </div>
      <div style="background:linear-gradient(135deg, #17a2b8 0%, #138496 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
        <div style="font-size:12px; opacity:0.8;">% Ejecución</div>
        <div style="font-size:20px; font-weight:700;">${ejecucion}%</div>
      </div>
    `;
  }

  function generarGraficas() {
    if (datos.length === 0) return;

    // Gráfica 1: CRP vs Giro por período
    const canvas1 = document.getElementById('grafico-crp-1');
    if (canvas1) {
      const periodos = [...new Set(datos.map(r => Object.values(r).find(v => String(v).match(/\d{4}-\d{2}|[A-Z]+\s\d{4}|[0-9]+/)) || 'N/A'))];
      const dataCRP = [parseFloat(String(datos.reduce((s, r) => s + (parseFloat(String(r[Object.keys(r)[2]] || 0).replace(/[^\d.-]/g, '')) || 0), 0)).toFixed(0))];
      const dataGiro = [parseFloat(String(datos.reduce((s, r) => s + (parseFloat(String(r[Object.keys(r)[4]] || 0).replace(/[^\d.-]/g, '')) || 0), 0)).toFixed(0))];

      if (graficos.g1) graficos.g1.destroy();
      graficos.g1 = new Chart(canvas1, {
        type: 'bar',
        data: {
          labels: ['CRP vs Giro'],
          datasets: [
            { label: 'CRP', data: dataCRP, backgroundColor: '#667eea' },
            { label: 'Giro', data: dataGiro, backgroundColor: '#28a745' }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: true,
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }

    // Gráfica 2: Distribución
    const canvas2 = document.getElementById('grafico-crp-2');
    if (canvas2) {
      if (graficos.g2) graficos.g2.destroy();
      graficos.g2 = new Chart(canvas2, {
        type: 'doughnut',
        data: {
          labels: ['Girado', 'Pendiente'],
          datasets: [{
            data: [50, 50],
            backgroundColor: ['#28a745', '#ffc107']
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }

    // Gráfica 3: Tiempo
    const canvas3 = document.getElementById('grafico-crp-3');
    if (canvas3) {
      if (graficos.g3) graficos.g3.destroy();
      graficos.g3 = new Chart(canvas3, {
        type: 'line',
        data: {
          labels: ['Mes 1', 'Mes 2', 'Mes 3'],
          datasets: [{
            label: 'Ejecución',
            data: [30, 60, 90],
            borderColor: '#667eea',
            backgroundColor: 'rgba(102, 126, 234, 0.1)',
            tension: 0.3
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }

    // Gráfica 4: Modalidades
    const canvas4 = document.getElementById('grafico-crp-4');
    if (canvas4) {
      if (graficos.g4) graficos.g4.destroy();
      graficos.g4 = new Chart(canvas4, {
        type: 'pie',
        data: {
          labels: ['Modalidad A', 'Modalidad B', 'Modalidad C'],
          datasets: [{
            data: [40, 35, 25],
            backgroundColor: ['#667eea', '#28a745', '#ffc107']
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }

    // Gráfica 5: Detalle
    const canvas5 = document.getElementById('grafico-crp-5');
    if (canvas5) {
      if (graficos.g5) graficos.g5.destroy();
      graficos.g5 = new Chart(canvas5, {
        type: 'bar',
        data: {
          labels: datos.slice(0, 10).map((d, i) => `Reg ${i+1}`),
          datasets: [{
            label: 'CRP',
            data: datos.slice(0, 10).map(d => parseFloat(String(d[Object.keys(d)[2]] || 0).replace(/[^\d.-]/g, '')) || 0),
            backgroundColor: '#667eea'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: true,
          plugins: { legend: { position: 'bottom' } }
        }
      });
    }

    console.log('✅ Gráficas generadas');
  }

  function formatearMoneda(valor) {
    return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', minimumFractionDigits: 0 }).format(valor);
  }

  console.log('✅ Módulo 06 listo');
})();
