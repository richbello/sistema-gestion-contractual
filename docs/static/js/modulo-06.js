(function() {
  'use strict';

  const datosModulo06 = [];
  let archivoActual = null;

  // Cargar archivo Excel
  document.getElementById('carga-06').addEventListener('change', async (e) => {
    const archivo = e.target.files[0];
    if (!archivo) return;

    archivoActual = archivo.name;
    datosModulo06.length = 0;

    try {
      const arrayBuffer = await archivo.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { header: 1 });
      
      if (!workbook.SheetNames.length) {
        console.error('Excel sin hojas');
        return;
      }

      const primeraHoja = workbook.Sheets[workbook.SheetNames[0]];
      const datos = XLSX.utils.sheet_to_json(primeraHoja);

      if (!datos || datos.length === 0) {
        mostrarMensaje('06', 'No hay datos en el Excel', 'error');
        return;
      }

      // Procesar datos
      datos.forEach((fila, idx) => {
        const registro = {};
        Object.keys(fila).forEach(col => {
          const colLimpia = col.trim();
          // Mapeo flexible de columnas
          if (colLimpia.toLowerCase().includes('ejercicio')) registro.ejercicio = fila[col];
          else if (colLimpia.toLowerCase().includes('período') || colLimpia.toLowerCase().includes('periodo')) registro.periodo = fila[col];
          else if (colLimpia.toLowerCase().includes('beneficiario')) registro.beneficiario = fila[col];
          else if (colLimpia.toLowerCase().includes('crp') && !colLimpia.toLowerCase().includes('nº')) {
            const valor = parseFloat(String(fila[col]).replace(/[^\d.-]/g, '')) || 0;
            registro.valorCRP = valor;
          }
          else if (colLimpia.toLowerCase().includes('giro')) {
            const valor = parseFloat(String(fila[col]).replace(/[^\d.-]/g, '')) || 0;
            registro.giro = valor;
          }
          else if (colLimpia.toLowerCase().includes('pend') || colLimpia.toLowerCase().includes('pendiente')) {
            const valor = parseFloat(String(fila[col]).replace(/[^\d.-]/g, '')) || 0;
            registro.pendiente = valor;
          }
          else if (colLimpia.toLowerCase().includes('nº') || colLimpia.toLowerCase().includes('compromiso')) {
            registro.compromiso = fila[col];
          }
          else if (colLimpia.toLowerCase().includes('descripción') || colLimpia.toLowerCase().includes('descripcion')) {
            registro.descripcion = fila[col];
          }
          else if (colLimpia.toLowerCase().includes('modalidad')) {
            registro.modalidad = fila[col];
          }
        });
        datosModulo06.push(registro);
      });

      const registrosCargados = datosModulo06.length;
      document.querySelector('#modulo-06 .panel-header').textContent = 
        `Reporte CRP ${archivoActual} - ${registrosCargados} registros`;

      // Mostrar KPIs
      actualizarKPIs();
      
      // Generar gráficas
      generarGraficas();
      
      mostrarMensaje('06', `✅ Cargados ${registrosCargados} registros`, 'exito');

    } catch (err) {
      console.error('Error al procesar Excel:', err);
      mostrarMensaje('06', 'Error al procesar el archivo: ' + err.message, 'error');
    }
  });

  function actualizarKPIs() {
    const totalCRP = datosModulo06.reduce((sum, r) => sum + (r.valorCRP || 0), 0);
    const totalGiro = datosModulo06.reduce((sum, r) => sum + (r.giro || 0), 0);
    const totalPendiente = datosModulo06.reduce((sum, r) => sum + (r.pendiente || 0), 0);
    const totalAnulaciones = datosModulo06.reduce((sum, r) => sum + (r.anulaciones || 0), 0);
    const totalReintegros = datosModulo06.reduce((sum, r) => sum + (r.reintegros || 0), 0);

    document.querySelector('#kpi-crp')?.parentElement.parentElement.querySelector('span').textContent = formatearMoneda(totalCRP);
    document.querySelector('#kpi-giro')?.parentElement.parentElement.querySelector('span').textContent = formatearMoneda(totalGiro);
    document.querySelector('#kpi-pendiente')?.parentElement.parentElement.querySelector('span').textContent = formatearMoneda(totalPendiente);
    document.querySelector('#kpi-anulaciones')?.parentElement.parentElement.querySelector('span').textContent = formatearMoneda(totalAnulaciones);
    document.querySelector('#kpi-reintegros')?.parentElement.parentElement.querySelector('span').textContent = formatearMoneda(totalReintegros);
  }

  function generarGraficas() {
    if (datosModulo06.length === 0) return;

    // Gráfica 1: CRP vs Giro
    const ctx1 = document.getElementById('grafica-06-1')?.getContext('2d');
    if (ctx1) {
      const periodos = [...new Set(datosModulo06.map(r => r.periodo))];
      const dataCRP = periodos.map(p => datosModulo06.filter(r => r.periodo === p).reduce((s, r) => s + (r.valorCRP || 0), 0));
      const dataGiro = periodos.map(p => datosModulo06.filter(r => r.periodo === p).reduce((s, r) => s + (r.giro || 0), 0));

      new Chart(ctx1, {
        type: 'line',
        data: {
          labels: periodos,
          datasets: [
            { label: 'CRP', data: dataCRP, borderColor: '#667eea', backgroundColor: 'rgba(102, 126, 234, 0.1)', tension: 0.3 },
            { label: 'Giro', data: dataGiro, borderColor: '#28a745', backgroundColor: 'rgba(40, 167, 69, 0.1)', tension: 0.3 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: true,
          plugins: { legend: { position: 'bottom' } },
          scales: { y: { beginAtZero: true } }
        }
      });
    }

    // Gráfica 2: Distribución por Modalidad
    const ctx2 = document.getElementById('grafica-06-2')?.getContext('2d');
    if (ctx2) {
      const modalidades = [...new Set(datosModulo06.map(r => r.modalidad))].filter(Boolean);
      const dataMod = modalidades.map(m => datosModulo06.filter(r => r.modalidad === m).length);

      new Chart(ctx2, {
        type: 'doughnut',
        data: {
          labels: modalidades.length ? modalidades : ['Sin datos'],
          datasets: [{
            data: dataMod.length ? dataMod : [1],
            backgroundColor: ['#667eea', '#28a745', '#ffc107', '#dc3545', '#17a2b8']
          }]
        },
        options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
      });
    }

    // Tabla
    mostrarTabla();
  }

  function mostrarTabla() {
    const tabla = document.querySelector('#modulo-06 table tbody');
    if (!tabla) return;

    tabla.innerHTML = datosModulo06.map(r => `
      <tr>
        <td>${r.ejercicio || '-'}</td>
        <td>${r.periodo || '-'}</td>
        <td>${r.beneficiario || '-'}</td>
        <td>${formatearMoneda(r.valorCRP || 0)}</td>
        <td>${formatearMoneda(r.giro || 0)}</td>
        <td>${formatearMoneda(r.pendiente || 0)}</td>
      </tr>
    `).join('');
  }

  function mostrarMensaje(modulo, texto, tipo) {
    console.log(`[Módulo ${modulo}] ${texto}`);
  }

  function formatearMoneda(valor) {
    return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', minimumFractionDigits: 0 }).format(valor);
  }

  // Exponer funciones globales
  window.modulo06 = { generarGraficas, actualizarKPIs };

  console.log('✅ Módulo 06 cargado correctamente');
})();
