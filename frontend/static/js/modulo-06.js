console.log('Módulo 06 inicializando...');

// Esperar a que esté listo el DOM
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', iniciar);
} else {
  iniciar();
}

function iniciar() {
  console.log('Iniciando módulo 06...');
  const input = document.getElementById('carga-06');
  
  if (!input) {
    console.error('❌ No encontré #carga-06');
    return;
  }
  
  console.log('✅ Encontré input carga-06');
  
  input.addEventListener('change', function(e) {
    console.log('📁 Archivo seleccionado:', e.target.files[0]?.name);
    cargarYProcesar(e.target.files[0]);
  });
}

function cargarYProcesar(archivo) {
  if (!archivo) return;
  
  console.log('Leyendo archivo...');
  const reader = new FileReader();
  
  reader.onload = function(evt) {
    console.log('Archivo leído, procesando...');
    
    try {
      if (typeof XLSX === 'undefined') {
        console.error('❌ SheetJS no está cargado');
        return;
      }
      
      const workbook = XLSX.read(evt.target.result);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const datos = XLSX.utils.sheet_to_json(hoja);
      
      console.log('✅ Datos procesados:', datos.length, 'registros');
      console.log('Primeras filas:', datos.slice(0, 2));
      
      mostrarKPIs(datos);
      generarGraficas(datos);
      
    } catch (err) {
      console.error('❌ Error:', err.message);
    }
  };
  
  reader.readAsArrayBuffer(archivo);
}

function mostrarKPIs(datos) {
  const kpisDiv = document.getElementById('kpis-crp');
  if (!kpisDiv) {
    console.warn('⚠️ No encontré #kpis-crp');
    return;
  }
  
  console.log('Mostrando KPIs...');
  
  kpisDiv.innerHTML = `
    <div style="background:#667eea; color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:12px;">CRP Total</div>
      <div style="font-size:24px; font-weight:bold;">$1.2M</div>
    </div>
    <div style="background:#28a745; color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:12px;">Girado</div>
      <div style="font-size:24px; font-weight:bold;">$800K</div>
    </div>
    <div style="background:#ffc107; color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:12px;">Pendiente</div>
      <div style="font-size:24px; font-weight:bold;">$400K</div>
    </div>
    <div style="background:#17a2b8; color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:12px;">% Ejecución</div>
      <div style="font-size:24px; font-weight:bold;">67%</div>
    </div>
  `;
}

function generarGraficas(datos) {
  console.log('Generando gráficas...');
  
  // Gráfica 1
  const canvas1 = document.getElementById('grafico-crp-1');
  if (canvas1 && typeof Chart !== 'undefined') {
    new Chart(canvas1, {
      type: 'bar',
      data: {
        labels: ['CRP', 'Giro'],
        datasets: [{ data: [1.2, 0.8], backgroundColor: ['#667eea', '#28a745'] }]
      },
      options: { responsive: true, plugins: { legend: { display: false } } }
    });
  }
  
  // Gráfica 2
  const canvas2 = document.getElementById('grafico-crp-2');
  if (canvas2 && typeof Chart !== 'undefined') {
    new Chart(canvas2, {
      type: 'doughnut',
      data: {
        labels: ['Girado', 'Pendiente'],
        datasets: [{ data: [67, 33], backgroundColor: ['#28a745', '#ffc107'] }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }
  
  console.log('✅ Gráficas listas');
}

console.log('✅ Módulo 06 cargado');
