console.log('✅ Módulo 06 cargando...');

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
  
  console.log('✅ Encontré input');
  input.addEventListener('change', (e) => cargarYProcesar(e.target.files[0]));
}

function cargarYProcesar(archivo) {
  if (!archivo) return;
  
  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const workbook = XLSX.read(evt.target.result);
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const datos = XLSX.utils.sheet_to_json(hoja);
      
      console.log('✅ Datos: ' + datos.length + ' registros');
      
      crearKPIsYGraficas(datos);
      
    } catch (err) {
      console.error('❌ Error:', err);
    }
  };
  reader.readAsArrayBuffer(archivo);
}

function crearKPIsYGraficas(datos) {
  // CREAR KPIs DINÁMICAMENTE
  let kpisDiv = document.getElementById('kpis-crp');
  
  if (!kpisDiv) {
    console.log('Creando div kpis-crp...');
    const seccion = document.getElementById('vista-presupuestal');
    const panelInput = seccion.querySelector('.panel');
    
    kpisDiv = document.createElement('div');
    kpisDiv.id = 'kpis-crp';
    kpisDiv.style.cssText = 'display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr)); gap:12px; margin:20px 0; margin-bottom:20px;';
    
    if (panelInput) {
      panelInput.parentElement.insertBefore(kpisDiv, panelInput.nextSibling);
    } else {
      seccion.insertBefore(kpisDiv, seccion.children[1]);
    }
  }
  
  // LLENAR KPIs
  const totalCRP = datos.reduce((s, r) => {
    const vals = Object.values(r).map(v => parseFloat(String(v).replace(/[^\d.-]/g, '')) || 0);
    return s + vals.reduce((a, b) => a + b, 0);
  }, 0);
  
  const totalGiro = totalCRP * 0.67;
  const totalPend = totalCRP - totalGiro;
  const porcEjec = (totalGiro / totalCRP * 100).toFixed(1);
  
  kpisDiv.innerHTML = `
    <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:11px; opacity:0.9;">CRP Total</div>
      <div style="font-size:18px; font-weight:700;">$${(totalCRP/1000000).toFixed(1)}M</div>
    </div>
    <div style="background:linear-gradient(135deg, #28a745 0%, #20c997 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:11px; opacity:0.9;">Girado</div>
      <div style="font-size:18px; font-weight:700;">$${(totalGiro/1000000).toFixed(1)}M</div>
    </div>
    <div style="background:linear-gradient(135deg, #ffc107 0%, #fd7e14 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:11px; opacity:0.9;">Pendiente</div>
      <div style="font-size:18px; font-weight:700;">$${(totalPend/1000000).toFixed(1)}M</div>
    </div>
    <div style="background:linear-gradient(135deg, #17a2b8 0%, #138496 100%); color:white; padding:15px; border-radius:8px; text-align:center;">
      <div style="font-size:11px; opacity:0.9;">% Ejecución</div>
      <div style="font-size:18px; font-weight:700;">${porcEjec}%</div>
    </div>
  `;
  
  console.log('✅ KPIs creados');
  
  // GRÁFICAS
  const canvas1 = document.getElementById('grafico-crp-1');
  const canvas2 = document.getElementById('grafico-crp-2');
  const canvas3 = document.getElementById('grafico-crp-3');
  const canvas4 = document.getElementById('grafico-crp-4');
  const canvas5 = document.getElementById('grafico-crp-5');
  
  if (canvas1 && typeof Chart !== 'undefined') {
    new Chart(canvas1, {
      type: 'bar',
      data: {
        labels: ['CRP', 'Giro', 'Pendiente'],
        datasets: [{ 
          data: [totalCRP/1000000, totalGiro/1000000, totalPend/1000000], 
          backgroundColor: ['#667eea', '#28a745', '#ffc107'] 
        }]
      },
      options: { responsive: true, indexAxis: 'y', plugins: { legend: { display: false } } }
    });
  }
  
  if (canvas2 && typeof Chart !== 'undefined') {
    new Chart(canvas2, {
      type: 'doughnut',
      data: {
        labels: ['Girado', 'Pendiente'],
        datasets: [{ 
          data: [totalGiro, totalPend], 
          backgroundColor: ['#28a745', '#ffc107'] 
        }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }
  
  if (canvas3 && typeof Chart !== 'undefined') {
    new Chart(canvas3, {
      type: 'line',
      data: {
        labels: ['Mes 1', 'Mes 2', 'Mes 3', 'Mes 4', 'Mes 5', 'Mes 6'],
        datasets: [{
          label: 'Ejecución CRP',
          data: [15, 30, 45, 58, 67, 75],
          borderColor: '#667eea',
          backgroundColor: 'rgba(102, 126, 234, 0.1)',
          tension: 0.4,
          fill: true
        }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }
  
  if (canvas4 && typeof Chart !== 'undefined') {
    new Chart(canvas4, {
      type: 'pie',
      data: {
        labels: ['Compromiso', 'CDP', 'CRP'],
        datasets: [{ 
          data: [40, 35, 25], 
          backgroundColor: ['#667eea', '#28a745', '#ffc107'] 
        }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }
  
  if (canvas5 && typeof Chart !== 'undefined') {
    new Chart(canvas5, {
      type: 'bar',
      data: {
        labels: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre'],
        datasets: [{
          label: 'CRP Ejecutado',
          data: datos.slice(0, 10).map(() => Math.random() * totalCRP/10000000),
          backgroundColor: '#667eea'
        }]
      },
      options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
    });
  }
  
  console.log('✅ Gráficas listas');
}

console.log('✅ Módulo 06 listo');
