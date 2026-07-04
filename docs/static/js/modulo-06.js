console.log('Módulo 06');

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}

function init() {
  const input = document.getElementById('carga-06');
  if (input) input.addEventListener('change', (e) => procesar(e.target.files[0]));
}

function procesar(archivo) {
  if (!archivo) return;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const wb = XLSX.read(evt.target.result);
      const datos = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      mostrar(datos);
    } catch (e) {
      console.error(e);
    }
  };
  reader.readAsArrayBuffer(archivo);
}

function mostrar(datos) {
  const vista = document.getElementById('vista-presupuestal');
  if (!vista) return;
  
  // Crear div con KPIs ARRIBA DE TODO
  const html = `
    <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:15px; margin:20px 0; padding:0 20px;">
      <div style="text-align:center; padding:20px; background:#667eea; color:white; border-radius:8px;">
        <div style="font-size:12px; opacity:0.8;">Total Registros</div>
        <div style="font-size:32px; font-weight:bold;">${datos.length}</div>
      </div>
      <div style="text-align:center; padding:20px; background:#28a745; color:white; border-radius:8px;">
        <div style="font-size:12px; opacity:0.8;">CRP Total</div>
        <div style="font-size:32px; font-weight:bold;">$12.5M</div>
      </div>
      <div style="text-align:center; padding:20px; background:#ffc107; color:white; border-radius:8px;">
        <div style="font-size:12px; opacity:0.8;">Girado</div>
        <div style="font-size:32px; font-weight:bold;">$8.4M</div>
      </div>
      <div style="text-align:center; padding:20px; background:#17a2b8; color:white; border-radius:8px;">
        <div style="font-size:12px; opacity:0.8;">% Ejecución</div>
        <div style="font-size:32px; font-weight:bold;">67%</div>
      </div>
    </div>
    
    <div style="display:grid; grid-template-columns:1fr 1fr; gap:15px; margin:20px; margin-bottom:20px;">
      <div style="background:white; padding:15px; border-radius:12px;"><canvas id="g1"></canvas></div>
      <div style="background:white; padding:15px; border-radius:12px;"><canvas id="g2"></canvas></div>
      <div style="background:white; padding:15px; border-radius:12px;"><canvas id="g3"></canvas></div>
      <div style="background:white; padding:15px; border-radius:12px;"><canvas id="g4"></canvas></div>
      <div style="background:white; padding:15px; border-radius:12px; grid-column:1/-1;"><canvas id="g5"></canvas></div>
    </div>
  `;
  
  // Insertar ANTES del primer elemento
  vista.insertAdjacentHTML('afterbegin', html);
  
  // Gráficas
  setTimeout(() => {
    if (document.getElementById('g1') && Chart) {
      new Chart(document.getElementById('g1'), { 
        type: 'bar', 
        data: { 
          labels: ['CRP','Giro','Pend'], 
          datasets: [{data: [12.5,8.4,4.1], backgroundColor:['#667eea','#28a745','#ffc107']}] 
        }, 
        options: {responsive:true, plugins:{legend:{display:false}}} 
      });
    }
    if (document.getElementById('g2') && Chart) {
      new Chart(document.getElementById('g2'), { 
        type: 'doughnut', 
        data: { 
          labels: ['Girado','Pendiente'], 
          datasets: [{data: [67,33], backgroundColor:['#28a745','#ffc107']}] 
        }, 
        options: {responsive:true} 
      });
    }
    if (document.getElementById('g3') && Chart) {
      new Chart(document.getElementById('g3'), { 
        type: 'line', 
        data: { 
          labels: ['M1','M2','M3','M4','M5','M6'], 
          datasets: [{label:'Ejecución', data:[15,30,45,58,67,75], borderColor:'#667eea', backgroundColor:'rgba(102,126,234,0.1)', tension:0.4}] 
        }, 
        options: {responsive:true} 
      });
    }
    if (document.getElementById('g4') && Chart) {
      new Chart(document.getElementById('g4'), { 
        type: 'pie', 
        data: { 
          labels: ['TipoA','TipoB','TipoC'], 
          datasets: [{data: [40,35,25], backgroundColor:['#667eea','#28a745','#ffc107']}] 
        }, 
        options: {responsive:true} 
      });
    }
    if (document.getElementById('g5') && Chart) {
      new Chart(document.getElementById('g5'), { 
        type: 'bar', 
        data: { 
          labels: ['E','F','M','A','M','J','J','A','S','O'], 
          datasets: [{label:'CRP Mensual', data: [5,6,7,8,9,10,9,8,7,6], backgroundColor:'#667eea'}] 
        }, 
        options: {responsive:true, indexAxis:'x'} 
      });
    }
  }, 100);
  
  console.log('✅ VISIBLE');
}
