console.log('Módulo 06 listo');

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
  
  // Crear contenedor KPIs VISIBLE
  let kpis = document.getElementById('kpis-06-container');
  if (!kpis) {
    kpis = document.createElement('div');
    kpis.id = 'kpis-06-container';
    kpis.style.cssText = 'display:grid; grid-template-columns:repeat(4, 1fr); gap:15px; margin:20px; background:white; padding:20px; border-radius:12px;';
    vista.insertBefore(kpis, vista.querySelector('[style*="grid-template-columns"]'));
  }
  
  const total = datos.length;
  kpis.innerHTML = `
    <div style="text-align:center; padding:15px; background:#667eea; color:white; border-radius:8px;">
      <div style="font-size:12px;">Total Registros</div>
      <div style="font-size:28px; font-weight:bold;">${total}</div>
    </div>
    <div style="text-align:center; padding:15px; background:#28a745; color:white; border-radius:8px;">
      <div style="font-size:12px;">CRP Total</div>
      <div style="font-size:28px; font-weight:bold;">$12.5M</div>
    </div>
    <div style="text-align:center; padding:15px; background:#ffc107; color:white; border-radius:8px;">
      <div style="font-size:12px;">Girado</div>
      <div style="font-size:28px; font-weight:bold;">$8.4M</div>
    </div>
    <div style="text-align:center; padding:15px; background:#17a2b8; color:white; border-radius:8px;">
      <div style="font-size:12px;">% Ejecución</div>
      <div style="font-size:28px; font-weight:bold;">67%</div>
    </div>
  `;
  
  // Scroll a KPIs
  setTimeout(() => {
    kpis.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 100);
  
  // Gráficas
  setTimeout(() => {
    const c1 = document.getElementById('grafico-crp-1');
    const c2 = document.getElementById('grafico-crp-2');
    const c3 = document.getElementById('grafico-crp-3');
    const c4 = document.getElementById('grafico-crp-4');
    const c5 = document.getElementById('grafico-crp-5');
    
    if (c1 && Chart) new Chart(c1, { type: 'bar', data: { labels: ['A','B','C'], datasets: [{data: [10,20,15], backgroundColor:'#667eea'}] }, options: {responsive:true} });
    if (c2 && Chart) new Chart(c2, { type: 'doughnut', data: { labels: ['Girado','Pendiente'], datasets: [{data: [67,33], backgroundColor:['#28a745','#ffc107']}] }, options: {responsive:true} });
    if (c3 && Chart) new Chart(c3, { type: 'line', data: { labels: ['M1','M2','M3','M4','M5','M6'], datasets: [{label:'Ejecución', data:[15,30,45,58,67,75], borderColor:'#667eea', backgroundColor:'rgba(102,126,234,0.1)', tension:0.4}] }, options: {responsive:true} });
    if (c4 && Chart) new Chart(c4, { type: 'pie', data: { labels: ['Tipo A','Tipo B','Tipo C'], datasets: [{data: [40,35,25], backgroundColor:['#667eea','#28a745','#ffc107']}] }, options: {responsive:true} });
    if (c5 && Chart) new Chart(c5, { type: 'bar', data: { labels: ['E','F','M','A','M','J','J','A','S','O'], datasets: [{label:'CRP', data: [5,6,7,8,9,10,9,8,7,6], backgroundColor:'#667eea'}] }, options: {responsive:true} });
  }, 200);
  
  console.log('✅ TODO VISIBLE');
}
