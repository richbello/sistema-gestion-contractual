document.addEventListener('DOMContentLoaded', function () {
  const API = window.API_BASE || '';
  const inputs = [
    { id: 'input-plantilla-ec', name: 'plantillaName', area: 'plantillaUploadArea' },
    { id: 'input-historico-ec', name: 'historicoName', area: 'historicoUploadArea' },
    { id: 'input-crp-ec', name: 'crpName', area: 'crpUploadArea' },
    { id: 'input-secop2-ec', name: 'secop2Name', area: 'secop2UploadArea' }
  ];

  inputs.forEach(inp => {
    const inputEl = document.getElementById(inp.id);
    const nameEl = document.getElementById(inp.name);
    const areaEl = document.getElementById(inp.area);

    if (inputEl && nameEl && areaEl) {
      areaEl.addEventListener('click', () => inputEl.click());
      inputEl.addEventListener('change', (e) => {
        nameEl.textContent = e.target.files[0]?.name || '';
      });
      areaEl.addEventListener('dragover', (e) => {
        e.preventDefault();
        areaEl.style.backgroundColor = '#e8f5ff';
      });
      areaEl.addEventListener('dragleave', () => {
        areaEl.style.backgroundColor = '';
      });
      areaEl.addEventListener('drop', (e) => {
        e.preventDefault();
        inputEl.files = e.dataTransfer.files;
        nameEl.textContent = inputEl.files[0]?.name || '';
      });
    }
  });

  document.getElementById('btn-generar-estadocuenta')?.addEventListener('click', async () => {
    const plantilla = document.getElementById('input-plantilla-ec').files[0];
    const historico = document.getElementById('input-historico-ec').files[0];
    const crp = document.getElementById('input-crp-ec').files[0];
    const secop = document.getElementById('input-secop2-ec').files[0];
    const contrato = document.getElementById('input-contrato-ec').value.trim();

    if (!plantilla || !historico || !crp || !secop || !contrato) {
      mostrarAlerta('alerta-estadocuenta', 
        'Debes adjuntar plantilla, histórico, Reporte CRP, Consolidado SECOP2, e indicar el número de contrato.',
        'error');
      return;
    }

    const spinner = document.getElementById('spinner-ec');
    const btn = document.getElementById('btn-generar-estadocuenta');
    spinner.style.display = 'block';
    btn.disabled = true;

    try {
      const fd = new FormData();
      fd.append('plantilla', plantilla);
      fd.append('reporte_crp', crp);
      fd.append('consolidado', secop);
      fd.append('historico', historico);
      fd.append('contrato', contrato);

      const resp = await fetch(API + '/api/estado-cuenta/procesar', { 
        method: 'POST', 
        body: fd 
      });

      if (resp.ok) {
        const blob = await resp.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Estado_de_Cuenta_${contrato}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        mostrarAlerta('alerta-estadocuenta', 
          `Estado de Cuenta del contrato ${contrato} generado correctamente.`, 
          'success');
      } else {
        const data = await resp.json();
        mostrarAlerta('alerta-estadocuenta', 
          data.mensaje || 'Error al generar el estado de cuenta.', 
          'error');
      }
    } catch (error) {
      mostrarAlerta('alerta-estadocuenta', 
        `Error: ${error.message}`, 
        'error');
    } finally {
      spinner.style.display = 'none';
      btn.disabled = false;
    }
  });
});

function mostrarAlerta(id, mensaje, tipo) {
  const elem = document.getElementById(id);
  if (elem) {
    elem.textContent = mensaje;
    elem.className = `alerta ${tipo}`;
    elem.style.display = 'block';
    setTimeout(() => {
      elem.style.display = 'none';
    }, 5000);
  }
}