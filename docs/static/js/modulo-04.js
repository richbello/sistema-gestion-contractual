document.addEventListener('DOMContentLoaded', function () {
  const form = document.getElementById('form-estadocuenta');
  if (!form) return;

  const API = (typeof window.BACKEND_URL === 'string' ? window.BACKEND_URL : '').replace(/\/$/, '');

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    mostrarAlerta('alerta-estadocuenta', '');
    document.getElementById('resultados-estadocuenta').classList.remove('visible');

    const plantilla = document.getElementById('input-plantilla-ec').files[0];
    const historico = document.getElementById('input-historico-ec').files[0];
    const contrato = document.getElementById('input-contrato-ec').value.trim();

    if (!plantilla || !historico || !contrato) {
      mostrarAlerta('alerta-estadocuenta', 'Debes adjuntar ambos archivos e indicar el número de contrato.');
      return;
    }

    const btn = document.getElementById('btn-estadocuenta');
    btn.disabled = true;
    mostrarEstado('estado-estadocuenta', true);

    try {
      // 1) Leer y filtrar el histórico EN EL NAVEGADOR (sin limite de memoria del servidor)
      const buf = await historico.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      const hoja = wb.Sheets[wb.SheetNames[0]];
      const crudo = XLSX.utils.sheet_to_json(hoja, { raw: true, defval: '' });
      const texto = XLSX.utils.sheet_to_json(hoja, { raw: false, defval: '' });

      const pagos = [];
      for (let i = 0; i < crudo.length; i++) {
        if (String(crudo[i]['Referencia'] || '').includes(contrato)) {
          const p = Object.assign({}, texto[i]);
          p['Valor Bruto'] = Number(crudo[i]['Valor Bruto']) || 0;
          p['VALOR FINAL DEL CONTRATO'] = Number(crudo[i]['VALOR FINAL DEL CONTRATO']) || 0;
          pagos.push(p);
        }
      }

      if (!pagos.length) {
        mostrarAlerta('alerta-estadocuenta', 'No se encontró información para: ' + contrato);
        return;
      }

      // 2) Enviar SOLO plantilla (23 KB) + pagos filtrados al backend
      const fd = new FormData();
      fd.append('plantilla', plantilla);
      fd.append('contrato', contrato);
      fd.append('pagos', JSON.stringify(pagos));

      const resp = await fetch(API + '/api/estado-cuenta/procesar-lite', { method: 'POST', body: fd });

      if (!resp.ok) {
        let msg = 'No se pudo generar el estado de cuenta.';
        try { const d = await resp.json(); if (d.mensaje) msg = d.mensaje; } catch (_) {}
        mostrarAlerta('alerta-estadocuenta', msg);
        return;
      }

      // 3) Descargar el Excel con el formato ORIGINAL intacto
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Estado_de_Cuenta_' + contrato + '.xlsx';
      a.click();

      const link = document.getElementById('link-descarga-estadocuenta');
      link.href = url;
      link.download = 'Estado_de_Cuenta_' + contrato + '.xlsx';

      // 4) Métricas
      const primer = pagos[0];
      let saldo = primer['VALOR FINAL DEL CONTRATO'];
      pagos.forEach(p => { saldo -= p['Valor Bruto']; });

      const cont = document.getElementById('metricas-estadocuenta');
      cont.innerHTML = '';
      cont.appendChild(crearMetrica(contrato, 'Contrato'));
      cont.appendChild(crearMetrica(pagos.length, 'Pagos encontrados', 'verde'));
      cont.appendChild(crearMetrica(formatearMoneda(primer['VALOR FINAL DEL CONTRATO']), 'Valor del contrato'));
      cont.appendChild(crearMetrica(formatearMoneda(saldo), 'Saldo final', 'ocre'));

      document.getElementById('nombre-descarga-ec').textContent = 'Estado_de_Cuenta_' + contrato + '.xlsx';
      document.getElementById('resultados-estadocuenta').classList.add('visible');

    } catch (err) {
      mostrarAlerta('alerta-estadocuenta', 'Error: ' + err.message);
      console.error(err);
    } finally {
      btn.disabled = false;
      mostrarEstado('estado-estadocuenta', false);
    }
  });
});
