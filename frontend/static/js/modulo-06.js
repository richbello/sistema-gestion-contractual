(function() {
  const datos = [];
  
  const inputCarga = document.getElementById('carga-06');
  if (!inputCarga) {
    console.error('No encontré input carga-06');
    return;
  }

  inputCarga.addEventListener('change', function(e) {
    const archivo = e.target.files[0];
    if (!archivo) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const workbook = XLSX.read(evt.target.result, { header: 1 });
        const hoja = workbook.Sheets[workbook.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja);

        datos.length = 0;
        datos.push(...filas);

        console.log('Datos cargados:', datos.length);
        actualizarVista();

      } catch (err) {
        console.error('Error Excel:', err);
      }
    };
    reader.readAsArrayBuffer(archivo);
  });

  function actualizarVista() {
    if (datos.length === 0) return;

    const totalCRP = datos.reduce((s, r) => {
      const v = Object.values(r).find(v => String(v).includes('CRP'));
      return s + (parseFloat(v) || 0);
    }, 0);

    console.log('Total CRP:', totalCRP);
    console.log('Registros:', datos.length);
  }

  console.log('Modulo 06 iniciado');
})();
