document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('form-estadocuenta');
    if (!form) return;

    const inputPlantilla = document.getElementById('input-plantilla-ec');
    const inputHistorico = document.getElementById('input-historico-ec');
    const inputContrato = document.getElementById('input-contrato-ec');
    const btnGenerar = document.getElementById('btn-estadocuenta');
    const alertaDiv = document.getElementById('alerta-estadocuenta');
    const estadoDiv = document.getElementById('estado-estadocuenta');

    function enableBtn() {
        btnGenerar.disabled = !(inputPlantilla.files.length && inputHistorico.files.length && inputContrato.value.trim());
    }

    inputPlantilla.addEventListener('change', enableBtn);
    inputHistorico.addEventListener('change', enableBtn);
    inputContrato.addEventListener('input', enableBtn);

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        alertaDiv.style.display = 'none';
        estadoDiv.style.display = 'block';

        try {
            if (!window.XLSX) throw new Error('XLSX no cargado');

            const contratoBuscado = inputContrato.value.trim();
            const historicoData = await inputHistorico.files[0].arrayBuffer();
            const historicoWB = XLSX.read(new Uint8Array(historicoData), { type: 'array' });
            const historico = XLSX.utils.sheet_to_json(historicoWB.Sheets[historicoWB.SheetNames[0]]);

            const pagos = historico.filter(r => String(r.Referencia || '').includes(contratoBuscado));
            if (!pagos.length) throw new Error(`No encontrado: ${contratoBuscado}`);

            const primer = pagos[0];
            const plantillaData = await inputPlantilla.files[0].arrayBuffer();
            
            // Leer manteniedo TODO el formato original
            const plantillaWB = XLSX.read(new Uint8Array(plantillaData), { type: 'array', cellStyles: true, cellFormula: true });
            const ws = plantillaWB.Sheets[plantillaWB.SheetNames[0]];

            // SOLO cambiar valores, preservar TODO lo demás
            if (ws['D5']) ws['D5'].v = contratoBuscado; else ws['D5'] = { v: contratoBuscado };
            if (ws['D6']) ws['D6'].v = String(primer['Nombre'] || ''); else ws['D6'] = { v: String(primer['Nombre'] || '') };
            if (ws['D7']) ws['D7'].v = Number(primer['VALOR FINAL DEL CONTRATO']) || 0; else ws['D7'] = { v: Number(primer['VALOR FINAL DEL CONTRATO']) || 0 };
            if (ws['D8']) ws['D8'].v = String(primer['FECHA INICIAL DE CONTRATO'] || ''); else ws['D8'] = { v: String(primer['FECHA INICIAL DE CONTRATO'] || '') };
            if (ws['H5']) ws['H5'].v = String(primer['Proveedor'] || ''); else ws['H5'] = { v: String(primer['Proveedor'] || '') };
            if (ws['H6']) ws['H6'].v = String(primer['Nº identificación'] || ''); else ws['H6'] = { v: String(primer['Nº identificación'] || '') };
            if (ws['H7']) ws['H7'].v = String(primer['Numero RP'] || ''); else ws['H7'] = { v: String(primer['Numero RP'] || '') };
            if (ws['H8']) ws['H8'].v = String(primer['FECHA DE TERMINACION FINAL'] || ''); else ws['H8'] = { v: String(primer['FECHA DE TERMINACION FINAL'] || '') };

            let fila = 17, saldo = Number(primer['VALOR FINAL DEL CONTRATO']) || 0;
            pagos.forEach((p, i) => {
                const monto = Number(p['Valor Bruto']) || 0;
                saldo -= monto;
                const cell_b = `B${fila}`, cell_c = `C${fila}`, cell_d = `D${fila}`, cell_e = `E${fila}`, cell_f = `F${fila}`, cell_g = `G${fila}`, cell_h = `H${fila}`, cell_i = `I${fila}`, cell_j = `J${fila}`;
                if (ws[cell_b]) ws[cell_b].v = i + 1; else ws[cell_b] = { v: i + 1 };
                if (ws[cell_c]) ws[cell_c].v = String(p['Texto cabecera documento'] || ''); else ws[cell_c] = { v: String(p['Texto cabecera documento'] || '') };
                if (ws[cell_d]) ws[cell_d].v = monto; else ws[cell_d] = { v: monto };
                if (ws[cell_e]) ws[cell_e].v = saldo; else ws[cell_e] = { v: saldo };
                if (ws[cell_f]) ws[cell_f].v = String(p['Doc.compensación'] || ''); else ws[cell_f] = { v: String(p['Doc.compensación'] || '') };
                if (ws[cell_g]) ws[cell_g].v = String(p['Fecha de pago'] || ''); else ws[cell_g] = { v: String(p['Fecha de pago'] || '') };
                if (ws[cell_h]) ws[cell_h].v = String(p['Numero RP'] || ''); else ws[cell_h] = { v: String(p['Numero RP'] || '') };
                if (ws[cell_i]) ws[cell_i].v = String(p['CDP Externo'] || ''); else ws[cell_i] = { v: String(p['CDP Externo'] || '') };
                if (ws[cell_j]) ws[cell_j].v = String(p['CRP Externo'] || ''); else ws[cell_j] = { v: String(p['CRP Externo'] || '') };
                fila++;
            });

            XLSX.writeFile(plantillaWB, `Estado_de_Cuenta_${contratoBuscado}.xlsx`);
            alertaDiv.style.display = 'none';
            estadoDiv.style.display = 'none';

        } catch (error) {
            alertaDiv.textContent = `❌ ${error.message}`;
            alertaDiv.style.display = 'block';
            estadoDiv.style.display = 'none';
        }
    });

    enableBtn();
});
