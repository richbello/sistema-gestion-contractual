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
            const plantillaWB = XLSX.read(new Uint8Array(plantillaData), { type: 'array' });
            const ws = plantillaWB.Sheets[plantillaWB.SheetNames[0]];

            // Escribir valores en celdas específicas sin perder formato
            ws['D5'] = { v: contratoBuscado, t: 's' };
            ws['D6'] = { v: String(primer['Nombre'] || ''), t: 's' };
            ws['D7'] = { v: Number(primer['VALOR FINAL DEL CONTRATO']) || 0, t: 'n' };
            ws['D8'] = { v: String(primer['FECHA INICIAL DE CONTRATO'] || ''), t: 's' };
            ws['H5'] = { v: String(primer['Proveedor'] || ''), t: 's' };
            ws['H6'] = { v: String(primer['Nº identificación'] || ''), t: 's' };
            ws['H7'] = { v: String(primer['Numero RP'] || ''), t: 's' };
            ws['H8'] = { v: String(primer['FECHA DE TERMINACION FINAL'] || ''), t: 's' };

            let fila = 17, saldo = Number(primer['VALOR FINAL DEL CONTRATO']) || 0;
            pagos.forEach((p, i) => {
                const monto = Number(p['Valor Bruto']) || 0;
                saldo -= monto;
                ws[`B${fila}`] = { v: i + 1, t: 'n' };
                ws[`C${fila}`] = { v: String(p['Texto cabecera documento'] || ''), t: 's' };
                ws[`D${fila}`] = { v: monto, t: 'n' };
                ws[`E${fila}`] = { v: saldo, t: 'n' };
                ws[`F${fila}`] = { v: String(p['Doc.compensación'] || ''), t: 's' };
                ws[`G${fila}`] = { v: String(p['Fecha de pago'] || ''), t: 's' };
                ws[`H${fila}`] = { v: String(p['Numero RP'] || ''), t: 's' };
                ws[`I${fila}`] = { v: String(p['CDP Externo'] || ''), t: 's' };
                ws[`J${fila}`] = { v: String(p['CRP Externo'] || ''), t: 's' };
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
