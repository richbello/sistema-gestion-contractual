document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('form-historico');
    if (!form) return;

    const inputHistorico = document.getElementById('input-historico-pagos');
    const inputFiltroContrato = document.getElementById('input-filtro-contrato');
    const btnCargar = document.getElementById('btn-cargar-historico');
    const alertaDiv = document.getElementById('alerta-historico');
    const estadoDiv = document.getElementById('estado-historico');
    const resultadosDiv = document.getElementById('resultados-historico');
    const metricasDiv = document.getElementById('metricas-historico');
    const filtroTabla = document.getElementById('filtro-tabla');
    const btnDescargar = document.getElementById('btn-descargar-filtrado');
    const contenedorTabla = document.getElementById('contenedor-tabla');

    let datosCompletos = [];
    let datosFiltrados = [];

    function enableBtn() {
        btnCargar.disabled = !inputHistorico.files.length;
    }

    function formatCurrency(value) {
        return new Intl.NumberFormat('es-CO', {style: 'currency', currency: 'COP', minimumFractionDigits: 0}).format(value || 0);
    }

    function mostrarAlerta(msg) {
        alertaDiv.textContent = msg;
        alertaDiv.style.display = 'block';
    }

    function mostrarMetricas(datos) {
        const totalPagado = datos.reduce((sum, r) => sum + (Number(r['Valor Bruto']) || 0), 0);
        const totalTransacciones = datos.length;
        const proveedoresUnicos = new Set(datos.map(r => r['Nombre'])).size;

        metricasDiv.innerHTML = `
            <div class="metrica-card"><div class="metrica-valor">${formatCurrency(totalPagado)}</div><div class="metrica-etiqueta">Valor Total Pagado</div></div>
            <div class="metrica-card"><div class="metrica-valor">${totalTransacciones}</div><div class="metrica-etiqueta">Transacciones</div></div>
            <div class="metrica-card"><div class="metrica-valor">${proveedoresUnicos}</div><div class="metrica-etiqueta">Proveedores Únicos</div></div>
        `;
    }

    function mostrarTabla(datos) {
        if (!datos.length) {
            contenedorTabla.innerHTML = '<p>No hay datos para mostrar.</p>';
            return;
        }

        const columnas = ['Nombre', 'Nº identificación', 'Valor Bruto', 'Fecha de pago', 'Referencia', 'Ejercicio'];
        let html = '<table class="tabla-historico" style="width:100%; border-collapse:collapse; font-size:12px;"><thead><tr style="background:#f0f0f0;">';

        columnas.forEach(col => {
            html += `<th style="padding:8px; text-align:left; border:1px solid #ddd; font-weight:bold;">${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        datos.forEach(fila => {
            html += '<tr>';
            columnas.forEach(col => {
                let valor = fila[col] || '';
                if (col === 'Valor Bruto') valor = formatCurrency(valor);
                html += `<td style="padding:8px; border:1px solid #ddd;">${valor}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        contenedorTabla.innerHTML = html;
    }

    function mostrarGrafico(datos) {
        const ctx = document.getElementById('grafico-pagos').getContext('2d');
        const porMes = {};

        datos.forEach(r => {
            const fecha = r['Fecha de pago'];
            if (fecha) {
                const mes = String(fecha).substring(0, 7); // YYYY-MM
                porMes[mes] = (porMes[mes] || 0) + (Number(r['Valor Bruto']) || 0);
            }
        });

        const meses = Object.keys(porMes).sort();
        const valores = meses.map(m => porMes[m]);

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: meses,
                datasets: [{
                    label: 'Pagos por Mes',
                    data: valores,
                    backgroundColor: '#667eea',
                    borderColor: '#667eea',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { display: true }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return '$' + (value / 1000000).toFixed(1) + 'M';
                            }
                        }
                    }
                }
            }
        });
    }

    inputHistorico.addEventListener('change', enableBtn);

    filtroTabla.addEventListener('keyup', function() {
        const busqueda = this.value.toLowerCase();
        const filtrados = datosFiltrados.filter(r =>
            Object.values(r).some(v => String(v).toLowerCase().includes(busqueda))
        );
        mostrarTabla(filtrados);
    });

    btnDescargar.addEventListener('click', function() {
        if (!datosFiltrados.length) return;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(datosFiltrados);
        XLSX.utils.book_append_sheet(wb, ws, 'Historico');
        XLSX.writeFile(wb, 'Historico_Pagos_Filtrado.xlsx');
    });

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        alertaDiv.style.display = 'none';
        estadoDiv.style.display = 'block';
        resultadosDiv.style.display = 'none';

        try {
            if (!window.XLSX) throw new Error('XLSX no cargado');

            const archivo = inputHistorico.files[0];
            const filtroContrato = inputFiltroContrato.value.trim();

            const buf = await archivo.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
            datosCompletos = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

            datosFiltrados = filtroContrato
                ? datosCompletos.filter(r => String(r['Referencia'] || '').includes(filtroContrato))
                : datosCompletos;

            if (!datosFiltrados.length) {
                throw new Error('No se encontraron datos');
            }

            mostrarMetricas(datosFiltrados);
            mostrarTabla(datosFiltrados);
            mostrarGrafico(datosFiltrados);

            resultadosDiv.style.display = 'block';
            estadoDiv.style.display = 'none';

        } catch (err) {
            mostrarAlerta(`❌ ${err.message}`);
            estadoDiv.style.display = 'none';
        }
    });

    enableBtn();
});
