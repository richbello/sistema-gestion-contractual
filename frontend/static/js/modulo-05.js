document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('form-historico');
    if (!form) return;

    const dropzone = document.getElementById('drop-historico-pagos');
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
    let chartMes = null;
    let chartProveedor = null;

    dropzone.addEventListener('click', () => inputHistorico.click());
    ['dragover', 'dragenter'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.style.background = '#e8f0ff'; }));
    ['dragleave', 'drop'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.style.background = ''; }));
    dropzone.addEventListener('drop', e => { if (e.dataTransfer.files.length) { inputHistorico.files = e.dataTransfer.files; enableBtn(); } });

    function enableBtn() { btnCargar.disabled = !inputHistorico.files.length; }

    function formatCurrency(value) {
        return new Intl.NumberFormat('es-CO', {style: 'currency', currency: 'COP', minimumFractionDigits: 0}).format(value || 0);
    }

    function formatDate(dateStr) {
        if (!dateStr) return '';
        try {
            const parts = String(dateStr).split('-');
            if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
            return dateStr;
        } catch { return dateStr; }
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
            <div class="metrica-card" style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div class="metrica-valor" style="font-size:24px; font-weight:bold;">${formatCurrency(totalPagado)}</div>
                <div class="metrica-etiqueta">Valor Total Pagado</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div class="metrica-valor" style="font-size:24px; font-weight:bold;">${totalTransacciones}</div>
                <div class="metrica-etiqueta">Transacciones</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div class="metrica-valor" style="font-size:24px; font-weight:bold;">${proveedoresUnicos}</div>
                <div class="metrica-etiqueta">Proveedores Únicos</div>
            </div>
        `;
    }

    function mostrarTabla(datos) {
        if (!datos.length) {
            contenedorTabla.innerHTML = '<p>No hay datos para mostrar.</p>';
            return;
        }

        const columnas = ['Nombre', 'Nº identificación', 'Valor Bruto', 'Fecha de pago', 'Referencia', 'Ejercicio'];
        let html = '<table style="width:100%; border-collapse:collapse; font-size:12px;"><thead><tr style="background:#667eea; color:white;">';

        columnas.forEach(col => {
            html += `<th style="padding:12px; text-align:left; border:1px solid #ddd; font-weight:bold;">${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        datos.forEach((fila, idx) => {
            html += `<tr style="background:${idx % 2 === 0 ? '#f9f9f9' : 'white'};">`;
            columnas.forEach(col => {
                let valor = fila[col] || '';
                if (col === 'Valor Bruto') valor = formatCurrency(valor);
                if (col === 'Fecha de pago') valor = formatDate(valor);
                html += `<td style="padding:12px; border:1px solid #ddd;">${valor}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        contenedorTabla.innerHTML = html;
    }

    function mostrarGraficos(datos) {
        const porMes = {}, porProveedor = {};
        
        datos.forEach(r => {
            const fecha = r['Fecha de pago'];
            if (fecha) {
                const mes = String(fecha).substring(0, 7);
                porMes[mes] = (porMes[mes] || 0) + (Number(r['Valor Bruto']) || 0);
            }
            const proveedor = r['Nombre'] || 'Sin nombre';
            porProveedor[proveedor] = (porProveedor[proveedor] || 0) + (Number(r['Valor Bruto']) || 0);
        });

        const meses = Object.keys(porMes).sort();
        const valoresMes = meses.map(m => porMes[m]);

        const coloresMes = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe', '#43e97b', '#fa7231'];
        const coloresProveedor = ['#ff6b6b', '#4ecdc4', '#45b7d1', '#ffa502', '#ff6348', '#ee5a6f', '#f9ca24', '#6c5ce7'];

        // Gráfico de barras por mes
        const ctxMes = document.createElement('canvas');
        ctxMes.id = 'grafico-mes';
        ctxMes.style.maxWidth = '100%';
        ctxMes.style.height = '300px';
        ctxMes.style.marginBottom = '40px';

        // Gráfico de pastel por proveedor (top 8)
        const ctxProveedor = document.createElement('canvas');
        ctxProveedor.id = 'grafico-proveedor';
        ctxProveedor.style.maxWidth = '100%';
        ctxProveedor.style.height = '300px';

        const contenedor = document.getElementById('grafico-pagos').parentElement;
        contenedor.innerHTML = '';
        contenedor.appendChild(ctxMes);
        contenedor.appendChild(ctxProveedor);

        if (window.Chart) {
            if (chartMes) chartMes.destroy();
            chartMes = new Chart(ctxMes, {
                type: 'bar',
                data: {
                    labels: meses,
                    datasets: [{
                        label: 'Pagos por Mes',
                        data: valoresMes,
                        backgroundColor: coloresMes.slice(0, meses.length),
                        borderColor: coloresMes.slice(0, meses.length),
                        borderWidth: 2,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: true, labels: { font: { size: 12, weight: 'bold' }, color: '#333' } },
                        title: { display: true, text: 'Pagos Realizados por Mes', font: { size: 14, weight: 'bold' }, color: '#333' }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: { callback: function(value) { return '$' + (value / 1000000).toFixed(0) + 'M'; }, font: { size: 11 } }
                        }
                    }
                }
            });

            const topProveedores = Object.entries(porProveedor).sort((a, b) => b[1] - a[1]).slice(0, 8);
            const nombresTop = topProveedores.map(p => p[0].substring(0, 15));
            const valoresTop = topProveedores.map(p => p[1]);

            if (chartProveedor) chartProveedor.destroy();
            chartProveedor = new Chart(ctxProveedor, {
                type: 'doughnut',
                data: {
                    labels: nombresTop,
                    datasets: [{
                        data: valoresTop,
                        backgroundColor: coloresProveedor,
                        borderColor: 'white',
                        borderWidth: 2
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: true, position: 'bottom', labels: { font: { size: 11 }, padding: 15 } },
                        title: { display: true, text: 'Top 8 Proveedores por Valor Pagado', font: { size: 14, weight: 'bold' }, color: '#333' }
                    }
                }
            });
        }
    }

    inputHistorico.addEventListener('change', enableBtn);

    filtroTabla.addEventListener('keyup', function() {
        const busqueda = this.value.toLowerCase();
        const filtrados = datosFiltrados.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(busqueda)));
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
            if (!archivo) throw new Error('Selecciona un archivo');

            const buf = await archivo.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
            datosCompletos = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

            datosFiltrados = filtroContrato
                ? datosCompletos.filter(r => String(r['Referencia'] || '').includes(filtroContrato))
                : datosCompletos;

            if (!datosFiltrados.length) throw new Error('No se encontraron datos');

            mostrarMetricas(datosFiltrados);
            mostrarTabla(datosFiltrados);
            mostrarGraficos(datosFiltrados);

            resultadosDiv.style.display = 'block';
            estadoDiv.style.display = 'none';

        } catch (err) {
            mostrarAlerta(`❌ ${err.message}`);
            estadoDiv.style.display = 'none';
        }
    });

    enableBtn();
});
