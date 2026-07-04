document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('form-historico');
    if (!form) return;

    const dropzone = document.getElementById('drop-historico-pagos');
    const inputHistorico = document.getElementById('input-historico-pagos');
    const btnCargar = document.getElementById('btn-cargar-historico');
    const alertaDiv = document.getElementById('alerta-historico');
    const estadoDiv = document.getElementById('estado-historico');
    const resultadosDiv = document.getElementById('resultados-historico');
    const metricasDiv = document.getElementById('metricas-historico');
    
    const filtroContratista = document.getElementById('filtro-contratista');
    const filtroReferencia = document.getElementById('filtro-referencia');
    const filtroFechaDesde = document.getElementById('filtro-fecha-desde');
    const filtroFechaHasta = document.getElementById('filtro-fecha-hasta');
    const filtroVigencia = document.getElementById('filtro-vigencia');
    const btnLimpiar = document.getElementById('btn-limpiar-filtros');
    const btnGraficas = document.getElementById('btn-vista-graficas');
    const btnTabla = document.getElementById('btn-vista-tabla');
    const btnDescargar = document.getElementById('btn-descargar-filtrado');
    const contenedorGraficas = document.getElementById('contenedor-graficas');
    const contenedorTabla = document.getElementById('contenedor-tabla');

    let datosCompletos = [];
    let datosFiltrados = [];
    let resumenContratistas = {};
    let chartCombo = null;
    let chartProveedores = null;

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

    function construirResumenContratistas() {
        resumenContratistas = {};
        datosCompletos.forEach(r => {
            const nombre = r['Nombre'] || 'Sin nombre';
            if (!resumenContratistas[nombre]) {
                resumenContratistas[nombre] = {
                    total: 0,
                    transacciones: 0,
                    ultimaFecha: null,
                    nit: r['Nº identificación'] || ''
                };
            }
            resumenContratistas[nombre].total += Number(r['Valor Bruto']) || 0;
            resumenContratistas[nombre].transacciones += 1;
            const fecha = r['Fecha de pago'];
            if (fecha && (!resumenContratistas[nombre].ultimaFecha || fecha > resumenContratistas[nombre].ultimaFecha)) {
                resumenContratistas[nombre].ultimaFecha = fecha;
            }
        });
    }

    function mostrarMetricas(datos) {
        const totalPagado = datos.reduce((sum, r) => sum + (Number(r['Valor Bruto']) || 0), 0);
        const totalTransacciones = datos.length;
        const proveedoresUnicos = new Set(datos.map(r => r['Nombre'])).size;

        metricasDiv.innerHTML = `
            <div class="metrica-card" style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${formatCurrency(totalPagado)}</div>
                <div style="font-size:12px; margin-top:5px;">Valor Total Pagado</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${totalTransacciones.toLocaleString()}</div>
                <div style="font-size:12px; margin-top:5px;">Transacciones</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${proveedoresUnicos.toLocaleString()}</div>
                <div style="font-size:12px; margin-top:5px;">Proveedores</div>
            </div>
        `;
    }

    function aplicarFiltros() {
        const busqContratista = filtroContratista.value.toLowerCase();
        const busqReferencia = filtroReferencia.value.toLowerCase();
        const fechaDesde = filtroFechaDesde.value ? new Date(filtroFechaDesde.value) : null;
        const fechaHasta = filtroFechaHasta.value ? new Date(filtroFechaHasta.value) : null;
        const vigencia = filtroVigencia.value;

        datosFiltrados = datosCompletos.filter(r => {
            const nombre = String(r['Nombre'] || '').toLowerCase();
            const nit = String(r['Nº identificación'] || '').toLowerCase();
            const referencia = String(r['Referencia'] || '').toLowerCase();
            const vig = String(r['Ejercicio'] || '');

            const cumpleContratista = !busqContratista || nombre.includes(busqContratista) || nit.includes(busqContratista);
            const cumpleReferencia = !busqReferencia || referencia.includes(busqReferencia);
            const cumpleVigencia = !vigencia || vig === vigencia;

            let cumpleFecha = true;
            if (fechaDesde || fechaHasta) {
                const fecha = r['Fecha de pago'] ? new Date(r['Fecha de pago']) : null;
                if (fecha) {
                    if (fechaDesde && fecha < fechaDesde) cumpleFecha = false;
                    if (fechaHasta && fecha > fechaHasta) cumpleFecha = false;
                }
            }

            return cumpleContratista && cumpleReferencia && cumpleVigencia && cumpleFecha;
        });

        btnTabla.textContent = `📋 Tabla (${datosFiltrados.length})`;
        mostrarGraficos();
        mostrarTabla();
    }

    function mostrarTabla() {
        if (!datosFiltrados.length) {
            contenedorTabla.innerHTML = '<p>No hay datos para mostrar.</p>';
            return;
        }

        const columnas = ['Nombre', 'Nº identificación', 'Valor Bruto', 'Fecha de pago', 'Referencia', 'Ejercicio'];
        let html = '<table style="width:100%; border-collapse:collapse; font-size:11px;"><thead><tr style="background:#667eea; color:white;">';

        columnas.forEach(col => {
            html += `<th style="padding:12px; text-align:left; border:1px solid #ddd; font-weight:bold;">${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        datosFiltrados.forEach((fila, idx) => {
            html += `<tr style="background:${idx % 2 === 0 ? '#f9f9f9' : 'white'};">`;
            columnas.forEach(col => {
                let valor = fila[col] || '';
                if (col === 'Valor Bruto') valor = formatCurrency(valor);
                if (col === 'Fecha de pago') valor = formatDate(valor);
                html += `<td style="padding:10px; border:1px solid #ddd;">${valor}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';
        contenedorTabla.innerHTML = html;
    }

    function mostrarGraficos() {
        if (!datosFiltrados.length) return;

        const porVigencia = {}, porProveedor = {};
        const vigencias = new Set();
        
        datosFiltrados.forEach(r => {
            const vig = r['Ejercicio'] || 'Sin año';
            vigencias.add(vig);
            porVigencia[vig] = (porVigencia[vig] || 0) + (Number(r['Valor Bruto']) || 0);

            const proveedor = r['Nombre'] || 'Sin nombre';
            porProveedor[proveedor] = (porProveedor[proveedor] || 0) + (Number(r['Valor Bruto']) || 0);
        });

        const vigenciasOrdenadas = Array.from(vigencias).sort();
        const valoresVigencia = vigenciasOrdenadas.map(v => porVigencia[v] || 0);
        const countVigencia = vigenciasOrdenadas.map(v => datosFiltrados.filter(r => r['Ejercicio'] === v).length);

        const ctxCombo = document.getElementById('grafico-combo');
        if (chartCombo) chartCombo.destroy();
        chartCombo = new Chart(ctxCombo, {
            type: 'bar',
            data: {
                labels: vigenciasOrdenadas,
                datasets: [
                    {
                        label: 'Valor Bruto',
                        data: valoresVigencia,
                        backgroundColor: '#667eea',
                        borderRadius: 6,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Nº Transacciones',
                        data: countVigencia,
                        type: 'line',
                        borderColor: '#f5576c',
                        backgroundColor: 'rgba(245, 87, 108, 0.1)',
                        borderWidth: 3,
                        fill: true,
                        tension: 0.4,
                        yAxisID: 'y1',
                        pointRadius: 5,
                        pointBackgroundColor: '#f5576c'
                    }
                ]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { display: true, labels: { font: { size: 12, weight: 'bold' } } },
                    title: { display: true, text: 'Pagos por Vigencia', font: { size: 14, weight: 'bold' } }
                },
                scales: {
                    y: { ticks: { callback: v => '$' + (v / 1000000).toFixed(0) + 'M' } },
                    y1: { position: 'right', grid: { drawOnChartArea: false } }
                }
            }
        });

        const topProveedores = Object.entries(porProveedor).sort((a, b) => b[1] - a[1]).slice(0, 10);
        const nombresTop = topProveedores.map(p => p[0].substring(0, 25));
        const valoresTop = topProveedores.map(p => p[1]);
        const colores = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe', '#43e97b', '#fa7231', '#ff6348', '#ee5a6f'];

        const ctxProveedores = document.getElementById('grafico-proveedores');
        if (chartProveedores) chartProveedores.destroy();
        chartProveedores = new Chart(ctxProveedores, {
            type: 'bar',
            data: {
                labels: nombresTop,
                datasets: [{
                    label: 'Valor Pagado',
                    data: valoresTop,
                    backgroundColor: colores
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                plugins: {
                    title: { display: true, text: 'Top 10 Proveedores por Monto', font: { size: 14, weight: 'bold' } }
                },
                scales: { x: { ticks: { callback: v => '$' + (v / 1000000).toFixed(0) + 'M' } } }
            }
        });
    }

    inputHistorico.addEventListener('change', enableBtn);
    btnLimpiar.addEventListener('click', () => {
        filtroContratista.value = '';
        filtroReferencia.value = '';
        filtroFechaDesde.value = '';
        filtroFechaHasta.value = '';
        filtroVigencia.value = '';
        aplicarFiltros();
    });

    btnGraficas.addEventListener('click', () => { contenedorGraficas.style.display = 'block'; contenedorTabla.style.display = 'none'; });
    btnTabla.addEventListener('click', () => { contenedorGraficas.style.display = 'none'; contenedorTabla.style.display = 'block'; });
    
    btnDescargar.addEventListener('click', () => {
        if (!datosFiltrados.length) return;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(datosFiltrados);
        XLSX.utils.book_append_sheet(wb, ws, 'Historico');
        XLSX.writeFile(wb, `Historico_${new Date().toISOString().split('T')[0]}.xlsx`);
    });

    filtroContratista.addEventListener('keyup', aplicarFiltros);
    filtroReferencia.addEventListener('keyup', aplicarFiltros);
    filtroFechaDesde.addEventListener('change', aplicarFiltros);
    filtroFechaHasta.addEventListener('change', aplicarFiltros);
    filtroVigencia.addEventListener('change', aplicarFiltros);

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        alertaDiv.style.display = 'none';
        estadoDiv.style.display = 'block';
        resultadosDiv.style.display = 'none';

        try {
            if (!window.XLSX) throw new Error('XLSX no cargado');
            const archivo = inputHistorico.files[0];
            if (!archivo) throw new Error('Selecciona un archivo');

            const buf = await archivo.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
            datosCompletos = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

            if (!datosCompletos.length) throw new Error('El archivo está vacío');

            construirResumenContratistas();
            datosFiltrados = [...datosCompletos];

            // Llenar select de vigencias
            const vigencias = [...new Set(datosCompletos.map(r => r['Ejercicio']))].sort();
            vigencias.forEach(v => {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = `📅 ${v}`;
                filtroVigencia.appendChild(opt);
            });

            mostrarMetricas(datosFiltrados);
            mostrarGraficos();
            mostrarTabla();

            resultadosDiv.style.display = 'block';
            estadoDiv.style.display = 'none';
            contenedorGraficas.style.display = 'block';
            contenedorTabla.style.display = 'none';

        } catch (err) {
            mostrarAlerta(`❌ ${err.message}`);
            estadoDiv.style.display = 'none';
        }
    });

    enableBtn();
});
