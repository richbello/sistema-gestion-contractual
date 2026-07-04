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
    const filtroEjercicio = document.getElementById('filtro-ejercicio');
    const btnLimpiar = document.getElementById('btn-limpiar-filtros');
    const btnGraficas = document.getElementById('btn-vista-graficas');
    const btnTabla = document.getElementById('btn-vista-tabla');
    const btnDescargar = document.getElementById('btn-descargar-filtrado');
    const contenedorGraficas = document.getElementById('contenedor-graficas');
    const contenedorTabla = document.getElementById('contenedor-tabla');

    let datosCompletos = [];
    let datosFiltrados = [];
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

    function mostrarMetricas(datos) {
        const totalPagado = datos.reduce((sum, r) => sum + (Number(r['Valor Bruto']) || 0), 0);
        const totalTransacciones = datos.length;
        const proveedoresUnicos = new Set(datos.map(r => r['Nombre'])).size;

        metricasDiv.innerHTML = `
            <div class="metrica-card" style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${formatCurrency(totalPagado)}</div>
                <div style="font-size:12px; margin-top:5px;">Valor Bruto Total</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${totalTransacciones.toLocaleString()}</div>
                <div style="font-size:12px; margin-top:5px;">Registros SAP</div>
            </div>
            <div class="metrica-card" style="background:linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color:white; padding:20px; border-radius:8px; text-align:center;">
                <div style="font-size:20px; font-weight:bold;">${proveedoresUnicos.toLocaleString()}</div>
                <div style="font-size:12px; margin-top:5px;">Proveedores Únicos</div>
            </div>
        `;
    }

    function aplicarFiltros() {
        const busqContratista = filtroContratista.value.toLowerCase();
        const busqReferencia = filtroReferencia.value.toLowerCase();
        const ejercicio = filtroEjercicio.value;

        datosFiltrados = datosCompletos.filter(r => {
            const nombre = String(r['Nombre'] || '').toLowerCase();
            const nit = String(r['Nº identificación'] || '').toLowerCase();
            const referencia = String(r['Referencia'] || '').toLowerCase();
            const ej = String(r['Ejercicio'] || '');

            const cumpleContratista = !busqContratista || nombre.includes(busqContratista) || nit.includes(busqContratista);
            const cumpleReferencia = !busqReferencia || referencia.includes(busqReferencia);
            const cumpleEjercicio = !ejercicio || ej === ejercicio;

            return cumpleContratista && cumpleReferencia && cumpleEjercicio;
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

        const porEjercicio = {}, porProveedor = {};
        const ejercicios = new Set();
        
        datosFiltrados.forEach(r => {
            const ej = r['Ejercicio'] || 'Sin año';
            ejercicios.add(ej);
            porEjercicio[ej] = (porEjercicio[ej] || 0) + (Number(r['Valor Bruto']) || 0);

            const proveedor = r['Nombre'] || 'Sin nombre';
            porProveedor[proveedor] = (porProveedor[proveedor] || 0) + (Number(r['Valor Bruto']) || 0);
        });

        // Actualizar select de ejercicio
        const ejerciciosSorted = Array.from(ejercicios).sort();
        const selectEjercicio = document.getElementById('filtro-ejercicio');
        const opcionesOriginales = selectEjercicio.innerHTML;
        selectEjercicio.innerHTML = opcionesOriginales + ejerciciosSorted.map(e => `<option value="${e}">${e}</option>`).join('');

        // Gráfica COMBO (barras + línea)
        const ctxCombo = document.getElementById('grafico-combo');
        const ejerciciosOrdenados = ejerciciosSorted;
        const valoresEjercicio = ejerciciosOrdenados.map(e => porEjercicio[e] || 0);
        const countEjercicio = ejerciciosOrdenados.map(e => 
            datosFiltrados.filter(r => r['Ejercicio'] === e).length
        );

        if (chartCombo) chartCombo.destroy();
        chartCombo = new Chart(ctxCombo, {
            type: 'bar',
            data: {
                labels: ejerciciosOrdenados,
                datasets: [
                    {
                        label: 'Valor Bruto por Ejercicio Fiscal',
                        data: valoresEjercicio,
                        backgroundColor: '#667eea',
                        borderColor: '#667eea',
                        borderWidth: 2,
                        borderRadius: 6,
                        yAxisID: 'y'
                    },
                    {
                        label: 'Nº Registros',
                        data: countEjercicio,
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
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: true, labels: { font: { size: 12, weight: 'bold' }, color: '#333' } },
                    title: { display: true, text: 'Valor Bruto por Ejercicio Fiscal', font: { size: 14, weight: 'bold' }, color: '#333' }
                },
                scales: {
                    y: {
                        type: 'linear',
                        position: 'left',
                        ticks: { callback: v => '$' + (v / 1000000).toFixed(0) + 'M' }
                    },
                    y1: {
                        type: 'linear',
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        ticks: { callback: v => v.toLocaleString() }
                    }
                }
            }
        });

        // Gráfica TOP 10 PROVEEDORES (horizontal)
        const ctxProveedores = document.getElementById('grafico-proveedores');
        const topProveedores = Object.entries(porProveedor).sort((a, b) => b[1] - a[1]).slice(0, 10);
        const nombresTop = topProveedores.map(p => p[0].substring(0, 30));
        const valoresTop = topProveedores.map(p => p[1]);

        const colores = ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe', '#00f2fe', '#43e97b', '#fa7231', '#ff6348', '#ee5a6f'];

        if (chartProveedores) chartProveedores.destroy();
        chartProveedores = new Chart(ctxProveedores, {
            type: 'bar',
            data: {
                labels: nombresTop,
                datasets: [{
                    label: 'Valor Pagado',
                    data: valoresTop,
                    backgroundColor: colores,
                    borderColor: colores,
                    borderWidth: 2
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: true },
                    title: { display: true, text: 'Top 10 Proveedores por Valor Bruto', font: { size: 14, weight: 'bold' }, color: '#333' }
                },
                scales: {
                    x: { ticks: { callback: v => '$' + (v / 1000000).toFixed(0) + 'M' } }
                }
            }
        });
    }

    inputHistorico.addEventListener('change', enableBtn);
    btnLimpiar.addEventListener('click', () => { filtroContratista.value = ''; filtroReferencia.value = ''; filtroEjercicio.value = ''; aplicarFiltros(); });
    btnGraficas.addEventListener('click', () => { contenedorGraficas.style.display = 'block'; contenedorTabla.style.display = 'none'; btnGraficas.style.opacity = '1'; btnTabla.style.opacity = '0.6'; });
    btnTabla.addEventListener('click', () => { contenedorGraficas.style.display = 'none'; contenedorTabla.style.display = 'block'; btnTabla.style.opacity = '1'; btnGraficas.style.opacity = '0.6'; });
    
    btnDescargar.addEventListener('click', () => {
        if (!datosFiltrados.length) return;
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(datosFiltrados);
        XLSX.utils.book_append_sheet(wb, ws, 'Historico');
        XLSX.writeFile(wb, `Historico_Pagos_Filtrado_${new Date().toISOString().split('T')[0]}.xlsx`);
    });

    filtroContratista.addEventListener('keyup', aplicarFiltros);
    filtroReferencia.addEventListener('keyup', aplicarFiltros);
    filtroEjercicio.addEventListener('change', aplicarFiltros);

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

            datosFiltrados = [...datosCompletos];
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
