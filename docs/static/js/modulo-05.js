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
    const tarjetaContratista = document.getElementById('tarjeta-contratista');
    
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
    ['dragover', 'dragenter'].forEach(ev => dropzone.addEventListener(ev, e => { e.preventDefault(); dropzone.style.background = '#1a1f2e'; }));
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
        const ejercicios = [...new Set(datos.map(r => r['Ejercicio']))].sort().join(', ');

        metricasDiv.innerHTML = `
            <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white; padding:25px; border-radius:8px;">
                <div style="font-size:28px; font-weight:bold;">${formatCurrency(totalPagado)}</div>
                <div style="font-size:12px; margin-top:8px; opacity:0.9;">Valor Bruto Total</div>
            </div>
            <div style="background:linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color:white; padding:25px; border-radius:8px;">
                <div style="font-size:28px; font-weight:bold;">${totalTransacciones.toLocaleString()}</div>
                <div style="font-size:12px; margin-top:8px; opacity:0.9;">Registros SAP</div>
            </div>
            <div style="background:linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); color:white; padding:25px; border-radius:8px;">
                <div style="font-size:28px; font-weight:bold;">${proveedoresUnicos}</div>
                <div style="font-size:12px; margin-top:8px; opacity:0.9;">Proveedores únicos</div>
            </div>
            <div style="background:linear-gradient(135deg, #43e97b 0%, #38f9d7 100%); color:white; padding:25px; border-radius:8px;">
                <div style="font-size:28px; font-weight:bold;">${ejercicios}</div>
                <div style="font-size:12px; margin-top:8px; opacity:0.9;">Mayor ejercicio</div>
            </div>
        `;
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
                    nit: r['Nº identificación'] || '',
                    asignaciones: new Set(),
                    vigencias: new Set()
                };
            }
            resumenContratistas[nombre].total += Number(r['Valor Bruto']) || 0;
            resumenContratistas[nombre].transacciones += 1;
            resumenContratistas[nombre].asignaciones.add(r['Referencia'] || 'Sin ref');
            resumenContratistas[nombre].vigencias.add(r['Ejercicio'] || 'Sin año');
            const fecha = r['Fecha de pago'];
            if (fecha && (!resumenContratistas[nombre].ultimaFecha || fecha > resumenContratistas[nombre].ultimaFecha)) {
                resumenContratistas[nombre].ultimaFecha = fecha;
            }
        });
    }

    function mostrarTarjetaContratista(nombreContratista) {
        const resumen = resumenContratistas[nombreContratista];
        if (!resumen) {
            tarjetaContratista.style.display = 'none';
            return;
        }

        const asignacionesArr = Array.from(resumen.asignaciones).slice(0, 3);
        const vigenciasArr = Array.from(resumen.vigencias).sort();

        document.getElementById('contratista-nombre').textContent = nombreContratista;
        document.getElementById('contratista-nit').textContent = `NIT: ${resumen.nit}`;

        const resumenHtml = `
            <div style="background:#0f1419; padding:12px; border-radius:6px; text-align:center;">
                <div style="font-size:20px; font-weight:bold; color:#667eea;">${formatCurrency(resumen.total)}</div>
                <div style="font-size:11px; color:#999; margin-top:5px;">Total pagado</div>
            </div>
            <div style="background:#0f1419; padding:12px; border-radius:6px; text-align:center;">
                <div style="font-size:20px; font-weight:bold; color:#f5576c;">${resumen.transacciones}</div>
                <div style="font-size:11px; color:#999; margin-top:5px;">Registros</div>
            </div>
            <div style="background:#0f1419; padding:12px; border-radius:6px; text-align:center;">
                <div style="font-size:20px; font-weight:bold; color:#43e97b;">${formatDate(resumen.ultimaFecha)}</div>
                <div style="font-size:11px; color:#999; margin-top:5px;">Último pago</div>
            </div>
            <div style="background:#0f1419; padding:12px; border-radius:6px; text-align:center;">
                <div style="font-size:20px; font-weight:bold; color:#4facfe;">${vigenciasArr.join(' - ')}</div>
                <div style="font-size:11px; color:#999; margin-top:5px;">Ejercicios activos</div>
            </div>
        `;
        document.getElementById('contratista-resumen').innerHTML = resumenHtml;

        const asignacionesHtml = `<strong>Asignaciones:</strong> ${asignacionesArr.map(a => `<span style="display:inline-block; background:#333; padding:6px 12px; border-radius:4px; margin-right:8px; margin-top:8px; font-size:11px;">${a}</span>`).join('')}`;
        document.getElementById('contratista-asignaciones').innerHTML = asignacionesHtml;

        tarjetaContratista.style.display = 'block';
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

        if (busqContratista && Object.keys(resumenContratistas).some(c => c.toLowerCase().includes(busqContratista))) {
            const contratista = Object.keys(resumenContratistas).find(c => c.toLowerCase().includes(busqContratista));
            mostrarTarjetaContratista(contratista);
        } else {
            tarjetaContratista.style.display = 'none';
        }

        btnTabla.textContent = `📋 Tabla (${datosFiltrados.length})`;
        mostrarGraficos();
        mostrarTabla();
    }

    function mostrarTabla() {
        if (!datosFiltrados.length) {
            contenedorTabla.innerHTML = '<p style="color:#999;">No hay datos para mostrar.</p>';
            return;
        }

        const columnas = ['Nombre', 'Nº identificación', 'Valor Bruto', 'Fecha de pago', 'Referencia', 'Ejercicio'];
        let html = '<table style="width:100%; border-collapse:collapse; font-size:11px; color:#ccc;"><thead><tr style="background:#333; border-bottom:2px solid #667eea;">';

        columnas.forEach(col => {
            html += `<th style="padding:12px; text-align:left; font-weight:bold;">${col}</th>`;
        });
        html += '</tr></thead><tbody>';

        datosFiltrados.forEach((fila, idx) => {
            html += `<tr style="border-bottom:1px solid #333;">`;
            columnas.forEach(col => {
                let valor = fila[col] || '';
                if (col === 'Valor Bruto') valor = formatCurrency(valor);
                if (col === 'Fecha de pago') valor = formatDate(valor);
                html += `<td style="padding:10px;">${valor}</td>`;
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
                    legend: { labels: { color: '#ccc', font: { size: 11 } } },
                    title: { display: true, text: 'Valor Bruto por Ejercicio Fiscal', color: '#ccc', font: { size: 12 } }
                },
                scales: {
                    y: { ticks: { color: '#999', callback: v => '$' + (v / 1000000).toFixed(0) + 'M' }, grid: { color: '#333' } },
                    y1: { position: 'right', ticks: { color: '#999' }, grid: { drawOnChartArea: false } },
                    x: { ticks: { color: '#999' }, grid: { color: '#333' } }
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
                    title: { display: true, text: 'Top 10 Proveedores por Monto', color: '#ccc', font: { size: 12 } },
                    legend: { labels: { color: '#ccc', font: { size: 11 } } }
                },
                scales: { x: { ticks: { color: '#999', callback: v => '$' + (v / 1000000).toFixed(0) + 'M' }, grid: { color: '#333' } }, y: { ticks: { color: '#999' }, grid: { color: '#333' } } }
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
        tarjetaContratista.style.display = 'none';
        aplicarFiltros();
    });

    btnGraficas.addEventListener('click', () => { contenedorGraficas.style.display = 'block'; contenedorTabla.style.display = 'none'; btnGraficas.style.opacity = '1'; btnTabla.style.opacity = '0.6'; });
    btnTabla.addEventListener('click', () => { contenedorGraficas.style.display = 'none'; contenedorTabla.style.display = 'block'; btnTabla.style.opacity = '1'; btnGraficas.style.opacity = '0.6'; });
    
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

            const vigencias = [...new Set(datosCompletos.map(r => r['Ejercicio']))].sort();
            vigencias.forEach(v => {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = v;
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
