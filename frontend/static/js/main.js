// ============================================
// MAIN.JS - GESTIÓN DE VISTAS Y NAVEGACIÓN
// ============================================

const CONFIG_VISTAS = {
  'extraccion': { eyebrow: 'Módulo 01', titulo: 'Extracción de causaciones' },
  'plantilla': { eyebrow: 'Módulo 02', titulo: 'Plantilla de pagos' },
  'pac': { eyebrow: 'Módulo 03', titulo: 'Cuadre PAC' },
  'estadocuenta': { eyebrow: 'Módulo 04', titulo: 'Estado de cuenta' },
  'historico': { eyebrow: 'Módulo 05', titulo: 'Histórico de Pagos' },
  'presupuestal': { eyebrow: 'Módulo 06', titulo: 'Análisis Presupuestal' },
  'contratacion': { eyebrow: 'Módulo 07', titulo: 'Contratación' },
  'ciberseguridad': { eyebrow: 'Módulo 08', titulo: 'Ciberseguridad' },
  'licencias': { eyebrow: 'Módulo 09', titulo: 'Licencias' }
};

function cambiarVista(nombreVista) {
  const config = CONFIG_VISTAS[nombreVista];
  if (!config) { console.error(`Vista no encontrada: ${nombreVista}`); return; }
  document.querySelectorAll('.vista').forEach(el => el.setAttribute('hidden', ''));
  const vistaElement = document.getElementById(`vista-${nombreVista}`);
  if (vistaElement) vistaElement.removeAttribute('hidden');
  document.getElementById('eyebrow-modulo').textContent = config.eyebrow;
  document.getElementById('titulo-modulo').textContent = config.titulo;
  document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('activo'));
  document.querySelector(`[data-vista="${nombreVista}"]`)?.classList.add('activo');
  window.scrollTo(0, 0);
}

document.querySelectorAll('.nav-item').forEach(btn => {
  btn.addEventListener('click', () => cambiarVista(btn.getAttribute('data-vista')));
});

document.addEventListener('DOMContentLoaded', () => {
  cambiarVista('extraccion');
  configurarDropzones();
});

function configurarDropzones() {
  document.querySelectorAll('.dropzone').forEach(zone => {
    const input = zone.querySelector('input[type="file"]');
    if (!input) return;
    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.style.background = 'rgba(102, 126, 234, 0.1)';
      zone.style.borderColor = '#667eea';
    });
    zone.addEventListener('dragleave', () => {
      zone.style.background = '';
      zone.style.borderColor = '';
    });
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.style.background = '';
      zone.style.borderColor = '';
      if (e.dataTransfer.files.length > 0) {
        input.files = e.dataTransfer.files;
        input.dispatchEvent(new Event('change', { bubbles: true }));
        mostrarArchivos(zone, input);
      }
    });
    zone.addEventListener('click', () => input.click());
    input.addEventListener('change', () => mostrarArchivos(zone, input));
  });
}

function mostrarArchivos(zone, input) {
  const listaElement = zone.nextElementSibling;
  if (!listaElement || !listaElement.classList.contains('lista-archivos')) return;
  listaElement.innerHTML = '';
  if (input.files.length === 0) return;
  const ul = document.createElement('ul');
  ul.style.cssText = 'list-style:none; padding:10px 0; margin:10px 0; border-top:1px solid #e8edf2; font-size:12px;';
  Array.from(input.files).forEach(file => {
    const li = document.createElement('li');
    li.style.cssText = 'padding:8px 0; color:#666; display:flex; align-items:center;';
    let icon = '📄';
    if (file.type.includes('pdf')) icon = '📕';
    else if (file.type.includes('spreadsheet')) icon = '📊';
    li.innerHTML = `${icon} <span style="margin-left:8px;">${file.name}</span>`;
    ul.appendChild(li);
  });
  listaElement.appendChild(ul);
}

window.descargarArchivo = function(blob, nombre) {
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = nombre;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
};

window.mostrarAlerta = function(elemento, mensaje, tipo = 'error') {
  if (!elemento) return;
  elemento.innerHTML = mensaje;
  elemento.style.display = 'block';
  if (tipo === 'error') {
    elemento.classList.add('alerta-error');
    elemento.classList.remove('alerta-aviso');
  } else {
    elemento.classList.add('alerta-aviso');
    elemento.classList.remove('alerta-error');
  }
};

window.formatearMoneda = function(valor) {
  return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', minimumFractionDigits: 0 }).format(valor);
};

window.formatearFecha = function(fecha) {
  if (!fecha) return '';
  const d = new Date(fecha);
  return d.toLocaleDateString('es-CO');
};

console.log('✅ Main.js cargado correctamente');
console.log('📍 Vistas disponibles:', Object.keys(CONFIG_VISTAS));
