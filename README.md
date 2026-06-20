# Sistema de Gestión Contractual y Financiera

Plataforma interna para apoyar el proceso de causación, pago y seguimiento presupuestal de contratos de prestación de servicios. Reemplaza el flujo manual de notebooks de Google Colab por una aplicación web con backend (Flask) y frontend (HTML/CSS/JS) desplegable de forma independiente.

## Módulos

| Módulo | Función |
|---|---|
| **01 · Extracción de causaciones** | Lee un lote de PDFs de causación, extrae contrato, contratista, valores y retenciones, y los cruza con la base general de contratistas (BASEGEN) para completar datos bancarios. |
| **02 · Plantilla de pagos** | Convierte el Excel de extracción en la plantilla de carga del sistema de pagos (registros contables C / P40 / P31), incluyendo retención por honorarios cuando aplica. |
| **03 · Cuadre PAC** | Cruza la plantilla de pagos contra el reporte PAC del mes para verificar disponibilidad presupuestal por rubro y fondo. |
| **04 · Estado de cuenta** | Genera el estado de cuenta oficial de un contrato específico a partir del histórico de pagos. |

## Estructura del proyecto

```
.
├── backend/
│   ├── app/
│   │   ├── __init__.py          # Fábrica de la app Flask
│   │   ├── routes/              # Endpoints de la API (uno por módulo)
│   │   └── services/            # Lógica de negocio (procesamiento de Excel/PDF)
│   ├── uploads/                 # Archivos recibidos (temporal, no versionado)
│   ├── outputs/                 # Archivos generados (temporal, no versionado)
│   ├── requirements.txt
│   └── run.py                   # Punto de entrada
└── frontend/
    ├── templates/index.html     # Interfaz (SPA de una sola página)
    └── static/
        ├── css/estilos.css
        └── js/main.js
```

## Requisitos

- Python 3.10 o superior

## Instalación y ejecución local

```bash
cd backend
python -m venv .venv
source .venv/bin/activate        # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

La aplicación queda disponible en `http://localhost:5000`. El backend sirve tanto la API (`/api/...`) como el frontend.

## Notas técnicas

- Cada módulo es independiente: recibe archivos vía `multipart/form-data`, procesa con `pandas` / `pdfplumber` / `openpyxl`, y devuelve un Excel descargable más un resumen en JSON.
- Los archivos subidos y generados se guardan con nombres únicos en `backend/uploads` y `backend/outputs`; estas carpetas no se versionan en git (ver `.gitignore`).
- La lógica de cada módulo está aislada en `backend/app/services/`, separada del manejo HTTP en `backend/app/routes/`, para facilitar pruebas unitarias futuras.

## Próximos pasos sugeridos

- Agregar autenticación si la plataforma se expone fuera de una red interna.
- Persistir un historial de procesos (base de datos) en lugar de archivos sueltos.
- Agregar pruebas automatizadas para los servicios de `backend/app/services/`.
