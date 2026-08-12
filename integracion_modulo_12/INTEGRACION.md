# Integración Módulo 12 · CDP-CRP de OXP (todo por Codespaces)

Cableado a tu convención real: `data-vista` + `MODULOS` en `main.js`, blueprint
con `url_prefix` en `__init__.py`, JS en `static/js/`. El parsing pesado del
reporte se hace en el navegador con SheetJS (ya cargado global) y solo viajan
~886 filas livianas al backend → sin OOM en Render.

## 1) Archivos NUEVOS (copiar tal cual)

| Origen (este paquete) | Destino en tu repo |
|---|---|
| `backend/app/routes/oxp_routes.py` | `backend/app/routes/oxp_routes.py` |
| `backend/app/services/oxp_service.py` | `backend/app/services/oxp_service.py` |
| `backend/app/plantillas/Formato_CRP_de_OXP.xlsx` | `backend/app/plantillas/Formato_CRP_de_OXP.xlsx` |
| `frontend/static/js/modulo-12-oxp.js` | `frontend/static/js/modulo-12-oxp.js` |

```bash
mkdir -p backend/app/plantillas
cp integracion_modulo_12/backend/app/routes/oxp_routes.py      backend/app/routes/
cp integracion_modulo_12/backend/app/services/oxp_service.py   backend/app/services/
cp integracion_modulo_12/backend/app/plantillas/Formato_CRP_de_OXP.xlsx backend/app/plantillas/
cp integracion_modulo_12/frontend/static/js/modulo-12-oxp.js   frontend/static/js/
```

## 2) Parchear los 3 archivos existentes (automatico, con respaldo .bak)

```bash
cp integracion_modulo_12/aplicar_integracion_modulo_12.py .
python aplicar_integracion_modulo_12.py \
  --index frontend/index.html \
  --main  frontend/static/js/main.js \
  --init  backend/app/__init__.py
```

Inserta, SIN tocar nada mas:
- index.html: boton nav-item (Modulo 12), la <section id="vista-oxp"> y el <script>.
- main.js: la entrada oxp: en MODULOS.
- __init__.py: import oxp_bp + register_blueprint(..., url_prefix="/api/oxp").

Es idempotente (si ya esta, lo omite) y deja .bak de cada archivo.

## 3) Verificar

```bash
node --check frontend/static/js/main.js
node --check frontend/static/js/modulo-12-oxp.js
python -c "import ast; ast.parse(open('backend/app/__init__.py').read())"
```

## 4) Probar backend local

```bash
cd backend && python -m flask run
curl -s -X POST http://localhost:5000/api/oxp/crp \
  -H "Content-Type: application/json" \
  -d '{"filas":[{"importe":100000,"objeto":"X","no_compromiso":"713-2023"}]}' \
  --output test_crp.xlsx && echo "ok test_crp.xlsx"
cd ..
```

## 5) Desplegar (tu flujo)

```bash
rm -rf docs && cp -r frontend docs
git add backend/app/routes/oxp_routes.py backend/app/services/oxp_service.py \
        backend/app/plantillas/Formato_CRP_de_OXP.xlsx \
        backend/app/__init__.py \
        frontend/static/js/modulo-12-oxp.js frontend/index.html frontend/static/js/main.js \
        docs/static/js/modulo-12-oxp.js docs/index.html docs/static/js/main.js
git commit -m "Modulo 12 CDP-CRP de OXP: backend openpyxl + frontend SheetJS"
git push
```

Ctrl+Shift+R (o incognito) tras desplegar, porque el navegador cachea el JS.

## Notas

- Endpoints nuevos: POST /api/oxp/crp y POST /api/oxp/cdp. Devuelven el .xlsx
  (cabecera X-OXP-Registros con el conteo).
- El CDP OXP deja vacios los campos # PENDIENTE (I, K, M, P, U, V) hasta definir
  su regla; el CRP OXP esta completo.
- SOCIEDAD_FIJA = 1001 en oxp_service.py — verificar por alcaldia.
