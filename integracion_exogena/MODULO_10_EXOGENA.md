# Módulo 10 · Reportes Exógena

Monta el módulo en el sidebar con dos pestañas (Distrital Bogotá / DIAN Nacional).
El generador se conecta cuando se defina la ficha técnica del formato distrital.

## Aplicar (Codespaces, raíz del repo)

```bash
python aplicar_modulo_10_exogena.py \
  --index frontend/index.html \
  --main  frontend/static/js/main.js \
  --jsdir frontend/static/js

# verificar
node --check frontend/static/js/main.js
node --check frontend/static/js/modulo-10-exogena.js

# desplegar
rm -rf docs && cp -r frontend docs
git add frontend/index.html frontend/static/js/main.js frontend/static/js/modulo-10-exogena.js \
        docs/index.html docs/static/js/main.js docs/static/js/modulo-10-exogena.js
git commit -m "Modulo 10 Reportes Exogena (andamiaje: pestanas Distrital/Nacional)"
git push
```

Ctrl+Shift+R tras desplegar.

## Estado
- Distrital Bogotá: lee el insumo con SheetJS y muestra filas/columnas. Botón
  "Generar archivo de cargue" DESHABILITADO hasta conectar la ficha técnica
  (Res. DDI-024115, entrega vía PIDO/Storm .STR).
- DIAN Nacional: "próximamente" (formato 1001 v10 XML).

Idempotente. Deja .bak de index.html y main.js.
