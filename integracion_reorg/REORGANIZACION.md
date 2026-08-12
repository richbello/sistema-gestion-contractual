# Reorganización de módulos + renombre del software

Cambia SOLO la presentación de la barra lateral. Los `data-vista` e IDs de
sección se conservan, así que la lógica de todos los módulos sigue intacta.

## Qué hace
- Software: "Gestión Contractual y Financiera" → "Gestión Presupuestal y Financiera"
- Módulo 1 "Trámite de Pagos" con submenú "Personas Naturales":
  Extracción · Plantilla de pagos · Cuadre PAC · Validación de plantilla · Cuentas bancarias
- Renumera: 2 Estados de Cuenta · 3 Históricos de Pagos · 4 Análisis Presupuestal ·
  5 Contratación · 6 Planeación · 7 Ciberseguridad · 8 Licencias · 9 CDP-CRP de OXP
- Pie: "Derechos Reservados · XBOG SMART 2026"
- CSS del submenú (en <head>) + toggle expandir/colapsar en main.js

## Cómo aplicarlo (Codespaces, desde la raíz del repo)

```bash
# (arrastra el zip y descomprímelo, o copia solo el .py)
python aplicar_reorganizacion.py \
  --index frontend/index.html \
  --main  frontend/static/js/main.js

# verificar
node --check frontend/static/js/main.js

# sincronizar a docs y desplegar
rm -rf docs && cp -r frontend docs
git add frontend/index.html frontend/static/js/main.js docs/index.html docs/static/js/main.js
git commit -m "Reorganizacion sidebar: Tramite de Pagos + renumeracion + XBOG SMART"
git push
```

Ctrl+Shift+R (o incógnito) tras desplegar, por el caché del JS/CSS.

## Notas
- Idempotente (si ya está aplicado, lo omite) y deja `.bak` de cada archivo.
- El CSS del submenú se inyecta inline en <head> (id="estilos-submenu-modulos").
  Si prefieres, luego lo mueves a static/css/estilos.css.
- El submenú "Personas Naturales" arranca ABIERTO por defecto (clase "abierto").
  Se colapsa/expande al hacer clic en "Trámite de Pagos".
