#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_integracion_modulo_12.py
================================
Inserta el Módulo 12 (CDP-CRP de OXP · submenú CRP OXP) en los 3 archivos del
SPA, SIN tocar nada más. Idempotente y con respaldo .bak.

Uso (desde la raíz del repo, en Codespaces):
    python aplicar_integracion_modulo_12.py \
        --index frontend/index.html \
        --main  frontend/static/js/main.js \
        --init  backend/app/__init__.py
"""

import argparse
import os
import shutil
import sys

NAV_ITEM = '''        <button class="nav-item" data-vista="oxp">
          <span class="nav-numero">12</span>
          <svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 4h16v16H4z"/><path d="M4 9h16M9 9v11"/></svg>
          CDP-CRP de OXP
        </button>

'''

SECCION = '''    <!-- ===================== VISTA 12: CDP-CRP de OXP ===================== -->
    <section class="vista" id="vista-oxp" hidden>
      <style>
        .oxp-tabs { display:flex; gap:8px; margin-bottom:20px; border-bottom:1px solid #e2e8f0; }
        .oxp-tab { background:none; border:0; padding:10px 18px; cursor:pointer; font-size:.95rem; font-weight:600; color:#64748b; border-bottom:2px solid transparent; margin-bottom:-1px; }
        .oxp-tab:hover { color:#334155; }
        .oxp-tab.activo { color:#4f46e5; border-bottom-color:#4f46e5; }
        .oxp-tab:disabled { color:#cbd5e1; cursor:not-allowed; }
        .oxp-panel { display:none; } .oxp-panel.activo { display:block; }
        #resumen-oxp { margin-top:10px; color:#475569; font-size:.9rem; }
      </style>

      <div class="oxp-tabs">
        <button type="button" class="oxp-tab activo" data-panel="oxp-crp">CRP OXP</button>
        <button type="button" class="oxp-tab" data-panel="oxp-cdp">CDP OXP</button>
      </div>

      <div id="oxp-crp" class="oxp-panel activo">
        <form class="panel" id="form-oxp" onsubmit="return false;">
          <div class="panel-acento-superior"></div>
          <div class="panel-cuerpo">
            <div class="grid-form una-col">
              <div class="campo">
                <label>Reporte CRP</label>
                <div class="ayuda">Export SAP/BOGDATA del CRP (cierre de vigencia). Se procesa en tu navegador y se filtran las obligaciones con saldo pendiente (Com.Sin.Aut.Giro &gt; 0). Un CRP con varias posiciones se agrupa automáticamente.</div>
                <div class="dropzone" id="drop-oxp">
                  <input type="file" id="input-oxp" accept=".xlsx,.xls" required>
                  <svg class="dropzone-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>
                  <div class="dropzone-texto"><strong>Arrastra el archivo</strong> o haz clic para seleccionarlo</div>
                  <div class="dropzone-detalle">.xlsx · un solo archivo</div>
                </div>
                <div class="lista-archivos" id="lista-oxp"></div>
                <div id="resumen-oxp"></div>
              </div>
            </div>

            <div class="barra-acciones">
              <button type="button" class="btn btn-primario" id="btn-oxp-crp" disabled>
                <svg class="icono-btn" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
                Generar CRP OXP
              </button>
              <div class="estado-proceso" id="estado-oxp"><span class="spinner"></span> Generando formato…</div>
            </div>

            <div class="alerta alerta-error" id="alerta-oxp"></div>

            <div class="resultados" id="resultados-oxp">
              <div class="resultados-cabecera">
                <svg class="sello-ok" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="10"/><path d="M8 12l3 3 5-6"/></svg>
                <h3>CRP OXP generado</h3>
              </div>
              <div class="metricas-grid" id="metricas-oxp"></div>
            </div>
          </div>
        </form>
      </div>

      <div id="oxp-cdp" class="oxp-panel">
        <form class="panel" id="form-oxp-cdp" onsubmit="return false;">
          <div class="panel-acento-superior"></div>
          <div class="panel-cuerpo">
            <div class="grid-form una-col">
              <div class="campo">
                <label>Reporte CRP</label>
                <div class="ayuda">Mismo archivo del CRP OXP. Los rubros se convierten automáticamente según la tabla de conversión definida (ej. O230689 → O230690, O219001 → O219002).</div>
                <div class="dropzone" id="drop-oxp-cdp">
                  <input type="file" id="input-oxp-cdp" accept=".xlsx,.xls" required>
                  <svg class="dropzone-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>
                  <div class="dropzone-texto"><strong>Arrastra el archivo</strong> o haz clic para seleccionarlo</div>
                  <div class="dropzone-detalle">.xlsx · un solo archivo</div>
                </div>
                <div class="lista-archivos" id="lista-oxp-cdp"></div>
                <div id="resumen-oxp-cdp"></div>
              </div>
            </div>

            <div class="barra-acciones">
              <button type="button" class="btn btn-primario" id="btn-oxp-cdp" disabled>
                <svg class="icono-btn" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
                Generar CDP OXP
              </button>
              <div class="estado-proceso" id="estado-oxp-cdp"><span class="spinner"></span> Generando formato…</div>
            </div>

            <div class="alerta alerta-error" id="alerta-oxp-cdp"></div>

            <div class="resultados" id="resultados-oxp-cdp">
              <div class="resultados-cabecera">
                <svg class="sello-ok" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="10"/><path d="M8 12l3 3 5-6"/></svg>
                <h3>CDP OXP generado</h3>
              </div>
              <div class="metricas-grid" id="metricas-oxp-cdp"></div>
            </div>
          </div>
        </form>
      </div>
    </section>

'''

SCRIPT_TAG = '  <script src="static/js/modulo-12-oxp.js"></script>\n'

MODULOS_ENTRY = '  oxp:           { eyebrow: "Módulo 12", titulo: "CDP-CRP de OXP" },\n'

INIT_IMPORT = "    from .routes.oxp_routes import oxp_bp\n"
INIT_REGISTER = '    app.register_blueprint(oxp_bp, url_prefix="/api/oxp")\n'


def _leer(ruta):
    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()


def _escribir(ruta, contenido):
    shutil.copy2(ruta, ruta + ".bak")
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(contenido)


def _insertar_antes(texto, ancla, fragmento, nombre):
    if texto.count(ancla) != 1:
        raise RuntimeError("[%s] ancla no encontrada exactamente 1 vez (%d)."
                           % (nombre, texto.count(ancla)))
    return texto.replace(ancla, fragmento + ancla, 1)


def _insertar_despues(texto, ancla, fragmento, nombre):
    if texto.count(ancla) != 1:
        raise RuntimeError("[%s] ancla no encontrada exactamente 1 vez (%d)."
                           % (nombre, texto.count(ancla)))
    return texto.replace(ancla, ancla + fragmento, 1)


def patch_index(ruta):
    txt = _leer(ruta)
    if 'data-vista="oxp"' in txt:
        print("  = index.html ya tiene el Módulo 12; se omite.")
        return
    ancla_nav = '        <a class="nav-item" href="cuentas_bancarias.html"'
    txt = _insertar_antes(txt, ancla_nav, NAV_ITEM, "index/nav")
    ancla_main = "  </main>\n</div>"
    txt = _insertar_antes(txt, ancla_main, SECCION, "index/seccion")
    ancla_script = '  <script src="static/js/estadocuenta-masivo.js"></script>\n'
    txt = _insertar_despues(txt, ancla_script, SCRIPT_TAG, "index/script")
    _escribir(ruta, txt)
    print("  ✔ index.html: sidebar + sección (pestañas CRP/CDP) + script.")


def patch_main(ruta):
    txt = _leer(ruta)
    if 'oxp:' in txt and 'Módulo 12' in txt:
        print("  = main.js ya tiene el Módulo 12; se omite.")
        return
    ancla = '  validacion:    { eyebrow: "Módulo 11", titulo: "Validación de plantilla" },\n'
    txt = _insertar_despues(txt, ancla, MODULOS_ENTRY, "main/MODULOS")
    _escribir(ruta, txt)
    print("  ✔ main.js: entrada MODULOS.oxp insertada.")


def patch_init(ruta):
    txt = _leer(ruta)
    if "oxp_routes" in txt:
        print("  = __init__.py ya tiene el Módulo 12; se omite.")
        return
    ancla_imp = "    from .routes.validacion_plantilla import validacion_bp\n"
    txt = _insertar_despues(txt, ancla_imp, INIT_IMPORT, "init/import")
    ancla_reg = '    app.register_blueprint(validacion_bp, url_prefix="/api/validacion-plantilla")\n'
    txt = _insertar_despues(txt, ancla_reg, INIT_REGISTER, "init/register")
    _escribir(ruta, txt)
    print("  ✔ __init__.py: import + register_blueprint insertados.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default="frontend/index.html")
    ap.add_argument("--main",  default="frontend/static/js/main.js")
    ap.add_argument("--init",  default="backend/app/__init__.py")
    args = ap.parse_args()

    errores = 0
    for etiqueta, ruta, fn in [
        ("index.html", args.index, patch_index),
        ("main.js",    args.main,  patch_main),
        ("__init__.py",args.init,  patch_init),
    ]:
        if not os.path.exists(ruta):
            print("  ✗ No existe: %s" % ruta); errores += 1; continue
        try:
            fn(ruta)
        except RuntimeError as e:
            print("  ✗ %s" % e); errores += 1

    if errores:
        print("\nTerminado con %d problema(s)." % errores); sys.exit(1)
    print("\n✅ Integración aplicada. Respaldos .bak junto a cada archivo.")


if __name__ == "__main__":
    main()
