#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_modulo_10_exogena.py
============================
Monta el "Módulo 10 · Reportes Exógena" en el SPA:
  - Item en el sidebar (data-vista="exogena")
  - Sección con dos pestañas: Distrital (Bogotá SHD) y DIAN Nacional
  - Entrada en MODULOS + carga del script del módulo
  - Archivo JS del módulo (tabs + carga de insumo con SheetJS, listo para conectar
    el generador cuando se defina la ficha técnica).

Idempotente y con respaldo .bak.

Uso (raíz del repo, en Codespaces):
    python aplicar_modulo_10_exogena.py \
        --index frontend/index.html \
        --main  frontend/static/js/main.js \
        --jsdir frontend/static/js
"""

import argparse
import os
import shutil
import sys

ICON_EXO = ('<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            'stroke-width="1.6"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>'
            '<path d="M14 2v6h6"/><path d="M9 15h6M9 18h4M9 12h2"/></svg>')

NAV_ITEM = ('      <button class="nav-item" data-vista="exogena"><span class="nav-numero">10</span>'
            + ICON_EXO + ' Reportes Exógena</button>\n')

SECCION = '''    <!-- ===================== VISTA 10: Reportes Exógena ===================== -->
    <section class="vista" id="vista-exogena" hidden>
      <style>
        .exo-tabs { display:flex; gap:8px; margin-bottom:20px; border-bottom:1px solid #e2e8f0; }
        .exo-tab { background:none; border:0; padding:10px 18px; cursor:pointer; font-size:.95rem; font-weight:600; color:#64748b; border-bottom:2px solid transparent; margin-bottom:-1px; }
        .exo-tab:hover { color:#334155; }
        .exo-tab.activo { color:#4f46e5; border-bottom-color:#4f46e5; }
        .exo-panel { display:none; } .exo-panel.activo { display:block; }
        .exo-nota { border:1px solid #f6ad55; background:#fffaf0; padding:10px 12px; border-radius:8px; font-size:.85rem; margin-top:12px; color:#7c4a03; }
        #resumen-exo { margin-top:10px; color:#475569; font-size:.9rem; }
      </style>

      <div class="exo-tabs">
        <button type="button" class="exo-tab activo" data-panel="exo-distrital">Distrital · Bogotá</button>
        <button type="button" class="exo-tab" data-panel="exo-nacional">DIAN · Nacional</button>
      </div>

      <div id="exo-distrital" class="exo-panel activo">
        <form class="panel" id="form-exo-distrital" onsubmit="return false;">
          <div class="panel-acento-superior"></div>
          <div class="panel-cuerpo">
            <div class="grid-form una-col">
              <div class="campo">
                <label>Insumo (Excel)</label>
                <div class="ayuda">Año gravable 2025 · Resolución DDI-024115. La entrega final se hace por PIDO / Storm User 4.4.2 (genera el .STR certificado). Este módulo produce el archivo de cargue con la estructura de la ficha técnica.</div>
                <div class="dropzone" id="drop-exo-d">
                  <input type="file" id="input-exo-d" accept=".xlsx,.xls">
                  <svg class="dropzone-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>
                  <div class="dropzone-texto"><strong>Arrastra el archivo</strong> o haz clic para seleccionarlo</div>
                  <div class="dropzone-detalle">.xlsx · un solo archivo</div>
                </div>
                <div class="lista-archivos" id="lista-exo-d"></div>
                <div id="resumen-exo"></div>
              </div>
            </div>
            <div class="barra-acciones">
              <button type="button" class="btn btn-primario" id="btn-exo-d" disabled>
                <svg class="icono-btn" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
                Generar archivo de cargue
              </button>
            </div>
            <div class="exo-nota">Pendiente conectar el generador: falta la ficha técnica / plantilla de Storm del formato distrital aplicable (Retenciones ICA o Compras). Al cargarla, se activa la generación.</div>
          </div>
        </form>
      </div>

      <div id="exo-nacional" class="exo-panel">
        <div class="panel"><div class="panel-cuerpo">
          <p style="color:#64748b;">Generador de <strong>DIAN Nacional · Formato 1001 v10</strong> (pagos y retenciones) en XML según el esquema oficial. <strong>Próximamente.</strong></p>
        </div></div>
      </div>
    </section>

'''

SCRIPT_TAG = '  <script src="static/js/modulo-10-exogena.js"></script>\n'

MODULOS_ENTRY = '  exogena:       { eyebrow: "Módulo 10", titulo: "Reportes Exógena" },\n'

MODULO_JS = '''/* ==========================================================================
 * Módulo 10 · Reportes Exógena  (Distrital Bogotá / DIAN Nacional)
 * --------------------------------------------------------------------------
 * Andamiaje: pestañas + carga de insumo con SheetJS en el navegador.
 * El generador se conecta cuando se defina la ficha técnica del formato.
 * ==========================================================================*/
(function () {
  "use strict";

  var input = document.getElementById("input-exo-d");
  if (!input) return;   // la vista no está en esta página

  // Pestañas Distrital / Nacional
  Array.prototype.forEach.call(document.querySelectorAll(".exo-tab"), function (tab) {
    tab.addEventListener("click", function () {
      document.querySelectorAll(".exo-tab").forEach(function (t) { t.classList.remove("activo"); });
      document.querySelectorAll(".exo-panel").forEach(function (p) { p.classList.remove("activo"); });
      tab.classList.add("activo");
      document.getElementById(tab.dataset.panel).classList.add("activo");
    });
  });

  if (typeof configurarDropzoneUnico === "function")
    configurarDropzoneUnico("drop-exo-d", "input-exo-d", "lista-exo-d");

  var resumen = document.getElementById("resumen-exo");
  var btn = document.getElementById("btn-exo-d");

  input.addEventListener("change", function (e) {
    var file = e.target.files[0];
    if (!file) return;
    resumen.textContent = "Leyendo insumo…";
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var wb = window.XLSX.read(ev.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var filas = aoa.length ? aoa.length - 1 : 0;
        var cols = aoa.length ? (aoa[0] || []).filter(function (c) { return c !== null && c !== ""; }).length : 0;
        resumen.innerHTML = "Insumo leído: <b>" + filas + "</b> filas · <b>" + cols + "</b> columnas. " +
          "(Generador pendiente de ficha técnica.)";
        // btn.disabled = false;  // se habilita al conectar el generador
      } catch (err) {
        resumen.textContent = "No se pudo leer el archivo: " + err.message;
      }
    };
    reader.readAsArrayBuffer(file);
  });
})();
'''


def _leer(p):
    with open(p, "r", encoding="utf-8") as f:
        return f.read()


def _escribir(p, c):
    shutil.copy2(p, p + ".bak")
    with open(p, "w", encoding="utf-8") as f:
        f.write(c)


def _antes(txt, ancla, frag, nombre):
    if txt.count(ancla) != 1:
        raise RuntimeError("[%s] ancla %d veces (esperaba 1)." % (nombre, txt.count(ancla)))
    return txt.replace(ancla, frag + ancla, 1)


def _despues(txt, ancla, frag, nombre):
    if txt.count(ancla) != 1:
        raise RuntimeError("[%s] ancla %d veces (esperaba 1)." % (nombre, txt.count(ancla)))
    return txt.replace(ancla, ancla + frag, 1)


def patch_index(ruta):
    txt = _leer(ruta)
    if 'data-vista="exogena"' in txt:
        print("  = index.html ya tiene el Módulo 10; se omite.")
        return
    # Item del sidebar: antes del cierre del nav
    txt = _antes(txt, "    </nav>", NAV_ITEM, "index/nav")
    # Sección: antes del cierre de <main>
    txt = _antes(txt, "  </main>", SECCION, "index/seccion")
    # Script del módulo: después del script del módulo 12 (oxp)
    ancla_script = '  <script src="static/js/modulo-12-oxp.js"></script>\n'
    if ancla_script in txt:
        txt = _despues(txt, ancla_script, SCRIPT_TAG, "index/script")
    else:
        # fallback: antes de </body>
        txt = _antes(txt, "</body>", SCRIPT_TAG, "index/script-fallback")
    _escribir(ruta, txt)
    print("  ✔ index.html: sidebar + sección (pestañas) + script.")


def patch_main(ruta):
    txt = _leer(ruta)
    if "exogena:" in txt:
        print("  = main.js ya tiene el Módulo 10; se omite.")
        return
    # Insertar la entrada MODULOS después de la línea del oxp (o antes del cierre del objeto)
    if 'titulo: "CDP-CRP de OXP" },\n' in txt:
        ancla = [l for l in txt.splitlines(keepends=True) if l.strip().startswith("oxp:")][0]
        txt = _despues(txt, ancla, MODULOS_ENTRY, "main/MODULOS")
    else:
        raise RuntimeError("[main/MODULOS] no se encontró la entrada 'oxp:' para anclar.")
    _escribir(ruta, txt)
    print("  ✔ main.js: entrada MODULOS.exogena insertada.")


def crear_js(jsdir):
    ruta = os.path.join(jsdir, "modulo-10-exogena.js")
    if os.path.exists(ruta):
        print("  = modulo-10-exogena.js ya existe; se omite.")
        return
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(MODULO_JS)
    print("  ✔ creado %s" % ruta)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default="frontend/index.html")
    ap.add_argument("--main",  default="frontend/static/js/main.js")
    ap.add_argument("--jsdir", default="frontend/static/js")
    args = ap.parse_args()

    errores = 0
    if not os.path.exists(args.index):
        print("  ✗ No existe: %s" % args.index); errores += 1
    else:
        try: patch_index(args.index)
        except RuntimeError as e: print("  ✗ %s" % e); errores += 1
    if not os.path.exists(args.main):
        print("  ✗ No existe: %s" % args.main); errores += 1
    else:
        try: patch_main(args.main)
        except RuntimeError as e: print("  ✗ %s" % e); errores += 1
    if os.path.isdir(args.jsdir):
        crear_js(args.jsdir)
    else:
        print("  ✗ No existe carpeta JS: %s" % args.jsdir); errores += 1

    if errores:
        print("\nTerminado con %d problema(s)." % errores); sys.exit(1)
    print("\n✅ Módulo 10 montado. Respaldos .bak junto a index.html y main.js.")


if __name__ == "__main__":
    main()
