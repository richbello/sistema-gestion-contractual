#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_reorganizacion.py
=========================
Reorganiza la barra lateral del SPA SIN tocar la lógica de los módulos
(los data-vista e IDs de sección se conservan; solo cambia la presentación).

Cambios:
  - Software: "Gestión Contractual y Financiera" -> "Gestión Presupuestal y Financiera"
  - Módulo 1 "Trámite de Pagos" con submenú "Personas Naturales" que agrupa:
      Extracción, Plantilla de pagos, Cuadre PAC, Validación de plantilla, Cuentas bancarias
  - Renumeración: Estados de Cuenta=2, Históricos de Pagos=3, Análisis Presupuestal=4,
      Contratación=5, Planeación=6, Ciberseguridad=7, Licencias=8, CDP-CRP de OXP=9
  - Pie: "Derechos Reservados · XBOG SMART 2026"
  - CSS del submenú inyectado en <head> + toggle en main.js

Idempotente y con respaldo .bak.

Uso (raíz del repo, en Codespaces):
    python aplicar_reorganizacion.py --index frontend/index.html --main frontend/static/js/main.js
"""

import argparse
import os
import shutil
import sys

# ---------------------------------------------------------------------------
ICON = {
    "tramite":   '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="4" y="3" width="16" height="18" rx="1.5"/><path d="M8 8h8M8 12h8M8 16h5"/></svg>',
    "extraccion":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><path d="M14 2v6h6"/><path d="M9 13h6M9 17h6M9 9h1"/></svg>',
    "plantilla": '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M3 9h18M9 9v11"/></svg>',
    "pac":       '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 3v18h18"/><path d="M7 14l4-4 3 3 5-6"/></svg>',
    "validacion":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M9 11l3 3 8-8"/><path d="M20 12v6a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h9"/></svg>',
    "cuentas":   '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20M6 15h4"/></svg>',
    "estadocuenta":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 4h16v16H4z"/><path d="M4 9h16M9 9v11"/><circle cx="6.3" cy="6.5" r="0.6" fill="currentColor" stroke="none"/></svg>',
    "historico": '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 9h18M3 15h18M5 3h14a2 2 0 012 2v14a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2z"/></svg>',
    "presupuestal":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 12h18M3 6h18M3 18h18M8 3v18M16 3v18"/></svg>',
    "contratacion":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M9 12h6M9 16h6M9 8h6M5 3h14a2 2 0 012 2v14a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2z"/></svg>',
    "planeacion":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>',
    "ciberseguridad":'<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 2l7 4v6c0 5-3 8-7 10-4-2-7-5-7-10V6z"/><path d="M9 12l2 2 4-4"/></svg>',
    "licencias": '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M15 7a4 4 0 11-8 0 4 4 0 018 0z"/><path d="M11 11l7 7 2-2-2-2 2-2-3-3"/></svg>',
    "oxp":       '<svg class="nav-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 4h16v16H4z"/><path d="M4 9h16M9 9v11"/></svg>',
    "chevron":   '<svg class="nav-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 6l6 6-6 6"/></svg>',
}


def _sub(vista, icono, label):
    return ('        <button class="nav-item nav-sub" data-vista="%s">'
            '<span class="nav-numero"></span>%s %s</button>\n'
            % (vista, ICON[icono], label))


def _top(vista, num, icono, label):
    return ('      <button class="nav-item" data-vista="%s"><span class="nav-numero">%s</span>%s %s</button>\n'
            % (vista, num, ICON[icono], label))


NAV_NUEVO = (
    '<nav class="nav-modulos">\n'
    '      <span class="nav-grupo-titulo">Módulos del proceso</span>\n\n'
    '      <button class="nav-item nav-parent expandido" data-submenu="tramite">'
    '<span class="nav-numero">01</span>' + ICON["tramite"] + ' Trámite de Pagos ' + ICON["chevron"] + '</button>\n'
    '      <div class="nav-submenu abierto" id="submenu-tramite">\n'
    '        <span class="nav-subgrupo-titulo">Personas Naturales</span>\n'
    + _sub("extraccion", "extraccion", "Extracción de causaciones")
    + _sub("plantilla", "plantilla", "Plantilla de pagos")
    + _sub("pac", "pac", "Cuadre PAC")
    + _sub("validacion", "validacion", "Validación de plantilla")
    + '        <a class="nav-item nav-sub" href="cuentas_bancarias.html" target="_blank" rel="noopener" style="text-decoration:none;">'
      '<span class="nav-numero"></span>' + ICON["cuentas"] + ' Cuentas bancarias</a>\n'
    '      </div>\n\n'
    + _top("estadocuenta", "02", "estadocuenta", "Estados de Cuenta")
    + _top("historico", "03", "historico", "Históricos de Pagos")
    + _top("presupuestal", "04", "presupuestal", "Análisis Presupuestal")
    + _top("contratacion", "05", "contratacion", "Contratación")
    + _top("planeacion", "06", "planeacion", "Planeación")
    + _top("ciberseguridad", "07", "ciberseguridad", "Ciberseguridad")
    + _top("licencias", "08", "licencias", "Licencias")
    + _top("oxp", "09", "oxp", "CDP-CRP de OXP")
    + '    </nav>'
)

CSS_SUBMENU = '''<style id="estilos-submenu-modulos">
  .nav-parent { position:relative; }
  .nav-chevron { margin-left:auto; width:14px; height:14px; transition:transform .2s ease; opacity:.55; flex:none; }
  .nav-parent.expandido .nav-chevron { transform:rotate(90deg); }
  .nav-submenu { display:none; margin:2px 0 8px; padding-left:6px; border-left:1px solid rgba(201,162,39,.28); }
  .nav-submenu.abierto { display:block; }
  .nav-subgrupo-titulo { display:block; font-size:.6rem; letter-spacing:1.2px; text-transform:uppercase; opacity:.5; margin:6px 0 4px 22px; }
  .nav-item.nav-sub { padding-left:22px; font-size:.85rem; }
  .sidebar-pie .copyright { margin-top:10px; opacity:.7; }
</style>
'''

MODULOS_NUEVO = '''const MODULOS = {
  extraccion:    { eyebrow: "Módulo 1 · Trámite de Pagos", titulo: "Extracción de causaciones" },
  plantilla:     { eyebrow: "Módulo 1 · Trámite de Pagos", titulo: "Plantilla de pagos" },
  pac:           { eyebrow: "Módulo 1 · Trámite de Pagos", titulo: "Cuadre PAC" },
  validacion:    { eyebrow: "Módulo 1 · Trámite de Pagos", titulo: "Validación de plantilla" },
  estadocuenta:  { eyebrow: "Módulo 2", titulo: "Estados de Cuenta" },
  historico:     { eyebrow: "Módulo 3", titulo: "Históricos de Pagos" },
  presupuestal:  { eyebrow: "Módulo 4", titulo: "Análisis Presupuestal" },
  contratacion:  { eyebrow: "Módulo 5", titulo: "Contratación" },
  planeacion:    { eyebrow: "Módulo 6", titulo: "Planeación PDL" },
  ciberseguridad:{ eyebrow: "Módulo 7", titulo: "Ciberseguridad" },
  licencias:     { eyebrow: "Módulo 8", titulo: "Licencias" },
  oxp:           { eyebrow: "Módulo 9", titulo: "CDP-CRP de OXP" },
};'''

HANDLER_VIEJO = '''document.querySelectorAll(".nav-item").forEach(btn => {
  btn.addEventListener("click", () => cambiarVista(btn.dataset.vista));
});'''

HANDLER_NUEVO = '''document.querySelectorAll(".nav-item").forEach(btn => {
  btn.addEventListener("click", () => {
    if (btn.dataset.submenu) {
      const sm = document.getElementById("submenu-" + btn.dataset.submenu);
      if (sm) sm.classList.toggle("abierto");
      btn.classList.toggle("expandido");
      return;
    }
    if (!btn.dataset.vista) return;   // enlaces externos (p.ej. Cuentas bancarias)
    cambiarVista(btn.dataset.vista);
  });
});'''


def _leer(p):
    with open(p, "r", encoding="utf-8") as f:
        return f.read()


def _escribir(p, c):
    shutil.copy2(p, p + ".bak")
    with open(p, "w", encoding="utf-8") as f:
        f.write(c)


def _reemplazar_unico(txt, viejo, nuevo, nombre):
    if txt.count(viejo) != 1:
        raise RuntimeError("[%s] patrón encontrado %d veces (esperaba 1)." % (nombre, txt.count(viejo)))
    return txt.replace(viejo, nuevo, 1)


def _reemplazar_entre(txt, ini, fin, nuevo, nombre):
    a = txt.find(ini)
    if a == -1:
        raise RuntimeError("[%s] no se encontró inicio %r." % (nombre, ini))
    b = txt.find(fin, a)
    if b == -1:
        raise RuntimeError("[%s] no se encontró fin %r." % (nombre, fin))
    return txt[:a] + nuevo + txt[b + len(fin):]


def patch_index(ruta):
    txt = _leer(ruta)
    if "Trámite de Pagos" in txt or 'id="estilos-submenu-modulos"' in txt:
        print("  = index.html ya reorganizado; se omite.")
        return

    # Título del documento
    txt = _reemplazar_unico(
        txt,
        "<title>Sistema de Gestión Contractual y Financiera</title>",
        "<title>Sistema de Gestión Presupuestal y Financiera</title>",
        "index/title")

    # Marca (nombre del software)
    txt = _reemplazar_unico(
        txt,
        '<p class="marca-titulo">Gestión Contractual<br>y Financiera</p>',
        '<p class="marca-titulo">Gestión Presupuestal<br>y Financiera</p>',
        "index/marca")

    # Eyebrow inicial del topbar
    txt = _reemplazar_unico(
        txt,
        '<span class="topbar-eyebrow" id="eyebrow-modulo">Módulo 01</span>',
        '<span class="topbar-eyebrow" id="eyebrow-modulo">Módulo 1 · Trámite de Pagos</span>',
        "index/eyebrow")

    # Reemplazo COMPLETO del nav (conserva data-vista e IDs)
    txt = _reemplazar_entre(txt, '<nav class="nav-modulos">', '</nav>', NAV_NUEVO, "index/nav")

    # Pie con copyright
    txt = _reemplazar_unico(
        txt,
        '''<div class="sidebar-pie">
      Plataforma interna de apoyo a la gestión presupuestal y de pagos a contratistas.
    </div>''',
        '''<div class="sidebar-pie">
      Plataforma interna de apoyo a la gestión presupuestal y de pagos a contratistas.
      <div class="copyright">Derechos Reservados · XBOG SMART 2026</div>
    </div>''',
        "index/pie")

    # CSS del submenú antes de </head>
    txt = _reemplazar_unico(txt, "</head>", CSS_SUBMENU + "</head>", "index/css")

    _escribir(ruta, txt)
    print("  ✔ index.html: sidebar reorganizado + software renombrado + pie + CSS.")


def patch_main(ruta):
    txt = _leer(ruta)
    if "Trámite de Pagos" in txt:
        print("  = main.js ya reorganizado; se omite.")
        return

    # Reemplazo COMPLETO del objeto MODULOS
    txt = _reemplazar_entre(txt, "const MODULOS = {", "};", MODULOS_NUEVO, "main/MODULOS")

    # Handler de navegación (soporta submenú y enlaces externos)
    txt = _reemplazar_unico(txt, HANDLER_VIEJO, HANDLER_NUEVO, "main/handler")

    _escribir(ruta, txt)
    print("  ✔ main.js: MODULOS renumerado + handler con submenú.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default="frontend/index.html")
    ap.add_argument("--main",  default="frontend/static/js/main.js")
    args = ap.parse_args()

    errores = 0
    for etiqueta, ruta, fn in [("index.html", args.index, patch_index),
                               ("main.js", args.main, patch_main)]:
        if not os.path.exists(ruta):
            print("  ✗ No existe: %s" % ruta); errores += 1; continue
        try:
            fn(ruta)
        except RuntimeError as e:
            print("  ✗ %s" % e); errores += 1

    if errores:
        print("\nTerminado con %d problema(s)." % errores); sys.exit(1)
    print("\n✅ Reorganización aplicada. Respaldos .bak junto a cada archivo.")


if __name__ == "__main__":
    main()
