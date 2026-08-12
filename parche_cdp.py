import sys, shutil, os
RUTA = "frontend/index.html"
TAB_VIEJA = '<button type="button" class="oxp-tab" data-panel="oxp-cdp" disabled>CDP OXP · próximamente</button>'
TAB_NUEVA = '<button type="button" class="oxp-tab" data-panel="oxp-cdp">CDP OXP</button>'
PANEL_VIEJO = '''      <div id="oxp-cdp" class="oxp-panel">
        <div class="panel"><div class="panel-cuerpo">
          <p style="color:#64748b;">El generador de <strong>CDP OXP</strong> estará disponible próximamente. Por ahora este módulo constituye únicamente el <strong>CRP OXP</strong>.</p>
        </div></div>
      </div>'''
PANEL_NUEVO = '''      <div id="oxp-cdp" class="oxp-panel">
        <form class="panel" id="form-oxp-cdp" onsubmit="return false;">
          <div class="panel-acento-superior"></div>
          <div class="panel-cuerpo">
            <div class="grid-form una-col">
              <div class="campo">
                <label>Reporte CRP</label>
                <div class="ayuda">Mismo archivo del CRP OXP. Los rubros se convierten automaticamente.</div>
                <div class="dropzone" id="drop-oxp-cdp">
                  <input type="file" id="input-oxp-cdp" accept=".xlsx,.xls" required>
                  <svg class="dropzone-icono" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>
                  <div class="dropzone-texto"><strong>Arrastra el archivo</strong> o haz clic para seleccionarlo</div>
                  <div class="dropzone-detalle">.xlsx un solo archivo</div>
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
              <div class="estado-proceso" id="estado-oxp-cdp"><span class="spinner"></span> Generando formato...</div>
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
      </div>'''
with open(RUTA, "r", encoding="utf-8") as f:
    html = f.read()
if 'id="input-oxp-cdp"' in html:
    print("= Ya presente, nada que hacer."); sys.exit(0)
c = 0
if TAB_VIEJA in html:
    html = html.replace(TAB_VIEJA, TAB_NUEVA, 1); c += 1; print("OK pestana")
else:
    print("WARN pestana no encontrada")
if PANEL_VIEJO in html:
    html = html.replace(PANEL_VIEJO, PANEL_NUEVO, 1); c += 1; print("OK panel")
else:
    print("WARN panel no encontrado")
if c:
    shutil.copy2(RUTA, RUTA + ".bak")
    with open(RUTA, "w", encoding="utf-8") as f:
        f.write(html)
    print("Guardado. Cambios:", c)
else:
    print("Sin cambios")
