(function() {
  "use strict";
  
  // ═══════════════════════════════════════════════════════════════
  // ASISTENTE IA GLOBAL - Conectado a los 10 módulos
  // ═══════════════════════════════════════════════════════════════
  
  const STORAGE_KEY = "claude_api_key_v1";
  const HISTORIAL_KEY = "asistente_historial_v1";
  let historial = [];
  
  // Cargar historial previo
  try {
    const raw = localStorage.getItem(HISTORIAL_KEY);
    if (raw) historial = JSON.parse(raw);
  } catch(e) {}
  
  function crearAsistente() {
    if (document.getElementById("asistente-ia-flotante")) return;
    
    const btn = document.createElement("div");
    btn.id = "asistente-ia-flotante";
    btn.innerHTML = `
      <div id="ia-boton" style="position:fixed;bottom:20px;right:20px;width:60px;height:60px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:50%;box-shadow:0 8px 24px rgba(102,126,234,.4);cursor:pointer;z-index:99999;display:flex;align-items:center;justify-content:center;transition:transform .2s;color:#fff;font-size:28px;">
        🤖
      </div>
      
      <div id="ia-panel" style="position:fixed;bottom:90px;right:20px;width:380px;max-width:calc(100vw - 40px);height:550px;max-height:calc(100vh - 120px);background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,.25);z-index:99998;display:none;flex-direction:column;overflow:hidden;">
        <div style="background:linear-gradient(135deg,#1a2742,#1e3a6e);color:#fff;padding:16px;">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <div>
              <div style="font-size:14px;font-weight:700;">🤖 Asistente Alcaldía</div>
              <div style="font-size:10px;opacity:.8;">Módulos 01-10 · Claude API</div>
            </div>
            <div style="display:flex;gap:8px;">
              <button id="ia-config" title="Configurar" style="background:transparent;border:none;color:#fff;font-size:16px;cursor:pointer;padding:4px 8px;">⚙️</button>
              <button id="ia-cerrar" style="background:transparent;border:none;color:#fff;font-size:20px;cursor:pointer;padding:0 6px;">×</button>
            </div>
          </div>
        </div>
        
        <div id="ia-mensajes" style="flex:1;overflow-y:auto;padding:16px;background:#f8f9fa;font-size:13px;line-height:1.5;">
          <div style="background:#e8edff;padding:12px;border-radius:12px;color:#1a2742;margin-bottom:10px;">
            <strong>👋 Hola Richard!</strong><br>
            Puedo ayudarte con los 10 módulos: Extracción, Plantilla, PAC, Estado de cuenta, Histórico, Presupuestal, Contratación, Planeación PDL, Ciberseguridad, Licencias.<br><br>
            Pregúntame sobre datos cargados, cálculos, procedimientos o cualquier duda del sistema.
          </div>
        </div>
        
        <div style="padding:12px;border-top:1px solid #e8edf2;background:#fff;">
          <div style="display:flex;gap:6px;margin-bottom:8px;flex-wrap:wrap;">
            <button class="ia-sugerencia" style="background:#f0f6ff;color:#667eea;border:1px solid #d0e0ff;border-radius:14px;padding:5px 10px;font-size:11px;cursor:pointer;">¿Cuánto está comprometido?</button>
            <button class="ia-sugerencia" style="background:#f0f6ff;color:#667eea;border:1px solid #d0e0ff;border-radius:14px;padding:5px 10px;font-size:11px;cursor:pointer;">Top proyectos por giros</button>
            <button class="ia-sugerencia" style="background:#f0f6ff;color:#667eea;border:1px solid #d0e0ff;border-radius:14px;padding:5px 10px;font-size:11px;cursor:pointer;">Riesgos de baja ejecución</button>
          </div>
          <div style="display:flex;gap:8px;">
            <input id="ia-input" placeholder="Pregunta algo..." style="flex:1;padding:10px;border:1.5px solid #e8edf2;border-radius:8px;font-size:13px;outline:none;">
            <button id="ia-enviar" style="background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;border:none;padding:10px 16px;border-radius:8px;cursor:pointer;font-weight:600;font-size:13px;">➤</button>
          </div>
        </div>
      </div>
    `;
    document.body.appendChild(btn);
    
    // Eventos
    document.getElementById("ia-boton").addEventListener("click", togglePanel);
    document.getElementById("ia-cerrar").addEventListener("click", () => togglePanel(false));
    document.getElementById("ia-config").addEventListener("click", configurarAPI);
    document.getElementById("ia-enviar").addEventListener("click", enviarMensaje);
    document.getElementById("ia-input").addEventListener("keypress", (e) => {
      if (e.key === "Enter") enviarMensaje();
    });
    document.querySelectorAll(".ia-sugerencia").forEach(btn => {
      btn.addEventListener("click", () => {
        document.getElementById("ia-input").value = btn.textContent;
        enviarMensaje();
      });
    });
    
    // Restaurar historial
    if (historial.length > 0) {
      historial.forEach(m => agregarMensajeUI(m.role, m.content));
    }
  }
  
  function togglePanel(force) {
    const panel = document.getElementById("ia-panel");
    const visible = panel.style.display === "flex";
    panel.style.display = (force === false || visible) ? "none" : "flex";
  }
  
  function configurarAPI() {
    const actual = localStorage.getItem(STORAGE_KEY) || "";
    const nueva = prompt("🔑 API Key de Claude (sk-ant-...):\n\nSe guarda en tu navegador. Consíguela en: https://console.anthropic.com/", actual);
    if (nueva !== null) {
      if (nueva.trim() === "") {
        localStorage.removeItem(STORAGE_KEY);
        alert("API Key eliminada");
      } else {
        localStorage.setItem(STORAGE_KEY, nueva.trim());
        alert("✅ API Key guardada");
      }
    }
  }
  
  function obtenerContextoModulos() {
    let ctx = "CONTEXTO DEL SISTEMA:\n";
    ctx += "Usuario: Richard Bello Roncancio - Economista, Alcaldía Local de Usme, Bogotá\n";
    ctx += "Sistema: Gestión Contractual y Financiera (10 módulos)\n\n";
    
    // Módulo 08 - Planeación PDL (variables globales del scope)
    try {
      const scripts = document.querySelectorAll("script");
      if (typeof window.PROYECTOS_08 !== "undefined" && window.PROYECTOS_08.length > 0) {
        ctx += "MÓDULO 08 - PLANEACIÓN PDL:\n";
        ctx += "Proyectos cargados: " + window.PROYECTOS_08.length + "\n";
        ctx += "POAI Total: " + window.TOTALS_08.poai + "\n";
        ctx += "Apropiaciones: " + window.TOTALS_08.apropiaciones + "\n";
        ctx += "Compromisos: " + window.TOTALS_08.compromisos + "\n";
        ctx += "Giros: " + window.TOTALS_08.giros + "\n";
        ctx += "Proyectos: " + JSON.stringify(window.PROYECTOS_08.slice(0, 10)) + "\n\n";
      }
    } catch(e) {}
    
    // Detectar qué módulos tienen dashboards visibles
    const modulos = [
      { id: "vista-extraccion", nombre: "01 Extracción de Causaciones" },
      { id: "vista-plantilla", nombre: "02 Plantilla de Pagos" },
      { id: "vista-pac", nombre: "03 Cuadre PAC" },
      { id: "vista-estadocuenta", nombre: "04 Estado de Cuenta" },
      { id: "vista-historico", nombre: "05 Histórico de Pagos" },
      { id: "vista-presupuestal", nombre: "06 Análisis Presupuestal" },
      { id: "vista-contratacion", nombre: "07 Contratación" },
      { id: "vista-planeacion", nombre: "08 Planeación PDL" },
      { id: "vista-ciberseguridad", nombre: "09 Ciberseguridad" },
      { id: "vista-licencias", nombre: "10 Licencias" }
    ];
    
    const activos = [];
    modulos.forEach(m => {
      const el = document.getElementById(m.id);
      if (el && el.offsetParent !== null) activos.push(m.nombre + " (ACTIVO)");
      else activos.push(m.nombre);
    });
    
    ctx += "MÓDULOS DISPONIBLES:\n" + activos.join("\n") + "\n\n";
    
    return ctx;
  }
  
  async function enviarMensaje() {
    const input = document.getElementById("ia-input");
    const texto = input.value.trim();
    if (!texto) return;
    
    const apiKey = localStorage.getItem(STORAGE_KEY);
    if (!apiKey) {
      alert("⚠️ Primero configura tu API Key\n\nClick en ⚙️ arriba a la derecha");
      configurarAPI();
      return;
    }
    
    input.value = "";
    agregarMensajeUI("user", texto);
    historial.push({ role: "user", content: texto });
    
    const cargandoId = "cargando-" + Date.now();
    const mensajes = document.getElementById("ia-mensajes");
    const cargando = document.createElement("div");
    cargando.id = cargandoId;
    cargando.style.cssText = "background:#fff;padding:12px;border-radius:12px;color:#888;font-style:italic;margin-bottom:10px;";
    cargando.textContent = "🤔 Pensando...";
    mensajes.appendChild(cargando);
    mensajes.scrollTop = mensajes.scrollHeight;
    
    try {
      const contexto = obtenerContextoModulos();
      const systemPrompt = "Eres un asistente experto en presupuesto público colombiano y gestión contractual para Alcaldías Locales. " +
        "Ayudas a Richard, economista del área de Presupuesto en la Alcaldía Local de Usme. " +
        "Conoces SAP Distrital, SIIF Nación, SECOP II, Ley 610/2000, Ley 358/1997, y el sistema de 10 módulos que él construyó. " +
        "Responde en español, de forma directa y práctica. Si tienes datos del sistema, úsalos.\n\n" + contexto;
      
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-5-20250929",
          max_tokens: 1024,
          system: systemPrompt,
          messages: historial.slice(-10)
        })
      });
      
      document.getElementById(cargandoId).remove();
      
      if (!response.ok) {
        const err = await response.text();
        throw new Error("Error " + response.status + ": " + err);
      }
      
      const data = await response.json();
      const respuesta = data.content[0].text;
      
      agregarMensajeUI("assistant", respuesta);
      historial.push({ role: "assistant", content: respuesta });
      
      // Guardar historial (últimos 20 mensajes)
      historial = historial.slice(-20);
      localStorage.setItem(HISTORIAL_KEY, JSON.stringify(historial));
      
    } catch (err) {
      document.getElementById(cargandoId).remove();
      agregarMensajeUI("assistant", "❌ Error: " + err.message + "\n\nVerifica tu API Key en ⚙️");
      console.error(err);
    }
  }
  
  function agregarMensajeUI(role, content) {
    const mensajes = document.getElementById("ia-mensajes");
    const div = document.createElement("div");
    if (role === "user") {
      div.style.cssText = "background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;padding:10px 14px;border-radius:12px;margin-bottom:10px;margin-left:40px;text-align:right;";
    } else {
      div.style.cssText = "background:#fff;color:#1a2742;padding:12px 14px;border-radius:12px;margin-bottom:10px;margin-right:40px;box-shadow:0 1px 3px rgba(0,0,0,.06);";
    }
    div.textContent = content;
    mensajes.appendChild(div);
    mensajes.scrollTop = mensajes.scrollHeight;
  }
  
  // Iniciar
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", crearAsistente);
  } else {
    crearAsistente();
  }
  
  console.log("🤖 Asistente IA global activado");
})();
