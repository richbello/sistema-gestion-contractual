/* ============================================================
   MÓDULO 09 — Gestor de Licencias ControlIA
   Vanilla JS. Conecta al Cloudflare Worker real.
   El ADMIN_SECRET NO va en el código: se pide al entrar y
   se guarda solo en memoria (se borra al cerrar la pestaña).
   Colores del módulo 07.
   ============================================================ */
(function () {
  "use strict";

  const WORKER_URL = "https://controlia-licencias.richbello.workers.dev";

  const C = {
    navy: "#1a2742", navy2: "#1e3a6e", card: "#ffffff", bg: "#f0f2f5",
    cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12", rojo: "#e74c3c",
    violet: "#8e44ad", muted: "#7f8c8d", border: "#e8edf2", texto: "#2c3e50",
  };

  const PLANES = {
    demo:     { label:"DEMO",        precio:0,       color:C.muted,  modulos:4  },
    basico:   { label:"EDIL BÁSICO", precio:390000,  color:C.cyan,   modulos:6  },
    pro:      { label:"EDIL PRO",    precio:690000,  color:C.verde,  modulos:10 },
    concejal: { label:"CONCEJAL",    precio:1200000, color:C.amar,   modulos:13 },
    premium:  { label:"PREMIUM",     precio:2200000, color:C.rojo,   modulos:13 },
  };

  const fmtCOP = (v) => v >= 1e6 ? "$" + (v/1e6).toFixed(1) + "M" : v >= 1e3 ? "$" + (v/1e3).toFixed(0) + "K" : "$" + v;

  let SECRET = null;         // solo en memoria
  let LICENCIAS = [];
  let TAB = "lista";
  let COPIADO = "";

  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();

  function init() {
    const vista = document.getElementById("vista-licencias");
    if (!vista || vista.dataset.ligado === "1") return;
    vista.dataset.ligado = "1";
    render();
  }

  function render() {
    const vista = document.getElementById("vista-licencias");
    if (!vista) return;
    const prev = document.getElementById("lic-root");
    if (prev) prev.remove();
    const root = document.createElement("div");
    root.id = "lic-root";
    vista.appendChild(root);
    if (!SECRET) renderAcceso(root);
    else renderPanel(root);
  }

  /* ---------- pantalla de acceso (pide secreto) ---------- */
  function renderAcceso(root) {
    root.innerHTML = `
      <div style="min-height:60vh;display:flex;align-items:center;justify-content:center;padding:24px;">
        <div style="width:100%;max-width:420px;text-align:center;">
          <div style="font-size:48px;margin-bottom:12px;">🔐</div>
          <div style="font-size:10px;color:${C.cyan};letter-spacing:4px;font-weight:700;text-transform:uppercase;">XBOG Technologies · Panel de Control</div>
          <h2 style="color:${C.navy};font-size:22px;font-weight:900;margin:8px 0 16px;">Gestor de Licencias ControlIA</h2>
          <div style="background:#fff;border-radius:16px;padding:28px;box-shadow:0 12px 32px rgba(26,39,66,.15);text-align:left;">
            <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Clave de administrador</label>
            <input id="lic-secret" type="password" placeholder="Ingresa el secreto de admin" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:10px;padding:12px 16px;color:${C.texto};font-size:13px;outline:none;box-sizing:border-box;margin-bottom:14px;">
            <div id="lic-acc-err" style="display:none;background:${C.rojo}18;border:1px solid ${C.rojo}44;border-radius:8px;padding:9px 12px;margin-bottom:14px;font-size:12px;color:${C.rojo};font-weight:600;"></div>
            <button id="lic-entrar" style="width:100%;background:linear-gradient(135deg,${C.navy},${C.navy2});border:none;border-radius:10px;padding:12px;color:#fff;font-weight:800;font-size:14px;cursor:pointer;">🔓 Acceder al panel</button>
            <div style="margin-top:14px;font-size:10px;color:${C.muted};text-align:center;">La clave no se guarda: solo vive mientras la pestaña esté abierta.</div>
          </div>
        </div>
      </div>
    `;
    const err = root.querySelector("#lic-acc-err");
    const entrar = async () => {
      const val = root.querySelector("#lic-secret").value.trim();
      if (!val) { err.textContent = "⚠️ Ingresa la clave."; err.style.display = "block"; return; }
      err.style.display = "none";
      root.querySelector("#lic-entrar").textContent = "⏳ Verificando...";
      // probar el secreto con /listar
      try {
        const r = await fetch(WORKER_URL + "/listar", { headers: { "X-Admin-Secret": val } });
        if (r.status === 401 || r.status === 403) {
          err.textContent = "⚠️ Clave incorrecta.";
          err.style.display = "block";
          root.querySelector("#lic-entrar").textContent = "🔓 Acceder al panel";
          return;
        }
        const d = await r.json();
        SECRET = val;
        LICENCIAS = (d.licencias || []).sort((a, b) => new Date(b.creada).getTime() - new Date(a.creada).getTime());
        render();
      } catch (e) {
        err.textContent = "⚠️ No se pudo conectar con el servidor.";
        err.style.display = "block";
        root.querySelector("#lic-entrar").textContent = "🔓 Acceder al panel";
      }
    };
    root.querySelector("#lic-entrar").addEventListener("click", entrar);
    root.querySelector("#lic-secret").addEventListener("keydown", (e) => { if (e.key === "Enter") entrar(); });
  }

  /* ---------- API ---------- */
  async function cargar() {
    try {
      const r = await fetch(WORKER_URL + "/listar", { headers: { "X-Admin-Secret": SECRET } });
      const d = await r.json();
      if (d.ok) LICENCIAS = d.licencias.sort((a, b) => new Date(b.creada).getTime() - new Date(a.creada).getTime());
    } catch { /* noop */ }
  }

  /* ---------- panel ---------- */
  function renderPanel(root) {
    const total = LICENCIAS.length;
    const activas = LICENCIAS.filter((l) => l.activa).length;
    const demo = LICENCIAS.filter((l) => l.plan === "demo").length;
    const pagas = LICENCIAS.filter((l) => l.plan !== "demo" && l.activa).length;
    const ingresos = LICENCIAS.filter((l) => l.activa && l.plan !== "demo").reduce((s, l) => s + (PLANES[l.plan] ? PLANES[l.plan].precio : 0), 0);

    root.innerHTML = `
      <div style="background:linear-gradient(135deg,${C.navy},${C.navy2});border-radius:12px;padding:18px 24px;margin-bottom:16px;color:#fff;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;">
          <div>
            <div style="font-size:10px;color:${C.cyan};letter-spacing:3px;font-weight:700;text-transform:uppercase;">🔐 XBOG Technologies · Panel de Control</div>
            <h2 style="margin:6px 0 4px;font-size:20px;font-weight:900;">Gestor de Licencias ControlIA</h2>
            <div style="font-size:11px;color:rgba(255,255,255,.7);">Administra claves de acceso para clientes de la plataforma</div>
          </div>
          <button id="lic-salir" style="background:${C.rojo}33;border:1px solid ${C.rojo}66;border-radius:8px;padding:7px 14px;color:#fff;cursor:pointer;font-size:12px;font-weight:700;">🚪 Salir</button>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-top:16px;">
          ${stat("🔑","Total licencias",total,C.cyan)}
          ${stat("✅","Activas",activas,C.verde)}
          ${stat("🎯","Demo",demo,C.amar)}
          ${stat("💰","Clientes pagos",pagas,C.rojo)}
          ${stat("📈","Ingresos/mes est.",fmtCOP(ingresos),C.violet)}
        </div>
      </div>

      <div id="lic-msg" style="display:none;border-radius:8px;padding:10px 16px;margin-bottom:12px;font-size:12px;font-weight:600;justify-content:space-between;align-items:center;"></div>

      <div style="background:#fff;border-radius:10px 10px 0 0;border-bottom:2px solid ${C.border};display:flex;padding:0 12px;">
        ${tabBtn("lista",`📋 Licencias (${total})`)}
        ${tabBtn("crear","➕ Nueva licencia")}
        <button id="lic-refresh" style="margin-left:auto;padding:12px 14px;background:transparent;border:none;color:${C.muted};cursor:pointer;font-size:12px;">🔄 Actualizar</button>
      </div>
      <div id="lic-tab" style="background:#fff;border-radius:0 0 10px 10px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);"></div>
    `;

    root.querySelector("#lic-salir").addEventListener("click", () => { SECRET = null; LICENCIAS = []; render(); });
    root.querySelector("#lic-refresh").addEventListener("click", async () => { await cargar(); render(); });
    root.querySelectorAll("[data-tab]").forEach((b) => b.addEventListener("click", () => { TAB = b.dataset.tab; pintarTab(); }));
    pintarTab();
  }

  function stat(icon, label, valor, color) {
    return `<div style="background:rgba(255,255,255,.1);border-radius:10px;padding:12px 14px;">
      <div style="font-size:18px;">${icon}</div>
      <div style="font-size:22px;font-weight:800;color:${color};margin-top:4px;">${valor}</div>
      <div style="font-size:11px;color:rgba(255,255,255,.7);margin-top:3px;">${label}</div>
    </div>`;
  }
  function tabBtn(id, label) {
    const a = TAB === id;
    return `<button data-tab="${id}" style="padding:12px 20px;background:transparent;border:none;border-bottom:${a?`2px solid ${C.cyan}`:"2px solid transparent"};color:${a?C.cyan:C.muted};cursor:pointer;font-size:12px;font-weight:${a?700:500};">${label}</button>`;
  }
  function msg(texto, color) {
    const m = document.getElementById("lic-msg");
    if (!m) return;
    m.style.display = "flex";
    m.style.background = color + "22";
    m.style.border = `1px solid ${color}44`;
    m.style.color = color;
    m.innerHTML = `<span>${texto}</span><button id="lic-msg-x" style="background:none;border:none;color:${color};cursor:pointer;font-size:16px;">✕</button>`;
    m.querySelector("#lic-msg-x").addEventListener("click", () => { m.style.display = "none"; });
  }

  function pintarTab() {
    const root = document.getElementById("lic-root");
    root.querySelectorAll("[data-tab]").forEach((b) => {
      const a = b.dataset.tab === TAB;
      b.style.borderBottom = a ? `2px solid ${C.cyan}` : "2px solid transparent";
      b.style.color = a ? C.cyan : C.muted;
      b.style.fontWeight = a ? 700 : 500;
    });
    const cont = document.getElementById("lic-tab");
    if (TAB === "lista") tabLista(cont);
    else tabCrear(cont);
  }

  /* ---------- lista ---------- */
  function tabLista(cont) {
    if (LICENCIAS.length === 0) {
      cont.innerHTML = `<div style="text-align:center;padding:40px;color:${C.muted};">No hay licencias creadas aún. Crea la primera.</div>`;
      return;
    }
    cont.innerHTML = `<div style="display:flex;flex-direction:column;gap:10px;">` + LICENCIAS.map((lic) => {
      const info = PLANES[lic.plan] || PLANES.demo;
      const expirada = lic.expira ? Date.now() > lic.expira : false;
      const diasRest = lic.expira ? Math.max(0, Math.ceil((lic.expira - Date.now()) / 86400000)) : null;
      const activoColor = lic.activa && !expirada ? info.color : C.muted;
      return `<div style="background:${C.bg};border-radius:10px;padding:14px 18px;border:1px solid ${lic.activa && !expirada ? info.color + "44" : C.border};border-left:3px solid ${activoColor};">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;">
          <div style="flex:2;min-width:200px;">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
              <span style="background:${info.color}22;color:${info.color};border:1px solid ${info.color}44;border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700;">${info.label}</span>
              ${!lic.activa ? `<span style="background:${C.rojo}22;color:${C.rojo};border-radius:20px;padding:2px 8px;font-size:10px;font-weight:700;">SUSPENDIDA</span>` : ""}
              ${expirada ? `<span style="background:${C.muted}22;color:${C.muted};border-radius:20px;padding:2px 8px;font-size:10px;font-weight:700;">EXPIRADA</span>` : ""}
            </div>
            <div style="font-weight:700;font-size:14px;color:${C.texto};">${lic.cliente}</div>
            <div style="font-size:11px;color:${C.muted};">${lic.localidad || ""}</div>
          </div>
          <div style="flex:1;min-width:170px;">
            <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px;">
              <span style="font-family:monospace;font-size:12px;color:${C.navy2};letter-spacing:1px;font-weight:700;">${lic.clave}</span>
              <button data-copiar="${lic.clave}" style="background:none;border:none;cursor:pointer;font-size:14px;">${COPIADO === lic.clave ? "✅" : "📋"}</button>
            </div>
            <div style="font-size:10px;color:${C.muted};line-height:1.6;">
              Creada: ${new Date(lic.creada).toLocaleDateString("es-CO")}<br>
              ${lic.expira ? `Vence: ${new Date(lic.expira).toLocaleDateString("es-CO")} (${diasRest}d)` : "Sin vencimiento"}<br>
              Accesos: ${lic.accesos || 0}
            </div>
          </div>
          <div style="display:flex;gap:6px;align-items:center;">
            ${info.precio > 0 ? `<span style="font-size:11px;font-weight:700;color:${C.verde};">${fmtCOP(info.precio)}/mes</span>` : ""}
            <button data-toggle="${lic.clave}" data-activa="${lic.activa}" style="background:${lic.activa ? C.rojo + "22" : C.verde + "22"};border:1px solid ${lic.activa ? C.rojo : C.verde}44;border-radius:6px;padding:5px 12px;color:${lic.activa ? C.rojo : C.verde};cursor:pointer;font-size:11px;font-weight:700;">${lic.activa ? "Suspender" : "Activar"}</button>
          </div>
        </div>
      </div>`;
    }).join("") + `</div>`;

    cont.querySelectorAll("[data-copiar]").forEach((b) => b.addEventListener("click", () => {
      navigator.clipboard.writeText(b.dataset.copiar);
      COPIADO = b.dataset.copiar;
      pintarTab();
      setTimeout(() => { COPIADO = ""; if (TAB === "lista") pintarTab(); }, 2000);
    }));
    cont.querySelectorAll("[data-toggle]").forEach((b) => b.addEventListener("click", async () => {
      const activa = b.dataset.activa === "true";
      try {
        await fetch(WORKER_URL + "/revocar", {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-Admin-Secret": SECRET },
          body: JSON.stringify({ clave: b.dataset.toggle, activa: !activa }),
        });
        await cargar();
        render();
      } catch { msg("Error actualizando licencia.", C.rojo); }
    }));
  }

  /* ---------- crear ---------- */
  let FORM = { plan: "demo", cliente: "", localidad: "", dias: "15" };

  function tabCrear(cont) {
    cont.innerHTML = `
      <div style="max-width:520px;">
        <div style="font-weight:700;font-size:14px;color:${C.texto};margin-bottom:20px;">Nueva licencia de acceso</div>
        <div style="margin-bottom:16px;">
          <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Plan</label>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
            ${Object.entries(PLANES).map(([key, info]) => `
              <div data-plan="${key}" style="padding:10px 14px;border-radius:8px;border:1.5px solid ${FORM.plan === key ? info.color : C.border};background:${FORM.plan === key ? info.color + "11" : "transparent"};cursor:pointer;">
                <div style="font-weight:700;font-size:12px;color:${FORM.plan === key ? info.color : C.muted};">${info.label}</div>
                <div style="font-size:10px;color:${C.muted};">${info.modulos} módulos · ${info.precio > 0 ? fmtCOP(info.precio) + "/mes" : "Gratis"}</div>
              </div>`).join("")}
          </div>
        </div>
        <div style="margin-bottom:14px;">
          <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Nombre del cliente *</label>
          <input id="fc-cliente" value="${FORM.cliente}" placeholder="Edil Juan García" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:8px;padding:10px 14px;color:${C.texto};font-size:12px;outline:none;box-sizing:border-box;">
        </div>
        <div style="margin-bottom:14px;">
          <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Localidad / municipio</label>
          <input id="fc-localidad" value="${FORM.localidad}" placeholder="Usme · Bogotá D.C." style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:8px;padding:10px 14px;color:${C.texto};font-size:12px;outline:none;box-sizing:border-box;">
        </div>
        <div style="margin-bottom:14px;">
          <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Días de vigencia (vacío = sin límite)</label>
          <input id="fc-dias" value="${FORM.dias}" placeholder="15" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:8px;padding:10px 14px;color:${C.texto};font-size:12px;outline:none;box-sizing:border-box;">
        </div>
        <div style="background:${C.bg};border-radius:8px;padding:12px 14px;margin-bottom:16px;font-size:11px;color:${C.muted};">
          <div style="font-weight:700;color:${C.texto};margin-bottom:4px;">La clave generada tendrá formato:</div>
          <div style="font-family:monospace;color:${C.navy2};font-size:13px;">CP-${FORM.plan.toUpperCase().slice(0,3)}-XXXX-XXXX-XXXX-XXXX</div>
          <div style="margin-top:4px;">16 caracteres aleatorios · encriptada en Cloudflare KV</div>
        </div>
        <button id="fc-crear" style="width:100%;background:linear-gradient(135deg,${C.navy},${C.navy2});border:none;border-radius:10px;padding:13px;color:#fff;font-weight:800;font-size:14px;cursor:pointer;">🔑 Generar licencia</button>
      </div>
    `;

    cont.querySelectorAll("[data-plan]").forEach((d) => d.addEventListener("click", () => {
      FORM.plan = d.dataset.plan;
      FORM.dias = d.dataset.plan === "demo" ? "15" : "";
      // preservar lo escrito
      FORM.cliente = cont.querySelector("#fc-cliente").value;
      FORM.localidad = cont.querySelector("#fc-localidad").value;
      tabCrear(cont);
    }));

    cont.querySelector("#fc-crear").addEventListener("click", async () => {
      FORM.cliente = cont.querySelector("#fc-cliente").value.trim();
      FORM.localidad = cont.querySelector("#fc-localidad").value.trim();
      FORM.dias = cont.querySelector("#fc-dias").value.trim();
      if (!FORM.cliente) { msg("Ingresa el nombre del cliente.", C.rojo); return; }
      const btn = cont.querySelector("#fc-crear");
      btn.textContent = "⏳ Creando...";
      try {
        const body = { plan: FORM.plan, cliente: FORM.cliente, localidad: FORM.localidad };
        if (FORM.dias && Number(FORM.dias) > 0) body.dias = Number(FORM.dias);
        const r = await fetch(WORKER_URL + "/crear", {
          method: "POST",
          headers: { "Content-Type": "application/json", "X-Admin-Secret": SECRET },
          body: JSON.stringify(body),
        });
        const d = await r.json();
        if (d.ok) {
          msg("✅ Licencia creada: " + d.clave, C.verde);
          FORM = { plan: "demo", cliente: "", localidad: "", dias: "15" };
          await cargar();
          TAB = "lista";
          render();
        } else {
          msg("❌ " + (d.error || "No se pudo crear"), C.rojo);
          btn.textContent = "🔑 Generar licencia";
        }
      } catch {
        msg("❌ Error de conexión.", C.rojo);
        btn.textContent = "🔑 Generar licencia";
      }
    });
  }

  console.log("✅ Módulo 09 (Licencias) listo");
})();
