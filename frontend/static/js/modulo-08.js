/* ============================================================
   MÓDULO 08 — Ciberseguridad (login + auditoría)
   Vanilla JS + localStorage. Colores del módulo 07.
   NOTA: login de demostración; las claves se guardan en el
   navegador sin cifrar. No es seguridad de producción.
   ============================================================ */
(function () {
  "use strict";

  // ---- Paleta (alineada al módulo 07) ----
  const C = {
    navy: "#1a2742", navy2: "#1e3a6e", card: "#ffffff", bg: "#f0f2f5",
    cyan: "#5dade2", verde: "#27ae60", amar: "#f39c12", rojo: "#e74c3c",
    violet: "#8e44ad", muted: "#7f8c8d", border: "#e8edf2", texto: "#2c3e50",
  };

  const KEY_USERS = "cp_ciberseg_usuarios";
  const KEY_LOGS = "cp_ciberseg_logs";
  const KEY_SES = "cp_ciberseg_sesion";
  const MAX_INTENTOS = 3;
  const SESION_MIN = 30;
  const DIAS = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"];

  const USUARIOS_DEFAULT = [
    { id:"u1", nombre:"Administrador XBOG", usuario:"admin", clave:"Admin2026#", rol:"admin", activo:true, ultimoAcceso:"", intentosFallidos:0, bloqueado:false },
    { id:"u2", nombre:"Edil Usme", usuario:"edil", clave:"Edil2026*", rol:"edil", activo:true, ultimoAcceso:"", intentosFallidos:0, bloqueado:false },
  ];

  let PAGINA = 0;
  const FILAS = 15;
  const F = { acc:"TODOS", user:"TODOS", fecha:"", busca:"" };

  // ---- helpers de estado ----
  const ahora = () => {
    const d = new Date();
    return {
      fecha: d.toLocaleDateString("es-CO",{day:"2-digit",month:"2-digit",year:"numeric"}),
      hora: d.toLocaleTimeString("es-CO",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),
      dia: DIAS[d.getDay()], ts: d.toISOString(),
    };
  };
  const uid = () => Math.random().toString(36).slice(2,10).toUpperCase();
  const getIP = () => "192.168.1." + Math.floor(Math.random()*254+1);
  const getDisp = () => {
    const ua = navigator.userAgent;
    if (/mobile/i.test(ua)) return "📱 Móvil";
    if (/tablet/i.test(ua)) return "📟 Tablet";
    return "🖥️ Escritorio";
  };

  const loadUsers = () => { try { return JSON.parse(localStorage.getItem(KEY_USERS)||"null")||USUARIOS_DEFAULT; } catch { return USUARIOS_DEFAULT; } };
  const saveUsers = (u) => localStorage.setItem(KEY_USERS, JSON.stringify(u));
  const loadLogs = () => { try { return JSON.parse(localStorage.getItem(KEY_LOGS)||"[]"); } catch { return []; } };
  const saveLogs = (l) => localStorage.setItem(KEY_LOGS, JSON.stringify(l));
  const loadSes = () => { try { const s=JSON.parse(localStorage.getItem(KEY_SES)||"null"); if(s&&Date.now()<s.expira) return s; localStorage.removeItem(KEY_SES); return null; } catch { return null; } };
  const saveSes = (s) => s ? localStorage.setItem(KEY_SES, JSON.stringify(s)) : localStorage.removeItem(KEY_SES);

  function addLog(logs, e) {
    const t = ahora();
    const nuevo = Object.assign({}, e, { id:uid(), timestamp:t.ts, fecha:t.fecha, hora:t.hora, dia:t.dia, ip:getIP(), dispositivo:getDisp() });
    return [nuevo, ...logs].slice(0, 500);
  }

  const accColor = (a) => ({
    LOGIN_OK:C.verde, LOGIN_FAIL:C.rojo, LOGOUT:C.cyan,
    CAMBIO_CLAVE:C.amar, SESION_EXPIRADA:"#c0392b",
  }[a] || C.muted);

  const badge = (txt, col) => `<span style="background:${col}22;color:${col};border:1px solid ${col}44;border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700;white-space:nowrap;">${txt}</span>`;

  // ---- init ----
  let SES = null, USERS = [], LOGS = [];

  document.addEventListener("DOMContentLoaded", init);
  if (document.readyState !== "loading") init();

  function init() {
    const vista = document.getElementById("vista-ciberseguridad");
    if (!vista || vista.dataset.ligado === "1") return;
    vista.dataset.ligado = "1";
    USERS = loadUsers(); LOGS = loadLogs(); SES = loadSes();
    render();
    // expiración
    setInterval(() => {
      if (SES && Date.now() > SES.expira) {
        LOGS = addLog(LOGS, { usuario:SES.usuario, nombre:SES.nombre, rol:SES.rol, accion:"SESION_EXPIRADA", exitoso:false, detalle:"Sesión expirada por inactividad" });
        saveLogs(LOGS); saveSes(null); SES = null; render();
      }
    }, 30000);
  }

  function render() {
    const vista = document.getElementById("vista-ciberseguridad");
    if (!vista) return;
    const prev = document.getElementById("cyber-root");
    if (prev) prev.remove();
    const root = document.createElement("div");
    root.id = "cyber-root";
    vista.appendChild(root);
    if (!SES) renderLogin(root);
    else renderPanel(root);
  }

  /* ---------------- LOGIN ---------------- */
  function renderLogin(root) {
    root.innerHTML = `
      <div style="min-height:70vh;display:flex;align-items:center;justify-content:center;padding:24px;">
        <div style="width:100%;max-width:420px;">
          <div style="text-align:center;margin-bottom:28px;">
            <div style="font-size:52px;margin-bottom:8px;">🔐</div>
            <div style="font-size:10px;color:${C.cyan};letter-spacing:4px;font-weight:700;text-transform:uppercase;">Control Político · XBOG Technologies</div>
            <h1 style="color:${C.navy};font-size:24px;font-weight:900;margin:8px 0 4px;">Acceso Seguro</h1>
            <div style="color:${C.muted};font-size:12px;">Fondo de Desarrollo Local - Sistema de Auditoría</div>
          </div>
          <div style="background:#fff;border-radius:16px;padding:32px;box-shadow:0 12px 32px rgba(26,39,66,.15);">
            <div style="margin-bottom:18px;">
              <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Usuario</label>
              <input id="cy-user" placeholder="Ingresa tu usuario" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:10px;padding:12px 16px;color:${C.texto};font-size:13px;outline:none;box-sizing:border-box;">
            </div>
            <div style="margin-bottom:22px;">
              <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Contraseña</label>
              <input id="cy-pass" type="password" placeholder="Ingresa tu contraseña" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:10px;padding:12px 16px;color:${C.texto};font-size:13px;outline:none;box-sizing:border-box;">
            </div>
            <div id="cy-error" style="display:none;background:${C.rojo}18;border:1px solid ${C.rojo}44;border-radius:8px;padding:10px 14px;margin-bottom:16px;font-size:12px;color:${C.rojo};font-weight:600;"></div>
            <button id="cy-btn-login" style="width:100%;background:linear-gradient(135deg,${C.navy},${C.navy2});border:none;border-radius:10px;padding:13px;color:#fff;font-weight:800;font-size:14px;cursor:pointer;">🔓 Ingresar al Sistema</button>
            <div style="margin-top:20px;padding:12px 14px;background:${C.bg};border-radius:8px;font-size:11px;color:${C.muted};">
              <div style="font-weight:700;margin-bottom:6px;">Roles disponibles:</div>
              <div>👤 <strong style="color:${C.cyan};">Edil</strong> — módulos de control político</div>
              <div style="margin-top:4px;">🛡️ <strong style="color:${C.amar};">Administrador</strong> — acceso total + gestión de usuarios</div>
            </div>
          </div>
          <div style="text-align:center;margin-top:18px;font-size:10px;color:${C.muted};">🔒 Sesión expira en ${SESION_MIN} min · XBOG Technologies © 2026</div>
        </div>
      </div>
    `;

    const errBox = root.querySelector("#cy-error");
    const showErr = (m) => { errBox.textContent = "⚠️ " + m; errBox.style.display = "block"; };

    function intentar() {
      const usuario = root.querySelector("#cy-user").value.trim();
      const clave = root.querySelector("#cy-pass").value;
      if (!usuario || !clave) return showErr("Completa todos los campos.");
      const user = USERS.find((u) => u.usuario.toLowerCase() === usuario.toLowerCase());
      if (!user || !user.activo) {
        LOGS = addLog(LOGS, { usuario, nombre:"DESCONOCIDO", rol:"DESCONOCIDO", accion:"LOGIN_FAIL", exitoso:false, detalle:"Usuario no encontrado: "+usuario });
        saveLogs(LOGS); return showErr("Usuario no encontrado o inactivo.");
      }
      if (user.bloqueado) return showErr("⛔ Cuenta bloqueada por intentos fallidos. Contacta al administrador.");
      if (user.clave !== clave) {
        const intentos = user.intentosFallidos + 1;
        const bloqueado = intentos >= MAX_INTENTOS;
        USERS = USERS.map((u) => u.id === user.id ? Object.assign({}, u, { intentosFallidos:intentos, bloqueado }) : u);
        saveUsers(USERS);
        LOGS = addLog(LOGS, { usuario:user.usuario, nombre:user.nombre, rol:user.rol, accion:"LOGIN_FAIL", exitoso:false, detalle:`Clave incorrecta. Intento ${intentos}/${MAX_INTENTOS}${bloqueado?" — BLOQUEADA":""}` });
        saveLogs(LOGS);
        return showErr(bloqueado ? "⛔ Cuenta bloqueada por 3 intentos fallidos." : `Clave incorrecta. Intento ${intentos}/${MAX_INTENTOS}.`);
      }
      const t = ahora();
      USERS = USERS.map((u) => u.id === user.id ? Object.assign({}, u, { intentosFallidos:0, bloqueado:false, ultimoAcceso:t.ts }) : u);
      saveUsers(USERS);
      LOGS = addLog(LOGS, { usuario:user.usuario, nombre:user.nombre, rol:user.rol, accion:"LOGIN_OK", exitoso:true, detalle:"Sesión iniciada correctamente" });
      saveLogs(LOGS);
      SES = { usuarioId:user.id, usuario:user.usuario, nombre:user.nombre, rol:user.rol, inicio:t.ts, expira:Date.now()+SESION_MIN*60*1000 };
      saveSes(SES);
      render();
    }

    root.querySelector("#cy-btn-login").addEventListener("click", intentar);
    root.querySelector("#cy-pass").addEventListener("keydown", (e) => { if (e.key === "Enter") intentar(); });
    root.querySelector("#cy-user").addEventListener("keydown", (e) => { if (e.key === "Enter") intentar(); });
  }

  /* ---------------- PANEL ---------------- */
  let TAB = "dashboard";

  function renderPanel(root) {
    const loginOk = LOGS.filter((l) => l.accion === "LOGIN_OK").length;
    const loginFail = LOGS.filter((l) => l.accion === "LOGIN_FAIL").length;
    const logouts = LOGS.filter((l) => l.accion === "LOGOUT").length;
    const hoy = ahora().fecha;
    const hoyCount = LOGS.filter((l) => l.fecha === hoy).length;
    const activos = USERS.filter((u) => u.activo).length;
    const bloqueados = USERS.filter((u) => u.bloqueado).length;

    root.innerHTML = `
      <div style="background:linear-gradient(135deg,${C.navy},${C.navy2});border-radius:12px;padding:18px 24px;margin-bottom:16px;color:#fff;">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;">
          <div>
            <div style="font-size:10px;color:${C.cyan};letter-spacing:3px;font-weight:700;text-transform:uppercase;">🛡️ Centro de Seguridad · XBOG</div>
            <h2 style="margin:6px 0 4px;font-size:20px;font-weight:900;">Módulo de Ciberseguridad</h2>
            <div style="font-size:11px;color:rgba(255,255,255,.7);">
              👤 ${SES.nombre} · <span style="color:${SES.rol==="admin"?C.amar:C.cyan};">${SES.rol==="admin"?"🛡️ Administrador":"👤 Edil"}</span> · Inició: ${new Date(SES.inicio).toLocaleTimeString("es-CO")}
            </div>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap;">
            <button id="cy-cambiar" style="background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);border-radius:8px;padding:7px 14px;color:${C.amar};cursor:pointer;font-size:12px;font-weight:600;">🔑 Cambiar clave</button>
            <button id="cy-logout" style="background:${C.rojo}33;border:1px solid ${C.rojo}66;border-radius:8px;padding:7px 14px;color:#fff;cursor:pointer;font-size:12px;font-weight:700;">🚪 Cerrar sesión</button>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-top:16px;">
          ${stat("✅","Accesos exitosos",loginOk,C.verde)}
          ${stat("❌","Intentos fallidos",loginFail,C.rojo)}
          ${stat("🚪","Cierres sesión",logouts,C.cyan)}
          ${stat("📅","Eventos hoy",hoyCount,C.amar)}
          ${stat("👥","Usuarios activos",activos,C.verde)}
          ${stat("⛔","Bloqueadas",bloqueados,"#c0392b")}
        </div>
      </div>

      <div style="background:#fff;border-radius:10px 10px 0 0;border-bottom:2px solid ${C.border};display:flex;padding:0 12px;overflow-x:auto;">
        ${tabBtn("dashboard","📊 Dashboard")}
        ${tabBtn("logs",`📋 Auditoría (${LOGS.length})`)}
        ${SES.rol==="admin"?tabBtn("usuarios","👥 Usuarios"):""}
        ${tabBtn("config","⚙️ Mi cuenta")}
      </div>
      <div id="cy-tabcontent" style="background:#fff;border-radius:0 0 10px 10px;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);"></div>
    `;

    root.querySelector("#cy-logout").addEventListener("click", () => {
      LOGS = addLog(LOGS, { usuario:SES.usuario, nombre:SES.nombre, rol:SES.rol, accion:"LOGOUT", exitoso:true, detalle:"Cierre de sesión manual" });
      saveLogs(LOGS); saveSes(null); SES = null; render();
    });
    root.querySelector("#cy-cambiar").addEventListener("click", modalClave);
    root.querySelectorAll("[data-tab]").forEach((b) =>
      b.addEventListener("click", () => { TAB = b.dataset.tab; PAGINA = 0; pintarTab(); })
    );
    pintarTab();
  }

  function stat(icon, label, valor, color) {
    return `<div style="background:rgba(255,255,255,.1);border-radius:10px;padding:12px 14px;">
      <div style="font-size:18px;">${icon}</div>
      <div style="font-size:22px;font-weight:800;color:${color};line-height:1;margin-top:4px;">${valor}</div>
      <div style="font-size:11px;color:rgba(255,255,255,.7);margin-top:3px;">${label}</div>
    </div>`;
  }
  function tabBtn(id, label) {
    const activo = TAB === id;
    return `<button data-tab="${id}" style="padding:12px 16px;background:transparent;border:none;border-bottom:${activo?`2px solid ${C.cyan}`:"2px solid transparent"};color:${activo?C.cyan:C.muted};cursor:pointer;font-size:12px;font-weight:${activo?700:500};white-space:nowrap;">${label}</button>`;
  }

  function pintarTab() {
    // refrescar barra de tabs (para el subrayado)
    const root = document.getElementById("cyber-root");
    root.querySelectorAll("[data-tab]").forEach((b) => {
      const activo = b.dataset.tab === TAB;
      b.style.borderBottom = activo ? `2px solid ${C.cyan}` : "2px solid transparent";
      b.style.color = activo ? C.cyan : C.muted;
      b.style.fontWeight = activo ? 700 : 500;
    });
    const cont = document.getElementById("cy-tabcontent");
    if (TAB === "dashboard") tabDashboard(cont);
    else if (TAB === "logs") tabLogs(cont);
    else if (TAB === "usuarios" && SES.rol === "admin") tabUsuarios(cont);
    else if (TAB === "config") tabConfig(cont);
  }

  /* ---- Dashboard ---- */
  function tabDashboard(cont) {
    const recientes = LOGS.slice(0, 10);
    cont.innerHTML = `<div style="font-weight:700;font-size:14px;color:${C.texto};margin-bottom:16px;">Actividad reciente — últimos 10 eventos</div>` +
      (recientes.length === 0
        ? `<div style="text-align:center;padding:40px;color:${C.muted};">No hay eventos registrados aún.</div>`
        : `<div style="display:flex;flex-direction:column;gap:8px;">` + recientes.map((l) => {
            const col = accColor(l.accion);
            const emoji = { LOGIN_OK:"✅", LOGIN_FAIL:"❌", LOGOUT:"🚪", CAMBIO_CLAVE:"🔑", SESION_EXPIRADA:"⏰" }[l.accion] || "📌";
            return `<div style="background:${C.bg};border-radius:10px;padding:12px 16px;border-left:3px solid ${col};display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
              <div style="display:flex;align-items:center;gap:12px;">
                <div style="width:36px;height:36px;border-radius:50%;background:${col}22;display:flex;align-items:center;justify-content:center;font-size:16px;">${emoji}</div>
                <div><div style="font-weight:700;font-size:13px;color:${C.texto};">${l.nombre}</div><div style="font-size:11px;color:${C.muted};">${l.detalle}</div></div>
              </div>
              <div style="text-align:right;font-size:11px;color:${C.muted};">
                <div style="color:${C.texto};font-weight:600;">${l.dia} ${l.fecha}</div><div>${l.hora}</div>
                <div style="margin-top:2px;">${badge(l.accion, col)}</div>
              </div>
            </div>`;
          }).join("") + `</div>`);
  }

  /* ---- Logs ---- */
  function tabLogs(cont) {
    const filt = LOGS.filter((l) => {
      const okA = F.acc === "TODOS" || l.accion === F.acc;
      const okU = F.user === "TODOS" || l.usuario === F.user;
      const okF = !F.fecha || l.fecha.includes(F.fecha.split("-").reverse().join("/"));
      const q = F.busca.toLowerCase();
      const okQ = !q || l.nombre.toLowerCase().includes(q) || l.usuario.toLowerCase().includes(q) || (l.detalle||"").toLowerCase().includes(q);
      return okA && okU && okF && okQ;
    });
    const paginas = Math.ceil(filt.length / FILAS) || 1;
    if (PAGINA >= paginas) PAGINA = paginas - 1;
    const pag = filt.slice(PAGINA*FILAS, (PAGINA+1)*FILAS);
    const usuariosUnicos = [...new Set(LOGS.map((l) => l.usuario))];

    cont.innerHTML = `
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;background:${C.bg};border-radius:10px;padding:14px 16px;">
        <input id="cl-busca" value="${F.busca}" placeholder="🔎 Buscar usuario, nombre, detalle..." style="flex:2;min-width:200px;border:1px solid ${C.border};border-radius:8px;padding:8px 12px;font-size:12px;">
        <select id="cl-acc" style="flex:1;min-width:150px;border:1px solid ${C.border};border-radius:8px;padding:8px 10px;font-size:12px;">
          <option value="TODOS">Todas las acciones</option>
          ${["LOGIN_OK","LOGIN_FAIL","LOGOUT","CAMBIO_CLAVE","SESION_EXPIRADA"].map((a)=>`<option value="${a}" ${F.acc===a?"selected":""}>${a}</option>`).join("")}
        </select>
        <select id="cl-user" style="flex:1;min-width:130px;border:1px solid ${C.border};border-radius:8px;padding:8px 10px;font-size:12px;">
          <option value="TODOS">Todos los usuarios</option>
          ${usuariosUnicos.map((u)=>`<option value="${u}" ${F.user===u?"selected":""}>${u}</option>`).join("")}
        </select>
        <input type="date" id="cl-fecha" value="${F.fecha}" style="border:1px solid ${C.border};border-radius:8px;padding:8px 10px;font-size:12px;">
        <button id="cl-limpiar" style="background:${C.rojo}18;border:1px solid ${C.rojo}44;border-radius:8px;padding:8px 14px;color:${C.rojo};cursor:pointer;font-size:12px;font-weight:700;">✕ Limpiar</button>
        ${SES.rol==="admin"?`<button id="cl-borrar" style="background:transparent;border:1px solid ${C.border};border-radius:8px;padding:8px 14px;color:${C.muted};cursor:pointer;font-size:12px;margin-left:auto;">🗑️ Limpiar logs</button>`:""}
      </div>
      <div style="font-size:11px;color:${C.muted};margin-bottom:10px;"><strong style="color:${C.texto};">${filt.length}</strong> eventos · Página ${PAGINA+1} de ${paginas}</div>
      <div style="overflow-x:auto;border-radius:10px;border:1px solid ${C.border};">
        <table style="width:100%;border-collapse:collapse;font-size:11px;">
          <thead><tr style="background:${C.navy};">
            ${["ID","Día","Fecha","Hora","Usuario","Nombre","Rol","Acción","Dispositivo","IP","Detalle","Estado"].map((h)=>`<th style="padding:10px;text-align:left;color:${C.cyan};font-weight:700;font-size:10px;white-space:nowrap;">${h}</th>`).join("")}
          </tr></thead>
          <tbody>
            ${pag.length===0
              ? `<tr><td colspan="12" style="text-align:center;padding:32px;color:${C.muted};">No hay eventos con los filtros aplicados</td></tr>`
              : pag.map((l,i)=>`<tr style="border-bottom:1px solid ${C.border};background:${i%2?C.bg:"#fff"};border-left:2px solid ${l.exitoso?C.verde+"66":C.rojo+"66"};">
                  <td style="padding:8px 10px;font-family:monospace;color:${C.muted};font-size:10px;">${l.id}</td>
                  <td style="padding:8px 10px;color:${C.muted};white-space:nowrap;">${l.dia}</td>
                  <td style="padding:8px 10px;white-space:nowrap;color:${C.texto};">${l.fecha}</td>
                  <td style="padding:8px 10px;white-space:nowrap;color:${C.cyan};font-family:monospace;">${l.hora}</td>
                  <td style="padding:8px 10px;font-weight:700;color:${C.amar};">${l.usuario}</td>
                  <td style="padding:8px 10px;color:${C.texto};max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${l.nombre}</td>
                  <td style="padding:8px 10px;">${badge(l.rol, l.rol==="admin"?C.amar:l.rol==="edil"?C.cyan:C.muted)}</td>
                  <td style="padding:8px 10px;">${badge(l.accion, accColor(l.accion))}</td>
                  <td style="padding:8px 10px;color:${C.muted};white-space:nowrap;">${l.dispositivo}</td>
                  <td style="padding:8px 10px;font-family:monospace;color:${C.muted};font-size:10px;">${l.ip}</td>
                  <td style="padding:8px 10px;color:${C.muted};max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${l.detalle}</td>
                  <td style="padding:8px 10px;">${badge(l.exitoso?"✓ OK":"✗ FAIL", l.exitoso?C.verde:C.rojo)}</td>
                </tr>`).join("")}
          </tbody>
        </table>
      </div>
      <div id="cl-pag" style="display:flex;justify-content:center;gap:6px;margin-top:14px;"></div>
    `;

    // listeners filtros
    const q = (id) => cont.querySelector(id);
    q("#cl-busca").addEventListener("input", (e) => { F.busca = e.target.value; PAGINA=0; tabLogs(cont); });
    q("#cl-acc").addEventListener("change", (e) => { F.acc = e.target.value; PAGINA=0; tabLogs(cont); });
    q("#cl-user").addEventListener("change", (e) => { F.user = e.target.value; PAGINA=0; tabLogs(cont); });
    q("#cl-fecha").addEventListener("change", (e) => { F.fecha = e.target.value; PAGINA=0; tabLogs(cont); });
    q("#cl-limpiar").addEventListener("click", () => { F.acc="TODOS"; F.user="TODOS"; F.fecha=""; F.busca=""; PAGINA=0; tabLogs(cont); });
    const borrar = q("#cl-borrar");
    if (borrar) borrar.addEventListener("click", () => {
      if (confirm("¿Eliminar todo el historial de logs? No se puede deshacer.")) { LOGS=[]; saveLogs(LOGS); tabLogs(cont); }
    });

    // paginación
    if (paginas > 1) {
      const cont2 = q("#cl-pag");
      let h = "";
      const b = (t, dest, dis) => `<button data-p="${dest}" ${dis?"disabled":""} style="padding:5px 10px;border-radius:6px;border:1px solid ${C.border};background:#fff;cursor:${dis?"default":"pointer"};opacity:${dis?.4:1};font-size:12px;">${t}</button>`;
      h += b("«",0,PAGINA===0) + b("‹",PAGINA-1,PAGINA===0);
      const ini = paginas<=7?0:Math.max(0,Math.min(PAGINA-3,paginas-7));
      for (let i=ini;i<Math.min(ini+7,paginas);i++)
        h += `<button data-p="${i}" style="padding:5px 10px;border-radius:6px;border:1px solid ${i===PAGINA?C.cyan:C.border};background:${i===PAGINA?C.cyan:"#fff"};color:${i===PAGINA?"#fff":C.texto};cursor:pointer;font-weight:${i===PAGINA?700:400};font-size:12px;">${i+1}</button>`;
      h += b("›",PAGINA+1,PAGINA>=paginas-1) + b("»",paginas-1,PAGINA>=paginas-1);
      cont2.innerHTML = h;
      cont2.querySelectorAll("button[data-p]").forEach((btn)=>btn.addEventListener("click",()=>{PAGINA=parseInt(btn.dataset.p);tabLogs(cont);}));
    }
  }

  /* ---- Usuarios (admin) ---- */
  function tabUsuarios(cont) {
    cont.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
        <div style="font-weight:700;font-size:14px;color:${C.texto};">Gestión de Usuarios</div>
        <button id="cu-nuevo" style="background:${C.cyan};border:none;border-radius:8px;padding:8px 16px;color:#fff;font-weight:700;font-size:12px;cursor:pointer;">+ Nuevo usuario</button>
      </div>
      <div style="display:flex;flex-direction:column;gap:10px;">
        ${USERS.map((u)=>`<div style="background:${C.bg};border-radius:10px;padding:14px 18px;border:1px solid ${u.bloqueado?C.rojo+"44":C.border};display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;">
          <div style="display:flex;align-items:center;gap:14px;">
            <div style="width:42px;height:42px;border-radius:50%;background:${u.rol==="admin"?C.amar+"22":C.cyan+"22"};display:flex;align-items:center;justify-content:center;font-size:20px;">${u.rol==="admin"?"🛡️":"👤"}</div>
            <div>
              <div style="font-weight:700;font-size:13px;color:${C.texto};">${u.nombre}</div>
              <div style="font-size:11px;color:${C.muted};">@${u.usuario} · ${badge(u.rol,u.rol==="admin"?C.amar:C.cyan)}</div>
              ${u.ultimoAcceso?`<div style="font-size:10px;color:${C.muted};margin-top:2px;">Último acceso: ${new Date(u.ultimoAcceso).toLocaleString("es-CO")}</div>`:""}
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
            ${u.intentosFallidos>0?badge(`⚠️ ${u.intentosFallidos} fallidos`,C.amar):""}
            ${badge(u.activo?"● Activo":"○ Inactivo",u.activo?C.verde:C.muted)}
            ${u.bloqueado?badge("⛔ Bloqueado",C.rojo):""}
            ${u.id!==SES.usuarioId?`
              <button data-act="${u.id}" style="background:${u.activo?C.rojo+"22":C.verde+"22"};border:1px solid ${u.activo?C.rojo:C.verde}44;border-radius:6px;padding:5px 12px;color:${u.activo?C.rojo:C.verde};cursor:pointer;font-size:11px;font-weight:700;">${u.activo?"Desactivar":"Activar"}</button>
              <button data-blq="${u.id}" style="background:${u.bloqueado?C.verde+"22":C.amar+"22"};border:1px solid ${u.bloqueado?C.verde:C.amar}44;border-radius:6px;padding:5px 12px;color:${u.bloqueado?C.verde:C.amar};cursor:pointer;font-size:11px;font-weight:700;">${u.bloqueado?"Desbloquear":"Bloquear"}</button>`
            :badge("Tu cuenta",C.cyan)}
          </div>
        </div>`).join("")}
      </div>
    `;
    cont.querySelector("#cu-nuevo").addEventListener("click", modalNuevoUsuario);
    cont.querySelectorAll("[data-act]").forEach((b)=>b.addEventListener("click",()=>{
      USERS = USERS.map((u)=>u.id===b.dataset.act?Object.assign({},u,{activo:!u.activo}):u); saveUsers(USERS); tabUsuarios(cont);
    }));
    cont.querySelectorAll("[data-blq]").forEach((b)=>b.addEventListener("click",()=>{
      USERS = USERS.map((u)=>u.id===b.dataset.blq?Object.assign({},u,{bloqueado:!u.bloqueado,intentosFallidos:0}):u); saveUsers(USERS); tabUsuarios(cont);
    }));
  }

  /* ---- Mi cuenta ---- */
  function tabConfig(cont) {
    cont.innerHTML = `
      <div style="max-width:480px;">
        <div style="font-weight:700;font-size:14px;color:${C.texto};margin-bottom:16px;">Mi cuenta</div>
        <div style="background:${C.bg};border-radius:12px;padding:20px;margin-bottom:20px;">
          <div style="display:flex;align-items:center;gap:14px;margin-bottom:16px;">
            <div style="width:52px;height:52px;border-radius:50%;background:${SES.rol==="admin"?C.amar+"22":C.cyan+"22"};display:flex;align-items:center;justify-content:center;font-size:26px;">${SES.rol==="admin"?"🛡️":"👤"}</div>
            <div>
              <div style="font-weight:800;font-size:16px;color:${C.texto};">${SES.nombre}</div>
              <div style="font-size:12px;color:${C.muted};">@${SES.usuario}</div>
              ${badge(SES.rol==="admin"?"Administrador":"Edil",SES.rol==="admin"?C.amar:C.cyan)}
            </div>
          </div>
          <div style="font-size:11px;color:${C.muted};border-top:1px solid ${C.border};padding-top:12px;">
            <div>🕐 Sesión iniciada: ${new Date(SES.inicio).toLocaleString("es-CO")}</div>
            <div style="margin-top:4px;">⏱️ Expira: ${new Date(SES.expira).toLocaleTimeString("es-CO")}</div>
            <div style="margin-top:4px;">💻 Dispositivo: ${getDisp()}</div>
          </div>
        </div>
        <button id="cf-cambiar" style="background:${C.amar}22;border:1px solid ${C.amar}44;border-radius:10px;padding:12px 20px;color:${C.amar};font-weight:700;font-size:13px;cursor:pointer;width:100%;">🔑 Cambiar contraseña</button>
      </div>
    `;
    cont.querySelector("#cf-cambiar").addEventListener("click", modalClave);
  }

  /* ---- Modales ---- */
  function overlay(inner) {
    const o = document.createElement("div");
    o.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.6);display:flex;align-items:center;justify-content:center;z-index:1000;padding:24px;";
    o.innerHTML = `<div style="background:#fff;border-radius:16px;padding:32px;width:100%;max-width:420px;box-shadow:0 24px 48px rgba(0,0,0,.3);">${inner}</div>`;
    document.body.appendChild(o);
    return o;
  }
  const inputHtml = (id,label,type,ph) => `<div style="margin-bottom:16px;">
    <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">${label}</label>
    <input id="${id}" type="${type||"text"}" placeholder="${ph||""}" style="width:100%;background:${C.bg};border:1.5px solid ${C.border};border-radius:8px;padding:10px 14px;color:${C.texto};font-size:12px;outline:none;box-sizing:border-box;">
  </div>`;

  function modalClave() {
    const o = overlay(`
      <div style="font-weight:800;font-size:16px;color:${C.texto};margin-bottom:20px;">🔑 Cambiar contraseña</div>
      ${inputHtml("mc-act","Contraseña actual","password")}
      ${inputHtml("mc-nue","Nueva contraseña","password","Mínimo 8 caracteres")}
      ${inputHtml("mc-con","Confirmar nueva","password")}
      <div id="mc-msg" style="display:none;padding:10px 14px;border-radius:8px;margin-bottom:16px;font-size:12px;font-weight:600;"></div>
      <div style="display:flex;gap:10px;">
        <button id="mc-cancel" style="flex:1;padding:10px;border-radius:8px;border:1px solid ${C.border};background:#fff;color:${C.muted};cursor:pointer;font-size:12px;">Cancelar</button>
        <button id="mc-ok" style="flex:1;padding:10px;border-radius:8px;border:none;background:${C.cyan};color:#fff;font-weight:700;cursor:pointer;font-size:12px;">Guardar</button>
      </div>
    `);
    const msg = o.querySelector("#mc-msg");
    const showMsg = (t, ok) => { msg.textContent = t; msg.style.display="block"; msg.style.background = ok?C.verde+"22":C.rojo+"22"; msg.style.color = ok?C.verde:C.rojo; };
    o.querySelector("#mc-cancel").addEventListener("click", () => o.remove());
    o.querySelector("#mc-ok").addEventListener("click", () => {
      const act = o.querySelector("#mc-act").value, nue = o.querySelector("#mc-nue").value, con = o.querySelector("#mc-con").value;
      const user = USERS.find((u) => u.id === SES.usuarioId);
      if (!user) return showMsg("❌ Usuario no encontrado.", false);
      if (user.clave !== act) return showMsg("❌ Contraseña actual incorrecta.", false);
      if (nue.length < 8) return showMsg("❌ Mínimo 8 caracteres.", false);
      if (nue !== con) return showMsg("❌ Las contraseñas no coinciden.", false);
      USERS = USERS.map((u) => u.id === SES.usuarioId ? Object.assign({}, u, { clave:nue }) : u);
      saveUsers(USERS);
      LOGS = addLog(LOGS, { usuario:SES.usuario, nombre:SES.nombre, rol:SES.rol, accion:"CAMBIO_CLAVE", exitoso:true, detalle:"Contraseña actualizada" });
      saveLogs(LOGS);
      showMsg("✅ Contraseña actualizada.", true);
      setTimeout(() => { o.remove(); render(); }, 1200);
    });
  }

  function modalNuevoUsuario() {
    const o = overlay(`
      <div style="font-weight:800;font-size:16px;color:${C.texto};margin-bottom:20px;">👤 Nuevo usuario</div>
      ${inputHtml("nu-nom","Nombre completo")}
      ${inputHtml("nu-usr","Usuario")}
      ${inputHtml("nu-cla","Contraseña","password","Mínimo 8 caracteres")}
      <div style="margin-bottom:16px;">
        <label style="display:block;font-size:11px;color:${C.muted};font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;">Rol</label>
        <select id="nu-rol" style="width:100%;background:${C.bg};border:1px solid ${C.border};border-radius:8px;padding:10px 14px;font-size:12px;">
          <option value="edil">👤 Edil</option><option value="admin">🛡️ Administrador</option>
        </select>
      </div>
      <div id="nu-msg" style="display:none;padding:10px 14px;border-radius:8px;margin-bottom:16px;font-size:12px;font-weight:600;"></div>
      <div style="display:flex;gap:10px;">
        <button id="nu-cancel" style="flex:1;padding:10px;border-radius:8px;border:1px solid ${C.border};background:#fff;color:${C.muted};cursor:pointer;font-size:12px;">Cancelar</button>
        <button id="nu-ok" style="flex:1;padding:10px;border-radius:8px;border:none;background:${C.cyan};color:#fff;font-weight:700;cursor:pointer;font-size:12px;">Crear usuario</button>
      </div>
    `);
    const msg = o.querySelector("#nu-msg");
    const showMsg = (t, ok) => { msg.textContent = t; msg.style.display="block"; msg.style.background = ok?C.verde+"22":C.rojo+"22"; msg.style.color = ok?C.verde:C.rojo; };
    o.querySelector("#nu-cancel").addEventListener("click", () => o.remove());
    o.querySelector("#nu-ok").addEventListener("click", () => {
      const nom = o.querySelector("#nu-nom").value.trim(), usr = o.querySelector("#nu-usr").value.trim(), cla = o.querySelector("#nu-cla").value, rol = o.querySelector("#nu-rol").value;
      if (!nom || !usr || !cla) return showMsg("❌ Completa todos los campos.", false);
      if (cla.length < 8) return showMsg("❌ Contraseña mínimo 8 caracteres.", false);
      if (USERS.find((u) => u.usuario === usr)) return showMsg("❌ Usuario ya existe.", false);
      USERS = [...USERS, { id:uid(), nombre:nom, usuario:usr, clave:cla, rol, activo:true, ultimoAcceso:"", intentosFallidos:0, bloqueado:false }];
      saveUsers(USERS);
      showMsg("✅ Usuario creado.", true);
      setTimeout(() => { o.remove(); render(); }, 1200);
    });
  }

  console.log("✅ Módulo 08 (Ciberseguridad) listo");
})();
