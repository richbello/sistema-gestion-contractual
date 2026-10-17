/* =====================================================================
   Clave de acceso al backend (X-API-Key)
   - Se pide una sola vez y se guarda en este navegador (localStorage).
   - Se agrega automáticamente a toda llamada fetch() hacia el backend.
   - Los enlaces de descarga del backend la llevan en ?k=.
   - Si el backend responde 401, se vuelve a pedir y se reintenta una vez.
   Debe cargarse justo después de config.js y antes de los módulos.
   ===================================================================== */
(function () {
  "use strict";

  const CLAVE_LS = "xbog_api_key";
  const BASE = (typeof window.BACKEND_URL === "string" ? window.BACKEND_URL : "").replace(/\/$/, "");
  const fetchOriginal = window.fetch.bind(window);

  function leerClave() {
    try { return localStorage.getItem(CLAVE_LS) || ""; } catch (e) { return window.__xbogClave || ""; }
  }
  function guardarClave(v) {
    window.__xbogClave = v;
    try { localStorage.setItem(CLAVE_LS, v); } catch (e) { /* modo privado */ }
  }
  function pedirClave(motivo) {
    const v = window.prompt((motivo ? motivo + "\n\n" : "") +
      "Clave de acceso al sistema (la entrega el administrador).\nSe guarda solo en este navegador.", "");
    if (v && v.trim()) { guardarClave(v.trim()); return v.trim(); }
    return "";
  }

  function esBackend(url) {
    if (!url) return false;
    if (BASE) return url.startsWith(BASE + "/api/");
    return url.startsWith("/api/") || url.startsWith(location.origin + "/api/");
  }

  function urlDe(input) {
    if (typeof input === "string") return input;
    if (input instanceof URL) return input.href;
    if (input && typeof input.url === "string") return input.url;
    return "";
  }

  function conClave(init, clave) {
    const nuevo = Object.assign({}, init || {});
    const h = new Headers((init && init.headers) || {});
    if (clave) h.set("X-API-Key", clave);
    nuevo.headers = h;
    return nuevo;
  }

  window.fetch = async function (input, init) {
    const url = urlDe(input);
    if (!esBackend(url) || url.endsWith("/api/salud")) return fetchOriginal(input, init);

    let resp = await fetchOriginal(input, conClave(init, leerClave()));
    if (resp.status === 401) {
      const nueva = pedirClave(leerClave() ? "La clave guardada no es válida." : "");
      if (nueva) resp = await fetchOriginal(input, conClave(init, nueva));
    }
    return resp;
  };

  // Enlaces de descarga <a href="https://backend/api/.../descargar/...">: agregar ?k=
  document.addEventListener("click", function (ev) {
    const a = ev.target && ev.target.closest ? ev.target.closest("a[href]") : null;
    if (!a || !esBackend(a.href)) return;
    const clave = leerClave() || pedirClave("");
    if (!clave) return;
    const u = new URL(a.href);
    u.searchParams.set("k", clave);
    a.href = u.href;
  }, true);

  // Permite cambiarla desde la consola: xbogCambiarClave()
  window.xbogCambiarClave = function () { return !!pedirClave(""); };
})();
