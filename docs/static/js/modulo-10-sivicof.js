/* ==========================================================================
   Módulo 10 — Reportes SIVICOF
   1) "Revisar cifras": sube Ejecución + Anulaciones CRP y muestra el CB-0104
      con los controles cruzados antes de generar.
   2) "Generar paquete": descarga SIVICOF <MES>.zip (Excel + 4 PDF + decreto).
   ========================================================================== */
(function () {
  const API_SIVICOF = `${API_BASE}/api/sivicof`;
  const CORREO_JEFE = "jorgea.gonzalez@gobiernobogota.gov.co";
  const $ = (id) => document.getElementById(id);
  if (!$("vista-sivicof")) return;

  configurarDropzoneUnico("drop-sivicof-ejecucion", "input-sivicof-ejecucion", "lista-sivicof-ejecucion");
  configurarDropzoneUnico("drop-sivicof-anulaciones", "input-sivicof-anulaciones", "lista-sivicof-anulaciones");
  configurarDropzoneUnico("drop-sivicof-decreto", "input-sivicof-decreto", "lista-sivicof-decreto");

  /* ---------- Valores por defecto: mes anterior y firma el día 9 ---------- */
  const hoy = new Date();
  const mesAnterior = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  const pad = (n) => String(n).padStart(2, "0");
  $("sivicof-mes").value = `${mesAnterior.getFullYear()}-${pad(mesAnterior.getMonth() + 1)}`;

  function actualizarPlazo() {
    const [a, m] = $("sivicof-mes").value.split("-").map(Number);
    if (!a || !m) return;
    const limite = new Date(a, m, 9);                 // día 9 del mes siguiente
    $("sivicof-plazo").textContent =
      `Corte ${pad(new Date(a, m, 0).getDate())}/${pad(m)}/${a} · plazo de envío ${pad(9)}/${pad(limite.getMonth() + 1)}/${limite.getFullYear()}`;
    if (!$("sivicof-firma").dataset.tocado) {
      const firma = hoy < limite && hoy.getMonth() === limite.getMonth() ? hoy : limite;
      $("sivicof-firma").value = `${firma.getFullYear()}-${pad(firma.getMonth() + 1)}-${pad(firma.getDate())}`;
    }
    invalidar();
  }
  $("sivicof-mes").addEventListener("change", actualizarPlazo);
  $("sivicof-firma").addEventListener("change", () => { $("sivicof-firma").dataset.tocado = "1"; invalidar(); });
  actualizarPlazo();

  /* ---------- Modificación presupuestal ---------- */
  function toggleDecreto() {
    const si = $("sivicof-modificacion").value === "SI";
    $("sivicof-bloque-decreto").hidden = !si;
    invalidar();
  }
  $("sivicof-modificacion").addEventListener("change", toggleDecreto);
  ["sivicof-num-decreto", "sivicof-fecha-decreto"].forEach((id) => $(id).addEventListener("input", invalidar));
  ["input-sivicof-ejecucion", "input-sivicof-anulaciones"].forEach((id) => $(id).addEventListener("change", invalidar));
  ["drop-sivicof-ejecucion", "drop-sivicof-anulaciones"].forEach((id) => $(id).addEventListener("drop", invalidar));
  toggleDecreto();

  function invalidar() {
    $("btn-sivicof-generar").disabled = true;
    $("resultados-sivicof").classList.remove("visible");
  }

  /* ---------- Formulario común ---------- */
  function armarFormData(paraGenerar) {
    const ejec = $("input-sivicof-ejecucion").files[0];
    const anul = $("input-sivicof-anulaciones").files[0];
    if (!ejec || !anul) throw new Error("Adjunte el reporte de Ejecución de Gastos y el de Anulaciones CRP.");
    const [anio, mes] = $("sivicof-mes").value.split("-");
    if (!anio) throw new Error("Seleccione el mes a reportar.");
    const hubo = $("sivicof-modificacion").value === "SI";

    const fd = new FormData();
    fd.append("anio", anio);
    fd.append("mes", String(Number(mes)));
    fd.append("fecha_firma", $("sivicof-firma").value);
    fd.append("hubo_modificacion", hubo ? "1" : "0");
    fd.append("numero_decreto", $("sivicof-num-decreto").value.trim());
    fd.append("fecha_decreto", $("sivicof-fecha-decreto").value);
    fd.append("ejecucion", ejec);
    fd.append("anulaciones", anul);

    if (paraGenerar) {
      if (!$("sivicof-firma").value) throw new Error("Indique la fecha de firma de los certificados.");
      if (hubo) {
        if (!$("sivicof-num-decreto").value.trim()) throw new Error("Indique el número del Decreto Local.");
        const dec = $("input-sivicof-decreto").files[0];
        if (!dec) throw new Error("Adjunte el PDF del decreto.");
        fd.append("decreto", dec);
      }
    }
    return fd;
  }

  const fmt = (n) => Number(n).toLocaleString("es-CO", { maximumFractionDigits: 0 });
  const pct = (n) => (Number(n) * 100).toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  /* ---------- 1) Revisar cifras ---------- */
  $("btn-sivicof-calcular").addEventListener("click", async () => {
    mostrarAlerta("alerta-sivicof", "");
    let fd;
    try { fd = armarFormData(false); } catch (e) { mostrarAlerta("alerta-sivicof", e.message); return; }

    const btn = $("btn-sivicof-calcular");
    btn.disabled = true;
    mostrarEstado("estado-sivicof", true);
    try {
      const resp = await fetch(`${API_SIVICOF}/calcular`, { method: "POST", body: fd });
      const data = await resp.json();
      if (!resp.ok || !data.ok) { mostrarAlerta("alerta-sivicof", data.mensaje || "No se pudo calcular."); return; }
      pintarResultado(data);
      $("btn-sivicof-generar").disabled = false;
    } catch (err) {
      mostrarAlerta("alerta-sivicof", "Error de conexión con el servidor.");
    } finally {
      btn.disabled = false;
      mostrarEstado("estado-sivicof", false);
    }
  });

  function pintarResultado(d) {
    const filas = d.filas.map((r) => {
      const okF = r.F === r.control_f_sap, okG = r.F + r.G === r.control_total_sap;
      return `<tr>
        <td><strong>${r.rubro}</strong><div style="font-size:11px;color:var(--tinta-500)">fila ${r.fila}</div></td>
        <td style="text-align:right">${fmt(r.E)}</td>
        <td style="text-align:right">${fmt(r.F)} <span class="pill-cobertura ${okF ? "pill-cubre" : "pill-sin-disp"}" title="CDP Mes en Ejecución: ${fmt(r.control_f_sap)}">${okF ? "✓" : "≠ SAP"}</span></td>
        <td style="text-align:right">${fmt(r.G)} <span class="pill-cobertura ${okG ? "pill-cubre" : "pill-sin-disp"}" title="F + G debería ser E − H en SAP: ${fmt(r.control_total_sap)}">${okG ? "✓" : "≠ SAP"}</span></td>
        <td style="text-align:right">${fmt(r.H)}</td>
        <td style="text-align:right;color:var(--tinta-500)">${pct(r.I)}</td>
        <td style="text-align:right">${fmt(r.J)}</td>
        <td style="text-align:right">${fmt(r.K)}</td>
        <td style="text-align:right;color:var(--tinta-500)">${pct(r.L)}</td>
        <td style="text-align:right;color:var(--tinta-500)">${fmt(r.M)}</td>
      </tr>`;
    }).join("");
    $("tabla-sivicof").innerHTML = filas;

    const m = $("metricas-sivicof");
    m.innerHTML = "";
    const tot = (k) => d.filas.reduce((s, r) => s + r[k], 0);
    m.appendChild(crearMetrica(d.fecha_corte.split("-").reverse().join("/"), "Fecha de corte (C5)"));
    m.appendChild(crearMetrica(fmt(tot("F")), "Anulaciones del mes", "ocre"));
    m.appendChild(crearMetrica(fmt(tot("J")), "Giro del mes", "verde"));
    m.appendChild(crearMetrica(d.alertas.length, "Controles a revisar", d.alertas.length ? "rojo" : "verde"));

    $("alertas-sivicof").innerHTML = d.alertas.length
      ? `<div class="alerta alerta-aviso visible" style="flex-direction:column;gap:4px;"><strong>Revise antes de enviar:</strong>${d.alertas.map((a) => `<div>• ${esc(a)}</div>`).join("")}</div>`
      : `<div class="alerta alerta-aviso visible" style="background:var(--verde-100);color:var(--verde-600);border-color:transparent;">Todos los controles cuadran con SAP.</div>`;

    $("textos-sivicof").innerHTML = `
      <div><span class="sivicof-etq">Fecha en los 4 certificados</span>Dado en Bogotá D.C. ${esc(d.texto_fecha_firma || "—")}.</div>
      <div><span class="sivicof-etq">CBN-1093</span>…del Fondo de Desarrollo Local de Usme ${esc(d.texto_1093 || "(falta el número del decreto)")}.</div>`;

    $("sivicof-nombre-zip").textContent = `${d.carpeta}.zip`;
    $("sivicof-para").value = CORREO_JEFE;
    $("sivicof-asunto").value = d.correo.asunto;
    $("sivicof-cuerpo").value = d.correo.cuerpo;
    $("sivicof-mailto").href = `mailto:${CORREO_JEFE}?subject=${encodeURIComponent(d.correo.asunto)}&body=${encodeURIComponent(d.correo.cuerpo)}`;
    $("resultados-sivicof").classList.add("visible");
  }

  /* ---------- 2) Generar paquete ---------- */
  $("btn-sivicof-generar").addEventListener("click", async () => {
    mostrarAlerta("alerta-sivicof", "");
    let fd;
    try { fd = armarFormData(true); } catch (e) { mostrarAlerta("alerta-sivicof", e.message); return; }

    const btn = $("btn-sivicof-generar");
    btn.disabled = true;
    mostrarEstado("estado-sivicof-generar", true);
    try {
      const resp = await fetch(`${API_SIVICOF}/generar`, { method: "POST", body: fd });
      if (!resp.ok) {
        let msg = "No se pudo generar el paquete.";
        try { msg = (await resp.json()).mensaje || msg; } catch (_) {}
        mostrarAlerta("alerta-sivicof", msg);
        return;
      }
      const blob = await resp.blob();
      const nombre = $("sivicof-nombre-zip").textContent || "SIVICOF.zip";
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = nombre; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
      $("sivicof-descargado").hidden = false;
    } catch (err) {
      mostrarAlerta("alerta-sivicof", "Error de conexión con el servidor.");
    } finally {
      btn.disabled = false;
      mostrarEstado("estado-sivicof-generar", false);
    }
  });

  /* ---------- Copiar textos del correo ---------- */
  document.querySelectorAll("[data-copiar]").forEach((b) => {
    b.addEventListener("click", async () => {
      const el = $(b.dataset.copiar);
      try { await navigator.clipboard.writeText(el.value); }
      catch (_) { el.select(); document.execCommand("copy"); }
      const t = b.textContent; b.textContent = "Copiado"; setTimeout(() => (b.textContent = t), 1200);
    });
  });
})();
