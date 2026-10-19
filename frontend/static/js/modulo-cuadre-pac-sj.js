/* ==========================================================================
   Módulo 1 · Cuadre PAC — Servicios Públicos y Personas Jurídicas
   Basado en el Cuadre PAC de personas naturales (main.js). El Reporte CRP
   se adjunta en cada cuadre; cada fila P40 se cruza por CRP + posición.
   Depende de: API_BASE, mostrarAlerta, mostrarEstado, crearMetrica,
   formatearMoneda, configurarDropzoneUnico (definidos en main.js).
   ========================================================================== */
(function () {
  const form = document.getElementById("form-pac-sj");
  if (!form) return;

  configurarDropzoneUnico("drop-plantilla-pac-sj", "input-plantilla-pac-sj", "lista-plantilla-pac-sj");
  configurarDropzoneUnico("drop-reporte-pac-sj", "input-reporte-pac-sj", "lista-reporte-pac-sj");
  configurarDropzoneUnico("drop-crp-pac-sj", "input-crp-pac-sj", "lista-crp-pac-sj");

  const esc = (v) => String(v ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  function clasePill(cob) {
    if (!cob) return "";
    if (cob.startsWith("OK") || cob.startsWith("CUBRE")) return "pill-cubre";
    if (cob.includes("NO CUBRE")) return "pill-no-cubre";
    if (cob.startsWith("SIN") || cob.startsWith("FALTA")) return "pill-sin-disp";
    return "";
  }

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    mostrarAlerta("alerta-pac-sj", "");
    mostrarAlerta("aviso-pac-sj", "");
    document.getElementById("resultados-pac-sj").classList.remove("visible");

    const plantilla = document.getElementById("input-plantilla-pac-sj").files[0];
    const reporte = document.getElementById("input-reporte-pac-sj").files[0];
    const crp = document.getElementById("input-crp-pac-sj").files[0];
    const mes = document.getElementById("input-mes-pac-sj").value.trim().toUpperCase();
    if (!plantilla || !reporte || !crp || !mes) {
      mostrarAlerta("alerta-pac-sj", "Debes adjuntar la plantilla, el reporte PAC y el Reporte CRP actualizado, e indicar el mes a evaluar.");
      return;
    }
    const fd = new FormData();
    fd.append("plantilla", plantilla);
    fd.append("pac", reporte);
    fd.append("crp", crp);
    fd.append("mes", mes);

    const btn = document.getElementById("btn-pac-sj");
    btn.disabled = true;
    mostrarEstado("estado-pac-sj", true);
    try {
      const resp = await fetch(`${API_BASE}/api/cuadre-pac-sj/procesar`, { method: "POST", body: fd });
      const data = await resp.json();
      if (!resp.ok || !data.ok) {
        mostrarAlerta("alerta-pac-sj", data.mensaje || "No se pudo realizar el cuadre PAC.");
        return;
      }

      const cont = document.getElementById("metricas-pac-sj");
      cont.innerHTML = "";
      cont.appendChild(crearMetrica(data.pagos, "Pagos"));
      cont.appendChild(crearMetrica(data.imputaciones, "CRP + posición (únicos)"));
      cont.appendChild(crearMetrica(formatearMoneda(data.total_p31), "Total pagos (P31)"));
      cont.appendChild(crearMetrica(formatearMoneda(data.total_importe), "Total imputado (P40)",
                                    data.total_importe === data.total_p31 ? "verde" : "rojo"));
      cont.appendChild(crearMetrica(data.cubre, "Rubros cubren en " + data.mes, "verde"));
      cont.appendChild(crearMetrica(data.no_cubre, "Rubros no cubren", data.no_cubre ? "rojo" : ""));
      cont.appendChild(crearMetrica(data.sin_disponibilidad, "Sin disponibilidad", data.sin_disponibilidad ? "ocre" : ""));

      const avisos = [`Hoja leída: "${data.hoja}".`];
      if (data.filas_ocultas > 0) avisos.push(`Esa hoja tiene ${data.filas_ocultas} fila(s) oculta(s) o filtrada(s) que SÍ se incluyeron en el cuadre; quita los filtros si no deben contar.`);
      if (!data.posicion_en_crp) avisos.push("El Reporte CRP no trae 'N° Posición CRP': el rubro se tomó solo por N° Interno CRP.");
      if (data.sin_match_en_pac > 0) avisos.push(`${data.sin_match_en_pac} rubro(s) no tienen coincidencia en el reporte PAC.`);
      if (data.sin_rubro && data.sin_rubro.length) {
        avisos.push(`${data.sin_rubro.length} CRP+posición no están en el Reporte CRP adjunto: ${data.sin_rubro.slice(0, 8).join(", ")}${data.sin_rubro.length > 8 ? "…" : ""}.`);
      }
      if (data.pagos_con_diferencia && data.pagos_con_diferencia.length) {
        avisos.push(`La suma de las P40 no coincide con el total P31 en el/los pago(s) N° ${data.pagos_con_diferencia.join(", ")}.`);
      }
      if (avisos.length) mostrarAlerta("aviso-pac-sj", avisos.join(" "));

      const tbody = document.getElementById("tabla-pac-sj-body");
      tbody.innerHTML = "";
      let tImp = 0, tDisp = 0, tDif = 0;
      data.tabla.forEach(f => {
        tImp += Number(f.Importe_Total) || 0; tDisp += Number(f.Disponibilidad_PAC) || 0; tDif += Number(f.Diferencia) || 0;
        const tr = document.createElement("tr");
        const dif = Number(f.Diferencia) < 0 ? ' style="color:#c0141c;font-weight:700;"' : "";
        tr.innerHTML = `
          <td>${esc(f.Rubro)}</td>
          <td>${esc(f.Fondos)}</td>
          <td>${formatearMoneda(f.Importe_Total)}</td>
          <td>${formatearMoneda(f.Disponibilidad_PAC)}</td>
          <td${dif}>${formatearMoneda(f.Diferencia)}</td>
          <td><span class="pill-cobertura ${clasePill(f.Cobertura)}">${esc(f.Cobertura)}</span></td>`;
        tbody.appendChild(tr);
      });
      const trT = document.createElement("tr");
      trT.style.fontWeight = "700";
      trT.innerHTML = `<td>TOTAL</td><td></td><td>${formatearMoneda(tImp)}</td><td>${formatearMoneda(tDisp)}</td>
        <td${tDif < 0 ? ' style="color:#c0141c;"' : ""}>${formatearMoneda(tDif)}</td><td></td>`;
      tbody.appendChild(trT);

      const tdet = document.getElementById("tabla-pac-sj-detalle");
      tdet.innerHTML = "";
      let tDet = 0;
      (data.detalle || []).forEach(f => {
        tDet += Number(f.importe) || 0;
        const val = f["Validación pago"];
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${esc(f["RP Doc Presupuestal"])}</td>
          <td>${esc(f["Posc Doc Pres"])}</td>
          <td>${esc(f.crp)}</td>
          <td>${formatearMoneda(f.importe)}${Number(f["Filas P40"]) > 1 ? ' <span style="font-family:var(--fuente-cuerpo);font-size:11px;color:var(--ocre-600)">(' + f["Filas P40"] + ' filas)</span>' : ""}</td>
          <td>${f.rubro ? esc(f.rubro) : '<span style="color:#c0141c">NO ESTÁ EN CRP</span>'}</td>
          <td>${esc(f.fondos)}</td>
          <td style="font-family:var(--fuente-cuerpo)">${esc(f["Pago N°"])} · ${esc(f["Referencia / servicio"])}</td>
          <td>${val === "OK" ? "✓" : '<span style="color:#c0141c;font-weight:700">' + esc(val) + "</span>"}</td>
          <td><span class="pill-cobertura ${clasePill(f["Cobertura rubro"])}">${esc(f["Cobertura rubro"])}</span></td>`;
        tdet.appendChild(tr);
      });
      const trD = document.createElement("tr");
      trD.style.fontWeight = "700";
      trD.innerHTML = `<td>TOTAL</td><td></td><td></td><td>${formatearMoneda(tDet)}</td><td colspan="5"></td>`;
      tdet.appendChild(trD);

      document.getElementById("link-descarga-pac-sj").href = API_BASE + data.descarga;
      document.getElementById("resultados-pac-sj").classList.add("visible");
    } catch (err) {
      mostrarAlerta("alerta-pac-sj", "Error de conexión con el servidor. Verifica que el backend esté en ejecución.");
    } finally {
      btn.disabled = false;
      mostrarEstado("estado-pac-sj", false);
    }
  });
})();
