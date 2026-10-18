/* =====================================================================
   MÓDULO 11 · CDP_CRP Vigencia — Elaboración de CRP desde memorandos
   ---------------------------------------------------------------------
   Flujo (todo en el navegador, ni el PDF ni el Excel salen del equipo):
     1) Lee uno o varios PDF de memorando "Solicitud expedición de CRP"
        → tabla CTO | CONTRATISTA | CC | RUBRO | No. CDP | VALOR CDP | VALOR ADICION
        → Radicado No. 2026 55 ........ (y fecha)
     2) Lee el Reporte CDP (BOGDATA) y cruza por No. CDP (col. H):
        No.Interno CDP (U) → reserva · No.Posición CDP → Pos
        Objeto → OBJETO · Fondo → FONDOS · Valor (Q) → IMPORTE
        Anulaciones → ANULACIONES · Cod.Concepto Gasto → POS PRES
     3) CÓDIGO: 145 = servicios profesionales / especializados
                148 = servicios técnicos / de apoyo a la gestión
     4) Genera el Excel MEMORANDO_SOLICITUD_CRP (hoja base1).
   Dependencias: SheetJS (XLSX) y ExcelJS ya cargados en index.html;
   PDF.js se carga bajo demanda desde cdnjs.
   ===================================================================== */
(function (root) {
  "use strict";

  /* ============================ NÚCLEO ============================ */

  function norm(t) {
    if (t === null || t === undefined) return "";
    return String(t).normalize("NFKD").replace(/[̀-ͯ]/g, "")
      .toLowerCase().replace(/[^a-z0-9]/g, "");
  }
  function sinTildesMayus(t) {
    return String(t || "").normalize("NFKD").replace(/[̀-ͯ]/g, "").toUpperCase();
  }
  // "15.998.528" | "$ 15,998,528" | 15998528 → 15998528
  function aNumero(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return v;
    var s = String(v).replace(/[$\s]/g, "");
    // decimales con coma al final (",00")
    s = s.replace(/,(\d{1,2})$/, "");
    var n = parseFloat(s.replace(/[.,]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  function claveCDP(v) {                      // "0000001586" | "1586" → "1586"
    var d = String(v === null || v === undefined ? "" : v).replace(/\D/g, "");
    return d.replace(/^0+/, "") || (d ? "0" : "");
  }

  /* ---------- Código 145 / 148 según el objeto ---------- */
  function clasificarCodigo(objeto) {
    var t = sinTildesMayus(objeto).replace(/\s+/g, "");
    if (!t) return "";
    // Mirar primero lo que sigue a "PRESTARSERVICIOS" (el tipo de servicio)
    var i = t.indexOf("SERVICIOS");
    var tramo = i >= 0 ? t.substr(i + 9, 45) : t.substr(0, 120);
    function decide(s) {
      if (/PROFESIONAL|ESPECIALIZAD/.test(s)) return "145";
      if (/APOYO|TECNIC|ASISTENCIAL|GESTION/.test(s)) return "148";
      return "";
    }
    return decide(tramo) || decide(t);
  }

  /* ---------- PDF: agrupar items en líneas ---------- */
  // items: [{str, x, y, w, h}] con y de arriba hacia abajo
  function tokens(items) {
    var out = [];
    items.forEach(function (it) {
      var s = String(it.str || "");
      if (!s.trim()) return;
      var partes = s.split(/(\s+)/), x = it.x;
      var cw = s.length ? (it.w || s.length * 5) / s.length : 5;
      partes.forEach(function (p) {
        if (p && !/^\s+$/.test(p)) out.push({ str: p, x: x, y: it.y, w: p.length * cw, h: it.h || 10 });
        x += p.length * cw;
      });
    });
    return out;
  }
  var RE_CTO = /^\d{1,4}-20\d{2}$/;
  var RE_NUM = /^\$?[\d][\d.,]*$/;
  var PALABRAS_ENCABEZADO = { CTO: 1, CONTRATISTA: 1, CC: 1, RUBRO: 1, NO: 1, "NO.": 1, CDP: 1, VALOR: 1, ADICION: 1, TOTAL: 1 };

  function extraerFilasPagina(items, gapPorDefecto) {
    var tk = tokens(items);
    var anclas = tk.filter(function (t) { return RE_CTO.test(t.str); })
                   .sort(function (a, b) { return a.y - b.y; });
    var filas = [];
    if (!anclas.length) return { filas: filas, gap: gapPorDefecto };

    var gaps = [];
    for (var g = 1; g < anclas.length; g++) gaps.push(anclas[g].y - anclas[g - 1].y);
    var gapTipico = gaps.length ? Math.min.apply(null, gaps) : (gapPorDefecto || 40);

    anclas.forEach(function (a, k) {
      var tol = Math.max(2.5, (a.h || 10) * 0.45);
      var arriba = k > 0 ? (anclas[k - 1].y + a.y) / 2 : a.y - gapTipico / 2;
      var abajo = k < anclas.length - 1 ? (anclas[k + 1].y + a.y) / 2 : a.y + gapTipico / 2;

      // Números de la misma línea, a la derecha del CTO
      var linea = tk.filter(function (t) { return Math.abs(t.y - a.y) <= tol && t.x > a.x; })
                    .sort(function (p, q) { return p.x - q.x; });
      var nums = linea.filter(function (t) { return RE_NUM.test(t.str); });
      if (nums.length < 4) return;              // no es una fila de la tabla
      var cc = nums[0], rubro = nums[1], cdp = nums[2], vcdp = nums[3], vad = nums[4] || nums[3];

      // Nombre del contratista: texto entre el CTO y la cédula dentro de la franja
      var limIzq = a.x + (a.w || 30) - 1, limDer = cc.x - 1;
      var nombre = tk.filter(function (t) {
        return t.y > arriba && t.y < abajo && t.x >= limIzq && t.x < limDer &&
               !RE_NUM.test(t.str) && !PALABRAS_ENCABEZADO[t.str.toUpperCase()];
      }).sort(function (p, q) { return Math.abs(p.y - q.y) > tol ? p.y - q.y : p.x - q.x; })
        .map(function (t) { return t.str; }).join(" ").replace(/\s+/g, " ").trim();

      filas.push({
        contrato: a.str,
        contratista: nombre,
        cc: cc.str.replace(/\D/g, ""),
        rubro: rubro.str.replace(/\D/g, ""),
        no_cdp: cdp.str.replace(/\D/g, ""),
        valor_cdp: aNumero(vcdp.str),
        valor_adicion: aNumero(vad.str)
      });
    });
    return { filas: filas, gap: gapTipico };
  }

  // paginas: [[items], [items], ...]
  function parsearMemorando(paginas, nombreArchivo) {
    var radicado = "", fecha = "", gap = 40, filas = [];
    var texto = paginas.map(function (items) {
      return items.map(function (i) { return i.str; }).join(" ");
    }).join("\n");

    var m = texto.match(/Radicado\s*No\.?\s*:?\s*\*?\s*(\d[\d\s]{10,20}\d)/i);
    if (m) radicado = m[1].replace(/\s/g, "");
    if (!/^20\d{2}55\d{6,10}$/.test(radicado)) {
      var m2 = texto.match(/\b(20\d{2}55\d{6,10})\b/);
      radicado = m2 ? m2[1] : (radicado || "");
    }
    var f = texto.match(/Fecha\s*:?\s*(\d{1,2}[-\/]\d{1,2}[-\/]20\d{2})/i);
    if (f) fecha = f[1];

    paginas.forEach(function (items) {
      var r = extraerFilasPagina(items, gap);
      gap = r.gap || gap;
      filas = filas.concat(r.filas);
    });
    filas.forEach(function (fl) { fl.radicado = radicado; fl.fecha_radicado = fecha; fl.archivo = nombreArchivo || ""; });
    return { radicado: radicado, fecha: fecha, filas: filas, archivo: nombreArchivo || "" };
  }

  /* ---------- Reporte CDP ---------- */
  var MAPEO_CDP = {
    no_cdp:       ["No. CDP", "Número de CDP", "No CDP"],
    objeto:       ["Objeto"],
    fondo:        ["Fondo", "Fondos"],
    valor:        ["Valor"],
    anulaciones:  ["Anulaciones"],
    reintegros:   ["Reintegros"],
    x_comprometer:["CDP X Comprometer"],
    interno:      ["No.Interno CDP", "N° Interno CDP", "No. Interno CDP"],
    posicion:     ["No.Posición CDP", "N° Posición CDP", "No. Posición CDP"],
    concepto:     ["Cod.Concepto Gasto"],
    cadena:       ["Cadena"],
    pep:          ["Elemento PEP"],
    id_responsable: ["ID Responsable"],
    id_solicitante: ["ID Solicitante"]
  };

  // aoa: matriz (primera fila = encabezados; si no, se busca la fila de encabezados)
  function indexarReporteCDP(aoa) {
    var hdr = -1;
    for (var r = 0; r < Math.min(aoa.length, 15); r++) {
      var fila = (aoa[r] || []).map(norm);
      if (fila.indexOf("nocdp") >= 0 || fila.indexOf("numerodecdp") >= 0) { hdr = r; break; }
    }
    if (hdr < 0) throw new Error("El reporte CDP no tiene la columna 'No. CDP'.");
    var idx = {};
    aoa[hdr].forEach(function (h, i) { var k = norm(h); if (idx[k] === undefined) idx[k] = i; });
    var col = {};
    Object.keys(MAPEO_CDP).forEach(function (c) {
      for (var k = 0; k < MAPEO_CDP[c].length; k++) {
        var i = idx[norm(MAPEO_CDP[c][k])];
        if (i !== undefined) { col[c] = i; break; }
      }
    });
    ["interno", "posicion", "objeto", "valor"].forEach(function (c) {
      if (col[c] === undefined) throw new Error("El reporte CDP no tiene la columna '" + MAPEO_CDP[c][0] + "'.");
    });
    var mapa = {}, total = 0;
    for (var j = hdr + 1; j < aoa.length; j++) {
      var row = aoa[j];
      if (!row) continue;
      var k2 = claveCDP(row[col.no_cdp]);
      if (!k2) continue;
      total++;
      var reg = {};
      Object.keys(col).forEach(function (c) { reg[c] = row[col[c]]; });
      (mapa[k2] = mapa[k2] || []).push(reg);
    }
    return { mapa: mapa, total: total, unicos: Object.keys(mapa).length };
  }

  function cruzar(filasPdf, indice) {
    return filasPdf.map(function (f) {
      var obs = [], cand = indice.mapa[claveCDP(f.no_cdp)] || [];
      var reg = null;
      if (!cand.length) obs.push("CDP " + f.no_cdp + " no está en el reporte CDP");
      else {
        reg = cand.filter(function (c) { return aNumero(c.valor) === f.valor_cdp; })[0];
        if (!reg && cand.length > 1) {
          reg = cand.filter(function (c) { return f.rubro && String(c.cadena || "").indexOf(f.rubro) >= 0; })[0];
        }
        if (!reg) reg = cand[0];
        if (cand.length > 1) obs.push("CDP con " + cand.length + " posiciones (se tomó la pos. " + reg.posicion + ")");
      }
      var valor = reg ? aNumero(reg.valor) : 0;
      var anul = reg ? aNumero(reg.anulaciones) : 0;
      var objeto = reg ? String(reg.objeto || "").trim() : "";
      if (reg) {
        if (f.valor_cdp && valor !== f.valor_cdp) obs.push("Valor CDP del memorando (" + f.valor_cdp.toLocaleString("es-CO") + ") ≠ reporte (" + valor.toLocaleString("es-CO") + ")");
        if (anul > 0) obs.push("CDP con anulaciones por " + anul.toLocaleString("es-CO"));
        if (f.valor_adicion > valor - anul) obs.push("Valor adición supera el saldo del CDP");
        if (f.rubro && reg.cadena && String(reg.cadena).indexOf(f.rubro) < 0) obs.push("Rubro " + f.rubro + " no coincide con la cadena del CDP");
      }
      var codigo = clasificarCodigo(objeto);
      if (!codigo) obs.push("No se pudo determinar el código 145/148");
      return {
        contrato: f.contrato, contratista: f.contratista, cc: f.cc, rubro: f.rubro,
        no_cdp: f.no_cdp, valor_cdp: f.valor_cdp, valor_adicion: f.valor_adicion,
        importe: reg ? valor : "", anulaciones: reg ? anul : "",
        reserva: reg ? String(reg.interno || "") : "",
        pos: reg ? reg.posicion : "",
        objeto: objeto,
        pos_pres: reg ? (reg.concepto === undefined ? "" : reg.concepto) : "",
        fondos: reg ? String(reg.fondo || "") : "",
        radicado: f.radicado || "",
        id_responsable: reg ? String(reg.id_responsable || "") : "",
        id_solicitante: reg ? String(reg.id_solicitante || "") : "",
        codigo: codigo,
        observaciones: obs.join(" · "),
        archivo: f.archivo || ""
      };
    });
  }

  var COLUMNAS = [
    { h: "CONTRATO",      k: "contrato",      w: 11, c: "FFFF0000", fc: "FFFFFFFF" },
    { h: "CONTRATISTA",   k: "contratista",   w: 22 },
    { h: "CC",            k: "cc",            w: 12 },
    { h: "RUBRO",         k: "rubro",         w: 8 },
    { h: "No. CDP",       k: "no_cdp",        w: 9 },
    { h: "VALOR CDP",     k: "valor_cdp",     w: 14, num: true },
    { h: "VALOR ADICION", k: "valor_adicion", w: 14, num: true },
    { h: "IMPORTE",       k: "importe",       w: 14, num: true, c: "FFC4D79B", cuerpo: "FFC4D79B" },
    { h: "ANULACIONES",   k: "anulaciones",   w: 13, num: true, c: "FFC4D79B" },
    { h: "reserva",       k: "reserva",       w: 14, c: "FFC4D79B", cuerpo: "FFC4D79B" },
    { h: "Pos",           k: "pos",           w: 6,  c: "FFC4D79B" },
    { h: "OBJETO",        k: "objeto",        w: 45, c: "FFFFFF00", wrap: true },
    { h: "POS PRES",      k: "pos_pres",      w: 9,  c: "FF92D050" },
    { h: "FONDOS",        k: "fondos",        w: 12, c: "FF92D050" },
    { h: "RADICADO",      k: "radicado",      w: 17, c: "FF92D050" },
    { h: "OBJETO",        k: "objeto",        w: 45, c: "FF92D050", wrap: true },
    { h: "CÓDIGO",        k: "codigo",        w: 9,  c: "FFFFFF00" },
    { h: "OBSERVACIONES", k: "observaciones", w: 40, wrap: true }
  ];

  async function generarExcel(ExcelJS, filas, opciones) {
    var wb = new ExcelJS.Workbook();
    wb.creator = "XBOG SMART · Módulo 11";
    var ws = wb.addWorksheet("base1", { views: [{ state: "frozen", ySplit: 1 }] });
    ws.columns = COLUMNAS.map(function (c) { return { header: c.h, width: c.w }; });
    var borde = { style: "thin", color: { argb: "FF999999" } };
    var bordes = { top: borde, left: borde, bottom: borde, right: borde };
    ws.getRow(1).eachCell(function (cell, n) {
      var c = COLUMNAS[n - 1];
      cell.font = { bold: true, size: 9, color: { argb: c.fc || "FF000000" } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = bordes;
      if (c.c) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.c } };
    });
    filas.forEach(function (f) {
      var row = ws.addRow(COLUMNAS.map(function (c) {
        var v = f[c.k];
        if (c.k === "pos" || c.k === "pos_pres") return v === "" || v === undefined ? "" : (isNaN(Number(v)) ? v : Number(v));
        if (c.num) return v === "" ? "" : Number(v);
        return v === undefined ? "" : v;
      }));
      row.eachCell({ includeEmpty: true }, function (cell, n) {
        var c = COLUMNAS[n - 1];
        cell.font = { size: 9, color: { argb: c.k === "observaciones" ? "FFC00000" : "FF000000" } };
        cell.border = bordes;
        cell.alignment = { vertical: "middle", horizontal: c.wrap || c.k === "contratista" || c.k === "observaciones" ? "left" : "center", wrapText: !!(c.wrap || c.k === "contratista" || c.k === "observaciones") };
        if (c.num) { cell.numFmt = "#,##0"; cell.alignment.horizontal = "right"; }
        if (c.cuerpo) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.cuerpo } };
        if (c.k === "reserva" || c.k === "radicado" || c.k === "cc") cell.numFmt = "@";
      });
    });
    // Totales
    var n = filas.length;
    if (n) {
      var tot = ws.addRow([]);
      tot.getCell(1).value = "TOTAL";
      ["valor_cdp", "valor_adicion", "importe", "anulaciones"].forEach(function (k) {
        var i = COLUMNAS.findIndex(function (c) { return c.k === k; }) + 1;
        var L = ws.getColumn(i).letter;
        tot.getCell(i).value = { formula: "SUM(" + L + "2:" + L + (n + 1) + ")" };
        tot.getCell(i).numFmt = "#,##0";
      });
      tot.eachCell(function (cell) { cell.font = { bold: true, size: 9 }; });
    }
    ws.autoFilter = { from: "A1", to: { row: 1, column: COLUMNAS.length } };
    hojaDatosSAP(wb, filas, opciones || {});
    return wb.xlsx.writeBuffer();
  }

  function dosDig(n) { return ("0" + n).slice(-2); }
  // Hoja con el mismo formato de la plantilla de la macro LlenarSAP (copiar A2:L y pegar en DATOS)
  function hojaDatosSAP(wb, filas, op) {
    var hoy = op.hoy || new Date();
    var fIni = dosDig(hoy.getDate()) + "." + dosDig(hoy.getMonth() + 1) + "." + hoy.getFullYear();
    var fFin = "31.12." + hoy.getFullYear();
    var ws = wb.addWorksheet("DATOS_SAP", { views: [{ state: "frozen", ySplit: 1 }] });
    var cab = ["CONTRATO", "Tipo de compromiso", "Número de Compromiso", "Fecha Inicial", "Fecha Final", "Tipo de pago",
               "Mod. de selección", "BP Beneficiario", "Tipo Doc.", "Identif.", "ID Responsable", "ID Solicitante", "ESTADO"];
    var anchos = [12, 18, 20, 13, 13, 13, 16, 16, 10, 15, 16, 16, 22];
    ws.columns = cab.map(function (h, i) { return { header: h, width: anchos[i] }; });
    ws.getRow(1).eachCell(function (c) {
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3864" } };
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
    filas.forEach(function (f) {
      var row = ws.addRow([f.contrato, f.codigo || "", f.contrato, fIni, fFin, op.tipoPago || "02", op.modSeleccion || "10",
                           "", "CC", f.cc || "", f.id_responsable || "", f.id_solicitante || "", ""]);
      row.eachCell({ includeEmpty: true }, function (c) { c.numFmt = "@"; });
    });
  }

  /* ================= ELABORACIÓN DE CDP (adiciones y prórrogas) ================= */

  // Rectángulos/líneas dibujados en la página (bordes de tablas) desde el operator list de pdf.js
  function bordesDesdeOperadores(ol, OPS, H) {
    function mul(m, n) { return [m[0]*n[0]+m[2]*n[1], m[1]*n[0]+m[3]*n[1], m[0]*n[2]+m[2]*n[3], m[1]*n[2]+m[3]*n[3], m[0]*n[4]+m[2]*n[5]+m[4], m[1]*n[4]+m[3]*n[5]+m[5]]; }
    function ap(m, x, y) { return [m[0]*x + m[2]*y + m[4], m[1]*x + m[3]*y + m[5]]; }
    function caja(p, q) { return { x0: Math.min(p[0], q[0]), x1: Math.max(p[0], q[0]), y0: H - Math.max(p[1], q[1]), y1: H - Math.min(p[1], q[1]) }; }
    var ctm = [1, 0, 0, 1, 0, 0], pila = [], rects = [];
    for (var i = 0; i < ol.fnArray.length; i++) {
      var fn = ol.fnArray[i], a = ol.argsArray[i];
      if (fn === OPS.save) pila.push(ctm.slice());
      else if (fn === OPS.restore) ctm = pila.pop() || ctm;
      else if (fn === OPS.transform) ctm = mul(ctm, a);
      else if (fn === OPS.constructPath) {
        var ops = a[0], c = a[1], k = 0, cur = null;
        for (var j = 0; j < ops.length; j++) {
          var op = ops[j];
          if (op === OPS.rectangle) { var x = c[k], y = c[k+1], w = c[k+2], h = c[k+3]; k += 4; rects.push(caja(ap(ctm, x, y), ap(ctm, x + w, y + h))); }
          else if (op === OPS.moveTo) { cur = ap(ctm, c[k], c[k+1]); k += 2; }
          else if (op === OPS.lineTo) { var p = ap(ctm, c[k], c[k+1]); k += 2; if (cur) rects.push(caja(cur, p)); cur = p; }
          else if (op === OPS.curveTo) k += 6;
          else if (op === OPS.curveTo2 || op === OPS.curveTo3) k += 4;
        }
      }
    }
    return rects;
  }

  function agrupar(vals, tol) {                 // agrupa números cercanos
    var s = vals.slice().sort(function (a, b) { return a - b; }), out = [];
    s.forEach(function (v) {
      var u = out[out.length - 1];
      if (u && v - u.max <= tol) { u.max = v; u.sum += v; u.n++; } else out.push({ min: v, max: v, sum: v, n: 1 });
    });
    return out.map(function (g) { return g.sum / g.n; });
  }

  // Une las líneas de una celda: "recreo-"+"deportivas" → "recreodeportivas"; "CPS-"+"049-2026" → "CPS-049-2026"
  function unirLineas(lineas) {
    var t = "";
    lineas.forEach(function (l) {
      l = l.trim(); if (!l) return;
      if (!t) { t = l; return; }
      if (/-$/.test(t)) {
        if (/^\d/.test(l) || /\d-$/.test(t) || /[A-Z]{2,}-$/.test(t) && /^\d/.test(l)) t += l;
        else t = t.replace(/-$/, "") + l;
      } else t += " " + l;
    });
    return t.replace(/\s+/g, " ").trim();
  }

  // Lee una tabla con bordes: devuelve filas de celdas (texto por columna)
  function tablaPorBordes(items, rects) {
    var tk = tokens(items);
    var hs = rects.filter(function (r) { return r.y1 - r.y0 < 2.5 && r.x1 - r.x0 > 15; });
    var vs = rects.filter(function (r) { return r.x1 - r.x0 < 2.5 && r.y1 - r.y0 > 4; });
    if (hs.length < 2 || vs.length < 2) return null;
    // columnas: líneas verticales con cobertura suficiente
    var gx = {};
    vs.forEach(function (r) { var k = Math.round(((r.x0 + r.x1) / 2) / 2); gx[k] = (gx[k] || 0) + (r.y1 - r.y0); });
    var maxCob = 0; Object.keys(gx).forEach(function (k) { maxCob = Math.max(maxCob, gx[k]); });
    var xs = agrupar(Object.keys(gx).filter(function (k) { return gx[k] >= maxCob * 0.6; }).map(function (k) { return k * 2; }), 3);
    var ys = agrupar(hs.map(function (r) { return (r.y0 + r.y1) / 2; }), 2.5);
    if (xs.length < 2 || ys.length < 2) return null;
    var x0 = xs[0], x1 = xs[xs.length - 1], filas = [];
    for (var i = 0; i < ys.length - 1; i++) {
      var top = ys[i], bot = ys[i + 1];
      if (bot - top < 6) continue;
      var dentro = tk.filter(function (t) { var cy = t.y - (t.h || 10) * 0.35, cx = t.x + t.w / 2; return cy > top && cy < bot && cx > x0 && cx < x1; });
      if (!dentro.length) continue;
      var celdas = [];
      for (var c = 0; c < xs.length - 1; c++) {
        var a = xs[c], b = xs[c + 1];
        var ts = dentro.filter(function (t) { var cx = t.x + Math.min(t.w, 8) / 2; return cx >= a && cx < b; })
                       .sort(function (p, q) { return Math.abs(p.y - q.y) > 2 ? p.y - q.y : p.x - q.x; });
        var lineas = [], yAct = null, prev = null;
        ts.forEach(function (t) {
          if (yAct === null || Math.abs(t.y - yAct) > 2) { lineas.push(t.str); yAct = t.y; }
          else {
            var hueco = t.x - (prev.x + prev.w);           // sin espacio real entre trozos → pegar
            lineas[lineas.length - 1] += (hueco < Math.max(0.9, (t.h || 10) * 0.12) ? "" : " ") + t.str;
          }
          prev = t;
        });
        celdas.push(unirLineas(lineas));
      }
      filas.push({ celdas: celdas, xs: xs });
    }
    return filas;
  }

  function columnaDe(h) {
    var n = norm(h);
    if (/proye/.test(n)) return "proyecto";
    if (/objeto/.test(n)) return "objeto";
    if (/contratista/.test(n)) return "contratista";
    if (/valor|adicion/.test(n)) return "valor_adicion";
    if (/meta/.test(n)) return "meta";
    if (/^con|contrat|cto/.test(n)) return "contrato";
    return "";
  }

  // paginas: [{items:[...], rects:[...]}]
  function parsearSolicitudCDP(paginas, nombreArchivo) {
    var texto = paginas.map(function (p) { return p.items.map(function (i) { return i.str; }).join(" "); }).join("\n");
    var radicado = "", fecha = "";
    var m = texto.match(/\b(20\d{2}55\d{6,10})\b/); if (m) radicado = m[1];
    var f = texto.match(/Fecha\s*:?\s*(\d{1,2}[-\/]\d{1,2}[-\/]20\d{2})/i); if (f) fecha = f[1];
    var mapa = null, filas = [], sinBordes = false;
    paginas.forEach(function (p) {
      var tabla = tablaPorBordes(p.items, p.rects || []);
      if (!tabla) { sinBordes = true; return; }
      tabla.forEach(function (fl) {
        var txt = norm(fl.celdas.join(" "));
        if (/objeto/.test(txt) && /contratista/.test(txt) && !/\d{3,}-?20\d{2}/.test(fl.celdas.join(" "))) {
          mapa = fl.celdas.map(columnaDe); return;      // fila de encabezado
        }
        if (!mapa) return;
        var r = {};
        mapa.forEach(function (k, i) { if (k) r[k] = fl.celdas[i] || ""; });
        var cto = String(r.contrato || "").replace(/\s/g, "").match(/(\d{1,4})-?(20\d{2})/);
        if (!cto) return;
        filas.push({
          proyecto: String(r.proyecto || "").replace(/\D/g, ""),
          meta: r.meta || "",
          contrato: ("000" + cto[1]).slice(-3) + "-" + cto[2],
          contratista: r.contratista || "",
          valor_adicion: aNumero(String(r.valor_adicion || "").match(/[\d.,]+/) ? String(r.valor_adicion).match(/[\d][\d.,]*/)[0] : 0),
          objeto: r.objeto || "",
          radicado: radicado, fecha_radicado: fecha, archivo: nombreArchivo || ""
        });
      });
    });
    return { radicado: radicado, fecha: fecha, filas: filas, archivo: nombreArchivo || "", sinBordes: sinBordes && !filas.length };
  }

  // Cruce por No. de contrato contra el Objeto del reporte CDP
  function cruzarPorContrato(filas, indice) {
    var todos = [];
    Object.keys(indice.mapa).forEach(function (k) { indice.mapa[k].forEach(function (reg) { reg._cdp = k; todos.push(reg); }); });
    return filas.map(function (f) {
      var obs = [], partes = f.contrato.split("-");
      var re = new RegExp("(^|\\D)0*" + parseInt(partes[0], 10) + "\\s*-\\s*" + partes[1] + "(\\D|$)");
      var cand = todos.filter(function (reg) { return re.test(String(reg.objeto || "")); });
      var reg = null;
      if (!cand.length) obs.push("Contrato " + f.contrato + " no aparece en el Objeto de ningún CDP del reporte");
      else {
        cand.forEach(function (c) {
          var o = sinTildesMayus(c.objeto);
          c._score = (aNumero(c.valor) === f.valor_adicion ? 2 : 0) + (/ADICION|PRORROGA/.test(o) ? 1 : 0) + (f.proyecto && String(c.pep || c.cadena || "").indexOf(f.proyecto) >= 0 ? 1 : 0);
        });
        cand.sort(function (a, b) { return b._score - a._score || Number(b._cdp) - Number(a._cdp); });
        reg = cand[0];
        var empatados = cand.filter(function (c) { return c._score === reg._score; });
        if (empatados.length > 1) obs.push("Hay " + empatados.length + " CDP para este contrato y valor (" + empatados.map(function (c) { return c._cdp; }).join(", ") + "); se tomó el " + reg._cdp);
        if (aNumero(reg.valor) !== f.valor_adicion) obs.push("Valor del CDP " + reg._cdp + " (" + aNumero(reg.valor).toLocaleString("es-CO") + ") ≠ valor adición");
      }
      var pep = reg ? String(reg.pep || "") : "";
      if (!pep && f.proyecto) pep = "PM/0005/0101/4599000" + f.proyecto;
      if (pep && f.proyecto && pep.slice(-4) !== f.proyecto) obs.push("El Elemento PEP no corresponde al proyecto " + f.proyecto);
      var codigo = clasificarCodigo(f.objeto);
      if (!codigo) obs.push("No se pudo determinar el código 145/148");
      return {
        proyecto: f.proyecto, meta: f.meta, contrato: f.contrato, contratista: f.contratista,
        valor_adicion: f.valor_adicion, objeto: f.objeto, codigo: codigo,
        importe: f.valor_adicion,
        pos_pres: reg && reg.concepto !== undefined ? reg.concepto : "",
        fondos: reg ? String(reg.fondo || "") : "",
        pep: pep, radicado: f.radicado, cdp_existente: reg ? reg._cdp : "",
        observaciones: obs.join(" · ")
      };
    });
  }

  async function generarExcelCDP(ExcelJS, filas) {
    var wb = new ExcelJS.Workbook();
    wb.creator = "XBOG SMART · Módulo 11";
    var ws = wb.addWorksheet("CDP Prorroga Adicion", { views: [{ state: "frozen", ySplit: 1 }] });
    var cols = [
      ["N° PROYECTO", "proyecto", 12], ["META", "meta", 30], ["CONTRATO", "contrato", 12], ["CONTRATISTA", "contratista", 30],
      ["VALOR ADICION", "valor_adicion", 15], ["OBJETO", "objeto", 60], ["CÓDIGO", "codigo", 10], [null, null, 4],
      ["IMPORTE", "importe", 12], ["POS PRES", "pos_pres", 10], ["FONDOS", "fondos", 12], ["ELEMENTO PEP", "pep", 28],
      ["RADICADO", "radicado", 16], ["OBSERVACIONES", "observaciones", 40]
    ];
    ws.columns = cols.map(function (c) { return { header: c[0] || "", width: c[2] }; });
    ws.getRow(1).eachCell(function (cell) { cell.font = { bold: true }; cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }; });
    filas.forEach(function (f) {
      var row = ws.addRow(cols.map(function (c) {
        if (!c[1]) return null;
        var v = f[c[1]];
        if (c[1] === "valor_adicion") return Number(v || 0).toLocaleString("es-CO");     // como en el formato: "16.974.048"
        if (c[1] === "importe") return v === "" ? "" : String(v);
        if (c[1] === "pos_pres" || c[1] === "codigo") return v === "" || v === undefined ? "" : (isNaN(Number(v)) ? v : Number(v));
        return v === undefined ? "" : v;
      }));
      row.eachCell(function (cell, n) {
        var k = cols[n - 1][1];
        cell.alignment = { vertical: "top", wrapText: k === "meta" || k === "objeto" || k === "contratista" || k === "observaciones" };
        if (k === "observaciones") cell.font = { color: { argb: "FFC00000" } };
      });
    });
    return wb.xlsx.writeBuffer();
  }

  var Core = { norm: norm, aNumero: aNumero, claveCDP: claveCDP, clasificarCodigo: clasificarCodigo,
               parsearMemorando: parsearMemorando, indexarReporteCDP: indexarReporteCDP,
               cruzar: cruzar, generarExcel: generarExcel, COLUMNAS: COLUMNAS,
               bordesDesdeOperadores: bordesDesdeOperadores, tablaPorBordes: tablaPorBordes, parsearSolicitudCDP: parsearSolicitudCDP,
               cruzarPorContrato: cruzarPorContrato, generarExcelCDP: generarExcelCDP };
  root.M11Core = Core;
  if (typeof module !== "undefined" && module.exports) { module.exports = Core; return; }

  /* ============================ INTERFAZ ============================ */

  var PDFJS_URL = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
  var PDFJS_WORKER = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  var pdfjsPromesa = null;
  function cargarPdfJs() {
    if (root.pdfjsLib) return Promise.resolve(root.pdfjsLib);
    if (pdfjsPromesa) return pdfjsPromesa;
    pdfjsPromesa = new Promise(function (ok, fail) {
      var s = document.createElement("script");
      s.src = PDFJS_URL;
      s.onload = function () { root.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER; ok(root.pdfjsLib); };
      s.onerror = function () { pdfjsPromesa = null; fail(new Error("No se pudo cargar el lector de PDF (pdf.js).")); };
      document.head.appendChild(s);
    });
    return pdfjsPromesa;
  }

  // Devuelve [{items, rects}] por página
  async function leerPaginas(file, conBordes) {
    var pdfjs = await cargarPdfJs();
    var doc = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
    var paginas = [];
    for (var p = 1; p <= doc.numPages; p++) {
      var page = await doc.getPage(p);
      var vp = page.getViewport({ scale: 1 });
      var tc = await page.getTextContent();
      var items = tc.items.map(function (it) {
        return { str: it.str, x: it.transform[4], y: vp.height - it.transform[5],
                 w: it.width, h: Math.abs(it.transform[3]) || it.height || 10 };
      });
      var rects = conBordes ? bordesDesdeOperadores(await page.getOperatorList(), pdfjs.OPS, vp.height) : [];
      paginas.push({ items: items, rects: rects });
    }
    return paginas;
  }

  function $(id) { return document.getElementById(id); }
  function fmt(n) { return n === "" || n === undefined ? "" : Number(n).toLocaleString("es-CO"); }
  function esc(s) { return String(s === undefined || s === null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function met(t, v, ac) { return '<div class="metrica-card' + (ac ? " acento-" + ac : "") + '"><div class="metrica-valor">' + v + '</div><div class="metrica-etiqueta">' + t + "</div></div>"; }
  function descargar(buf, nombre) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = nombre; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
  }
  function dropMultiple(dropId, inp) {
    var drop = $(dropId);
    ["dragover", "dragenter"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("arrastrando"); }); });
    ["dragleave", "drop"].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("arrastrando"); }); });
    drop.addEventListener("drop", function (e) { if (e.dataTransfer.files.length) { inp.files = e.dataTransfer.files; inp.dispatchEvent(new Event("change")); } });
  }

  // Configura un panel (CRP o CDP). cfg: {s: sufijo de ids, leerMemo(file), procesar(filas, indice), excel(filas), nombre(memos), pintar(r)}
  function montarPanel(cfg) {
    var s = cfg.s, st = { memos: [], indice: null };
    function alerta(msg) { if (typeof mostrarAlerta === "function") mostrarAlerta("alerta-" + s, msg); else if (msg) alert(msg); }
    function estado(on) { if (typeof mostrarEstado === "function") mostrarEstado("estado-" + s, on); }
    function boton() { $("btn-" + s).disabled = !(st.memos.length && st.indice); }
    if (typeof configurarDropzoneUnico === "function") configurarDropzoneUnico("drop-" + s + "-cdp", "input-" + s + "-cdp", "lista-" + s + "-cdp");
    var inp = $("input-" + s + "-pdf");
    dropMultiple("drop-" + s + "-pdf", inp);

    inp.addEventListener("change", async function () {
      alerta(""); st.memos = []; boton();
      var files = Array.prototype.slice.call(inp.files).filter(function (f) { return /\.pdf$/i.test(f.name); });
      $("lista-" + s + "-pdf").innerHTML = files.map(function (f) { return '<div class="archivo-chip"><span>' + esc(f.name) + "</span></div>"; }).join("");
      if (!files.length) return;
      $("resumen-" + s + "-pdf").textContent = "Leyendo " + files.length + " PDF…";
      try {
        for (var i = 0; i < files.length; i++) st.memos.push(await cfg.leerMemo(files[i]));
        var n = st.memos.reduce(function (a, m) { return a + m.filas.length; }, 0);
        $("resumen-" + s + "-pdf").innerHTML = st.memos.map(function (m) {
          return "📄 " + esc(m.archivo) + " · Radicado <b>" + (esc(m.radicado) || "<span style='color:#c00'>no encontrado</span>") + "</b> · <b>" + m.filas.length + "</b> contratos";
        }).join("<br>") + (st.memos.length > 1 ? "<br>Total: <b>" + n + "</b> contratos" : "");
        if (!n) alerta("No se encontraron contratos en los PDF. Verifique que sean memorandos con texto (no escaneados) y con la tabla de contratos.");
      } catch (err) { $("resumen-" + s + "-pdf").textContent = ""; alerta(err.message); st.memos = []; }
      boton();
    });

    $("input-" + s + "-cdp").addEventListener("change", function (e) {
      var file = e.target.files[0]; if (!file) return;
      alerta(""); st.indice = null; boton();
      $("resumen-" + s + "-cdp").textContent = "Leyendo reporte CDP…";
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var wb = root.XLSX.read(ev.target.result, { type: "array" });
          var aoa = root.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: "" });
          st.indice = indexarReporteCDP(aoa);
          $("resumen-" + s + "-cdp").innerHTML = "Filas: <b>" + st.indice.total + "</b> · CDP únicos: <b>" + st.indice.unicos + "</b>";
        } catch (err) { $("resumen-" + s + "-cdp").textContent = ""; alerta(err.message); }
        boton();
      };
      reader.readAsArrayBuffer(file);
    });

    $("btn-" + s).addEventListener("click", async function () {
      alerta(""); estado(true); this.disabled = true;
      try {
        var filas = [];
        st.memos.forEach(function (m) { filas = filas.concat(m.filas); });
        var r = cfg.procesar(filas, st.indice);
        descargar(await cfg.excel(r), cfg.nombre(st.memos));
        cfg.pintar(r);
        $("resultados-" + s).classList.add("visible");
      } catch (err) { alerta(err.message || String(err)); }
      estado(false); boton();
    });
  }

  function pintarCRP(r) {
    var ok = r.filter(function (x) { return x.reserva; }).length;
    var obs = r.filter(function (x) { return x.observaciones; }).length;
    var tot = r.reduce(function (a, x) { return a + (Number(x.valor_adicion) || 0); }, 0);
    $("metricas-m11").innerHTML = met("Contratos", r.length) + met("CDP cruzados", ok + " / " + r.length, ok === r.length ? "verde" : "rojo") +
      met("Código 145", r.filter(function (x) { return x.codigo === "145"; }).length, "ocre") +
      met("Código 148", r.filter(function (x) { return x.codigo === "148"; }).length, "ocre") +
      met("Con observaciones", obs, obs ? "rojo" : "verde") + met("Total adiciones", "$ " + fmt(tot));
    $("tabla-m11").innerHTML = '<table class="tabla-m11"><thead><tr><th>Contrato</th><th>Contratista</th><th>No. CDP</th><th>Valor adición</th><th>Importe</th><th>Anul.</th><th>reserva</th><th>Pos</th><th>Fondo</th><th>Pos pres</th><th>Radicado</th><th>Cód.</th><th>Observaciones</th></tr></thead><tbody>' +
      r.map(function (x) {
        return "<tr" + (x.observaciones ? ' class="m11-obs"' : "") + "><td>" + esc(x.contrato) + "</td><td>" + esc(x.contratista) + "</td><td>" + esc(x.no_cdp) +
          '</td><td class="num">' + fmt(x.valor_adicion) + '</td><td class="num">' + fmt(x.importe) + '</td><td class="num">' + fmt(x.anulaciones) + "</td><td>" + esc(x.reserva) +
          "</td><td>" + esc(x.pos) + "</td><td>" + esc(x.fondos) + "</td><td>" + esc(x.pos_pres) + "</td><td>" + esc(x.radicado) + "</td><td><b>" + esc(x.codigo) + "</b></td><td>" + esc(x.observaciones) + "</td></tr>";
      }).join("") + "</tbody></table>";
  }

  function pintarCDP(r) {
    var ok = r.filter(function (x) { return x.fondos; }).length;
    var obs = r.filter(function (x) { return x.observaciones; }).length;
    var tot = r.reduce(function (a, x) { return a + (Number(x.valor_adicion) || 0); }, 0);
    $("metricas-m11c").innerHTML = met("Contratos", r.length) + met("Cruzados por contrato", ok + " / " + r.length, ok === r.length ? "verde" : "rojo") +
      met("Código 145", r.filter(function (x) { return x.codigo === "145"; }).length, "ocre") +
      met("Código 148", r.filter(function (x) { return x.codigo === "148"; }).length, "ocre") +
      met("Con observaciones", obs, obs ? "rojo" : "verde") + met("Total adiciones", "$ " + fmt(tot));
    $("tabla-m11c").innerHTML = '<table class="tabla-m11"><thead><tr><th>Proyecto</th><th>Contrato</th><th>Contratista</th><th>Valor adición</th><th>Cód.</th><th>Pos pres</th><th>Fondo</th><th>Elemento PEP</th><th>Radicado</th><th>Observaciones</th></tr></thead><tbody>' +
      r.map(function (x) {
        return "<tr" + (x.observaciones ? ' class="m11-obs"' : "") + "><td>" + esc(x.proyecto) + "</td><td>" + esc(x.contrato) + "</td><td>" + esc(x.contratista) +
          '</td><td class="num">' + fmt(x.valor_adicion) + "</td><td><b>" + esc(x.codigo) + "</b></td><td>" + esc(x.pos_pres) + "</td><td>" + esc(x.fondos) +
          "</td><td>" + esc(x.pep) + "</td><td>" + esc(x.radicado) + "</td><td>" + esc(x.observaciones) + "</td></tr>";
      }).join("") + "</tbody></table>";
  }

  function nombreConRadicado(base, memos) {
    return base + (memos.length === 1 && memos[0].radicado ? "_" + memos[0].radicado : "") + ".xlsx";
  }

  function iniciar() {
    var vista = $("vista-cdpcrp-vigencia");
    if (!vista) return;
    // Pestañas CRP / CDP
    Array.prototype.forEach.call(vista.querySelectorAll(".m11-tab"), function (tab) {
      tab.addEventListener("click", function () {
        vista.querySelectorAll(".m11-tab").forEach(function (t) { t.classList.remove("activo"); });
        vista.querySelectorAll(".m11-panel").forEach(function (p) { p.classList.remove("activo"); });
        tab.classList.add("activo");
        $(tab.dataset.panel).classList.add("activo");
        if ($("titulo-modulo")) $("titulo-modulo").textContent = "CDP_CRP Vigencia · " + tab.textContent.trim();
      });
    });
    montarPanel({
      s: "m11",
      leerMemo: async function (f) { var p = await leerPaginas(f, false); return parsearMemorando(p.map(function (x) { return x.items; }), f.name); },
      procesar: cruzar,
      excel: function (r) {
        return generarExcel(root.ExcelJS, r, {
          tipoPago: ($("m11-tipopago") && $("m11-tipopago").value.trim()) || "02",
          modSeleccion: ($("m11-modsel") && $("m11-modsel").value.trim()) || "10"
        });
      },
      nombre: function (m) { return nombreConRadicado("MEMORANDO_SOLICITUD_CRP", m); },
      pintar: pintarCRP
    });
    if ($("btn-m11c")) montarPanel({
      s: "m11c",
      leerMemo: async function (f) { return parsearSolicitudCDP(await leerPaginas(f, true), f.name); },
      procesar: cruzarPorContrato,
      excel: function (r) { return generarExcelCDP(root.ExcelJS, r); },
      nombre: function (m) { return nombreConRadicado("SOLICITUD_CDP_PRORROGA_ADICION", m); },
      pintar: pintarCDP
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar); else iniciar();
})(typeof window !== "undefined" ? window : globalThis);
