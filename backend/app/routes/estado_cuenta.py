"""
Módulo 04 — Estado de Cuenta (versión integrada)
=================================================
Genera el estado de cuenta oficial de un contrato combinando tres fuentes:

  1. Reporte CRP Histórico  -> valor inicial (RP base), RP SAP1 y ADICIONES
  2. Consolidado SECOP2      -> contratista, CC/NIT y fechas
  3. Histórico de pagos      -> relación de pagos (opcional)

Reglas de negocio (confirmadas con el usuario):
  - CTO Y VIG se cruza contra "No. Compromiso" por PREFIJO
    (ej. 008-2022 agarra 008-2022, 008-20221, 008-20222 = adiciones 1, 2, ...).
  - Se EXCLUYEN las reservas: objeto con REEMPLAZA / OBLIGACIÓN POR PAGAR /
    CONSTITUIRSE (referencian CRP de años anteriores, no son la adición).
  - La adición se consolida por N° Interno CRP; si tiene varias N° Posición CRP
    se SUMAN. Cada N° Interno CRP distinto = una fila de adición.
  - D7 (VALOR CONTRATO) = valor INICIAL (RP base), para que
    E18 = D7 + adiciones - pago cuadre sin duplicar.
  - Saldo RP: E(prim) = (D7 + suma adiciones) - D(prim); luego E = E_anterior - D.
  - TOTAL (columna E de la fila TOTAL) = último saldo, en negrilla.
  - La tabla de pagos se ajusta al número de pagos: inserta filas si sobran,
    elimina las filas en blanco si faltan, conservando fórmulas y formato.

Cualquier nombre de columna o coordenada se ajusta en la sección CONFIG.
"""

from flask import Blueprint, request, jsonify, send_file
from flask_cors import cross_origin
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import unicodedata, copy, re, os, tempfile, json

estado_cuenta_bp = Blueprint('estado_cuenta', __name__)

# ============================ CONFIG ============================
# --- Reporte CRP (nombres de columna) ---
CRP_COMPROMISO   = 'No. Compromiso'
CRP_OBJETO       = 'Objeto'
CRP_VALOR        = 'Valor CRP'          # OJO: por nombre, no por letra (está en AK)
CRP_INTERNO      = 'N° Interno CRP'     # por nombre, no por letra (está en AQ)
CRP_POSICION     = 'N° Posición CRP'
CRP_BENEF_NOMBRE = 'Nombre BP Beneficiario'
CRP_BENEF_DOC    = 'Número Doc. BP Beneficiario'
CRP_FECHA_INI    = 'Fecha Inicial'
CRP_FECHA_FIN    = 'Fecha Final'

# --- Consolidado SECOP2 ---
CON_CONTRATO   = 'CONTRATO'
CON_NOMBRE     = 'NOMBRE DEL CONTRATISTA'
CON_DOC        = 'NUMERO DE IDENTIFICACION'
CON_VALOR_INI  = 'VALOR INICIAL DEL CONTRATO'
CON_FECHA_INI  = 'FECHA INICIAL DE CONTRATO'
CON_FECHA_FIN  = 'FECHA DE TERMINACION FINAL'

# --- Histórico de pagos ---
HIS_REFERENCIA = 'Referencia'
HIS_VALOR      = 'Valor Bruto'
HIS_PERIODO    = 'Texto cabecera documento'
HIS_DOC        = 'Doc.compensación'
HIS_FECHA      = 'Fecha de pago'
HIS_RP         = 'Numero RP'
HIS_CDP        = 'CDP Externo'
HIS_CRP        = 'CRP Externo'

HIS_STATUS     = 'Estatus'      # columna BH del histórico
HIS_STATUS_OK  = 'PAGADA'





# --- Coordenadas base del formato (antes de insertar adiciones) ---
C_CTO, C_CONTRATISTA, C_CCNIT      = 'D5', 'D6', 'H6'
C_VALOR, C_RPSAP1                  = 'D7', 'H7'
C_FECHA_INI, C_FECHA_FIN           = 'D8', 'H8'
FILA_ADICION_1                     = 9      # D9 valor / H9 RP de la 1ª adición
FILA_PAGOS_INI_BASE                = 17     # 1er pago (formato original)
SLOTS_PAGOS_BASE                   = 12     # filas de pago que trae el formato
REEMPLAZO = re.compile(r'REEMPLAZA|OBLIGACION POR PAGAR|CONSTITUIRSE')
# ===============================================================


def _sinac(s):
    """Quita tildes/ñ y pasa a mayúsculas para comparar textos de forma robusta."""
    return ''.join(c for c in unicodedata.normalize('NFKD', str(s))
                   if not unicodedata.combining(c)).upper()


def _es_match(nc, contrato):
    """No. Compromiso pertenece al contrato si es igual o es <contrato><dígitos>."""
    nc = str(nc).strip()
    if nc == contrato:
        return True
    return nc.startswith(contrato) and nc[len(contrato):].isdigit()


def _copiar_estilo(origen, destino):
    if origen.has_style:
        destino.font = copy.copy(origen.font)
        destino.fill = copy.copy(origen.fill)
        destino.border = copy.copy(origen.border)
        destino.alignment = copy.copy(origen.alignment)
        destino.number_format = origen.number_format
        destino.protection = copy.copy(origen.protection)


def _bump_formula(f, R, k):
    """Suma k a toda referencia de fila >= R dentro de una fórmula."""
    return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                  lambda m: f"{m.group(1)}{int(m.group(2)) + k}" if int(m.group(2)) >= R else m.group(0),
                  f)


def insertar_filas(ws, R, k):
    """Inserta k filas en la posición R ajustando fórmulas, combinaciones y alturas."""
    if k <= 0:
        return
    max_row, max_col = ws.max_row, ws.max_column
    readd = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= R:
            readd.append((mr.min_row + k, mr.min_col, mr.max_row + k, mr.max_col)); ws.unmerge_cells(str(mr))
        elif mr.max_row >= R:
            readd.append((mr.min_row, mr.min_col, mr.max_row + k, mr.max_col)); ws.unmerge_cells(str(mr))
    for row in range(max_row, R - 1, -1):
        for col in range(1, max_col + 1):
            s = ws.cell(row=row, column=col); d = ws.cell(row=row + k, column=col)
            v = s.value
            if isinstance(v, str) and v.startswith('='):
                v = _bump_formula(v, R, k)
            d.value = v
            _copiar_estilo(s, d)
            s.value = None
    for (r1, c1, r2, c2) in readd:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for row in range(max_row, R - 1, -1):
        if row in ws.row_dimensions:
            ws.row_dimensions[row + k].height = ws.row_dimensions[row].height


def eliminar_filas(ws, R, k):
    """Elimina k filas desde R ajustando fórmulas, combinaciones y alturas."""
    if k <= 0:
        return
    max_row, max_col = ws.max_row, ws.max_column
    readd = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= R + k:              # totalmente debajo del bloque borrado
            readd.append((mr.min_row - k, mr.min_col, mr.max_row - k, mr.max_col)); ws.unmerge_cells(str(mr))
        elif mr.min_row >= R:                # dentro del bloque borrado -> se descarta
            ws.unmerge_cells(str(mr))
        elif mr.max_row >= R:                # atraviesa el bloque -> se encoge
            readd.append((mr.min_row, mr.min_col, max(mr.min_row, mr.max_row - k), mr.max_col)); ws.unmerge_cells(str(mr))

    def bump_down(f):
        return re.sub(r'(\$?[A-Z]{1,3}\$?)(\d+)',
                      lambda m: f"{m.group(1)}{int(m.group(2)) - k}" if int(m.group(2)) >= R + k else m.group(0),
                      f)
    for row in range(R + k, max_row + 1):
        for col in range(1, max_col + 1):
            s = ws.cell(row=row, column=col); d = ws.cell(row=row - k, column=col)
            v = s.value
            if isinstance(v, str) and v.startswith('='):
                v = bump_down(v)
            d.value = v
            _copiar_estilo(s, d)
    # limpiar las últimas k filas que quedaron duplicadas al final
    for row in range(max_row - k + 1, max_row + 1):
        for col in range(1, max_col + 1):
            c = ws.cell(row=row, column=col)
            c.value = None
    for (r1, c1, r2, c2) in readd:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _extraer_datos_crp(df_crp, contrato):
    """Devuelve dict con base (valor inicial + RP) y lista de adiciones consolidadas."""
    df = df_crp.copy()
    df['_nc'] = df[CRP_COMPROMISO].astype(str).str.strip()
    df['_obj'] = df[CRP_OBJETO].map(_sinac)
    df = df[df['_nc'].apply(lambda x: _es_match(x, contrato))]

    es_reserva = df['_obj'].str.contains(REEMPLAZO, na=False)
    es_adicion = df['_obj'].str.contains('ADICION Y PRORROGA', na=False)
    df = df[~es_reserva]                       # fuera reservas siempre

    base = df[~es_adicion.reindex(df.index, fill_value=False)]
    adic = df[es_adicion.reindex(df.index, fill_value=False)]

    # Valor inicial (RP base) y RP SAP1 (interno base principal)
    valor_inicial, rp_sap1 = 0.0, ''
    if not base.empty:
        valor_inicial = float(base[CRP_VALOR].fillna(0).sum())
        por_interno = base.groupby(CRP_INTERNO)[CRP_VALOR].sum()
        if not por_interno.empty:
            rp_sap1 = int(por_interno.idxmax())

    # Adiciones: una fila por N° Interno CRP, sumando posiciones
    adiciones = []
    if not adic.empty:
        for interno, g in adic.groupby(CRP_INTERNO, sort=True):
            adiciones.append({'valor': float(g[CRP_VALOR].fillna(0).sum()),
                              'interno': int(interno)})

    # Nombre/CC de respaldo desde el CRP (por si no hay consolidado)
    nombre = doc = ''
    if not df.empty:
        n = df[CRP_BENEF_NOMBRE].dropna()
        d = df[CRP_BENEF_DOC].dropna()
        if len(n): nombre = str(n.iloc[0])
        if len(d): doc = str(d.iloc[0])

    return {'valor_inicial': valor_inicial, 'rp_sap1': rp_sap1,
            'adiciones': adiciones, 'nombre_crp': nombre, 'doc_crp': doc}


def _extraer_datos_consolidado(df_con, contrato):
    df = df_con.copy()
    df[CON_CONTRATO] = df[CON_CONTRATO].astype(str).str.strip()
    fila = df[df[CON_CONTRATO] == contrato]
    if fila.empty:
        return None
    r = fila.iloc[0]
    def _fecha(v):
        try:
            return pd.to_datetime(v).to_pydatetime()
        except Exception:
            return None
    return {
        'nombre': None if pd.isna(r.get(CON_NOMBRE)) else str(r.get(CON_NOMBRE)),
        'doc':    None if pd.isna(r.get(CON_DOC)) else str(r.get(CON_DOC)),
        'valor_inicial': None if pd.isna(r.get(CON_VALOR_INI)) else float(r.get(CON_VALOR_INI)),
        'fecha_ini': _fecha(r.get(CON_FECHA_INI)),
        'fecha_fin': _fecha(r.get(CON_FECHA_FIN)),
    }


def _extraer_pagos(df_his, contrato):
    if df_his is None:
        return []
    df = df_his[df_his[HIS_REFERENCIA].astype(str).str.contains(contrato, na=False, regex=False)]

    if HIS_STATUS in df.columns:
        df = df[df[HIS_STATUS].astype(str).str.strip().str.upper() == HIS_STATUS_OK]
    pagos = []

    for _, p in df.iterrows():
        pagos.append({
            'periodo': '' if pd.isna(p.get(HIS_PERIODO)) else str(p.get(HIS_PERIODO)),
            'valor':   float(p.get(HIS_VALOR, 0) or 0),
            'doc':     '' if pd.isna(p.get(HIS_DOC)) else str(p.get(HIS_DOC)),
            'fecha':   '' if pd.isna(p.get(HIS_FECHA)) else str(p.get(HIS_FECHA)),
            'rp':      '' if pd.isna(p.get(HIS_RP)) else str(p.get(HIS_RP)),
            'cdp':     '' if pd.isna(p.get(HIS_CDP)) else str(p.get(HIS_CDP)),
            'crp':     '' if pd.isna(p.get(HIS_CRP)) else str(p.get(HIS_CRP)),
        })
    return pagos


def generar_estado_cuenta(plantilla_path, crp_path, consolidado_path,
                          historico_path, contrato, salida_path):
    """Genera el estado de cuenta completo. Devuelve dict de resumen."""
    contrato = str(contrato).strip()

    df_crp = pd.read_excel(crp_path, sheet_name=0, engine="calamine")
    df_con = pd.read_excel(consolidado_path, sheet_name=0, engine="calamine") if consolidado_path else None
    _cols_his = [HIS_REFERENCIA, HIS_VALOR, HIS_PERIODO, HIS_DOC,
                 HIS_FECHA, HIS_RP, HIS_CDP, HIS_CRP, HIS_STATUS]
    if historico_path:
        df_his = pd.read_excel(historico_path, sheet_name=0,
                               usecols=lambda c: c in _cols_his,
                               engine="calamine")
    else:
        df_his = None

    crp = _extraer_datos_crp(df_crp, contrato)
    con = _extraer_datos_consolidado(df_con, contrato) if df_con is not None else None
    pagos = _extraer_pagos(df_his, contrato)

    if crp['valor_inicial'] == 0 and not crp['adiciones'] and con is None:
        return {'ok': False, 'mensaje': f'No se encontró información para el contrato {contrato}'}

    # Fuentes finales (consolidado con prioridad; CRP de respaldo)
    nombre = (con and con['nombre']) or crp['nombre_crp'] or ''
    doc    = (con and con['doc'])    or crp['doc_crp'] or ''
    valor_inicial = crp['valor_inicial'] or (con and con['valor_inicial']) or 0.0
    fecha_ini = con['fecha_ini'] if con else None
    fecha_fin = con['fecha_fin'] if con else None
    adiciones = crp['adiciones']
    n_ad = len(adiciones)

    wb = load_workbook(plantilla_path)
    ws = wb.active

    # ---------- Cabecera ----------
    ws[C_CTO] = contrato
    ws[C_CONTRATISTA] = nombre
    ws[C_CCNIT] = doc
    ws[C_VALOR] = valor_inicial
    if crp['rp_sap1'] != '':
        ws[C_RPSAP1] = crp['rp_sap1']; ws[C_RPSAP1].number_format = 'General'
    if fecha_ini is not None:
        ws[C_FECHA_INI] = fecha_ini; ws[C_FECHA_INI].number_format = 'DD/MM/YYYY'
    if fecha_fin is not None:
        ws[C_FECHA_FIN] = fecha_fin; ws[C_FECHA_FIN].number_format = 'DD/MM/YYYY'

    # ---------- Adiciones ----------
    if n_ad >= 1:
        ws[f'D{FILA_ADICION_1}'] = adiciones[0]['valor']
        ws[f'H{FILA_ADICION_1}'] = adiciones[0]['interno']
        ws[f'H{FILA_ADICION_1}'].number_format = 'General'
    extra = adiciones[1:]
    if extra:
        insertar_filas(ws, FILA_ADICION_1 + 1, len(extra))
        for i, ad in enumerate(extra):
            r = FILA_ADICION_1 + 1 + i
            for cc in ['C', 'D', 'G', 'H']:
                _copiar_estilo(ws[f'{cc}{FILA_ADICION_1}'], ws[f'{cc}{r}'])
            ws[f'C{r}'] = f'VALOR ADICIÓN {i + 2}'
            ws[f'D{r}'] = ad['valor']
            ws[f'G{r}'] = f'RP ADICIÓN {i + 2}'
            ws[f'H{r}'] = ad['interno']; ws[f'H{r}'].number_format = 'General'

    # Desplazamiento provocado por adiciones extra
    k_ad = len(extra)
    pago_ini   = FILA_PAGOS_INI_BASE + k_ad
    total_row0 = FILA_PAGOS_INI_BASE + SLOTS_PAGOS_BASE + k_ad   # fila TOTAL con 12 slots

    # ---------- Ajuste de filas de pago ----------
    n_pagos = len(pagos)
    if n_pagos > SLOTS_PAGOS_BASE:
        insertar_filas(ws, pago_ini + SLOTS_PAGOS_BASE, n_pagos - SLOTS_PAGOS_BASE)
    elif 0 < n_pagos < SLOTS_PAGOS_BASE:
        eliminar_filas(ws, pago_ini + n_pagos, SLOTS_PAGOS_BASE - n_pagos)

    filas_pago = max(n_pagos, SLOTS_PAGOS_BASE if n_pagos == 0 else n_pagos)
    total_row = pago_ini + filas_pago

    # Base para el saldo: D7 + bloque de adiciones (D9:D(8+n_ad))
    if n_ad >= 1:
        base_saldo = f"{C_VALOR}+SUM(D{FILA_ADICION_1}:D{FILA_ADICION_1 + n_ad - 1})"
    else:
        base_saldo = C_VALOR

    fmt_valor = ws[f'D{pago_ini}'].number_format

    # ---------- Relación de pagos ----------
    for i in range(filas_pago):
        r = pago_ini + i
        if i < n_pagos:
            p = pagos[i]
            ws[f'B{r}'] = i + 1
            ws[f'C{r}'] = p['periodo']
            ws[f'D{r}'] = p['valor']
            ws[f'F{r}'] = p['doc']
            ws[f'G{r}'] = p['fecha']
            ws[f'H{r}'] = p['rp']
            ws[f'I{r}'] = p['cdp']
            ws[f'J{r}'] = p['crp']
        # Saldo RP (fórmula viva) para todas las filas de pago
        if i == 0:
            ws[f'E{r}'] = f"={base_saldo}-D{r}"
        else:
            ws[f'E{r}'] = f"=E{r - 1}-D{r}"
        ws[f'E{r}'].number_format = fmt_valor

    # ---------- Fila TOTAL ----------
    ult = total_row - 1
    ws[f'D{total_row}'] = f"=SUM(D{pago_ini}:D{ult})"
    ws[f'D{total_row}'].number_format = fmt_valor
    ws[f'E{total_row}'] = f"=E{ult}"
    f = ws[f'E{total_row}'].font
    ws[f'E{total_row}'].font = Font(name=f.name, size=f.size, bold=True)
    ws[f'E{total_row}'].number_format = fmt_valor

    wb.save(salida_path)
    return {'ok': True, 'contrato': contrato, 'contratista': nombre,
            'valor_inicial': valor_inicial, 'n_adiciones': n_ad,
            'n_pagos': n_pagos, 'valor_final': valor_inicial + sum(a['valor'] for a in adiciones)}


# ============================ ENDPOINTS ============================
@estado_cuenta_bp.route('/procesar', methods=['POST'])
@cross_origin()
def procesar():
    try:
        plantilla    = request.files.get('plantilla')
        reporte_crp  = request.files.get('reporte_crp')
        consolidado  = request.files.get('consolidado')
        historico    = request.files.get('historico')     # opcional
        contrato     = (request.form.get('contrato') or '').strip()

        if not plantilla or not reporte_crp or not contrato:
            return jsonify({"ok": False,
                            "mensaje": "Faltan datos: se requieren plantilla, reporte_crp y contrato"}), 400

        tmp = tempfile.mkdtemp()
        p_plantilla = os.path.join(tmp, 'plantilla.xlsx');   plantilla.save(p_plantilla)
        p_crp       = os.path.join(tmp, 'crp.xlsx');          reporte_crp.save(p_crp)
        p_con = p_his = None
        if consolidado:
            p_con = os.path.join(tmp, 'consolidado.xlsx');    consolidado.save(p_con)
        if historico:
            p_his = os.path.join(tmp, 'historico.xlsx');      historico.save(p_his)

        salida = os.path.join(tmp, f'Estado_de_Cuenta_{contrato}.xlsx')
        res = generar_estado_cuenta(p_plantilla, p_crp, p_con, p_his, contrato, salida)
        if not res.get('ok'):
            return jsonify(res), 400
        return send_file(salida, as_attachment=True,
                         download_name=f'Estado_de_Cuenta_{contrato}.xlsx')
    except Exception as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 500


@estado_cuenta_bp.route('/procesar-lite', methods=['POST'])
def procesar_lite():
    """Ruta previa conservada (recibe pagos ya armados en JSON)."""
    try:
        from app.services.estado_cuenta_service import generar_estado_cuenta_desde_datos
    except ImportError:
        from ..services.estado_cuenta_service import generar_estado_cuenta_desde_datos
    try:
        plantilla = request.files.get('plantilla')
        contrato = (request.form.get('contrato') or '').strip()
        pagos_raw = request.form.get('pagos')
        if not plantilla or not contrato or not pagos_raw:
            return jsonify({"ok": False, "mensaje": "Faltan datos"}), 400
        pagos = json.loads(pagos_raw)
        tmp_dir = tempfile.mkdtemp()
        ruta_plantilla = os.path.join(tmp_dir, 'plantilla.xlsx')
        plantilla.save(ruta_plantilla)
        ruta_salida = os.path.join(tmp_dir, 'Estado_de_Cuenta_' + contrato + '.xlsx')
        resultado = generar_estado_cuenta_desde_datos(ruta_plantilla, pagos, contrato, ruta_salida)
        if not resultado.get('ok'):
            return jsonify(resultado), 400
        return send_file(ruta_salida, as_attachment=True,
                         download_name='Estado_de_Cuenta_' + contrato + '.xlsx')
    except Exception as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 500