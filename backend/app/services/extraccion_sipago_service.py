"""
Módulo — Extracción SIPAGO.

Toma el "Reporte del Periodo" (listado mensual de contratistas a pagar) y lo
enriquece con dos bases de referencia que quedan guardadas en el servidor:

  BASE CRP        -> CRP (N° Interno CRP, col AR) y SALDO CRP (Com.Sin.Aut.Giro, col AQ)
                     Regla: si un contratista tiene varios CRP, se toma el de MAYOR saldo.
  BASE BANCARIA   -> Código Bco, Tipo Cta, No Cuenta, extraídos de los históricos de pago.
                     Regla: si aparece en varios meses, se toma la cuenta MÁS RECIENTE.

El cruce en ambos casos es por CÉDULA (No Documento), no por nombre: la cédula es
única y no se ve afectada por tildes ni por variaciones de escritura.

La salida replica exactamente el esquema de 20 columnas del Excel del Módulo 01.
"""
import os
import re
import json
import itertools
import openpyxl
import pandas as pd

# ── Esquema de salida (idéntico al Módulo 01) ────────────────────────────────
COL_ORDEN_FINAL = [
    'Contrato', 'Documento No.', 'Contratista', 'NIT_CC',
    'Valor Bruto', 'BASE RETEICA', 'Valor Reteica',
    'TOTAL DESCUENTOS', 'Neto a Pagar',
    'PERIODO', 'DEL', 'AL', 'PAGO No.',
    'HONOR_BASE', 'HONOR_VALOR',
    'Código Bco', 'Tipo Cta', 'No Cuenta', 'CRP', 'SALDO CRP',
]

# ── Posiciones en el Reporte CRP (por índice: el nombre puede variar) ─────────
CRP_IDX_DOC     = 31   # AF - Número Doc. BP Beneficiario
CRP_IDX_NOMBRE  = 32   # AG - Nombre BP Beneficiario
CRP_IDX_SALDO   = 42   # AQ - Com.Sin.Aut.Giro      -> SALDO CRP
CRP_IDX_INTERNO = 43   # AR - N° Interno CRP        -> CRP

# ── Bloques de retención del Reporte del Periodo ─────────────────────────────
# Cada retención viene en un bloque de 3 columnas: CRÉDITO (valor) / % / BASE.
# Se ubican por NOMBRE de encabezado (no por posición), así funciona aunque el
# reporte traiga una columna de más o de menos. Regla:
#   - CRÉDITO = 0  -> se deja en 0 (sin % ni base).
#   - CRÉDITO > 0  -> se toman valor, % y base de ese mismo bloque.
# (clave, etiqueta de salida, función que reconoce el encabezado normalizado)
BLOQUES_RETENCION = [
    ('RF_RENTAS',  'RETEFUENTE RENTAS',  lambda h: 'RENTAS DE TRABAJO' in h),
    ('RF_HONOR',   'RETEFUENTE HONOR',   lambda h: 'HONORARIOS' in h and 'RENTAS' not in h
                                                   and 'DEBITO' not in h and 'MENSUALIDAD' not in h),
    ('RF_SERV',    'RETEFUENTE SERV',    lambda h: 'SERVICIOS' in h),
    ('RETEIVA',    'RETEIVA',            lambda h: h.endswith(' IVA')),
    ('RETEICA',    'RETEICA',            lambda h: h.endswith(' ICA')),
    ('EST_UD',     'ESTAMPILLA UD',      lambda h: 'UNIVERSIDAD' in h),
    ('EST_PROCUL', 'ESTAMPILLA PROCULT', lambda h: 'PROCULTURA' in h),
    ('EST_ADULTO', 'ESTAMPILLA ADULTO M',lambda h: 'PROADULTO' in h),
]
# Retenciones en la fuente: van a la plantilla como filas código 44 (HONOR_*/RET2_*).
CLAVES_RETEFUENTE = ('RF_RENTAS', 'RF_HONOR', 'RF_SERV')

# Posiciones de respaldo (formato original) por si el reporte llega sin encabezados.
REP_IDX_HONOR_VALOR = 16   # Q
REP_IDX_HONOR_PCT   = 17   # R
REP_IDX_HONOR_BASE  = 18   # S
REP_IDX_RETEICA_VAL = 28   # AC
REP_IDX_RETEICA_PCT = 29   # AD
REP_IDX_RETEICA_BASE= 30   # AE

TIPO_CTA = {'AHORROS': '02', 'CORRIENTE': '01'}
RE_CUENTA = re.compile(r'^\s*(\S+)\s+(AHORROS|CORRIENTE)\s+(.+)$', re.IGNORECASE)
RE_PAGO   = re.compile(r'PAGO\s*(?:N[°º.]?\s*)?(\d+)', re.IGNORECASE)

MESES = {
    'ENERO': '01', 'FEBRERO': '02', 'MARZO': '03', 'ABRIL': '04',
    'MAYO': '05', 'JUNIO': '06', 'JULIO': '07', 'AGOSTO': '08',
    'SEPTIEMBRE': '09', 'OCTUBRE': '10', 'NOVIEMBRE': '11', 'DICIEMBRE': '12',
}


# ── Utilidades ───────────────────────────────────────────────────────────────
def _digitos(valor):
    """Deja solo dígitos. Quita el '.0' de los floats leídos desde Excel."""
    s = str(valor).strip()
    s = re.sub(r'\.0$', '', s)
    return re.sub(r'[^\d]', '', s)


def _num(valor):
    """Convierte a número; devuelve 0 si no se puede."""
    return pd.to_numeric(valor, errors='coerce')


def _pct(valor):
    """'1,104%' -> 1.104 ; '0,766%' -> 0.766. Devuelve '' si no aplica."""
    s = str(valor).strip().replace('%', '').replace(' ', '').replace(',', '.')
    if not s or s in ('nan', '.'):
        return ""
    try:
        return float(s)
    except ValueError:
        return ""


def _col(df, idx):
    """Serie de la columna por índice; vacía si el archivo no llega tan lejos."""
    if idx < len(df.columns):
        return df.iloc[:, idx]
    return pd.Series([None] * len(df), index=df.index)


# Palabras que descartan una columna aunque coincida con la clave buscada:
# 'Clase identificación' contiene el tipo de documento ("CC", "NIT"), no el número.
_EXCLUIR_COL = ('clase', 'tipo doc', 'tipo de doc')


def _buscar_col(df, *claves):
    """Primera columna cuyo nombre contenga alguna de las claves.
    Se ignoran las columnas descriptivas (Clase/Tipo de documento)."""
    for c in df.columns:
        cl = str(c).strip().lower()
        if any(x in cl for x in _EXCLUIR_COL):
            continue
        if any(k in cl for k in claves):
            return c
    return None


# Claves para ubicar la cédula, EN ORDEN DE PRIORIDAD.
# 'documento' va de último a propósito: el histórico de pagos trae 'Nº documento'
# (el consecutivo contable de SAP) mucho antes que 'Nº identificación', y buscar
# columna por columna hacía que ganara el consecutivo y la base quedara con
# llaves equivocadas. Probando clave por clave, 'identificación' gana siempre.
_ID_CLAVES = ('identificación', 'identificacion', 'nit', 'cedula', 'cédula', 'documento')


def _idx_por_prioridad(encabezados, claves, excluir=_EXCLUIR_COL):
    """Índice de columna probando las claves EN ORDEN (no columna por columna).
    `encabezados` es la lista de nombres ya en minúscula."""
    for clave in claves:
        for i, h in enumerate(encabezados):
            if any(x in h for x in excluir):
                continue
            if clave in h:
                return i
    return None


# Formato B: cuenta en columnas separadas. Claves en orden de prioridad.
_NUM_CTA_CLAVES     = ('no cuenta', 'nº cuenta', 'n° cuenta', 'numero de cuenta',
                       'número de cuenta', 'numero cuenta')
_TIPO_CTA_CLAVES    = ('tipo cta', 'tipo de cuenta', 'tipo cuenta')
_COD_BANCO_CLAVES   = ('codigo de banco', 'código de banco', 'codigo banco', 'codigo', 'código')
_NOMBRE_BANCO_CLAVES = ('nombre del banco', 'nombre banco', 'banco')


def _solo_digitos_num(valor):
    """Número de cuenta sin separadores ni el '.0' que deja Excel en los floats."""
    if valor is None:
        return ""
    s = re.sub(r'\.0$', '', str(valor).strip())
    return re.sub(r'\D', '', s)


def _tipo_cuenta_b(valor):
    """Tipo de cuenta del Formato B: acepta '01'/'02', 'AHORROS'/'CORRIENTE' o un número."""
    if valor is None:
        return ""
    s = re.sub(r'\.0$', '', str(valor).strip().upper())
    if s in ('01', '02'):
        return s
    if s in TIPO_CTA:
        return TIPO_CTA[s]
    return s.zfill(2) if s.isdigit() else ""


def _celda(fila, idx):
    """Valor de la celda, tolerando filas más cortas que el encabezado."""
    if idx is None or idx >= len(fila):
        return None
    return fila[idx]


def _detectar_col_cuenta(filas_muestra):
    """Índice de la columna de cuenta, detectada por CONTENIDO y no por nombre.

    El histórico trae varias columnas llamadas 'Cuenta bancaria'; solo una viene
    con el formato '<numero> AHORROS|CORRIENTE <banco>'. Se elige la que más
    coincidencias tenga en la muestra.
    """
    conteo = {}
    for fila in filas_muestra:
        for i, v in enumerate(fila):
            if v is None:
                continue
            if RE_CUENTA.match(str(v)):
                conteo[i] = conteo.get(i, 0) + 1
    if not conteo:
        return None
    return max(conteo, key=conteo.get)


def _fmt_contrato(valor):
    """'310 2026' -> '310-2026' ; '1 2026' -> '001-2026'.
    El número de contrato se rellena a 3 dígitos, como en el Excel de extracción."""
    s = str(valor).strip()
    if not s or s.lower() == 'nan':
        return ""
    s = re.sub(r'[\s/]+', '-', s)
    m = re.match(r'^(\d+)-(\d{4})$', s)
    if m:
        return f"{m.group(1).zfill(3)}-{m.group(2)}"
    return s


def _norm_hdr(valor):
    """Encabezado en mayúsculas, sin tildes ni saltos de línea ni espacios dobles."""
    import unicodedata
    s = unicodedata.normalize('NFKD', str(valor)).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', s).strip().upper()


def _ubicar_bloques(df):
    """{clave: {'valor': col, 'pct': col, 'base': col}} buscando por encabezado.
    Si no encuentra RETEICA ni RETEFUENTE RENTAS por nombre, usa las posiciones
    del formato original."""
    hdrs = [(c, _norm_hdr(c)) for c in df.columns]
    bloques = {}
    for clave, _, reconoce in BLOQUES_RETENCION:
        b = {}
        for col, h in hdrs:
            if not reconoce(h):
                continue
            if h.startswith('CREDITO') and 'valor' not in b:
                b['valor'] = col
            elif h.startswith('%') and 'pct' not in b:
                b['pct'] = col
            elif h.startswith('BASE') and 'base' not in b:
                b['base'] = col
        if 'valor' in b:
            bloques[clave] = b

    def _por_idx(i):
        return df.columns[i] if i < len(df.columns) else None

    if 'RF_RENTAS' not in bloques and _por_idx(REP_IDX_HONOR_VALOR) is not None:
        bloques['RF_RENTAS'] = {'valor': _por_idx(REP_IDX_HONOR_VALOR),
                                'pct': _por_idx(REP_IDX_HONOR_PCT),
                                'base': _por_idx(REP_IDX_HONOR_BASE)}
    if 'RETEICA' not in bloques and _por_idx(REP_IDX_RETEICA_VAL) is not None:
        bloques['RETEICA'] = {'valor': _por_idx(REP_IDX_RETEICA_VAL),
                              'pct': _por_idx(REP_IDX_RETEICA_PCT),
                              'base': _por_idx(REP_IDX_RETEICA_BASE)}
    return bloques


def _txt_pct(valor):
    """Porcentaje tal como viene ('0,766%'); vacío si la celda está vacía."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    s = str(valor).strip()
    return "" if s.lower() == 'nan' else s


def _retenciones_fila(rep, bloques):
    """DataFrame con VALOR/%/BASE de cada bloque aplicando la regla:
    CRÉDITO en 0 se deja en 0; CRÉDITO con valor arrastra su % y su base."""
    out = {}
    for clave, etiqueta, _ in BLOQUES_RETENCION:
        b = bloques.get(clave)
        if b is None:
            valor = pd.Series([0.0] * len(rep), index=rep.index)
            pct = pd.Series([""] * len(rep), index=rep.index)
            base = pd.Series([None] * len(rep), index=rep.index, dtype=float)
        else:
            valor = _num(rep[b['valor']]).fillna(0)
            pct = rep[b['pct']].apply(_txt_pct) if b.get('pct') is not None \
                else pd.Series([""] * len(rep), index=rep.index)
            base = _num(rep[b['base']]) if b.get('base') is not None \
                else pd.Series([None] * len(rep), index=rep.index, dtype=float)
        hay = valor > 0
        out[f'{etiqueta} VALOR'] = valor
        out[f'{etiqueta} %'] = pct.where(hay, "")
        out[f'{etiqueta} BASE'] = base.where(hay)
    return pd.DataFrame(out, index=rep.index)


def _fechas_periodo(texto):
    """Deriva DEL y AL del texto del periodo.
    'HONORARIOS AGOSTO 01-31 DE 2026' -> ('01/08/2026', '31/08/2026')"""
    s = str(texto).upper()
    mes = next((MESES[m] for m in MESES if m in s), None)
    anio = re.search(r'\b(20\d{2})\b', s)
    rango = re.search(r'\b(\d{1,2})\s*[-aAlL]{1,2}\s*(\d{1,2})\b', s)
    if not (mes and anio and rango):
        return "", ""
    a = anio.group(1)
    d1, d2 = rango.group(1).zfill(2), rango.group(2).zfill(2)
    return f"{d1}/{mes}/{a}", f"{d2}/{mes}/{a}"


# ── Construcción de la BASE CRP ──────────────────────────────────────────────
def construir_base_crp(ruta_reporte_crp):
    """Lee el Reporte CRP y deja una fila por cédula con el CRP de mayor saldo."""
    df = pd.read_excel(ruta_reporte_crp)

    base = pd.DataFrame({
        'cedula':    _col(df, CRP_IDX_DOC).apply(_digitos),
        'nombre':    _col(df, CRP_IDX_NOMBRE).astype(str).str.strip(),
        'SALDO CRP': _num(_col(df, CRP_IDX_SALDO)).fillna(0),
        'CRP':       _col(df, CRP_IDX_INTERNO).apply(_digitos),
    })
    base = base[base['cedula'].str.len() > 0]

    # Regla acordada: cuando hay varios CRP para la misma cédula, gana el de mayor saldo.
    base = (base.sort_values('SALDO CRP', ascending=False)
                .drop_duplicates(subset=['cedula'], keep='first')
                .reset_index(drop=True))
    return base


# ── Construcción de la BASE BANCARIA ─────────────────────────────────────────
COLS_BASE_BANCARIA = ['cedula', 'Código Bco', 'Tipo Cta', 'No Cuenta']


def extraer_cuentas_historico(ruta, muestra=300):
    """Extrae las cuentas de UN histórico de pago sin cargarlo entero en memoria.

    openpyxl en modo read_only recorre la hoja fila por fila y aquí solo se
    conservan 3 celdas de cada una. Medido contra un histórico real de 35.441
    filas x 80 columnas: ~106 MB de pico, frente a ~225 MB de pd.read_excel y
    ~440 MB del motor calamine (calamine es más rápido pero gasta MÁS memoria,
    por eso no sirve como solución en el free tier de 512 MB).

    Devuelve (registros, motivo_omision). Si todo sale bien, motivo es None.
    """
    try:
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    except Exception as e:
        return [], f"no se pudo abrir ({e})"

    try:
        ws = wb[wb.sheetnames[0]]
        filas = ws.iter_rows(values_only=True)

        try:
            encabezados = [str(h).strip().lower() if h is not None else ""
                           for h in next(filas)]
        except StopIteration:
            return [], "el archivo está vacío"

        i_id = _idx_por_prioridad(encabezados, _ID_CLAVES)
        if i_id is None:
            return [], "sin columna de identificación"
        i_cod = _idx_por_prioridad(encabezados, ('codigo de banco', 'código de banco'))

        # La columna de cuenta se detecta leyendo una muestra, porque el nombre
        # no la distingue (hay varias 'Cuenta bancaria' y solo una trae el formato).
        buffer = list(itertools.islice(filas, muestra))
        i_cta = _detectar_col_cuenta(buffer)
        todas = itertools.chain(buffer, filas)

        if i_cta is not None:
            # Formato A: una sola columna "<numero> AHORROS|CORRIENTE <banco>"
            registros = []
            for fila in todas:
                bruto = _celda(fila, i_cta)
                if bruto is None:
                    continue
                m = RE_CUENTA.match(str(bruto))
                if not m:
                    continue

                cedula = _digitos(_celda(fila, i_id))
                cuenta = re.sub(r'[^\d]', '', m.group(1))
                if not cedula or not cuenta:
                    continue

                registros.append({
                    'cedula':    cedula,
                    'No Cuenta': cuenta,
                    'Tipo Cta':  TIPO_CTA.get(m.group(2).upper()),
                    'banco_txt': m.group(3).strip(),
                    'cod_banco': _celda(fila, i_cod) if i_cod is not None else None,
                })
            return registros, None

        # Formato B: la cuenta viene en columnas separadas (número, tipo, código, banco)
        i_num = _idx_por_prioridad(encabezados, _NUM_CTA_CLAVES)
        if i_num is None:
            return [], "sin columna de cuenta bancaria"
        i_tipo = _idx_por_prioridad(encabezados, _TIPO_CTA_CLAVES)
        i_cod_b = _idx_por_prioridad(encabezados, _COD_BANCO_CLAVES)
        i_bco = _idx_por_prioridad(encabezados, _NOMBRE_BANCO_CLAVES)

        registros = []
        for fila in todas:
            cuenta = _solo_digitos_num(_celda(fila, i_num))
            cedula = _digitos(_celda(fila, i_id))
            if not cedula or not cuenta:
                continue
            banco = _celda(fila, i_bco) if i_bco is not None else None
            registros.append({
                'cedula':    cedula,
                'No Cuenta': cuenta,
                'Tipo Cta':  _tipo_cuenta_b(_celda(fila, i_tipo)) if i_tipo is not None else '02',
                'banco_txt': '' if banco is None else str(banco).strip(),
                'cod_banco': _celda(fila, i_cod_b) if i_cod_b is not None else None,
            })
        return registros, None
    except Exception as e:
        return [], f"no se pudo leer ({e})"
    finally:
        try:
            wb.close()
        except Exception:
            pass


def _codigo_banco(valor, banco_txt, mapa):
    """Código de banco a 3 dígitos. Si la fila no lo trae, se deduce del nombre
    del banco usando el mapa aprendido de los archivos que sí lo incluyen."""
    v = valor
    if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
        v = mapa.get(banco_txt)
    if v is None:
        return ""
    try:
        return str(int(float(v))).zfill(3)
    except (TypeError, ValueError):
        return ""


def consolidar_registros_bancarios(registros, mapa_bancos=None):
    """Arma la base bancaria final a partir de registros ya extraídos.

    `registros` viene de `extraer_cuentas_historico` (servidor) o del navegador
    vía SheetJS; cada uno debe traer '_orden' (el índice del archivo del que
    salió). Regla acordada: gana la cuenta MÁS RECIENTE, es decir la del archivo
    de mayor '_orden'; dentro de un mismo archivo gana la primera aparición.
    """
    if not registros:
        return pd.DataFrame(columns=COLS_BASE_BANCARIA)

    # Mapa banco -> código, aprendido de las filas que sí traen 'Codigo de Banco'.
    mapa = {}
    conteos = {}
    for r in registros:
        v = r.get('cod_banco')
        if v is None or str(v).strip() == "":
            continue
        try:
            cod = int(float(v))
        except (TypeError, ValueError):
            continue
        banco = r.get('banco_txt') or ""
        if not banco:
            continue  # sin nombre de banco no hay de qué aprender el código
        conteos.setdefault(banco, {})
        conteos[banco][cod] = conteos[banco].get(cod, 0) + 1
    for banco, votos in conteos.items():
        # Más frecuente; si empatan, el código menor (igual que .mode() de pandas).
        mapa[banco] = min(sorted(votos), key=lambda c: -votos[c])
    if mapa_bancos:
        for banco, cod in mapa_bancos.items():
            mapa.setdefault(banco, cod)

    # Deduplicación: nos quedamos con la cuenta del archivo más reciente.
    elegidos = {}
    for r in registros:
        ced = str(r.get('cedula') or "").strip()
        if not ced:
            continue
        orden = r.get('_orden', 0) or 0
        previo = elegidos.get(ced)
        if previo is None or orden > previo.get('_orden', 0):
            elegidos[ced] = {**r, '_orden': orden}

    filas = [{
        'cedula':     ced,
        'Código Bco': _codigo_banco(r.get('cod_banco'), r.get('banco_txt') or "", mapa),
        'Tipo Cta':   r.get('Tipo Cta') or "",
        'No Cuenta':  r.get('No Cuenta') or "",
    } for ced, r in elegidos.items()]

    return pd.DataFrame(filas, columns=COLS_BASE_BANCARIA)


def construir_base_bancaria(rutas_historicos):
    """Lee uno o varios históricos de pago y extrae la cuenta de cada contratista.

    Los archivos se procesan en el orden recibido; el ÚLTIMO gana, de modo que
    basta con enviarlos de más antiguo a más reciente.

    Ojo: esta ruta lee los Excel en el servidor. Funciona para archivos
    moderados, pero para históricos completos el camino recomendado es que el
    navegador haga la extracción (endpoint /cargar-base-bancaria-json).
    """
    registros = []
    omitidos = []

    for orden, ruta in enumerate(rutas_historicos):
        nombre_arch = os.path.basename(ruta)
        filas, motivo = extraer_cuentas_historico(ruta)
        if motivo:
            omitidos.append(f"{nombre_arch}: {motivo}")
            continue
        for r in filas:
            r['_orden'] = orden
        registros.extend(filas)

    return consolidar_registros_bancarios(registros), omitidos


def base_bancaria_desde_registros(registros, mapa_bancos=None):
    """Arma la base bancaria con los registros que el navegador ya extrajo."""
    limpios = []
    for r in registros or []:
        cedula = _digitos(r.get('cedula'))
        cuenta = re.sub(r'[^\d]', '', str(r.get('No Cuenta') or r.get('no_cuenta') or ''))
        if not cedula or not cuenta:
            continue
        limpios.append({
            'cedula':    cedula,
            'No Cuenta': cuenta,
            'Tipo Cta':  r.get('Tipo Cta') or r.get('tipo_cta') or "",
            'banco_txt': (r.get('banco_txt') or "").strip(),
            'cod_banco': r.get('cod_banco'),
            '_orden':    r.get('_orden', 0) or 0,
        })
    return consolidar_registros_bancarios(limpios, mapa_bancos)


# ── Consolidación ────────────────────────────────────────────────────────────
def consolidar(ruta_reporte, base_crp, base_bancaria, ruta_salida):
    """Cruza el Reporte del Periodo con ambas bases y genera el Excel final."""
    rep = pd.read_excel(ruta_reporte)

    col_doc    = _buscar_col(rep, 'no documento', 'nº documento', 'n° documento')
    col_nombre = _buscar_col(rep, 'contratista', 'nombre')
    if col_doc is None or col_nombre is None:
        return {"ok": False,
                "mensaje": "El reporte no tiene las columnas 'No Documento' y 'Contratista'."}

    col_bruto  = _buscar_col(rep, 'valor bruto')
    col_desc   = _buscar_col(rep, 'total descuentos')
    col_neto   = _buscar_col(rep, 'neto pagar', 'neto a pagar')
    col_per    = _buscar_col(rep, 'periodo', 'período')
    col_contr  = _buscar_col(rep, 'contrato')
    col_pago   = _buscar_col(rep, 'numero pago', 'número pago')

    cedulas = rep[col_doc].apply(_digitos)
    periodo_txt = rep[col_per].astype(str) if col_per else pd.Series([""] * len(rep))
    fechas = periodo_txt.apply(_fechas_periodo)

    # Retenciones: cada bloque CRÉDITO / % / BASE ubicado por encabezado.
    bloques = _ubicar_bloques(rep)
    rets = _retenciones_fila(rep, bloques)

    # Retenciones en la fuente con valor (en orden Q, T, W) -> HONOR_* y RET2_*,
    # que la plantilla convierte en filas código 44.
    rf = [(rets[f'{e} VALOR'], rets[f'{e} BASE'])
          for k, e, _ in BLOQUES_RETENCION if k in CLAVES_RETEFUENTE]
    honor_valor, honor_base, ret2_valor, ret2_base = [], [], [], []
    for i in range(len(rep)):
        con_valor = [(v.iloc[i], b.iloc[i]) for v, b in rf if v.iloc[i] > 0]
        v1, b1 = con_valor[0] if len(con_valor) > 0 else (None, None)
        v2, b2 = con_valor[1] if len(con_valor) > 1 else (None, None)
        honor_valor.append(v1); honor_base.append(b1)
        ret2_valor.append(v2); ret2_base.append(b2)

    ica_valor = rets['RETEICA VALOR']
    hay_ica = ica_valor > 0

    out = pd.DataFrame({
        'Contrato':         rep[col_contr].apply(_fmt_contrato) if col_contr else "",
        'Documento No.':    rep[col_pago].astype(str).str.strip() if col_pago else "",
        'Contratista':      rep[col_nombre].astype(str).str.strip(),
        'NIT_CC':           cedulas,
        'Valor Bruto':      _num(rep[col_bruto]) if col_bruto else 0,
        'BASE RETEICA':     rets['RETEICA BASE'],
        # 'Valor Reteica' lleva el % (mismo criterio del Módulo 01: la plantilla
        # lo convierte en el indicador de retención SAP).
        'Valor Reteica':    rets['RETEICA %'].apply(_pct).where(hay_ica, ""),
        'TOTAL DESCUENTOS': _num(rep[col_desc]) if col_desc else 0,
        'Neto a Pagar':     _num(rep[col_neto]) if col_neto else 0,
        # PERIODO en el Módulo 01 es el TOTAL de pagos del contrato ('PAGO 8 DE 12' -> 12).
        # El Reporte del Periodo no trae ese total, así que la columna queda vacía.
        'PERIODO':          "",
        'DEL':              fechas.apply(lambda t: t[0]),
        'AL':               fechas.apply(lambda t: t[1]),
        'PAGO No.':         (rep[col_pago].astype(str).str.extract(RE_PAGO)[0]
                             if col_pago else ""),
        'HONOR_BASE':       honor_base,
        'HONOR_VALOR':      honor_valor,
    })
    out['RET2_BASE'] = ret2_base
    out['RET2_VALOR'] = ret2_valor
    # Importe real del ICA: la plantilla lo usa en lugar de TOTAL DESCUENTOS,
    # que aquí es la SUMA de todas las retenciones.
    out['RETEICA_VALOR'] = ica_valor
    for c in rets.columns:
        out[c] = rets[c].values

    # Control: suma de créditos vs. TOTAL DESCUENTOS del reporte.
    suma_creditos = rets[[c for c in rets.columns if c.endswith(' VALOR')]].sum(axis=1)
    out['SUMA RETENCIONES'] = suma_creditos.values
    out['No. RETENCIONES'] = (rets[[c for c in rets.columns if c.endswith(' VALOR')]] > 0).sum(axis=1).values
    total_rep = pd.to_numeric(out['TOTAL DESCUENTOS'], errors='coerce').fillna(0)
    out['CUADRE'] = ['OK' if abs(s - t) < 1 else f'DIFERENCIA {t - s:,.0f}'
                     for s, t in zip(out['SUMA RETENCIONES'], total_rep)]

    # Cruce 1: CRP + SALDO CRP
    out = out.merge(
        base_crp[['cedula', 'CRP', 'SALDO CRP']].rename(columns={'cedula': 'NIT_CC'}),
        on='NIT_CC', how='left'
    )

    # Cruce 2: datos bancarios
    out = out.merge(
        base_bancaria.rename(columns={'cedula': 'NIT_CC'}),
        on='NIT_CC', how='left'
    )

    # Orden final por número de contrato
    out['_num'] = pd.to_numeric(out['Contrato'].astype(str).str.extract(r'^(\d+)')[0],
                                errors='coerce')
    out = (out.sort_values('_num', na_position='last')
              .drop(columns=['_num'])
              .reset_index(drop=True))

    for c in COL_ORDEN_FINAL:
        if c not in out.columns:
            out[c] = ""
    # Las 20 columnas del Módulo 01 primero; luego el detalle de retenciones.
    extras = ['RET2_BASE', 'RET2_VALOR', 'RETEICA_VALOR'] + list(rets.columns) + \
             ['SUMA RETENCIONES', 'No. RETENCIONES', 'CUADRE']
    out = out[COL_ORDEN_FINAL + extras]

    _escribir_excel(out, ruta_salida)

    sin_crp    = out['CRP'].isna() | (out['CRP'].astype(str).str.strip() == '')
    sin_cuenta = out['No Cuenta'].isna() | (out['No Cuenta'].astype(str).str.strip() == '')

    # Listado de pendientes para que el usuario los revise antes de pagar
    pendientes = out[sin_crp | sin_cuenta][['Contrato', 'Contratista', 'NIT_CC']].copy()
    pendientes['falta'] = [
        ', '.join(f for f, v in (('CRP', a), ('cuenta bancaria', b)) if v)
        for a, b in zip(sin_crp[sin_crp | sin_cuenta], sin_cuenta[sin_crp | sin_cuenta])
    ]

    return {
        "ok": True,
        "registros":      int(len(out)),
        "con_crp":        int((~sin_crp).sum()),
        "sin_crp":        int(sin_crp.sum()),
        "con_cuenta":     int((~sin_cuenta).sum()),
        "sin_cuenta":     int(sin_cuenta.sum()),
        "completos":      int((~sin_crp & ~sin_cuenta).sum()),
        "multi_retencion": int((out['No. RETENCIONES'] >= 2).sum()),
        "descuadres":     out.loc[out['CUADRE'] != 'OK',
                                  ['Contrato', 'Contratista', 'CUADRE']].to_dict(orient='records'),
        "pendientes":     pendientes.head(200).to_dict(orient='records'),
        "archivo_salida": ruta_salida,
    }


# ── Escritura del Excel preservando ceros a la izquierda ─────────────────────
# Código Bco ('007'), Tipo Cta ('02') y No Cuenta ('0550488...') deben quedar como
# TEXTO: si Excel los interpreta como número se pierden los ceros iniciales y el
# archivo deja de servir para el cargue.
COL_TEXTO = ['Contrato', 'NIT_CC', 'Código Bco', 'Tipo Cta', 'No Cuenta', 'CRP']


def _escribir_excel(df, ruta_salida):
    df = df.copy()
    for c in COL_TEXTO:
        if c in df.columns:
            df[c] = df[c].apply(lambda v: "" if pd.isna(v) else str(v).strip())

    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
        hoja = writer.sheets['Consolidado']
        indices = [df.columns.get_loc(c) + 1 for c in COL_TEXTO if c in df.columns]
        for fila in hoja.iter_rows(min_row=2, max_row=hoja.max_row):
            for i in indices:
                fila[i - 1].number_format = '@'


# ── Persistencia de las bases guardadas ──────────────────────────────────────
def guardar_base(df, ruta_pickle):
    df.to_pickle(ruta_pickle)


def cargar_base(ruta_pickle):
    if os.path.isfile(ruta_pickle):
        return pd.read_pickle(ruta_pickle)
    return None


def info_base(ruta_pickle):
    """Metadatos para mostrar en la interfaz qué base hay guardada."""
    df = cargar_base(ruta_pickle)
    if df is None:
        return {"existe": False}
    import datetime
    ts = os.path.getmtime(ruta_pickle)
    return {
        "existe": True,
        "registros": int(len(df)),
        "actualizada": datetime.datetime.fromtimestamp(ts).strftime("%d/%m/%Y %H:%M"),
    }


# ── Función principal ────────────────────────────────────────────────────────
def ejecutar_consolidado(ruta_reporte, ruta_base_crp_pkl, ruta_base_banco_pkl, ruta_salida):
    """Consolida usando las bases ya guardadas en el servidor."""
    base_crp = cargar_base(ruta_base_crp_pkl)
    if base_crp is None:
        return {"ok": False, "mensaje": "No hay base CRP guardada. Cárgala primero."}

    base_banco = cargar_base(ruta_base_banco_pkl)
    if base_banco is None:
        return {"ok": False, "mensaje": "No hay base bancaria guardada. Cárgala primero."}

    return consolidar(ruta_reporte, base_crp, base_banco, ruta_salida)