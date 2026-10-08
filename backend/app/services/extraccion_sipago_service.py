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

# ── Posiciones en el Reporte del Periodo (columnas sin encabezado) ───────────
REP_IDX_HONOR_VALOR = 16   # retención adicional (valor)
REP_IDX_HONOR_BASE  = 18   # retención adicional (base)
REP_IDX_RETEICA_VAL = 28   # valor ReteICA
REP_IDX_RETEICA_PCT = 29   # % ReteICA  -> va en 'Valor Reteica' (igual que Módulo 01)
REP_IDX_RETEICA_BASE= 30   # base ReteICA

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
def construir_base_bancaria(rutas_historicos):
    """Lee uno o varios históricos de pago y extrae la cuenta de cada contratista.

    La columna de cuenta trae el formato: '<numero> AHORROS|CORRIENTE <banco>'.
    Los archivos se procesan en el orden recibido; el ÚLTIMO gana, de modo que
    basta con enviarlos de más antiguo a más reciente.
    """
    bloques = []
    omitidos = []

    for orden, ruta in enumerate(rutas_historicos):
        nombre_arch = os.path.basename(ruta)
        try:
            df = pd.read_excel(ruta)
        except Exception as e:
            omitidos.append(f"{nombre_arch}: no se pudo leer ({e})")
            continue

        # Columna de cuenta = la que contenga el patrón AHORROS/CORRIENTE
        col_cuenta = None
        for c in df.columns:
            v = df[c].dropna().astype(str)
            if len(v) and v.str.match(RE_CUENTA).sum() > 0:
                col_cuenta = c
                break
        col_id = _buscar_col(df, 'identificación', 'identificacion', 'nit', 'cedula', 'documento')
        if col_id is None:
            omitidos.append(f"{nombre_arch}: sin columna de identificación")
            continue

        # Formato A: columna unica "<numero> AHORROS|CORRIENTE <banco>"
        if col_cuenta is not None:
            col_cod = _buscar_col(df, 'codigo de banco', 'código de banco')
            sub = df[df[col_cuenta].notna()].copy()
            partes = sub[col_cuenta].astype(str).str.extract(RE_CUENTA)
            bloque = pd.DataFrame({
                'cedula':     sub[col_id].apply(_digitos),
                'No Cuenta':  partes[0].astype(str).str.replace(r'[^\d]', '', regex=True),
                'Tipo Cta':   partes[1].str.upper().map(TIPO_CTA),
                'banco_txt':  partes[2].astype(str).str.strip(),
                'cod_banco':  _num(sub[col_cod]) if col_cod else pd.NA,
                '_orden':     orden,
            })
        else:
            # Formato B: columnas separadas (No Cuenta, Tipo Cta, Codigo, Banco)
            col_num = _buscar_col(df, 'no cuenta', 'nº cuenta', 'n° cuenta',
                                  'numero de cuenta', 'número de cuenta', 'numero cuenta')
            if col_num is None:
                omitidos.append(f"{nombre_arch}: sin columna de cuenta bancaria")
                continue
            col_tipo = _buscar_col(df, 'tipo cta', 'tipo de cuenta', 'tipo cuenta')
            col_cod  = _buscar_col(df, 'codigo', 'código')
            col_bco  = _buscar_col(df, 'nombre del banco', 'banco')

            def _tipo(v):
                s = str(v).strip().upper().replace('.0', '')
                if s in ('01', '02'):
                    return s
                if s in TIPO_CTA:
                    return TIPO_CTA[s]
                return s.zfill(2) if s.isdigit() else ''

            bloque = pd.DataFrame({
                'cedula':     df[col_id].apply(_digitos),
                'No Cuenta':  df[col_num].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'[^\d]', '', regex=True),
                'Tipo Cta':   df[col_tipo].apply(_tipo) if col_tipo is not None else '02',
                'banco_txt':  df[col_bco].astype(str).str.strip() if col_bco is not None else '',
                'cod_banco':  _num(df[col_cod]) if col_cod is not None else pd.NA,
                '_orden':     orden,
            })

        bloque = bloque[(bloque['cedula'].str.len() > 0) &
                        (bloque['No Cuenta'].str.len() > 0)]
        bloques.append(bloque)

    if not bloques:
        return pd.DataFrame(columns=['cedula', 'Código Bco', 'Tipo Cta', 'No Cuenta']), omitidos

    base = pd.concat(bloques, ignore_index=True)

    # Mapa banco -> código, aprendido de los archivos que sí traen 'Codigo de Banco'.
    # Sirve para completar el código en los archivos que no lo incluyen.
    con_cod = base.dropna(subset=['cod_banco'])
    mapa = {}
    if len(con_cod):
        mapa = (con_cod.groupby('banco_txt')['cod_banco']
                       .agg(lambda s: int(s.mode().iloc[0]))
                       .to_dict())

    def _cod(row):
        v = row['cod_banco']
        if pd.isna(v):
            v = mapa.get(row['banco_txt'])
        if v is None or pd.isna(v):
            return ""
        try:
            return str(int(float(v))).zfill(3)
        except (TypeError, ValueError):
            return ""

    base['Código Bco'] = base.apply(_cod, axis=1)

    # Regla acordada: la cuenta MÁS RECIENTE (último archivo en el que aparece).
    base = (base.sort_values('_orden', ascending=False)
                .drop_duplicates(subset=['cedula'], keep='first')
                .reset_index(drop=True))

    return base[['cedula', 'Código Bco', 'Tipo Cta', 'No Cuenta']], omitidos


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

    honor_base  = _num(_col(rep, REP_IDX_HONOR_BASE))
    honor_valor = _num(_col(rep, REP_IDX_HONOR_VALOR))
    # Solo se reporta honorario cuando efectivamente hay un valor retenido.
    hay_honor = honor_valor.fillna(0) > 0

    out = pd.DataFrame({
        'Contrato':         rep[col_contr].apply(_fmt_contrato) if col_contr else "",
        'Documento No.':    rep[col_pago].astype(str).str.strip() if col_pago else "",
        'Contratista':      rep[col_nombre].astype(str).str.strip(),
        'NIT_CC':           cedulas,
        'Valor Bruto':      _num(rep[col_bruto]) if col_bruto else 0,
        'BASE RETEICA':     _num(_col(rep, REP_IDX_RETEICA_BASE)),
        'Valor Reteica':    _col(rep, REP_IDX_RETEICA_PCT).apply(_pct),
        'TOTAL DESCUENTOS': _num(rep[col_desc]) if col_desc else 0,
        'Neto a Pagar':     _num(rep[col_neto]) if col_neto else 0,
        # PERIODO en el Módulo 01 es el TOTAL de pagos del contrato ('PAGO 8 DE 12' -> 12).
        # El Reporte del Periodo no trae ese total, así que la columna queda vacía.
        'PERIODO':          "",
        'DEL':              fechas.apply(lambda t: t[0]),
        'AL':               fechas.apply(lambda t: t[1]),
        'PAGO No.':         (rep[col_pago].astype(str).str.extract(RE_PAGO)[0]
                             if col_pago else ""),
        'HONOR_BASE':       honor_base.where(hay_honor),
        'HONOR_VALOR':      honor_valor.where(hay_honor),
    })

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
    out = out[COL_ORDEN_FINAL]

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