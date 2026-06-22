"""
Módulo 1 — Extracción de causaciones desde PDFs.
Versión actualizada: detecta retenciones adicionales (HONOR_BASE/HONOR_VALOR)
además del ReteICA, para que el Módulo 02 genere filas extra con código 44.
"""
import os
import re
import pandas as pd
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── Constantes ────────────────────────────────────────────────────────────────
IDX_CONCEPTO = 0
IDX_BASE     = 4
IDX_PCT      = 5
IDX_VALOR    = 6

KEYS_RETENCION_ADICIONAL = ['RENTA', 'HONORARIO', 'RETEFUENTE', 'CONTRIBUCION']
KEY_RETEICA  = 'RETEICA'
KEY_RETEIVA  = 'RETEIVA'

MESES = {
    'enero':'01','febrero':'02','marzo':'03','abril':'04',
    'mayo':'05','junio':'06','julio':'07','agosto':'08',
    'septiembre':'09','octubre':'10','noviembre':'11','diciembre':'12'
}

RE_CONTRATO = re.compile(r'\bCPS\s*(\d{3}[-\s]\d{4})', re.IGNORECASE)
_F          = r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})'
RE_DEL      = re.compile(r'\bDEL\b\s*:?\s*' + _F, re.IGNORECASE)
RE_AL       = re.compile(r'\bAL\b\s*:?\s*'  + _F, re.IGNORECASE)
RE_PAGO_DE  = re.compile(r'PAGO\s+(?:N[°º.]?\s*)?(\d+)\s+DE\s+(\d+)', re.IGNORECASE)
RE_PAGO_ALT = re.compile(r'(?:CUOTA|MES|ENTREGA)\s+(\d+)\s+DE\s+(\d+)', re.IGNORECASE)

COLUMNAS_BASE = [
    "No. IDENTIFICACIÓN DEL CONTRATISTA",
    "Código Bco", "Tipo Cta", "No Cuenta", "CRP", "SALDO CRP"
]

COL_ORDEN_FINAL = [
    'Contrato', 'Documento No.', 'Contratista', 'NIT_CC',
    'Valor Bruto', 'BASE RETEICA', 'Valor Reteica',
    'TOTAL DESCUENTOS', 'Neto a Pagar',
    'PERIODO', 'DEL', 'AL', 'PAGO No.',
    'HONOR_BASE', 'HONOR_VALOR',
    'Código Bco', 'Tipo Cta', 'No Cuenta', 'CRP', 'SALDO CRP'
]

# ── Funciones auxiliares ──────────────────────────────────────────────────────
def _normalizar(texto):
    return (texto.upper()
            .replace("Á","A").replace("É","E")
            .replace("Í","I").replace("Ó","O").replace("Ú","U"))

def _fmt_fecha(m1, m2, m3):
    a = ("20" + m3) if len(m3) == 2 else m3
    return f"{m1.zfill(2)}/{m2.zfill(2)}/{a}"

def _fecha_texto(dia, mes_texto, anio):
    return f"{dia.zfill(2)}/{MESES.get(mes_texto.lower().strip(),'00')}/{anio}"

def _monto(celda):
    if celda is None:
        return ""
    s = re.sub(r'[\$\s]', '', str(celda).strip())
    if not s or s == '-':
        return ""
    if re.search(r'\d\.\d{3}', s) and ',' not in s:
        s = s.replace('.', '')
    elif ',' in s and '.' in s:
        s = (s.replace('.','').replace(',','.') if s.rfind(',') > s.rfind('.')
             else s.replace(',',''))
    elif ',' in s:
        s = s.replace(',', '')
    try:
        val = float(s)
        return "" if val == 0 else (str(int(val)) if val == int(val) else str(val))
    except ValueError:
        return ""

def _pct(celda):
    if celda is None:
        return ""
    s = str(celda).strip().replace('%','').replace(' ','').replace(',','.')
    try:
        return float(s)
    except ValueError:
        return ""

def _gcol(fila_l, idx):
    return fila_l[idx] if len(fila_l) > idx else ""

# ── Procesamiento de un PDF ───────────────────────────────────────────────────
def _procesar_causacion(pdf_path):
    datos = {
        'Archivo'         : os.path.basename(pdf_path),
        'Contrato'        : "N/A",
        'Documento No.'   : "N/A",
        'Contratista'     : "N/A",
        'NIT_CC'          : "N/A",
        'Valor Bruto'     : "",
        'BASE RETEICA'    : "",
        'Valor Reteica'   : "",
        'TOTAL DESCUENTOS': "",
        'Neto a Pagar'    : "",
        'PERIODO'         : 0,
        'DEL'             : "N/A",
        'AL'              : "N/A",
        'PAGO No.'        : 0,
        'HONOR_BASE'      : "",
        'HONOR_VALOR'     : "",
    }

    with pdfplumber.open(pdf_path) as pdf:
        tablas     = pdf.pages[0].extract_tables()
        texto_full = pdf.pages[0].extract_text() or ""

    texto_norm = re.sub(r'[ \t]*\n[ \t]*', ' ', texto_full)

    # NIT / CC
    m = re.search(
        r'(?:NIT\.?\s*[oO]\s*C\.?C\.?|C\.?C\.?|NIT)\s*[:\s#]*([\d][,\.\d]{6,14})',
        texto_full, re.IGNORECASE
    )
    if m:
        datos['NIT_CC'] = re.sub(r'[^\d]', '', m.group(1))
    else:
        m2 = re.search(r'\b\d{1,3}(?:[.,]\d{3}){2,3}\b', texto_full)
        if m2:
            datos['NIT_CC'] = re.sub(r'[^\d]', '', m2.group(0))

    # Contrato
    m = RE_CONTRATO.search(texto_full)
    if m:
        datos['Contrato'] = re.sub(r'\s+', '-', m.group(1).strip())

    # Contratista
    for tabla in tablas:
        for fila in tabla:
            fila_l = [str(c).replace('\n',' ').strip() if c else "" for c in fila]
            if len(fila_l) > 1 and "CONTRATISTA" in fila_l[0].upper():
                candidato = fila_l[1].strip()
                if re.search(r'[A-Z]{3,}', candidato):
                    datos['Contratista'] = candidato
                    break
    if datos['Contratista'] == "N/A":
        m = re.search(
            r'CONTRATISTA\s*:\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{5,60}?)(?:\s{2,}|\n|NIT|$)',
            texto_full, re.IGNORECASE
        )
        if m:
            datos['Contratista'] = m.group(1).strip()

    # PAGO No. / PERIODO
    m = RE_PAGO_DE.search(texto_norm) or RE_PAGO_ALT.search(texto_norm)
    if m:
        datos['PAGO No.']      = int(m.group(1))
        datos['PERIODO']       = int(m.group(2))
        datos['Documento No.'] = m.group(0).upper().strip()

    # Fechas DEL / AL
    md = RE_DEL.search(texto_norm)
    ma = RE_AL.search(texto_norm)
    if md:
        datos['DEL'] = _fmt_fecha(md.group(1), md.group(2), md.group(3))
    if ma:
        datos['AL']  = _fmt_fecha(ma.group(1), ma.group(2), ma.group(3))
    if datos['DEL'] == "N/A":
        m = re.search(
            r'(\d{1,2})\s+de\s+(\w+)\s+(?:de\s+)?(\d{4})',
            texto_norm, re.IGNORECASE
        )
        if m and m.group(2).lower() in MESES:
            datos['DEL'] = _fecha_texto(m.group(1), m.group(2), m.group(3))

    # Parseo de tabla de retenciones
    rets_adicionales = []

    for tabla in tablas:
        for fila in tabla:
            fila_l    = [str(c).replace('\n',' ').strip() if c else "" for c in fila]
            linea_raw = " ".join(fila_l)
            linea     = _normalizar(linea_raw)

            base_v  = _monto(_gcol(fila_l, IDX_BASE))
            pct_v   = _pct(_gcol(fila_l, IDX_PCT))
            valor_v = _monto(_gcol(fila_l, IDX_VALOR))

            if re.search(r'VALOR\s+BRUTO', linea):
                if valor_v:
                    datos['Valor Bruto'] = valor_v

            elif re.search(r'NETO\s+A\s+PAGAR', linea):
                if valor_v:
                    datos['Neto a Pagar'] = valor_v

            elif re.search(r'TOTAL\s+DESCUENTOS', linea):
                if valor_v and not datos['TOTAL DESCUENTOS']:
                    datos['TOTAL DESCUENTOS'] = valor_v

            elif (KEY_RETEICA in linea
                  and KEY_RETEIVA not in linea
                  and 'RETEFUENTE' not in linea):
                if base_v:
                    datos['BASE RETEICA'] = base_v
                if pct_v != "":
                    datos['Valor Reteica'] = pct_v
                if valor_v:
                    datos['TOTAL DESCUENTOS'] = valor_v

            elif any(k in linea for k in KEYS_RETENCION_ADICIONAL):
                if base_v and valor_v:
                    rets_adicionales.append((base_v, valor_v))

    if len(rets_adicionales) >= 1:
        datos['HONOR_BASE']  = rets_adicionales[0][0]
        datos['HONOR_VALOR'] = rets_adicionales[0][1]

    return datos


def _procesar_safe(pdf_path):
    try:
        return _procesar_causacion(pdf_path), None
    except Exception as e:
        return None, f"{os.path.basename(pdf_path)}: {e}"


# ── Función principal exportada ───────────────────────────────────────────────
def ejecutar_extraccion(ruta_base_excel, rutas_pdfs, ruta_salida_excel, max_workers=8):
    resultados = []
    errores    = []
    total      = len(rutas_pdfs)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futuros = {executor.submit(_procesar_safe, p): p for p in rutas_pdfs}
        for futuro in as_completed(futuros):
            info, err = futuro.result()
            if info:
                resultados.append(info)
            if err:
                errores.append(err)

    if not resultados:
        return {"ok": False, "mensaje": "No se extrajo ningún dato de los PDFs.", "errores": errores}

    col_orden_base = [
        'Contrato', 'Documento No.', 'Contratista', 'NIT_CC',
        'Valor Bruto', 'BASE RETEICA', 'Valor Reteica',
        'TOTAL DESCUENTOS', 'Neto a Pagar',
        'PERIODO', 'DEL', 'AL', 'PAGO No.',
        'HONOR_BASE', 'HONOR_VALOR',
    ]
    df_out = pd.DataFrame(resultados)
    for col in col_orden_base:
        if col not in df_out.columns:
            df_out[col] = ""
    df_out = df_out[col_orden_base].copy()

    df_out['NIT_CC'] = (
        pd.to_numeric(
            df_out['NIT_CC'].astype(str).str.replace(r'[^\d]', '', regex=True),
            errors='coerce'
        ).fillna(0).astype('Int64')
    )

    df_base = pd.read_excel(
        ruta_base_excel,
        usecols=COLUMNAS_BASE,
        dtype={
            "No. IDENTIFICACIÓN DEL CONTRATISTA": str,
            "Código Bco": str,
            "No Cuenta": str,
            "CRP": str,
        }
    )
    df_base['Código Bco'] = (
        df_base['Código Bco'].astype(str).str.replace(r'[^\d]', '', regex=True).str.zfill(3)
    )
    df_base['Tipo Cta'] = "02"
    df_base['No Cuenta'] = (
        df_base['No Cuenta'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    )
    df_base['CRP'] = (
        df_base['CRP'].astype(str).str.replace(r'\.0$', '', regex=True)
        .str.replace(r'[^\d]', '', regex=True).str.strip()
    )
    df_base['SALDO CRP'] = (
        pd.to_numeric(df_base['SALDO CRP'], errors='coerce').fillna(0).astype('Int64')
    )
    df_base['No. IDENTIFICACIÓN DEL CONTRATISTA'] = (
        df_base['No. IDENTIFICACIÓN DEL CONTRATISTA']
        .astype(str).str.replace(r'[^\d]', '', regex=True).str.strip()
    )

    duplicados = df_base['No. IDENTIFICACIÓN DEL CONTRATISTA'].value_counts()
    duplicados = duplicados[duplicados > 1]
    df_base = df_base.drop_duplicates(
        subset=['No. IDENTIFICACIÓN DEL CONTRATISTA'], keep='first'
    )

    df_out['NIT_CC'] = df_out['NIT_CC'].astype(str).str.replace(r'[^\d]', '', regex=True)

    df_final = pd.merge(
        df_out, df_base,
        left_on='NIT_CC', right_on='No. IDENTIFICACIÓN DEL CONTRATISTA',
        how='left'
    )
    if 'No. IDENTIFICACIÓN DEL CONTRATISTA' in df_final.columns:
        df_final.drop(columns=['No. IDENTIFICACIÓN DEL CONTRATISTA'], inplace=True)

    df_final['_num'] = pd.to_numeric(
        df_final['Contrato'].str.extract(r'^(\d+)')[0], errors='coerce'
    )
    df_final = (df_final.sort_values('_num', na_position='last')
                .drop(columns=['_num']).reset_index(drop=True))

    col_final = [c for c in COL_ORDEN_FINAL if c in df_final.columns]
    df_final  = df_final[col_final]

    df_final.to_excel(ruta_salida_excel, index=False)

    con_honor = (
        df_final['HONOR_VALOR'].astype(str).str.strip()
        .replace('', pd.NA).notna().sum()
        if 'HONOR_VALOR' in df_final.columns else 0
    )

    return {
        "ok"                  : True,
        "total_pdfs"          : total,
        "procesados"          : len(resultados),
        "errores"             : errores,
        "nits_duplicados_base": int(len(duplicados)),
        "registros"           : len(df_final),
        "con_retencion_44"    : int(con_honor),
        "archivo_salida"      : ruta_salida_excel,
    }
