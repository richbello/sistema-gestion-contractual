"""
Módulo 1 — Extracción de datos de causación desde PDFs.

Adaptado del script original de Colab: lee un lote de PDFs de "actas de pago" /
causación, extrae los campos clave (contrato, contratista, valores, fechas,
retenciones) y los cruza contra una base general (BASEGEN) en Excel para
completar datos bancarios y de CRP. Produce un Excel consolidado.
"""
import os
import re
import pandas as pd
import pdfplumber
from concurrent.futures import ThreadPoolExecutor, as_completed

RE_CONTRATO = re.compile(r'\bCPS\s*(\d{3}[-\s]\d{4})', re.IGNORECASE)
_F = r'(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{2,4})'
RE_DEL = re.compile(r'\bDEL\b\s*:?\s*' + _F, re.IGNORECASE)
RE_AL = re.compile(r'\bAL\b\s*:?\s*' + _F, re.IGNORECASE)
RE_PAGO_DE = re.compile(r'PAGO\s+(?:N[°º.]?\s*)?(\d+)\s+DE\s+(\d+)', re.IGNORECASE)
RE_PAGO_ALT = re.compile(r'(?:CUOTA|MES|ENTREGA)\s+(\d+)\s+DE\s+(\d+)', re.IGNORECASE)

MESES = {
    'enero': '01', 'febrero': '02', 'marzo': '03', 'abril': '04',
    'mayo': '05', 'junio': '06', 'julio': '07', 'agosto': '08',
    'septiembre': '09', 'octubre': '10', 'noviembre': '11', 'diciembre': '12'
}

IDX_BASE = 4
IDX_PCT = 5
IDX_VALOR = 6

COLUMNAS_BASE = [
    "No. IDENTIFICACIÓN DEL CONTRATISTA",
    "Código Bco", "Tipo Cta", "No Cuenta", "CRP", "SALDO CRP"
]

COL_ORDEN = [
    'Contrato', 'Documento No.', 'Contratista', 'NIT_CC',
    'Valor Bruto', 'BASE RETEICA', 'Valor Reteica',
    'TOTAL DESCUENTOS', 'Neto a Pagar',
    'PERIODO', 'DEL', 'AL', 'PAGO No.',
    'Código Bco', 'Tipo Cta', 'No Cuenta', 'CRP', 'SALDO CRP'
]


def _normalizar_fecha(m1, m2, m3):
    a = ("20" + m3) if len(m3) == 2 else m3
    return f"{m1.zfill(2)}/{m2.zfill(2)}/{a}"


def _fecha_texto_a_numero(dia, mes_texto, anio):
    mes_num = MESES.get(mes_texto.lower().strip(), '00')
    return f"{dia.zfill(2)}/{mes_num}/{anio}"


def _extraer_monto_limpio(celda):
    if celda is None:
        return ""
    s = re.sub(r'[\$\s]', '', str(celda).strip())
    if not s or s == '-':
        return ""
    if re.search(r'\d\.\d{3}', s) and ',' not in s:
        s = s.replace('.', '')
    elif ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '')
    try:
        val = float(s)
        if val == 0:
            return ""
        return str(int(val)) if val == int(val) else str(val)
    except ValueError:
        return ""


def _extraer_porcentaje(celda):
    if celda is None:
        return ""
    s = str(celda).strip().replace('%', '').replace(' ', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return ""


def _get_col(fila_l, idx):
    return fila_l[idx] if len(fila_l) > idx else ""


def _procesar_causacion(pdf_path):
    datos = {
        'Archivo': os.path.basename(pdf_path),
        'Contrato': "N/A",
        'Documento No.': "N/A",
        'Contratista': "N/A",
        'NIT_CC': "N/A",
        'Valor Bruto': "",
        'BASE RETEICA': "",
        'Valor Reteica': "",
        'TOTAL DESCUENTOS': "",
        'Neto a Pagar': "",
        'PERIODO': 0,
        'DEL': "N/A",
        'AL': "N/A",
        'PAGO No.': 0,
    }

    with pdfplumber.open(pdf_path) as pdf:
        tablas = pdf.pages[0].extract_tables()
        texto_full = pdf.pages[0].extract_text() or ""

    texto_norm = re.sub(r'[ \t]*\n[ \t]*', ' ', texto_full)

    m_nit = re.search(
        r'(?:NIT\.?\s*[oO]\s*C\.?C\.?|C\.?C\.?|NIT)\s*[:\s#]*([\d][,\.\d]{6,14})',
        texto_full, re.IGNORECASE
    )
    if m_nit:
        datos['NIT_CC'] = re.sub(r'[^\d]', '', m_nit.group(1))
    else:
        m_nit2 = re.search(r'\b\d{1,3}(?:[.,]\d{3}){2,3}\b', texto_full)
        if m_nit2:
            datos['NIT_CC'] = re.sub(r'[^\d]', '', m_nit2.group(0))

    m_cont = RE_CONTRATO.search(texto_full)
    if m_cont:
        datos['Contrato'] = re.sub(r'\s+', '-', m_cont.group(1).strip())

    for tabla in tablas:
        for fila in tabla:
            fila_l = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
            if len(fila_l) > 1 and "CONTRATISTA" in fila_l[0].upper():
                candidato = fila_l[1].strip()
                if re.search(r'[A-Z]{3,}', candidato):
                    datos['Contratista'] = candidato
                    break

    if datos['Contratista'] == "N/A":
        m_c = re.search(
            r'CONTRATISTA\s*:\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{5,60}?)(?:\s{2,}|\n|NIT|$)',
            texto_full, re.IGNORECASE
        )
        if m_c:
            datos['Contratista'] = m_c.group(1).strip()

    m_pago = RE_PAGO_DE.search(texto_norm) or RE_PAGO_ALT.search(texto_norm)
    if m_pago:
        datos['PAGO No.'] = int(m_pago.group(1))
        datos['PERIODO'] = int(m_pago.group(2))
        datos['Documento No.'] = m_pago.group(0).upper().strip()

    md = RE_DEL.search(texto_norm)
    ma = RE_AL.search(texto_norm)
    if md:
        datos['DEL'] = _normalizar_fecha(md.group(1), md.group(2), md.group(3))
    if ma:
        datos['AL'] = _normalizar_fecha(ma.group(1), ma.group(2), ma.group(3))

    if datos['DEL'] == "N/A":
        m_ini = re.search(
            r'(?:del?\s+)?(\d{1,2})\s+de\s+(\w+)\s+(?:de\s+)?(\d{4})',
            texto_norm, re.IGNORECASE
        )
        if m_ini and m_ini.group(2).lower() in MESES:
            datos['DEL'] = _fecha_texto_a_numero(m_ini.group(1), m_ini.group(2), m_ini.group(3))

    for tabla in tablas:
        for fila in tabla:
            fila_l = [str(c).replace('\n', ' ').strip() if c else "" for c in fila]
            linea_txt = " ".join(fila_l).upper()
            linea_txt = (linea_txt
                         .replace("Á", "A").replace("É", "E")
                         .replace("Í", "I").replace("Ó", "O").replace("Ú", "U"))

            if ("RETEICA" in linea_txt and
                    "RETEIVA" not in linea_txt and
                    "RETEFUENTE" not in linea_txt):
                base_val = _extraer_monto_limpio(_get_col(fila_l, IDX_BASE))
                pct_val = _extraer_porcentaje(_get_col(fila_l, IDX_PCT))
                valor_ret = _extraer_monto_limpio(_get_col(fila_l, IDX_VALOR))
                if base_val:
                    datos['BASE RETEICA'] = base_val
                if pct_val != "":
                    datos['Valor Reteica'] = pct_val
                if valor_ret:
                    datos['TOTAL DESCUENTOS'] = valor_ret

            if re.search(r'VALOR\s+BRUTO', linea_txt):
                v = _extraer_monto_limpio(_get_col(fila_l, IDX_VALOR))
                if v:
                    datos['Valor Bruto'] = v

            if re.search(r'NETO\s+A\s+PAGAR', linea_txt):
                v = _extraer_monto_limpio(_get_col(fila_l, IDX_VALOR))
                if v:
                    datos['Neto a Pagar'] = v

    return datos


def _procesar_safe(pdf_path):
    try:
        return _procesar_causacion(pdf_path), None
    except Exception as e:
        return None, f"{os.path.basename(pdf_path)}: {e}"


def ejecutar_extraccion(ruta_base_excel, rutas_pdfs, ruta_salida_excel, max_workers=8):
    """
    Procesa una lista de PDFs en paralelo, cruza contra BASEGEN y guarda el
    Excel consolidado en ruta_salida_excel.

    Retorna un dict con el resumen del proceso (para mostrar en la interfaz).
    """
    resultados = []
    errores = []
    total = len(rutas_pdfs)

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

    df_out = pd.DataFrame(resultados)[[
        'Contrato', 'Documento No.', 'Contratista', 'NIT_CC',
        'Valor Bruto', 'BASE RETEICA', 'Valor Reteica',
        'TOTAL DESCUENTOS', 'Neto a Pagar',
        'PERIODO', 'DEL', 'AL', 'PAGO No.'
    ]].copy()

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
    df_base['SALDO CRP'] = pd.to_numeric(df_base['SALDO CRP'], errors='coerce').fillna(0).astype('Int64')

    df_base['No. IDENTIFICACIÓN DEL CONTRATISTA'] = (
        df_base['No. IDENTIFICACIÓN DEL CONTRATISTA'].astype(str).str.replace(r'[^\d]', '', regex=True).str.strip()
    )

    duplicados = df_base['No. IDENTIFICACIÓN DEL CONTRATISTA'].value_counts()
    duplicados = duplicados[duplicados > 1]
    df_base = df_base.drop_duplicates(subset=['No. IDENTIFICACIÓN DEL CONTRATISTA'], keep='first')

    df_out['NIT_CC'] = df_out['NIT_CC'].astype(str).str.replace(r'[^\d]', '', regex=True)

    df_final = pd.merge(
        df_out, df_base,
        left_on='NIT_CC', right_on='No. IDENTIFICACIÓN DEL CONTRATISTA',
        how='left'
    )

    if 'No. IDENTIFICACIÓN DEL CONTRATISTA' in df_final.columns:
        df_final.drop(columns=['No. IDENTIFICACIÓN DEL CONTRATISTA'], inplace=True)

    df_final['_num_contrato'] = pd.to_numeric(
        df_final['Contrato'].str.extract(r'^(\d+)-\d+$')[0], errors='coerce'
    )
    df_final = (df_final
                .sort_values('_num_contrato', na_position='last')
                .drop(columns=['_num_contrato'])
                .reset_index(drop=True))

    col_orden = [c for c in COL_ORDEN if c in df_final.columns]
    df_final = df_final[col_orden]

    df_final.to_excel(ruta_salida_excel, index=False)

    return {
        "ok": True,
        "total_pdfs": total,
        "procesados": len(resultados),
        "errores": errores,
        "nits_duplicados_base": int(len(duplicados)),
        "registros": len(df_final),
        "archivo_salida": ruta_salida_excel,
    }
