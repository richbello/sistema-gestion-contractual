"""
Cuadre PAC · Servicios Públicos y Personas Jurídicas (módulo aparte).

Formato real de la plantilla (consolidado de servicios públicos):
    C   -> encabezado del pago (factura / referencia y servicio en J-K)
    P40 -> una imputación: 'RP Doc Presupuestal' (N° Interno CRP 5001...),
           'Posc Doc Pres' (posición 1, 2, 3...) e 'importe'
    P31 -> tercero (NIT) e importe TOTAL del pago = suma de sus P40
           (p.ej. 77.906 + 54.374 = 132.280)

Proceso:
 1. Se extrae cada P40: CRP + posición + importe. Llave de cruce
    crp = RP Doc Presupuestal & Posc Doc Pres  (5001034181 & 1 -> 50010341811).
 2. Con el Reporte CRP (adjunto en cada cuadre, porque cambia) se trae Rubro y
    Fondos por N° Interno CRP + N° Posición CRP.
 3. Se valida que la suma de las P40 de cada pago sea igual al total de la P31.
 4. Se agrupa por Rubro+Fondos y se cruza contra el PAC con la MISMA regla del
    Cuadre PAC de personas naturales (falta / sobra / cubre acumulado).

Salida Excel: hoja 'Cuadre' (Rubro+Fondos vs PAC) y hoja 'Detalle CRP'
(RP Doc Presupuestal | Posc Doc Pres | crp | importe | rubro | fondos | ...).
"""
import pandas as pd

from .cuadre_pac_service import (
    _norm_rubro, _normalizar_periodo, _parse_mes_numero, _convertir_tipos,
    COLS_PAC_REQUERIDAS,
)
from .enriquecer_rubros import _buscar_col

COLS_DETALLE = ['RP Doc Presupuestal', 'Posc Doc Pres', 'crp', 'importe', 'rubro', 'fondos',
                'Pago N°', 'Referencia / servicio', 'NIT', 'Total pago (P31)', 'Validación pago']


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def _vacio(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ('', 'nan', 'None', 'NaT')


def _txt(v):
    return None if _vacio(v) else str(v).strip()


def _digitos(v):
    """Número como texto de solo dígitos (sin '.0'): 5001034181.0 -> '5001034181'."""
    if _vacio(v):
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return ''.join(ch for ch in s if ch.isdigit())


def _pos(v):
    d = _digitos(v).lstrip('0')
    return d or ''


def _num(v):
    """Importe robusto: 1437330 | '1437330' | '1.437.330' | '$ 1.437.330' | '1437330,00'."""
    if _vacio(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace('$', '').replace(' ', '').strip()
    if ',' in s:                                   # formato es-CO: 1.437.330,50
        s = s.replace('.', '').replace(',', '.')
    elif s.count('.') > 1 or (s.count('.') == 1 and len(s.split('.')[1]) == 3):
        s = s.replace('.', '')                     # miles con punto
    try:
        return float(s)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Lectura de archivos
# --------------------------------------------------------------------------
def _hojas_visibles(ruta):
    """(hojas visibles con la ACTIVA primero, {hoja: n filas ocultas}).
    Las hojas ocultas se ignoran: suelen ser consolidados viejos."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(ruta, read_only=False, data_only=True)
    except Exception:
        return None, {}
    visibles = [ws.title for ws in wb.worksheets if ws.sheet_state == 'visible']
    activa = wb.active.title if wb.active is not None else None
    if activa in visibles:
        visibles.remove(activa)
        visibles.insert(0, activa)
    ocultas = {ws.title: sum(1 for d in ws.row_dimensions.values() if d.hidden)
               for ws in wb.worksheets if ws.title in visibles}
    return visibles, ocultas


def _leer_plantilla(ruta):
    """Lee la hoja ACTIVA y visible que tenga la fila de encabezados con
    'Tipo Registro' e 'importe' (puede no estar en la fila 1).
    Devuelve (df, hoja, filas_ocultas)."""
    xls = pd.ExcelFile(ruta)
    visibles, ocultas = _hojas_visibles(ruta)
    hojas = visibles if visibles else xls.sheet_names
    for hoja in hojas:
        crudo = pd.read_excel(xls, sheet_name=hoja, header=None, nrows=15)
        for i in range(len(crudo)):
            fila = [str(x).upper() for x in crudo.iloc[i].tolist()]
            if any('TIPO REGISTRO' in x for x in fila) and any('IMPORTE' in x for x in fila):
                df = pd.read_excel(xls, sheet_name=hoja, header=i, dtype=object)
                df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
                return df, hoja, ocultas.get(hoja, 0)
    raise KeyError("No se encontró en ninguna hoja visible la fila de encabezados con "
                   "'Tipo Registro P' e 'importe'.")


def _leer_pac(ruta):
    xls = pd.ExcelFile(ruta)
    if 'Data' not in xls.sheet_names:
        raise KeyError("El reporte PAC no tiene la hoja 'Data'.")
    return pd.read_excel(xls, sheet_name='Data')


def _mapa_crp(ruta_crp):
    """{(crp, posicion): (rubro, fondos)} y {crp: (rubro, fondos)} (primera posición)."""
    df = pd.read_excel(ruta_crp, dtype=object)
    c_int = _buscar_col(df.columns, 'N Interno CRP') or _buscar_col(df.columns, 'Interno CRP')
    c_pos = _buscar_col(df.columns, 'N Posición CRP') or _buscar_col(df.columns, 'Posición CRP')
    c_rub = _buscar_col(df.columns, 'Rubro')
    c_fon = _buscar_col(df.columns, 'Fondos')
    faltan = [n for n, c in [('N° Interno CRP', c_int), ('Rubro', c_rub), ('Fondos', c_fon)] if c is None]
    if faltan:
        raise KeyError(f"No se encontraron estas columnas en el Reporte CRP: {faltan}")

    por_pos, por_crp = {}, {}
    for _, r in df.iterrows():
        crp = _digitos(r[c_int])
        if not crp:
            continue
        par = (_txt(r[c_rub]), _txt(r[c_fon]))
        if c_pos is not None:
            por_pos.setdefault((crp, _pos(r[c_pos])), par)
        por_crp.setdefault(crp, par)
    return por_pos, por_crp, c_pos is not None


# --------------------------------------------------------------------------
# Extracción de imputaciones P40
# --------------------------------------------------------------------------
def _col(df, *nombres):
    for n in nombres:
        c = _buscar_col(df.columns, n)
        if c is not None:
            return c
    return None


def _extraer(plantilla, por_pos, por_crp, hay_pos):
    c_tipo = _col(plantilla, 'Tipo Registro P', 'Tipo Registro')
    c_clave = _col(plantilla, 'Clave Contab', 'Clave')
    c_imp = _col(plantilla, 'importe')
    c_rp = _col(plantilla, 'RP Doc Presupuestal', 'Doc Presupuestal')
    c_pos = _col(plantilla, 'Posc Doc Pres', 'Pos Doc Pres', 'Posc Doc')
    if c_pos is None and c_rp is not None:
        # en el consolidado la posición no tiene encabezado: es la columna
        # inmediatamente a la derecha de 'RP Doc Presupuestal'
        i = list(plantilla.columns).index(c_rp)
        if i + 1 < len(plantilla.columns):
            c_pos = plantilla.columns[i + 1]
    c_nit = _col(plantilla, 'No Identificación', 'Identificación')
    faltan = [n for n, c in [('Tipo Registro P', c_tipo), ('Clave Contab.', c_clave), ('importe', c_imp),
                             ('RP Doc Presupuestal', c_rp), ('Posc Doc Pres', c_pos)] if c is None]
    if faltan:
        raise KeyError(f"Faltan columnas en la plantilla: {faltan}")

    filas, bloque, n_pago, totales_p31 = [], None, 0, []

    def cerrar():
        if not bloque or not bloque['p40']:
            return
        totales_p31.append(bloque['p31'])
        suma = sum(f['importe'] for f in bloque['p40'])
        total = bloque['p31']
        if total is None:
            val = 'SIN P31'
        elif abs(suma - total) < 1:
            val = 'OK'
        else:
            val = f'DIFERENCIA {suma - total:,.0f}'
        for f in bloque['p40']:
            f['NIT'] = bloque['nit'] or ''
            f['Total pago (P31)'] = total
            f['Validación pago'] = val
        filas.extend(bloque['p40'])

    for _, r in plantilla.iterrows():
        tipo = str(r[c_tipo]).strip().upper()
        clave = _digitos(r[c_clave])
        if tipo == 'C':
            cerrar()
            n_pago += 1
            ref = ' '.join(x for x in (_txt(r[c_rp]), _txt(r[c_pos])) if x)
            bloque = {'n': n_pago, 'ref': ref, 'nit': None, 'p31': None, 'p40': []}
        elif bloque is None:
            continue
        elif clave == '40':
            crp, pos, imp = _digitos(r[c_rp]), _pos(r[c_pos]), _num(r[c_imp])
            if not crp or imp is None:
                continue
            rub, fon = por_pos.get((crp, pos), (None, None)) if hay_pos else (None, None)
            if rub is None and not hay_pos:
                rub, fon = por_crp.get(crp, (None, None))
            bloque['p40'].append({
                'RP Doc Presupuestal': crp, 'Posc Doc Pres': pos, 'crp': crp + pos,
                'importe': imp, 'rubro': rub, 'fondos': fon,
                'Pago N°': bloque['n'], 'Referencia / servicio': bloque['ref'],
            })
        elif clave == '31':
            if bloque['nit'] is None and c_nit is not None:
                bloque['nit'] = _digitos(r[c_nit]) or None
            if bloque['p31'] is None:
                bloque['p31'] = _num(r[c_imp])
    cerrar()
    total_p31 = sum(t for t in totales_p31 if t is not None)
    return _consolidar(pd.DataFrame(filas, columns=COLS_DETALLE)), n_pago, total_p31


def _unir(valores):
    vistos = []
    for v in valores:
        v = '' if _vacio(v) else str(v)
        if v and v not in vistos:
            vistos.append(v)
    return ' | '.join(vistos)


def _consolidar(detalle):
    """Un mismo CRP + posición (aunque venga en varias filas P40 o en varios
    pagos) queda en UNA fila con el importe sumado."""
    if detalle.empty:
        return detalle
    detalle = detalle.copy()
    detalle['_filas'] = 1
    g = detalle.groupby('crp', sort=False, as_index=False).agg({
        'RP Doc Presupuestal': 'first', 'Posc Doc Pres': 'first', 'importe': 'sum',
        'rubro': 'first', 'fondos': 'first', 'Pago N°': _unir,
        'Referencia / servicio': _unir, 'NIT': _unir,
        'Total pago (P31)': lambda x: x.drop_duplicates().sum(min_count=1),
        'Validación pago': lambda x: 'OK' if (x == 'OK').all() else _unir(v for v in x if v != 'OK'),
        '_filas': 'sum',
    })
    g['Filas P40'] = g.pop('_filas')
    return g[COLS_DETALLE + ['Filas P40']]


# --------------------------------------------------------------------------
# Cobertura PAC (misma regla del Cuadre PAC de personas naturales)
# --------------------------------------------------------------------------
def _preparar_pac(pac):
    falt = COLS_PAC_REQUERIDAS - set(pac.columns)
    if falt:
        raise KeyError(f"Faltan columnas en el PAC: {falt}")
    pac['Disponibilidad PAC'] = pd.to_numeric(pac['Disponibilidad PAC'], errors='coerce').fillna(0)
    pac['_periodo_norm'] = pac['Descripción período presupuesto'].apply(_normalizar_periodo)
    pac['_progr_norm'] = pac['Progr.financiación'].apply(_norm_rubro)
    pac['_pos_norm'] = pac['Pos.presupuestaria'].apply(_norm_rubro) if 'Pos.presupuestaria' in pac.columns else ''
    pac['_fondo_norm'] = pac['Fondos'].astype(str).str.strip().str.upper()
    pac['_es_mes'] = pac['_periodo_norm'].apply(lambda p: _parse_mes_numero(p) is not None)
    return pac


def _cobertura(pac, consolidado, mes_pac, periodos):
    mes_num = _parse_mes_numero(mes_pac)
    mes_norm = _normalizar_periodo(mes_pac)
    tiene_pos = 'Pos.presupuestaria' in pac.columns
    salida = []
    for _, row in consolidado.iterrows():
        rub, fon, imp = str(row['Rubro']).strip(), str(row['Fondos']).strip(), float(row['Importe_Total'])
        rn = _norm_rubro(rub)
        bf = pac['_fondo_norm'] == fon.upper()
        pac_rf = pac[(pac['_progr_norm'] == rn) & bf]
        if pac_rf.empty and tiene_pos:
            pac_rf = pac[(pac['_pos_norm'] == rn) & bf]

        meses = pac_rf[pac_rf['_es_mes']]
        pac_mes = float(meses[meses['_periodo_norm'] == mes_norm]['Disponibilidad PAC'].sum())
        acumulable = {}
        for p in periodos:
            n = _parse_mes_numero(p)
            if n is not None and n >= mes_num:
                acumulable[p] = float(meses[meses['_periodo_norm'] == _normalizar_periodo(p)]['Disponibilidad PAC'].sum())

        if pac_rf.empty:
            cob = "SIN MATCH EN PAC"
        elif float(meses['Disponibilidad PAC'].sum()) <= 0:
            cob = "SIN DISPONIBILIDAD PAC"
        elif pac_mes >= imp:
            cob = f"OK - CUBRE EN {mes_pac.upper()}"
        else:
            acum, usados = 0.0, []
            for p, monto in sorted(acumulable.items(), key=lambda x: _parse_mes_numero(x[0]) or 999):
                if monto <= 0:
                    continue
                acum += monto
                usados.append(p.upper())
                if acum >= imp:
                    break
            if acum >= imp:
                cob = (f"FALTA ${imp - pac_mes:,.0f} EN {mes_pac.upper()}, CUBRE acum. {' + '.join(usados)}"
                       if pac_mes > 0 else f"OK - CUBRE acum. {' + '.join(usados)}")
            else:
                cob = f"FALTA ${imp - acum:,.0f}, NO CUBRE"

        salida.append({'Rubro': rub, 'Fondos': fon, 'Importe_Total': int(round(imp)),
                       'Disponibilidad_PAC': int(round(pac_mes)),
                       'Diferencia': int(round(pac_mes - imp)), 'Cobertura': cob})
    return pd.DataFrame(salida, columns=['Rubro', 'Fondos', 'Importe_Total',
                                         'Disponibilidad_PAC', 'Diferencia', 'Cobertura'])


# --------------------------------------------------------------------------
# Punto de entrada
# --------------------------------------------------------------------------
def cuadrar_pac_sj(ruta_plantilla, ruta_pac, mes_pac, ruta_salida_excel, ruta_crp):
    if _parse_mes_numero(mes_pac) is None:
        raise ValueError(f"Mes '{mes_pac}' no reconocido.")

    plantilla, hoja, filas_ocultas = _leer_plantilla(ruta_plantilla)
    pac = _preparar_pac(_leer_pac(ruta_pac))
    por_pos, por_crp, hay_pos = _mapa_crp(ruta_crp)

    detalle, n_pagos, total_p31 = _extraer(plantilla, por_pos, por_crp, hay_pos)
    if detalle.empty:
        raise ValueError(f"No se encontró ninguna fila P40 con CRP e importe en la hoja '{hoja}'.")

    sin_rubro = detalle[detalle['rubro'].isna() | detalle['fondos'].isna()]
    con_rubro = detalle.drop(sin_rubro.index)
    if con_rubro.empty:
        raise ValueError("Ningún CRP+posición de la plantilla se encontró en el Reporte CRP adjunto. "
                         "Verifica que sea el reporte actualizado del periodo.")

    consolidado = (con_rubro.groupby(['rubro', 'fondos'], as_index=False)['importe'].sum()
                   .rename(columns={'rubro': 'Rubro', 'fondos': 'Fondos', 'importe': 'Importe_Total'}))
    periodos = list(pac['Descripción período presupuesto'].dropna().unique())
    resultado = _cobertura(pac, consolidado, mes_pac, periodos)

    mapa = {(r['Rubro'], r['Fondos']): r['Cobertura'] for _, r in resultado.iterrows()}
    detalle['Cobertura rubro'] = [mapa.get((r, f), 'SIN RUBRO EN CRP')
                                  for r, f in zip(detalle['rubro'], detalle['fondos'])]

    # filas de total para que el Excel se lea igual que el cuadro manual
    det_xls = detalle.copy()
    for c in ('RP Doc Presupuestal', 'Posc Doc Pres', 'crp'):     # enteros, como el cuadro manual
        det_xls[c] = pd.to_numeric(det_xls[c], errors='coerce').astype('Int64')
    det_xls['importe'] = det_xls['importe'].round(0).astype('int64')
    tot_det = pd.DataFrame([{'RP Doc Presupuestal': 'TOTAL', 'importe': detalle['importe'].sum()}])
    tot_res = pd.DataFrame([{'Rubro': 'TOTAL', 'Importe_Total': resultado['Importe_Total'].sum(),
                             'Disponibilidad_PAC': resultado['Disponibilidad_PAC'].sum(),
                             'Diferencia': resultado['Diferencia'].sum()}])
    with pd.ExcelWriter(ruta_salida_excel, engine='openpyxl') as w:
        pd.concat([resultado, tot_res], ignore_index=True).to_excel(w, sheet_name='Cuadre', index=False)
        pd.concat([det_xls, tot_det], ignore_index=True).to_excel(w, sheet_name='Detalle CRP', index=False)
        for hoja_out in ('Cuadre', 'Detalle CRP'):
            ws = w.sheets[hoja_out]
            for col in ws.columns:
                ancho = max(len(str(c.value or '')) for c in col) + 2
                ws.column_dimensions[col[0].column_letter].width = min(ancho, 45)
                if str(col[0].value) in ('importe', 'Importe_Total', 'Disponibilidad_PAC',
                                         'Diferencia', 'Total pago (P31)'):
                    for c in col[1:]:
                        c.number_format = '#,##0'

    sin_match = int((resultado['Cobertura'] == 'SIN MATCH EN PAC').sum())
    sin_disp = int((resultado['Cobertura'] == 'SIN DISPONIBILIDAD PAC').sum())
    no_cubre = int(resultado['Cobertura'].str.contains('NO CUBRE').sum())
    pagos_dif = sorted({p for v in detalle.loc[detalle['Validación pago'] != 'OK', 'Pago N°']
                        for p in str(v).split(' | ') if p}, key=lambda x: int(x) if x.isdigit() else 0)

    det_json = detalle.astype(object).where(detalle.notna(), None).to_dict(orient='records')
    return {
        "ok": True,
        "mes": mes_pac,
        "hoja": hoja,
        "filas_ocultas": filas_ocultas,
        "total_p31": int(round(total_p31)),
        "posicion_en_crp": hay_pos,
        "pagos": n_pagos,
        "imputaciones": int(len(detalle)),
        "total_importe": int(round(detalle['importe'].sum())),
        "combinaciones": int(len(resultado)),
        "cubre": int(len(resultado) - sin_match - sin_disp - no_cubre),
        "no_cubre": no_cubre,
        "sin_disponibilidad": sin_disp,
        "sin_match_en_pac": sin_match,
        "sin_rubro": _convertir_tipos(sin_rubro['crp'].tolist()),
        "pagos_con_diferencia": _convertir_tipos(pagos_dif),
        "tabla": _convertir_tipos(resultado.to_dict(orient='records')),
        "detalle": _convertir_tipos(det_json),
        "archivo_salida": ruta_salida_excel,
    }
