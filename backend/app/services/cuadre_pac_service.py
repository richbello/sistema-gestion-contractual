"""
Módulo 3 — Cuadre PAC Acumulado.
Agrupa pagos por Rubro+Fondo, cruza contra PAC de varios meses e indica
en qué mes se cubre si falta en el mes evaluado.

CORRECCIONES (sep 2026):
 1. _importe_por_bloque: el Rubro y Fondos se leen de la fila P40 (donde vienen
    realmente), no de la fila C. Se conserva compatibilidad si algún formato los
    trae en la fila C.
 2. Cruce de rubro robusto (_norm_rubro): ignora ceros a la izquierda, la letra
    'O' inicial y lecturas float ("230689.0").
 3. El rubro se cruza contra 'Progr.financiación' y, si no hay match, contra
    'Pos.presupuestaria' (los rubros de cuentas por pagar como O230689 viven en
    Pos.presupuestaria, no en Progr.financiación).
 4. Se elimina el candado `if pac_mes > 0` y se distingue "SIN MATCH EN PAC" de
    "SIN DISPONIBILIDAD PAC". Las sumas solo usan filas con período de mes válido
    (excluye REZAGO / PAC NO EJECUTADO / Total → sin doble conteo).
"""
import pandas as pd
import numpy as np
import unicodedata

from .enriquecer_rubros import construir_mapa_rubros, enriquecer_plantilla

COLS_PLANTILLA_REQUERIDAS = {'Tipo Registro P', 'Clave Contab.', 'importe'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


def _norm_rubro(x):
    """Clave de cruce robusta: solo dígitos, sin ceros a la izquierda.
    Así 'O230689', '0230689' y '230689.0' normalizan todos a '230689'.
    """
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s = ''.join(ch for ch in s if ch.isdigit())
    return s.lstrip('0')


def _celda(plantilla, col, i):
    if col not in plantilla.columns:
        return None
    v = str(plantilla[col].iat[i]).strip()
    return None if v in ('', 'nan', 'None', 'NaT') else v


def _importe_por_bloque(plantilla):
    tipo  = plantilla['Tipo Registro P'].astype(str).str.upper().str.strip()
    clave = pd.to_numeric(plantilla['Clave Contab.'], errors='coerce')
    imp   = pd.to_numeric(plantilla['importe'], errors='coerce')
    filas = []
    rubro = fondo = valor = None
    abierto = False
    def cerrar():
        if abierto and str(rubro) not in ('None', 'nan') and valor is not None:
            filas.append((rubro, fondo, valor))
    for i in range(len(plantilla)):
        if tipo.iat[i] == 'C':
            cerrar()
            rubro = fondo = valor = None
            abierto = True
            # compatibilidad: algunos formatos traen rubro/fondos en la fila C
            if _celda(plantilla, 'Rubro', i):
                rubro = _celda(plantilla, 'Rubro', i)
            if _celda(plantilla, 'Fondos', i):
                fondo = _celda(plantilla, 'Fondos', i)
        elif clave.iat[i] == 40:
            # el Rubro y Fondos vienen normalmente en la fila P40
            if rubro is None and _celda(plantilla, 'Rubro', i):
                rubro = _celda(plantilla, 'Rubro', i)
            if fondo is None and _celda(plantilla, 'Fondos', i):
                fondo = _celda(plantilla, 'Fondos', i)
            if valor is None and pd.notna(imp.iat[i]):
                if 'RP Doc Presupuestal' in plantilla.columns:
                    doc = pd.to_numeric(plantilla['RP Doc Presupuestal'].iat[i], errors='coerce')
                    if imp.iat[i] != doc:
                        valor = imp.iat[i]
                else:
                    valor = imp.iat[i]
        elif clave.iat[i] == 31 and pd.notna(imp.iat[i]):
            valor = imp.iat[i]
    cerrar()
    return pd.DataFrame(filas, columns=['Rubro', 'Fondos', 'importe'])


def _normalizar_periodo(periodo):
    if periodo is None:
        return ''
    s = str(periodo).strip().upper()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def _parse_mes_numero(periodo_str):
    meses = {'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
             'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12}
    return meses.get(_normalizar_periodo(periodo_str), None)


def _convertir_tipos(obj):
    if isinstance(obj, dict):
        return {k: _convertir_tipos(v) for k, v in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [_convertir_tipos(x) for x in obj]
    elif isinstance(obj, (np.integer, np.floating)):
        return obj.item()
    elif isinstance(obj, pd.Series):
        return _convertir_tipos(obj.to_dict())
    else:
        return obj


def cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida_excel, ruta_crp=None):
    plantilla = pd.read_excel(ruta_plantilla, sheet_name='Hoja1')
    pac = pd.read_excel(ruta_pac, sheet_name='Data')

    # La plantilla de SAP NO trae Rubro/Fondos (BOGDATA los rechaza).
    # Se traen del Reporte CRP cruzando por No. Compromiso, en memoria, sin tocar el archivo.
    if ruta_crp:
        d_rubro, d_fondos = construir_mapa_rubros(ruta_crp)
        encontrados, sin_match = enriquecer_plantilla(plantilla, d_rubro, d_fondos)
        if encontrados == 0:
            raise ValueError(
                'No se pudo cruzar ningun compromiso de la plantilla contra el CRP. '
                'Verifica que el Reporte CRP cargado corresponde al periodo.')

    faltantes_p = COLS_PLANTILLA_REQUERIDAS - set(plantilla.columns)
    faltantes_pac = COLS_PAC_REQUERIDAS - set(pac.columns)
    if faltantes_p:
        raise KeyError(f"Faltan columnas en la plantilla: {faltantes_p}")
    if faltantes_pac:
        raise KeyError(f"Faltan columnas en el PAC: {faltantes_pac}")

    pac['Disponibilidad PAC'] = pd.to_numeric(pac['Disponibilidad PAC'], errors='coerce').fillna(0)

    bloques = _importe_por_bloque(plantilla)
    if bloques.empty:
        raise ValueError("No se pudo determinar el importe de ningún contrato.")

    consolidado = (
        bloques.dropna(subset=['Rubro', 'Fondos'])
        .groupby(['Rubro', 'Fondos'], as_index=False)['importe']
        .sum().rename(columns={'importe': 'Importe_Total'})
    )

    periodos = list(pac['Descripción período presupuesto'].dropna().unique())
    mes_num_actual = _parse_mes_numero(mes_pac)
    if mes_num_actual is None:
        raise ValueError(f"Mes '{mes_pac}' no reconocido.")

    tiene_pos = 'Pos.presupuestaria' in pac.columns
    pac['_periodo_norm'] = pac['Descripción período presupuesto'].apply(_normalizar_periodo)
    pac['_progr_norm']   = pac['Progr.financiación'].apply(_norm_rubro)
    pac['_pos_norm']     = pac['Pos.presupuestaria'].apply(_norm_rubro) if tiene_pos else ''
    pac['_fondo_norm']   = pac['Fondos'].astype(str).str.strip().str.upper()
    pac['_es_mes'] = pac['_periodo_norm'].apply(lambda p: _parse_mes_numero(p) is not None)

    mes_norm = _normalizar_periodo(mes_pac)

    resultado_filas = []
    for _, row in consolidado.iterrows():
        rub = str(row['Rubro']).strip()
        fon = str(row['Fondos']).strip()
        imp = float(row['Importe_Total'])

        rub_norm = _norm_rubro(rub)
        fon_norm = fon.upper()
        base_fondo = pac['_fondo_norm'] == fon_norm

        # 1º cruce por Progr.financiación; si no hay match, por Pos.presupuestaria
        pac_rf = pac[(pac['_progr_norm'] == rub_norm) & base_fondo]
        if pac_rf.empty and tiene_pos:
            pac_rf = pac[(pac['_pos_norm'] == rub_norm) & base_fondo]

        pac_rf_meses = pac_rf[pac_rf['_es_mes']]
        pac_mes = float(pac_rf_meses[pac_rf_meses['_periodo_norm'] == mes_norm]['Disponibilidad PAC'].sum())

        pac_acumulado = {}
        for periodo in periodos:
            mes_n = _parse_mes_numero(periodo)
            if mes_n is not None and mes_n >= mes_num_actual:
                pac_acumulado[periodo] = float(
                    pac_rf_meses[pac_rf_meses['_periodo_norm'] == _normalizar_periodo(periodo)]['Disponibilidad PAC'].sum()
                )

        disp_meses = float(pac_rf_meses['Disponibilidad PAC'].sum())

        if pac_rf.empty:
            cobertura = "SIN MATCH EN PAC"
        elif disp_meses <= 0:
            cobertura = "SIN DISPONIBILIDAD PAC"
        elif pac_mes >= imp:
            cobertura = f"OK - CUBRE EN {mes_pac.upper()}"
        else:
            acum = 0.0
            meses_cub = []
            for periodo, monto in sorted(pac_acumulado.items(),
                                         key=lambda x: _parse_mes_numero(x[0]) or 999):
                if monto <= 0:
                    continue
                acum += monto
                meses_cub.append(periodo.upper())
                if acum >= imp:
                    break
            if acum >= imp:
                if pac_mes > 0:
                    falta = imp - pac_mes
                    cobertura = (f"FALTA ${falta:,.0f} EN {mes_pac.upper()}, "
                                 f"CUBRE acum. {' + '.join(meses_cub)}")
                else:
                    cobertura = f"OK - CUBRE acum. {' + '.join(meses_cub)}"
            else:
                cobertura = f"FALTA ${imp - acum:,.0f}, NO CUBRE"

        resultado_filas.append({
            'Rubro': rub,
            'Fondos': fon,
            'Importe_Total': int(imp),
            'Disponibilidad_PAC': int(pac_mes),
            'Diferencia': int(pac_mes - imp),
            'Cobertura': cobertura
        })

    resultado = pd.DataFrame(resultado_filas)
    resultado.to_excel(ruta_salida_excel, index=False)

    tabla_json = _convertir_tipos(resultado.to_dict(orient="records"))
    sin_match = int((resultado['Cobertura'] == 'SIN MATCH EN PAC').sum())
    sin_disp  = int((resultado['Cobertura'] == 'SIN DISPONIBILIDAD PAC').sum())
    no_cubre  = int(resultado['Cobertura'].str.contains('NO CUBRE').sum())
    cubre     = int(len(resultado) - sin_match - sin_disp - no_cubre)

    return {
        "ok": True,
        "mes": mes_pac,
        "periodos_disponibles": periodos,
        "combinaciones": int(len(resultado)),
        "cubre": cubre,
        "no_cubre": no_cubre,
        "sin_disponibilidad": sin_disp,
        "sin_match_en_pac": sin_match,
        "tabla": tabla_json,
        "archivo_salida": ruta_salida_excel,
    }