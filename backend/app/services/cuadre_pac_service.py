"""
Módulo 3 — Cuadre PAC Acumulado.
Agrupa pagos por Rubro+Fondo, cruza contra PAC de varios meses e indica
en qué mes se cubre si falta en el mes evaluado.

CORRECCIONES (sep 2026):
 1. La llave de cruce del rubro se normaliza (_norm_rubro): ignora ceros a la
    izquierda y lecturas tipo float ("230689.0"), para que "0230689" de la
    plantilla cruce con "230689" de BOGDATA/PAC y viceversa.
 2. Se elimina el candado `if pac_mes > 0`: ahora, si el mes evaluado está en 0
    pero hay PAC en meses siguientes, se reporta la cobertura acumulada en lugar
    de marcar "SIN DISPONIBILIDAD PAC".
 3. Se distingue "SIN MATCH EN PAC" (no hubo cruce) de "SIN DISPONIBILIDAD PAC"
    (cruzó pero el PAC programado en meses es 0), y la disponibilidad se suma
    solo de filas con período de mes válido (evita doble conteo con filas Total).
"""
import pandas as pd
import numpy as np
import unicodedata

COLS_PLANTILLA_REQUERIDAS = {'Tipo Registro P', 'Clave Contab.', 'importe', 'Rubro', 'Fondos'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


def _norm_rubro(x):
    """Clave de cruce robusta para el rubro.
    Ignora ceros a la izquierda y lecturas tipo float ('230689.0'),
    de modo que '0230689' (plantilla) cruce con '230689' (PAC/BOGDATA).
    Los códigos de 24 dígitos quedan iguales en ambos lados (pierden el mismo
    cero inicial), así que no se rompe ningún cruce que hoy funciona.
    """
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s = ''.join(ch for ch in s if ch.isdigit())
    return s.lstrip('0')


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
            rubro = str(plantilla['Rubro'].iat[i]).strip()
            fondo = str(plantilla['Fondos'].iat[i]).strip()
            valor = None
            abierto = True
        elif clave.iat[i] == 31 and pd.notna(imp.iat[i]):
            valor = imp.iat[i]
        elif clave.iat[i] == 40 and valor is None and pd.notna(imp.iat[i]):
            if 'RP Doc Presupuestal' in plantilla.columns:
                doc = pd.to_numeric(plantilla['RP Doc Presupuestal'].iat[i], errors='coerce')
                if imp.iat[i] != doc:
                    valor = imp.iat[i]
            else:
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


def cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida_excel):
    plantilla = pd.read_excel(ruta_plantilla, sheet_name='Hoja1')
    pac = pd.read_excel(ruta_pac, sheet_name='Data')

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

    # --- Llaves normalizadas para el cruce (defecto 1) ---
    pac['_periodo_norm'] = pac['Descripción período presupuesto'].apply(_normalizar_periodo)
    pac['_rubro_norm']   = pac['Progr.financiación'].apply(_norm_rubro)
    pac['_fondo_norm']   = pac['Fondos'].astype(str).str.strip().str.upper()
    # marca de fila con período de MES válido (excluye filas Total / vacías)
    pac['_es_mes'] = pac['_periodo_norm'].apply(lambda p: _parse_mes_numero(p) is not None)

    mes_norm = _normalizar_periodo(mes_pac)

    resultado_filas = []
    for _, row in consolidado.iterrows():
        rub = str(row['Rubro']).strip()
        fon = str(row['Fondos']).strip()
        imp = float(row['Importe_Total'])

        rub_norm = _norm_rubro(rub)
        fon_norm = fon.upper()
        pac_rf = pac[(pac['_rubro_norm'] == rub_norm) & (pac['_fondo_norm'] == fon_norm)]
        pac_rf_meses = pac_rf[pac_rf['_es_mes']]

        pac_mes = float(pac_rf_meses[pac_rf_meses['_periodo_norm'] == mes_norm]['Disponibilidad PAC'].sum())

        # disponibilidad por período, solo del mes actual en adelante
        pac_acumulado = {}
        for periodo in periodos:
            mes_n = _parse_mes_numero(periodo)
            if mes_n is not None and mes_n >= mes_num_actual:
                pac_acumulado[periodo] = float(
                    pac_rf_meses[pac_rf_meses['_periodo_norm'] == _normalizar_periodo(periodo)]['Disponibilidad PAC'].sum()
                )

        # disponibilidad total del rubro (solo filas de mes → sin doble conteo)
        disp_meses = float(pac_rf_meses['Disponibilidad PAC'].sum())

        # --- Lógica de cobertura (defecto 2 y 3) ---
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