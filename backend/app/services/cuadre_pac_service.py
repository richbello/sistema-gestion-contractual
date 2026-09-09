"""
Módulo 3 — Cuadre PAC Acumulado.
Agrupa pagos por Rubro+Fondo, cruza contra PAC de varios meses e indica
en qué mes se cubre si falta en el mes evaluado.
"""
import pandas as pd
import numpy as np
import unicodedata

COLS_PLANTILLA_REQUERIDAS = {'Tipo Registro P', 'Clave Contab.', 'importe', 'Rubro', 'Fondos'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


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

    pac['_periodo_norm'] = pac['Descripción período presupuesto'].apply(_normalizar_periodo)
    pac['Progr.financiación'] = pac['Progr.financiación'].astype(str).str.strip()
    pac['Fondos'] = pac['Fondos'].astype(str).str.strip()

    resultado_filas = []
    for _, row in consolidado.iterrows():
        rub = str(row['Rubro']).strip()
        fon = str(row['Fondos']).strip()
        imp = float(row['Importe_Total'])

        pac_rf = pac[(pac['Progr.financiación'] == rub) & (pac['Fondos'] == fon)]
        pac_mes = float(pac_rf[pac_rf['_periodo_norm'] == _normalizar_periodo(mes_pac)]['Disponibilidad PAC'].sum())

        pac_acumulado = {}
        for periodo in periodos:
            mes_n = _parse_mes_numero(periodo)
            if mes_n is not None and mes_n >= mes_num_actual:
                pac_acumulado[periodo] = float(pac_rf[pac_rf['_periodo_norm'] == _normalizar_periodo(periodo)]['Disponibilidad PAC'].sum())

        cobertura = "SIN DISPONIBILIDAD PAC"
        if pac_mes > 0:
            if pac_mes >= imp:
                cobertura = f"OK - CUBRE EN {mes_pac.upper()}"
            else:
                acum = 0; meses_cub = []
                for periodo, monto in sorted(pac_acumulado.items(), key=lambda x: _parse_mes_numero(x[0]) or 999):
                    acum += monto; meses_cub.append(periodo.upper())
                    if acum >= imp:
                        break
                falta = imp - pac_mes
                if acum >= imp:
                    cobertura = f"FALTA ${falta:,.0f} EN {mes_pac.upper()}, CUBRE acum. {' + '.join(meses_cub)}"
                else:
                    cobertura = f"FALTA ${falta:,.0f}, NO CUBRE"

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
    cubre = int(resultado['Cobertura'].str.startswith('OK').sum())
    no_cubre = int(resultado['Cobertura'].str.startswith('FALTA').sum())
    sin_disp = int((resultado['Cobertura'] == 'SIN DISPONIBILIDAD PAC').sum())

    return {
        "ok": True,
        "mes": mes_pac,
        "periodos_disponibles": periodos,
        "combinaciones": int(len(resultado)),
        "cubre": cubre,
        "no_cubre": no_cubre,
        "sin_disponibilidad": sin_disp,
        "sin_match_en_pac": sin_disp,
        "tabla": tabla_json,
        "archivo_salida": ruta_salida_excel,
    }
