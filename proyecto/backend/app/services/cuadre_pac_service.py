"""
Módulo 3 — Cuadre PAC.

Adaptado del script original de Colab: cruza la plantilla de pagos
(Módulo 2) contra el reporte PAC del mes, agrupando por Rubro + Fondos,
para verificar si el PAC disponible cubre el valor a pagar.
"""
import pandas as pd

COLS_PLANTILLA_REQUERIDAS = {'Clave Contab.', 'importe', 'Rubro', 'Fondos'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


def cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida_excel):
    """
    Realiza el cuadre PAC y guarda el resultado en ruta_salida_excel.
    Retorna un resumen con métricas para mostrar en la interfaz.
    """
    plantilla = pd.read_excel(ruta_plantilla, sheet_name='Hoja1')
    pac = pd.read_excel(ruta_pac, sheet_name='Data')

    faltantes_p = COLS_PLANTILLA_REQUERIDAS - set(plantilla.columns)
    faltantes_pac = COLS_PAC_REQUERIDAS - set(pac.columns)
    if faltantes_p:
        raise KeyError(f"Faltan columnas en la plantilla: {faltantes_p}")
    if faltantes_pac:
        raise KeyError(f"Faltan columnas en el PAC: {faltantes_pac}")

    plantilla['importe'] = pd.to_numeric(plantilla['importe'], errors='coerce').fillna(0)
    pac['Disponibilidad PAC'] = pd.to_numeric(pac['Disponibilidad PAC'], errors='coerce').fillna(0)

    p40 = plantilla[plantilla['Clave Contab.'] == 40].copy()
    if p40.empty:
        raise ValueError("No se encontraron filas con Clave Contab. = 40 en la plantilla.")

    consolidado = (
        p40.dropna(subset=['Rubro', 'Fondos'])
        .groupby(['Rubro', 'Fondos'], as_index=False)['importe']
        .sum()
        .rename(columns={'importe': 'Importe_Total'})
    )

    periodos = list(pac['Descripción período presupuesto'].dropna().unique())
    pac_mes = pac[pac['Descripción período presupuesto'].str.upper().str.strip() == mes_pac.upper()]

    if pac_mes.empty:
        raise ValueError(
            f'No se encontraron filas para el mes "{mes_pac}". Períodos disponibles: {periodos}'
        )

    pac_disp = (
        pac_mes
        .groupby(['Progr.financiación', 'Fondos'], as_index=False)['Disponibilidad PAC']
        .sum()
        .rename(columns={'Progr.financiación': 'Rubro', 'Disponibilidad PAC': 'Disponibilidad_PAC'})
    )

    resultado = consolidado.merge(pac_disp, on=['Rubro', 'Fondos'], how='left')
    sin_match = int(resultado['Disponibilidad_PAC'].isna().sum())
    resultado['Disponibilidad_PAC'] = resultado['Disponibilidad_PAC'].fillna(0)

    resultado['Diferencia'] = resultado['Disponibilidad_PAC'] - resultado['Importe_Total']

    def evaluar_cobertura(row):
        if row['Disponibilidad_PAC'] == 0:
            return 'SIN DISPONIBILIDAD PAC'
        elif row['Diferencia'] >= 0:
            return f'CUBRE EN {mes_pac}'
        else:
            return f'NO CUBRE EN {mes_pac}'

    resultado['Cobertura'] = resultado.apply(evaluar_cobertura, axis=1)

    for col in ['Importe_Total', 'Disponibilidad_PAC', 'Diferencia']:
        resultado[col] = resultado[col].astype(int)

    resultado.to_excel(ruta_salida_excel, index=False)

    cubre = int((resultado['Cobertura'] == f'CUBRE EN {mes_pac}').sum())
    no_cubre = int((resultado['Cobertura'] == f'NO CUBRE EN {mes_pac}').sum())
    sin_disp = int((resultado['Cobertura'] == 'SIN DISPONIBILIDAD PAC').sum())

    return {
        "ok": True,
        "mes": mes_pac,
        "periodos_disponibles": periodos,
        "combinaciones": len(resultado),
        "cubre": cubre,
        "no_cubre": no_cubre,
        "sin_disponibilidad": sin_disp,
        "sin_match_en_pac": sin_match,
        "tabla": resultado.to_dict(orient="records"),
        "archivo_salida": ruta_salida_excel,
    }
