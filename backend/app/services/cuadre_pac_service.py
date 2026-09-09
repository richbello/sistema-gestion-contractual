"""
Módulo 3 — Cuadre PAC.

Adaptado del script original de Colab: cruza la plantilla de pagos
(Módulo 2) contra el reporte PAC del mes, agrupando por Rubro + Fondos,
para verificar si el PAC disponible cubre el valor a pagar.

CAMBIO vs. versión anterior:
  El importe a pagar YA NO se suma de la fila P40. BogData a veces exporta
  la P40 con el número de documento (RP Doc Presupuestal) en la columna
  'importe' (valores ~5.000.000.000), lo que inflaba el cuadre. El valor
  real de pago está en la fila P31 de cada bloque; de ahí se toma ahora.
  El Rubro y los Fondos se siguen leyendo de la P40.
"""
import pandas as pd

COLS_PLANTILLA_REQUERIDAS = {'Tipo Registro P', 'Clave Contab.', 'importe', 'Rubro', 'Fondos'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


def _importe_por_bloque(plantilla):
    """
    Recorre la plantilla por bloques (cabecera 'C' -> P40 -> P31) y devuelve
    un DataFrame Rubro/Fondos/importe con el importe REAL de cada contrato:
      - Rubro y Fondos: de la fila P40 del bloque.
      - importe: de la fila P31 (manda porque siempre viene limpio). Si no
        hay P31, cae a la P40 solo si su importe no es el # de documento.
    """
    tipo  = plantilla['Tipo Registro P'].astype(str).str.upper().str.strip()
    clave = pd.to_numeric(plantilla['Clave Contab.'], errors='coerce')
    imp   = pd.to_numeric(plantilla['importe'], errors='coerce')
    if 'RP Doc Presupuestal' in plantilla.columns:
        doc = pd.to_numeric(plantilla['RP Doc Presupuestal'], errors='coerce')
    else:
        doc = pd.Series([pd.NA] * len(plantilla))

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
        elif clave.iat[i] == 40:
            rubro = str(plantilla['Rubro'].iat[i]).strip()
            fondo = str(plantilla['Fondos'].iat[i]).strip()
            # fallback: usar P40 solo si su importe NO es el # de documento
            if valor is None and pd.notna(imp.iat[i]) and imp.iat[i] != doc.iat[i]:
                valor = imp.iat[i]
        elif clave.iat[i] == 31 and pd.notna(imp.iat[i]):
            valor = imp.iat[i]          # la P31 manda: importe real de pago
    cerrar()

    return pd.DataFrame(filas, columns=['Rubro', 'Fondos', 'importe'])


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

    pac['Disponibilidad PAC'] = pd.to_numeric(pac['Disponibilidad PAC'], errors='coerce').fillna(0)

    # --- Importe real por Rubro + Fondos (tomado de la P31, no de la P40) ---
    bloques = _importe_por_bloque(plantilla)
    if bloques.empty:
        raise ValueError("No se pudo determinar el importe de ningún contrato "
                         "(¿faltan filas P31 o P40 con Rubro?).")

    consolidado = (
        bloques.dropna(subset=['Rubro', 'Fondos'])
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
        resultado[col] = resultado[col].round(0).astype(int)

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
