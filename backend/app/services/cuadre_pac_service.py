"""
Módulo 3 — Cuadre PAC Acumulado.

Versión mejorada que agrupa los pagos por Rubro+Fondo (unificados),
y cruce contra PAC de múltiples meses para indicar en qué mes se cubre
completamente si falta en septiembre.
"""
import pandas as pd

COLS_PLANTILLA_REQUERIDAS = {'Tipo Registro P', 'Clave Contab.', 'importe', 'Rubro', 'Fondos'}
COLS_PAC_REQUERIDAS = {'Descripción período presupuesto', 'Progr.financiación', 'Fondos', 'Disponibilidad PAC'}


def _importe_por_bloque(plantilla):
    """Extrae Rubro/Fondos/importe por bloque (cabecera C -> P31)."""
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
    """Normaliza nombre del período a mayúsculas sin acentos."""
    import unicodedata
    if periodo is None:
        return ''
    s = str(periodo).strip().upper()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def _parse_mes_numero(periodo_str):
    """Convierte 'ENERO' -> 1, 'FEBRERO' -> 2, etc."""
    meses = {
        'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
        'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
    }
    norm = _normalizar_periodo(periodo_str)
    return meses.get(norm, None)


def cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida_excel):
    """
    Cuadre PAC acumulado: agrupa por Rubro+Fondo, cruza contra PAC de múltiples
    meses, e indica en qué mes(es) se cubre si falta en el mes actual.
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

    # --- Importe unificado por Rubro+Fondos ---
    bloques = _importe_por_bloque(plantilla)
    if bloques.empty:
        raise ValueError("No se pudo determinar el importe de ningún contrato.")

    consolidado = (
        bloques.dropna(subset=['Rubro', 'Fondos'])
        .groupby(['Rubro', 'Fondos'], as_index=False)['importe']
        .sum()
        .rename(columns={'importe': 'Importe_Total'})
    )

    # --- Obtener lista de períodos disponibles ---
    periodos = list(pac['Descripción período presupuesto'].dropna().unique())
    mes_num_actual = _parse_mes_numero(mes_pac)
    if mes_num_actual is None:
        raise ValueError(f"Mes '{mes_pac}' no reconocido.")

    # --- PAC por período: agrupar ---
    pac_norm = pac.copy()
    pac_norm['_periodo_norm'] = pac_norm['Descripción período presupuesto'].map(_normalizar_periodo)
    pac_norm['_mes_num'] = pac_norm['_periodo_norm'].map(_parse_mes_numero)

    # --- Para cada Rubro+Fondo en consolidado, obtener PAC de meses desde mes_pac en adelante ---
    resultado_filas = []
    for _, row in consolidado.iterrows():
        rub = row['Rubro']
        fon = row['Fondos']
        imp = row['Importe_Total']

        # Filtrar PAC para este Rubro+Fondo en todos los meses
        pac_rub_fon = pac_norm[
            (pac_norm['Progr.financiación'].astype(str).str.strip() == str(rub).strip()) &
            (pac_norm['Fondos'].astype(str).str.strip() == str(fon).strip())
        ]

        # PAC del mes actual
        pac_mes_actual = pac_rub_fon[pac_rub_fon['_periodo_norm'] == _normalizar_periodo(mes_pac)][
            'Disponibilidad PAC'
        ].sum()

        # PAC acumulado desde mes actual hacia adelante
        pac_acumulado = {}
        for periodo in periodos:
            mes_n = _parse_mes_numero(periodo)
            if mes_n is not None and mes_n >= mes_num_actual:
                pac_acumulado[periodo] = pac_rub_fon[
                    pac_rub_fon['_periodo_norm'] == _normalizar_periodo(periodo)
                ]['Disponibilidad PAC'].sum()

        # Calcular cobertura
        cobertura_texto = "SIN DISPONIBILIDAD PAC"
        if pac_mes_actual > 0:
            if pac_mes_actual >= imp:
                cobertura_texto = f"OK - CUBRE EN {mes_pac.upper()}"
            else:
                # Buscar en qué mes(es) se cubre acumulando
                acum = 0
                meses_cubrimiento = []
                meses_ord = sorted(pac_acumulado.items(), key=lambda x: _parse_mes_numero(x[0]) or 999)
                for periodo, monto in meses_ord:
                    acum += monto
                    meses_cubrimiento.append(periodo.upper())
                    if acum >= imp:
                        break

                if acum >= imp:
                    meses_str = " + ".join(meses_cubrimiento)
                    falta = imp - pac_mes_actual
                    cobertura_texto = f"FALTA ${falta:,.0f} EN {mes_pac.upper()}, CUBRE acum. {meses_str}"
                else:
                    falta = imp - pac_mes_actual
                    cobertura_texto = f"FALTA ${falta:,.0f}, NO CUBRE NI EN {meses_cubrimiento[-1] if meses_cubrimiento else 'AÑO'}"

        resultado_filas.append({
            'Rubro': rub,
            'Fondos': fon,
            'Importe_Total': imp,
            'PAC_Septiembre': pac_mes_actual,
            'Diferencia': pac_mes_actual - imp,
            'Cobertura': cobertura_texto
        })

    resultado = pd.DataFrame(resultado_filas)

    # Guardar
    resultado.to_excel(ruta_salida_excel, index=False)

    # Métricas
    cubre = (resultado['Diferencia'] >= 0).sum()
    no_cubre = (resultado['Diferencia'] < 0).sum()
    sin_disp = (resultado['PAC_Septiembre'] == 0).sum()

    return {
        "ok": True,
        "mes": mes_pac,
        "periodos_disponibles": periodos,
        "combinaciones": len(resultado),
        "cubre_septiembre": cubre,
        "no_cubre_septiembre": no_cubre,
        "sin_disponibilidad": sin_disp,
        "tabla": resultado.to_dict(orient="records"),
        "archivo_salida": ruta_salida_excel,
    }
