"""
Enriquecimiento de Rubro/Fondos para el Cuadre PAC.
Porta la macro VBA TraerRubroFondos_POR_COMPROMISO a Python.

La plantilla de SAP NO trae Rubro/Fondos (si los llevara, BOGDATA la rechaza).
Aquí se traen del Reporte CRP cruzando por No. Compromiso y se inyectan EN MEMORIA
en las columnas Rubro/Fondos de las filas cabecera 'C'. El archivo de la plantilla
nunca se modifica.
"""
import pandas as pd


def _norm_header(s):
    s = str(s).upper().strip()
    for ch in (' ', '.', 'º', '°', 'ª'):
        s = s.replace(ch, '')
    return s


def _buscar_col(cols, objetivo):
    """Puerto de BuscarCol (VBA): match exacto y luego 'contiene', sobre
    encabezados normalizados. Cruza por NOMBRE, nunca por posición."""
    obj = _norm_header(objetivo)
    norm = {c: _norm_header(c) for c in cols}
    for c, n in norm.items():
        if n == obj:
            return c
    for c, n in norm.items():
        if obj and obj in n:
            return c
    return None


def _norm_comp(v):
    """Puerto de NormComp (VBA): espacio->guion, colapsa guiones dobles.
    Reproduce exactamente el match de la macro."""
    s = str(v).strip()
    if s in ('', 'nan', 'None', 'NaT'):
        return ''
    s = s.replace('.0', '') if s.endswith('.0') else s
    s = s.replace(' ', '-')
    while '--' in s:
        s = s.replace('--', '-')
    return s


def construir_mapa_rubros(ruta_crp):
    """Lee el Reporte CRP y devuelve dos dicts {compromiso_norm: valor}
    para Rubro y Fondos. Cruza por nombre de columna."""
    df = pd.read_excel(ruta_crp)
    col_comp   = _buscar_col(df.columns, 'No Compromiso')
    col_rubro  = _buscar_col(df.columns, 'Rubro')
    col_fondos = _buscar_col(df.columns, 'Fondos')
    faltan = [n for n, c in [('No. Compromiso', col_comp),
                             ('Rubro', col_rubro), ('Fondos', col_fondos)] if c is None]
    if faltan:
        raise KeyError(f"No se encontraron estas columnas en el CRP: {faltan}")

    d_rubro, d_fondos = {}, {}
    for _, r in df.iterrows():
        k = _norm_comp(r[col_comp])
        if k and k not in d_rubro:          # primera aparición gana (consistente)
            d_rubro[k]  = r[col_rubro]
            d_fondos[k] = r[col_fondos]
    return d_rubro, d_fondos


def enriquecer_plantilla(plantilla, d_rubro, d_fondos, col_comp_idx=9):
    """Inyecta Rubro/Fondos en las filas cabecera 'C' (col J = idx 9),
    cruzando por compromiso. Devuelve (encontrados, sin_match)."""
    if 'Rubro' not in plantilla.columns:
        plantilla['Rubro'] = None
    if 'Fondos' not in plantilla.columns:
        plantilla['Fondos'] = None
    iR = plantilla.columns.get_loc('Rubro')
    iF = plantilla.columns.get_loc('Fondos')

    tipo = plantilla['Tipo Registro P'].astype(str).str.upper().str.strip()
    comp_col = plantilla.iloc[:, col_comp_idx]

    encontrados, sin_match = 0, []
    for i in range(len(plantilla)):
        if tipo.iat[i] == 'C':
            k = _norm_comp(comp_col.iat[i])
            if not k:
                continue
            if k in d_rubro:
                plantilla.iat[i, iR] = d_rubro[k]
                plantilla.iat[i, iF] = d_fondos[k]
                encontrados += 1
            else:
                sin_match.append(k)
    return encontrados, sin_match