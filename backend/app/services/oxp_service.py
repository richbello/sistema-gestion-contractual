# -*- coding: utf-8 -*-
"""
Módulo 12 · CDP-CRP de OXP  ─  Servicio backend (solo CRP por ahora)
====================================================================
Rellena la hoja 'CRP_vf' de la plantilla a partir de filas ya FILTRADAS
(saldo > 0) y MAPEADAS en el navegador (claves canónicas). El backend agrupa
por CRP, aplica constantes y escribe → RAM mínima (Render free tier).

AGRUPAMIENTO (según formato diligenciado):
    Un CRP puede tener varias POSICIONES. Se agrupa por 'interno_crp'
    (N° Interno CRP del reporte); cada grupo recibe UN número de CRP
    consecutivo (1..N) y sus filas conservan la 'pos_crp' (N° Posición CRP).

Contrato de datos (cada fila = dict con claves canónicas; opcionales):
    importe, objeto, tipo_compromiso, no_compromiso, modo_seleccion,
    tipo_doc_benef, id_benef, id_solicitante, id_responsable,
    interno_crp, pos_crp, num_crp
"""

import os
from io import BytesIO
from datetime import date

import openpyxl

# ----------------------------------------------------------------------------
# Constantes (revisar por alcaldía)
# ----------------------------------------------------------------------------
PLANTILLA_PATH = os.environ.get(
    "OXP_PLANTILLA_PATH",
    os.path.join(os.path.dirname(__file__), "..", "plantillas",
    "Formato_CRP_de_OXP.xlsx"),
)

SOCIEDAD_FIJA   = 1001            # Usme = 1001 (confirmado)
CLASE_DOC_CRP   = "RP"
MONEDA          = "COP"
POSICION_CDP    = 1
TIPO_DE_PAGO    = "02"
FECHA_INICIAL   = "01/01/2026"
FECHA_FINAL     = "31/12/2026"

# Num. Ext. Entidad (columna V). Modos:
#   "num_crp" -> número de CRP original del reporte (constante por grupo)  [segun imagen]
#   "grupo"   -> mismo consecutivo del CRP (1..N por grupo)
#   "fila"    -> consecutivo por fila (1..N)   [texto original del spec]
NUM_EXT_ENTIDAD_MODO = "num_crp"


def _hoy():
    return date.today().strftime("%d.%m.%Y")


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, str):
        limpio = v.replace(".", "").replace(",", "").strip()
        try:
            v = float(limpio)
        except ValueError:
            return v
    try:
        return int(round(float(v)))
    except (TypeError, ValueError):
        return v


def _g(fila, clave):
    v = fila.get(clave)
    return None if v in ("", None) else v


def _pos_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return 999999   # posiciones sin dato van al final del grupo


def _agrupar_por_crp(filas):
    """Agrupa por interno_crp preservando el orden de primera aparicion.
    Cada grupo: filas ordenadas por pos_crp. Devuelve lista de grupos."""
    grupos = {}
    orden = []
    for i, fila in enumerate(filas):
        clave = _g(fila, "interno_crp")
        if clave is None:
            clave = "__sin_interno_%d" % i   # fila suelta = su propio grupo
        if clave not in grupos:
            grupos[clave] = []
            orden.append(clave)
        grupos[clave].append(fila)
    return [sorted(grupos[k], key=lambda f: _pos_int(_g(f, "pos_crp"))) for k in orden]


# ----------------------------------------------------------------------------
# CRP OXP  -> hoja CRP_vf
# ----------------------------------------------------------------------------
def llenar_crp_vf(filas, plantilla_path=None):
    plantilla_path = plantilla_path or PLANTILLA_PATH
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb["CRP_vf"]
    hoy = _hoy()

    grupos = _agrupar_por_crp(filas)

    r = 2
    for crp_consecutivo, grupo in enumerate(grupos, start=1):
        for pos_idx, fila in enumerate(grupo, start=1):
            pos = _g(fila, "pos_crp")
            posicion = _num(pos) if pos not in (None, "") else pos_idx

            if NUM_EXT_ENTIDAD_MODO == "num_crp":
                num_ext = _g(fila, "num_crp")
            elif NUM_EXT_ENTIDAD_MODO == "grupo":
                num_ext = crp_consecutivo
            else:
                num_ext = r - 1

            ws.cell(row=r, column=1,  value=crp_consecutivo)            # A CRP (grupo)
            ws.cell(row=r, column=2,  value=posicion)                   # B Posicion
            ws.cell(row=r, column=3,  value=hoy)                        # C Fecha Documento
            ws.cell(row=r, column=4,  value=hoy)                        # D Fecha Contab.
            ws.cell(row=r, column=5,  value=SOCIEDAD_FIJA)             # E Sociedad
            ws.cell(row=r, column=6,  value=CLASE_DOC_CRP)            # F Clase Documento
            ws.cell(row=r, column=7,  value=MONEDA)                    # G Moneda
            ws.cell(row=r, column=8,  value=_num(_g(fila, "importe")))  # H importe
            # I CDP -> en blanco (por ahora)
            ws.cell(row=r, column=10, value=POSICION_CDP)             # J Posicion CDP
            ws.cell(row=r, column=11, value=_g(fila, "objeto"))         # K Objeto
            ws.cell(row=r, column=12, value=_g(fila, "tipo_compromiso"))# L Tipo compromiso
            ws.cell(row=r, column=13, value=_g(fila, "no_compromiso"))  # M No. Compromiso
            ws.cell(row=r, column=14, value=FECHA_INICIAL)            # N Fecha Inicial
            ws.cell(row=r, column=15, value=FECHA_FINAL)              # O Fecha Final
            cp = ws.cell(row=r, column=16, value=TIPO_DE_PAGO)       # P Tipo de Pago
            cp.number_format = "@"
            ws.cell(row=r, column=17, value=_g(fila, "modo_seleccion"))  # Q Modo Seleccion
            ws.cell(row=r, column=18, value=_g(fila, "tipo_doc_benef"))  # R Tipo Doc Benef
            ws.cell(row=r, column=19, value=_g(fila, "id_benef"))        # S Identif. Benef
            ws.cell(row=r, column=20, value=_g(fila, "id_solicitante"))  # T ID Solicitante
            ws.cell(row=r, column=21, value=_g(fila, "id_responsable"))  # U ID Responsable
            ws.cell(row=r, column=22, value=num_ext)                    # V Num. Ext. Entidad
            r += 1

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, (r - 2)
