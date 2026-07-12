# -*- coding: utf-8 -*-
"""Validaciones 2 (BP creado) y 3 (recursos CRP) contra bp_crp.json."""
import os, re, json
from collections import defaultdict

_JSON_PATH = os.path.abspath(os.environ.get(
    "BP_CRP_JSON_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "bp_crp.json")
))
_CACHE = None


def _digitos(v):
    return re.sub(r"\D", "", str(v or ""))


def _norm(v):
    return re.sub(r"\s+", " ", str(v or "").strip().upper())


def _cargar():
    global _CACHE
    if _CACHE is None:
        with open(_JSON_PATH, "r", encoding="utf-8") as f:
            _CACHE = json.load(f)
    return _CACHE


def _pesos(n):
    try:
        return f"${n:,.0f}"
    except (ValueError, TypeError):
        return str(n)


def validar_bp_crp(pagos):
    data = _cargar()
    docs_bp = data.get("documentos_bp", {})
    lineas = data.get("lineas", [])
    por_crp = defaultdict(float)          # crp -> saldo total disponible
    info_crp = {}                          # crp -> (contrato, doc, nombre)
    for ln in lineas:
        por_crp[ln["crp"]] += ln["saldo"]
        if ln["crp"] not in info_crp:
            info_crp[ln["crp"]] = (ln["contrato"], ln["doc"], ln["nombre"])

    res_bp = []
    res_crp = []

    for p in pagos:
        contrato = _norm(p.get("contrato"))
        doc = _digitos(p.get("cedula"))
        importe = float(p.get("importe") or 0)
        nombre = p.get("contratista") or docs_bp.get(doc, {}).get("nombre", "")

        tiene_bp = doc in docs_bp
        res_bp.append({
            "contrato": contrato, "doc": doc, "contratista": nombre,
            "estado": "OK" if tiene_bp else "SIN_BP",
            "mensaje": "OK" if tiene_bp
                       else f"Contrato {contrato} con documento {doc} no tiene BP creado",
        })

        # --- V3: recursos contra el CRP del P40 (Opcion A: sin reparto) ---
        crp = _digitos(p.get("crp"))
        if not crp or crp not in por_crp:
            res_crp.append({
                "contrato": contrato, "doc": doc, "contratista": nombre, "crp": crp,
                "importe": importe, "estado": "SIN_CRP",
                "mensaje": f"CRP {crp or '(vacio)'} del contrato {contrato} no esta en la base",
            })
            continue

        saldo = por_crp[crp]
        if importe <= saldo:
            res_crp.append({
                "contrato": contrato, "doc": doc, "contratista": nombre, "crp": crp,
                "importe": importe, "estado": "OK",
                "mensaje": f"OK - importe {_pesos(importe)} cubierto por CRP {crp} (saldo {_pesos(saldo)})",
            })
        else:
            faltante = importe - saldo
            res_crp.append({
                "contrato": contrato, "doc": doc, "contratista": nombre, "crp": crp,
                "importe": importe, "estado": "INSUFICIENTE",
                "mensaje": (f"Fondos insuficientes para el CRP {crp}: importe {_pesos(importe)}, "
                            f"saldo {_pesos(saldo)}, falta {_pesos(faltante)}"),
            })

    return {"bp": res_bp, "crp": res_crp}