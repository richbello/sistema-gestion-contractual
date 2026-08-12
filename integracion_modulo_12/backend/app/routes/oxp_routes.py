# -*- coding: utf-8 -*-
"""
Módulo 12 · CDP-CRP de OXP  ─  Rutas Flask (solo CRP por ahora)
==============================================================
Recibe filas ya filtradas/mapeadas por el navegador (JSON liviano) y devuelve
el .xlsx de la plantilla diligenciada (hoja CRP_vf).

Endpoint:
    POST /api/oxp/crp   body: {"filas": [ {canónicas...}, ... ]}  -> xlsx
"""

from datetime import date

from flask import Blueprint, request, jsonify, send_file

from ..services.oxp_service import llenar_crp_vf, llenar_cdp_oxp

# El url_prefix ("/api/oxp") se define al registrar en app/__init__.py,
# igual que los demás blueprints del proyecto.
oxp_bp = Blueprint("oxp", __name__)


@oxp_bp.route("/crp", methods=["POST"])
def generar_crp():
    data = request.get_json(silent=True) or {}
    filas = data.get("filas")
    if not isinstance(filas, list) or not filas:
        return jsonify({"error": "Se requiere 'filas' (lista no vacía)."}), 400

    buf, n = llenar_crp_vf(filas)
    nombre = "CRP_de_OXP_%s.xlsx" % date.today().strftime("%Y%m%d")
    resp = send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=nombre,
    )
    resp.headers["X-OXP-Registros"] = str(n)
    return resp


@oxp_bp.route("/cdp", methods=["POST"])
def generar_cdp():
    data = request.get_json(silent=True) or {}
    filas = data.get("filas")
    if not isinstance(filas, list) or not filas:
        return jsonify({"error": "Se requiere 'filas' (lista no vacía)."}), 400

    buf, n = llenar_cdp_oxp(filas)
    nombre = "CDP_de_OXP_%s.xlsx" % date.today().strftime("%Y%m%d")
    resp = send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=nombre,
    )
    resp.headers["X-OXP-Registros"] = str(n)
    return resp
