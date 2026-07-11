# -*- coding: utf-8 -*-
"""
Controlador (Blueprint): Registro de cuentas bancarias.
Se monta bajo /api/cuentas-bancarias
"""
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.cuentas_bancarias_service import (
    buscar_cuenta,
    agregar_cuenta,
    editar_cuenta,
    eliminar_cuenta,
    listar,
    total,
    ruta_json,
    recargar,
)

cuentas_bp = Blueprint("cuentas_bancarias", __name__)


@cuentas_bp.route("/buscar/<cedula>", methods=["GET"])
def buscar(cedula):
    try:
        reg = buscar_cuenta(cedula)
        if reg is None:
            return jsonify({"ok": False, "encontrado": False,
                            "mensaje": f"No hay cuenta registrada para {cedula}."}), 404
        return jsonify({"ok": True, "encontrado": True, "cuenta": reg})
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error consultando: {e}"}), 500


@cuentas_bp.route("/agregar", methods=["POST"])
def agregar():
    try:
        data = request.get_json(silent=True) or request.form
        if not data:
            return jsonify({"ok": False, "mensaje": "Faltan datos (JSON)."}), 400
        resultado = agregar_cuenta(
            cedula=data.get("cedula", ""),
            nombre=data.get("nombre", ""),
            tipo_cuenta=data.get("tipo_cuenta", ""),
            codigo_banco=data.get("codigo_banco", ""),
            numero_cuenta=data.get("numero_cuenta", ""),
            sobreescribir=bool(data.get("sobreescribir", True)),
        )
        return jsonify(resultado), (200 if resultado.get("ok") else 409)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error agregando: {e}"}), 500


@cuentas_bp.route("/editar", methods=["POST", "PUT"])
def editar():
    try:
        data = request.get_json(silent=True) or request.form
        if not data or not data.get("cedula"):
            return jsonify({"ok": False, "mensaje": "Falta la cédula a modificar."}), 400
        resultado = editar_cuenta(
            cedula=data.get("cedula", ""),
            nombre=data.get("nombre"),
            tipo_cuenta=data.get("tipo_cuenta"),
            codigo_banco=data.get("codigo_banco"),
            numero_cuenta=data.get("numero_cuenta"),
            nueva_cedula=data.get("nueva_cedula"),
        )
        return jsonify(resultado), (200 if resultado.get("ok") else 409)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error modificando: {e}"}), 500


@cuentas_bp.route("/eliminar/<cedula>", methods=["DELETE", "POST"])
def eliminar(cedula):
    try:
        resultado = eliminar_cuenta(cedula)
        return jsonify(resultado), (200 if resultado.get("ok") else 404)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error eliminando: {e}"}), 500


@cuentas_bp.route("/listar", methods=["GET"])
def listar_rt():
    try:
        q = request.args.get("q", "").strip() or None
        datos = listar(filtro=q)
        return jsonify({"ok": True, "total": len(datos), "cuentas": datos})
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error listando: {e}"}), 500


@cuentas_bp.route("/total", methods=["GET"])
def total_rt():
    return jsonify({"ok": True, "total": total()})


@cuentas_bp.route("/exportar", methods=["GET"])
def exportar():
    try:
        return send_file(ruta_json(), as_attachment=True,
                         download_name="cuentas_bancarias.json")
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error exportando: {e}"}), 500


@cuentas_bp.route("/recargar", methods=["POST"])
def recargar_rt():
    try:
        n = recargar()
        return jsonify({"ok": True, "mensaje": "Registro recargado.", "total": n})
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error recargando: {e}"}), 500
