# -*- coding: utf-8 -*-
"""
Controlador (Blueprint): Validacion de plantilla de pagos.
Se monta bajo /api/validacion-plantilla
"""
import os
import traceback
from flask import Blueprint, request, jsonify, current_app
from ..services.validacion_plantilla_service import validar_plantilla
from .utils import guardar_archivo

validacion_bp = Blueprint("validacion_plantilla", __name__)


@validacion_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        if "plantilla" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo de plantilla."}), 400
        ruta = guardar_archivo(request.files["plantilla"], upload_dir)
        resultado = validar_plantilla(ruta)
        # limpiar el archivo temporal subido
        try:
            os.remove(ruta)
        except OSError:
            pass
        return jsonify(resultado)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error validando: {e}"}), 500
