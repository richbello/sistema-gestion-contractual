import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.plantilla_pagos_service import generar_plantilla_pagos
from .utils import guardar_archivo, nombre_salida

plantilla_bp = Blueprint("plantilla_pagos", __name__)


@plantilla_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "extraccion" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo de extracción (Excel)."}), 400

        ruta_entrada = guardar_archivo(request.files["extraccion"], upload_dir)
        ruta_salida = nombre_salida(output_dir, "Plantilla_Pagos")

        resultado = generar_plantilla_pagos(ruta_entrada, ruta_salida)
        resultado["descarga"] = f"/api/plantilla-pagos/descargar/{os.path.basename(ruta_salida)}"
        return jsonify(resultado)

    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando la solicitud: {e}"}), 500


@plantilla_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
