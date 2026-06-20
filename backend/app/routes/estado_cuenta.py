import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.estado_cuenta_service import generar_estado_cuenta
from .utils import guardar_archivo, nombre_salida

estado_cuenta_bp = Blueprint("estado_cuenta", __name__)


@estado_cuenta_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "plantilla" not in request.files or "historico" not in request.files:
            return jsonify({"ok": False, "mensaje": "Se requieren los archivos 'plantilla' e 'historico'."}), 400

        contrato = request.form.get("contrato", "").strip()
        if not contrato:
            return jsonify({"ok": False, "mensaje": "Debe indicar el número de contrato (ej. 713-2023)."}), 400

        ruta_plantilla = guardar_archivo(request.files["plantilla"], upload_dir)
        ruta_historico = guardar_archivo(request.files["historico"], upload_dir)
        ruta_salida = nombre_salida(output_dir, f"Estado_Cuenta_{contrato}")

        resultado = generar_estado_cuenta(ruta_plantilla, ruta_historico, contrato, ruta_salida)

        if not resultado.get("ok"):
            return jsonify(resultado), 404

        resultado["descarga"] = f"/api/estado-cuenta/descargar/{os.path.basename(ruta_salida)}"
        return jsonify(resultado)

    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando la solicitud: {e}"}), 500


@estado_cuenta_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
