import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.cuadre_pac_service import cuadrar_pac
from .utils import guardar_archivo, nombre_salida

cuadre_bp = Blueprint("cuadre_pac", __name__)


@cuadre_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "plantilla" not in request.files or "pac" not in request.files:
            return jsonify({"ok": False, "mensaje": "Se requieren los archivos 'plantilla' y 'pac'."}), 400

        mes_pac = request.form.get("mes", "").strip()
        if not mes_pac:
            return jsonify({"ok": False, "mensaje": "Debe indicar el mes a evaluar."}), 400

        ruta_plantilla = guardar_archivo(request.files["plantilla"], upload_dir)
        ruta_pac = guardar_archivo(request.files["pac"], upload_dir)
        ruta_salida = nombre_salida(output_dir, "Cuadre_PAC")

        resultado = cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida)
        resultado["descarga"] = f"/api/cuadre-pac/descargar/{os.path.basename(ruta_salida)}"
        return jsonify(resultado)

    except (KeyError, ValueError) as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 422
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando la solicitud: {e}"}), 500


@cuadre_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
