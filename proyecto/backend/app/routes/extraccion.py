import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.extraccion_service import ejecutar_extraccion
from .utils import guardar_archivo, nombre_salida

extraccion_bp = Blueprint("extraccion", __name__)


@extraccion_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "basegen" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo BASEGEN (Excel)."}), 400

        pdfs = request.files.getlist("pdfs")
        if not pdfs:
            return jsonify({"ok": False, "mensaje": "No se recibieron PDFs."}), 400

        ruta_base = guardar_archivo(request.files["basegen"], upload_dir)
        rutas_pdfs = [guardar_archivo(f, upload_dir) for f in pdfs if f.filename.lower().endswith(".pdf")]

        if not rutas_pdfs:
            return jsonify({"ok": False, "mensaje": "Ninguno de los archivos enviados es un PDF válido."}), 400

        ruta_salida = nombre_salida(output_dir, "Extraccion_Entrega")

        resultado = ejecutar_extraccion(ruta_base, rutas_pdfs, ruta_salida)

        if not resultado.get("ok"):
            return jsonify(resultado), 422

        resultado["descarga"] = f"/api/extraccion/descargar/{os.path.basename(ruta_salida)}"
        return jsonify(resultado)

    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando la solicitud: {e}"}), 500


@extraccion_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
