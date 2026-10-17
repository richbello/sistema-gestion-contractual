import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.cruce_extraccion_crp import cruzar_extraccion_con_crp
from .utils import guardar_archivo, nombre_salida

cruce_bp = Blueprint("cruce_crp", __name__)


@cruce_bp.route("/procesar", methods=["POST"])
def procesar():
    """
    Cruce de Extracción contra Reporte CRP.

    Recibe:
      - extraccion: Excel de Extracción_Entrega
      - reporte_crp: Excel de Reporte_CRP

    Retorna:
      - Extracción con N° Interno CRP asignado
      - Columna nueva: N° Interno CRP
      - Filas divididas cuando el saldo es insuficiente
      - Hoja "Cruce_CRP" con el detalle de asignaciones
    """
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "extraccion" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo de Extracción (Excel)."}), 400
        if "reporte_crp" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo de Reporte CRP (Excel)."}), 400

        ruta_extraccion = guardar_archivo(request.files["extraccion"], upload_dir)
        ruta_crp = guardar_archivo(request.files["reporte_crp"], upload_dir)
        ruta_salida = nombre_salida(output_dir, "Extraccion_Entrega_con_CRP")

        resumen = cruzar_extraccion_con_crp(ruta_extraccion, ruta_crp, ruta_salida)

        resultado = {
            "ok": True,
            "mensaje": "Cruce completado exitosamente.",
            "filas_entrada": resumen["filas_entrada"],
            "filas_salida": resumen["filas_salida"],
            "filas_divididas": resumen["filas_divididas"],
            "sin_cruce": resumen["sin_cruce"],
            "saldo_insuficiente": resumen["saldo_insuficiente"],
            "descarga": f"/api/cruce-crp/descargar/{os.path.basename(ruta_salida)}"
        }

        if "alertas" in resumen and resumen["alertas"]:
            resultado["alertas"] = resumen["alertas"]

        return jsonify(resultado)

    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error en el cruce CRP: {e}"}), 500


@cruce_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    """Descarga el Excel con el cruce realizado."""
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
