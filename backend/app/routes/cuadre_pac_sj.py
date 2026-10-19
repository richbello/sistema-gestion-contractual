"""
Cuadre PAC · Servicios Públicos y Personas Jurídicas.
Recibe en cada solicitud: plantilla, reporte PAC y Reporte CRP actualizado
(el CRP cambia con frecuencia, por eso no se usa una base guardada).
"""
import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.cuadre_pac_sj_service import cuadrar_pac_sj
from .utils import guardar_archivo, nombre_salida

cuadre_sj_bp = Blueprint("cuadre_pac_sj", __name__)


@cuadre_sj_bp.route("/procesar", methods=["POST"])
def procesar():
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]

        if "plantilla" not in request.files or "pac" not in request.files:
            return jsonify({"ok": False, "mensaje": "Se requieren los archivos 'plantilla' y 'pac'."}), 400

        mes_pac = request.form.get("mes", "").strip()
        if not mes_pac:
            return jsonify({"ok": False, "mensaje": "Debe indicar el mes a evaluar."}), 400

        if "crp" not in request.files or not request.files["crp"].filename:
            return jsonify({"ok": False, "mensaje":
                "Adjunta el Reporte CRP actualizado (cambia con cada periodo)."}), 400
        ruta_crp = guardar_archivo(request.files["crp"], upload_dir)

        ruta_plantilla = guardar_archivo(request.files["plantilla"], upload_dir)
        ruta_pac = guardar_archivo(request.files["pac"], upload_dir)
        ruta_salida = nombre_salida(output_dir, "Cuadre_PAC_SJ")

        resultado = cuadrar_pac_sj(ruta_plantilla, ruta_pac, mes_pac, ruta_salida, ruta_crp)
        resultado["descarga"] = f"/api/cuadre-pac-sj/descargar/{os.path.basename(ruta_salida)}"
        resultado.pop("archivo_salida", None)
        return jsonify(resultado)

    except (KeyError, ValueError) as e:
        return jsonify({"ok": False, "mensaje": str(e).strip('"')}), 422
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando la solicitud: {e}"}), 500


@cuadre_sj_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    nombre = os.path.basename(nombre_archivo)
    ruta = os.path.join(current_app.config["OUTPUT_FOLDER"], nombre)
    if not nombre.startswith("Cuadre_PAC_SJ_") or not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True, download_name="Cuadre_PAC_Servicios_Juridicos.xlsx")
