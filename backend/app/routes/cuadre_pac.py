import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.cuadre_pac_service import cuadrar_pac
from .utils import guardar_archivo, nombre_salida

cuadre_bp = Blueprint("cuadre_pac", __name__)

ARCHIVO_BASE_CRP = "cuadre_base_crp.xlsx"


def _ruta_base_crp():
    return os.path.join(current_app.config["OUTPUT_FOLDER"], ARCHIVO_BASE_CRP)


@cuadre_bp.route("/estado-base-crp", methods=["GET"])
def estado_base_crp():
    ruta = _ruta_base_crp()
    if os.path.isfile(ruta):
        import datetime
        ts = datetime.datetime.fromtimestamp(os.path.getmtime(ruta))
        return jsonify({"ok": True, "existe": True,
                        "actualizado": ts.strftime("%Y-%m-%d %H:%M")})
    return jsonify({"ok": True, "existe": False})


@cuadre_bp.route("/cargar-base-crp", methods=["POST"])
def cargar_base_crp():
    try:
        if "crp" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo 'crp'."}), 400
        ruta_tmp = guardar_archivo(request.files["crp"], current_app.config["UPLOAD_FOLDER"])

        from ..services.enriquecer_rubros import construir_mapa_rubros
        d_rubro, _ = construir_mapa_rubros(ruta_tmp)

        import shutil
        shutil.copyfile(ruta_tmp, _ruta_base_crp())
        return jsonify({"ok": True,
                        "mensaje": f"Base CRP actualizada. {len(d_rubro)} compromisos disponibles."})
    except KeyError as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 422
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error cargando la base CRP: {e}"}), 500


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

        ruta_crp = None
        if "crp" in request.files and request.files["crp"].filename:
            ruta_crp = guardar_archivo(request.files["crp"], upload_dir)
        elif os.path.isfile(_ruta_base_crp()):
            ruta_crp = _ruta_base_crp()
        else:
            return jsonify({"ok": False, "mensaje":
                "No hay Reporte CRP. Cargue la base CRP primero (una vez por periodo) "
                "o adjunte el CRP en esta solicitud."}), 400

        resultado = cuadrar_pac(ruta_plantilla, ruta_pac, mes_pac, ruta_salida, ruta_crp=ruta_crp)
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
