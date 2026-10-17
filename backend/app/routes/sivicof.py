"""
Módulo 10 — Reportes SIVICOF
  POST /api/sivicof/calcular  -> vista previa del CB-0104 con controles
  POST /api/sivicof/generar   -> ZIP "SIVICOF <MES>.zip" listo para enviar

Todo se procesa en memoria: no se guardan archivos en el servidor.
"""
import io
import traceback

from flask import Blueprint, current_app, jsonify, request, send_file
from werkzeug.utils import secure_filename

from ..services.sivicof_service import (
    ErrorSivicof,
    armar_paquete,
    calcular,
    correo_sugerido,
    fecha_corte,
    nombre_carpeta,
    parse_fecha,
    texto_1093,
    texto_fecha_firma,
)

sivicof_bp = Blueprint("sivicof", __name__)


def _periodo():
    try:
        anio = int(request.form.get("anio", ""))
        mes = int(request.form.get("mes", ""))
        if not (2000 <= anio <= 2099 and 1 <= mes <= 12):
            raise ValueError
        return anio, mes
    except ValueError:
        raise ErrorSivicof("Seleccione el mes a reportar.")


def _archivos():
    ejec = request.files.get("ejecucion")
    anul = request.files.get("anulaciones")
    if not ejec or not anul:
        raise ErrorSivicof("Adjunte el reporte de Ejecución de Gastos y el reporte de Anulaciones CRP.")
    return io.BytesIO(ejec.read()), io.BytesIO(anul.read())


def _modificacion():
    hubo = request.form.get("hubo_modificacion", "").lower() in ("1", "true", "si", "sí", "on")
    numero = request.form.get("numero_decreto", "").strip()
    fdec = parse_fecha(request.form.get("fecha_decreto", ""))
    return hubo, numero, fdec


@sivicof_bp.route("/calcular", methods=["POST"])
def calcular_endpoint():
    try:
        anio, mes = _periodo()
        ejec, anul = _archivos()
        res = calcular(ejec, anul, anio, mes)
        corte = fecha_corte(anio, mes)

        firma = parse_fecha(request.form.get("fecha_firma", ""))
        hubo, numero, fdec = _modificacion()
        res["texto_fecha_firma"] = texto_fecha_firma(firma) if firma else ""
        res["texto_1093"] = texto_1093(hubo, corte, numero, fdec) if (not hubo or numero) else ""
        res["carpeta"] = nombre_carpeta(corte)
        res["correo"] = correo_sugerido(corte)
        return jsonify({"ok": True, **res})
    except ErrorSivicof as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 422
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando los reportes: {e}"}), 500


@sivicof_bp.route("/generar", methods=["POST"])
def generar_endpoint():
    try:
        anio, mes = _periodo()
        ejec, anul = _archivos()
        firma = parse_fecha(request.form.get("fecha_firma", ""))
        if not firma:
            raise ErrorSivicof("Indique la fecha de firma de los certificados.")
        hubo, numero, fdec = _modificacion()

        decreto = None
        archivo_dec = request.files.get("decreto")
        if hubo:
            if not archivo_dec:
                raise ErrorSivicof("Hubo modificación presupuestal: adjunte el PDF del decreto.")
            nombre = archivo_dec.filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1].strip() \
                or secure_filename(archivo_dec.filename) or "Decreto.pdf"
            decreto = (nombre, archivo_dec.read())

        zip_bytes, nombre_zip, _ = armar_paquete(
            ejec, anul, anio, mes, firma,
            hubo_modificacion=hubo, numero_decreto=numero,
            fecha_decreto=fdec, decreto=decreto,
        )
        return send_file(io.BytesIO(zip_bytes), mimetype="application/zip",
                         as_attachment=True, download_name=nombre_zip)
    except ErrorSivicof as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 422
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error generando el paquete: {e}"}), 500
