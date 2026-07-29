import os
import json
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file
from ..services.extraccion_service import ejecutar_extraccion, extraer_pdfs, consolidar
from .utils import guardar_archivo, nombre_salida

extraccion_bp = Blueprint("extraccion", __name__)


@extraccion_bp.route("/procesar", methods=["POST"])
def procesar():
    """Endpoint clásico: procesa todos los PDFs en una sola petición.
    Se conserva para lotes pequeños y compatibilidad."""
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


@extraccion_bp.route("/extraer-lote", methods=["POST"])
def extraer_lote():
    """Etapa 1 (troceable): recibe un lote pequeño de PDFs y devuelve
    los registros crudos en JSON. NO cruza con BASEGEN ni genera Excel.
    El frontend llama este endpoint varias veces y acumula los registros."""
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        pdfs = request.files.getlist("pdfs")
        if not pdfs:
            return jsonify({"ok": False, "mensaje": "No se recibieron PDFs en el lote."}), 400
        rutas_pdfs = [guardar_archivo(f, upload_dir) for f in pdfs if f.filename.lower().endswith(".pdf")]
        if not rutas_pdfs:
            return jsonify({"ok": False, "mensaje": "Ninguno de los archivos del lote es un PDF válido."}), 400

        registros, errores = extraer_pdfs(rutas_pdfs)

        # Limpieza inmediata de los PDFs del lote para no acumular en disco/memoria.
        for r in rutas_pdfs:
            try:
                os.remove(r)
            except OSError:
                pass

        return jsonify({
            "ok": True,
            "registros": registros,
            "procesados": len(registros),
            "errores": errores,
        })
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando el lote: {e}"}), 500


@extraccion_bp.route("/consolidar", methods=["POST"])
def consolidar_endpoint():
    """Etapa 2 (una sola vez): recibe BASEGEN + todos los registros acumulados
    (como JSON en el campo 'registros') y genera el Excel final."""
    try:
        upload_dir = current_app.config["UPLOAD_FOLDER"]
        output_dir = current_app.config["OUTPUT_FOLDER"]
        if "basegen" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el archivo BASEGEN (Excel)."}), 400

        registros_raw = request.form.get("registros")
        if not registros_raw:
            return jsonify({"ok": False, "mensaje": "No se recibieron registros para consolidar."}), 400
        try:
            registros = json.loads(registros_raw)
        except json.JSONDecodeError:
            return jsonify({"ok": False, "mensaje": "El campo 'registros' no es JSON válido."}), 400
        if not isinstance(registros, list) or not registros:
            return jsonify({"ok": False, "mensaje": "La lista de registros está vacía."}), 400

        total_pdfs = request.form.get("total_pdfs", type=int) or len(registros)
        errores_raw = request.form.get("errores")
        errores = json.loads(errores_raw) if errores_raw else []

        ruta_base = guardar_archivo(request.files["basegen"], upload_dir)
        ruta_salida = nombre_salida(output_dir, "Extraccion_Entrega")

        resultado = consolidar(registros, ruta_base, ruta_salida,
                               total_pdfs=total_pdfs, errores=errores)
        if not resultado.get("ok"):
            return jsonify(resultado), 422
        resultado["descarga"] = f"/api/extraccion/descargar/{os.path.basename(ruta_salida)}"
        return jsonify(resultado)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error consolidando: {e}"}), 500


@extraccion_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    output_dir = current_app.config["OUTPUT_FOLDER"]
    ruta = os.path.join(output_dir, nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)
