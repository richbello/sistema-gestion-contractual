import os
import traceback
from flask import Blueprint, request, jsonify, current_app, send_file

from ..services.extraccion_sipago_service import (
    construir_base_crp,
    construir_base_bancaria,
    base_bancaria_desde_registros,
    guardar_base,
    info_base,
    ejecutar_consolidado,
)
from ..services.cruce_extraccion_crp import cruzar_extraccion_con_crp
from .utils import guardar_archivo, nombre_salida

extraccion_sipago_bp = Blueprint("extraccion_sipago", __name__)

# Las bases quedan guardadas en el servidor y solo se reemplazan cuando el
# usuario sube archivos nuevos desde la pestaña "Bases de referencia".
ARCHIVO_BASE_CRP   = "sipago_base_crp.pkl"
ARCHIVO_BASE_BANCO = "sipago_base_bancaria.pkl"
# Reporte CRP completo (todos los CRP de cada contrato) para dividir pagos
# cuando el saldo de un CRP no alcanza. La base .pkl solo guarda uno por cédula.
ARCHIVO_REPORTE_CRP = "sipago_reporte_crp.xlsx"


def _ruta_base(nombre):
    carpeta = current_app.config.get("BASES_FOLDER") or current_app.config["OUTPUT_FOLDER"]
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, nombre)


@extraccion_sipago_bp.route("/estado-bases", methods=["GET"])
def estado_bases():
    """Informa qué bases hay guardadas, cuántos registros y cuándo se actualizaron."""
    return jsonify({
        "ok": True,
        "crp":     info_base(_ruta_base(ARCHIVO_BASE_CRP)),
        "bancaria": info_base(_ruta_base(ARCHIVO_BASE_BANCO)),
    })


@extraccion_sipago_bp.route("/cargar-base-crp", methods=["POST"])
def cargar_base_crp():
    """Reemplaza la base CRP con un Reporte CRP nuevo."""
    try:
        if "reporte_crp" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el Reporte CRP."}), 400

        ruta = guardar_archivo(request.files["reporte_crp"],
                               current_app.config["UPLOAD_FOLDER"])
        base = construir_base_crp(ruta)

        if base.empty:
            return jsonify({"ok": False,
                            "mensaje": "El archivo no tiene datos en las columnas "
                                       "esperadas (AF documento, AQ saldo, AR CRP)."}), 422

        guardar_base(base, _ruta_base(ARCHIVO_BASE_CRP))
        try:
            os.replace(ruta, _ruta_base(ARCHIVO_REPORTE_CRP))
        except OSError:
            pass

        return jsonify({"ok": True, "registros": int(len(base)),
                        "mensaje": f"Base CRP actualizada con {len(base)} contratistas."})
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error cargando la base CRP: {e}"}), 500


@extraccion_sipago_bp.route("/cargar-base-bancaria", methods=["POST"])
def cargar_base_bancaria():
    """Reemplaza la base bancaria con uno o varios históricos de pago.
    Enviar los archivos de más antiguo a más reciente: la cuenta que queda
    guardada es la del archivo más reciente en el que aparece el contratista."""
    try:
        archivos = request.files.getlist("historicos")
        if not archivos:
            return jsonify({"ok": False, "mensaje": "No se recibieron históricos de pago."}), 400

        upload_dir = current_app.config["UPLOAD_FOLDER"]
        rutas = [guardar_archivo(f, upload_dir) for f in archivos
                 if f.filename.lower().endswith((".xlsx", ".xls"))]
        if not rutas:
            return jsonify({"ok": False, "mensaje": "Ninguno de los archivos es un Excel válido."}), 400

        base, omitidos = construir_base_bancaria(rutas)

        for r in rutas:
            try:
                os.remove(r)
            except OSError:
                pass

        if base.empty:
            return jsonify({"ok": False,
                            "mensaje": "No se encontraron cuentas bancarias en los archivos.",
                            "omitidos": omitidos}), 422

        guardar_base(base, _ruta_base(ARCHIVO_BASE_BANCO))
        sin_codigo = int((base["Código Bco"].astype(str).str.strip() == "").sum())

        return jsonify({
            "ok": True,
            "registros": int(len(base)),
            "sin_codigo_banco": sin_codigo,
            "omitidos": omitidos,
            "mensaje": f"Base bancaria actualizada con {len(base)} contratistas.",
        })
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error cargando la base bancaria: {e}"}), 500


@extraccion_sipago_bp.route("/cargar-base-bancaria-json", methods=["POST"])
def cargar_base_bancaria_json():
    """Reemplaza la base bancaria con los registros que extrajo el navegador.

    El navegador lee los históricos con SheetJS, se queda solo con cédula +
    cuenta + código de banco y manda eso ya deduplicado. Así el servidor nunca
    toca el Excel completo: entran unos cientos de KB de JSON en vez de 13 MB
    de Excel que al parsearse se convertían en cientos de MB y tumbaban al
    worker (SIGKILL por falta de memoria en el free tier de 512 MB).
    """
    try:
        payload = request.get_json(force=True, silent=True) or {}
        registros = payload.get("registros") or []
        if not registros:
            return jsonify({"ok": False,
                            "mensaje": "No se recibieron cuentas bancarias."}), 400

        base = base_bancaria_desde_registros(registros, payload.get("mapa_bancos"))

        if base.empty:
            return jsonify({"ok": False,
                            "mensaje": "No se encontraron cuentas bancarias válidas.",
                            "omitidos": payload.get("omitidos") or []}), 422

        guardar_base(base, _ruta_base(ARCHIVO_BASE_BANCO))
        sin_codigo = int((base["Código Bco"].astype(str).str.strip() == "").sum())

        return jsonify({
            "ok": True,
            "registros": int(len(base)),
            "sin_codigo_banco": sin_codigo,
            "omitidos": payload.get("omitidos") or [],
            "mensaje": f"Base bancaria actualizada con {len(base)} contratistas.",
        })
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error cargando la base bancaria: {e}"}), 500


@extraccion_sipago_bp.route("/procesar", methods=["POST"])
def procesar():
    """Consolida el Reporte del Periodo contra las bases guardadas."""
    try:
        if "reporte" not in request.files:
            return jsonify({"ok": False, "mensaje": "Falta el Reporte del Periodo."}), 400

        ruta_reporte = guardar_archivo(request.files["reporte"],
                                       current_app.config["UPLOAD_FOLDER"])
        ruta_salida = nombre_salida(current_app.config["OUTPUT_FOLDER"],
                                    "Extraccion_SIPAGO")

        resultado = ejecutar_consolidado(
            ruta_reporte,
            _ruta_base(ARCHIVO_BASE_CRP),
            _ruta_base(ARCHIVO_BASE_BANCO),
            ruta_salida,
        )

        try:
            os.remove(ruta_reporte)
        except OSError:
            pass

        if not resultado.get("ok"):
            return jsonify(resultado), 422

        # Cruce por Contrato con el Reporte CRP completo: asigna el CRP más
        # antiguo con saldo y divide la fila cuando el saldo no alcanza.
        ruta_crp = _ruta_base(ARCHIVO_REPORTE_CRP)
        if os.path.isfile(ruta_crp):
            try:
                resultado["cruce_crp"] = cruzar_extraccion_con_crp(
                    ruta_salida, ruta_crp, ruta_salida,
                    conservar_sin_cruce=True, actualizar_saldo=True)
            except Exception as e:
                current_app.logger.error(traceback.format_exc())
                resultado["cruce_crp"] = {"error": f"No se pudo cruzar por contrato: {e}"}
        else:
            resultado["cruce_crp"] = None

        resultado["descarga"] = (
            f"/api/extraccion-sipago/descargar/{os.path.basename(ruta_salida)}"
        )
        return jsonify(resultado)
    except Exception as e:
        current_app.logger.error(traceback.format_exc())
        return jsonify({"ok": False, "mensaje": f"Error procesando el reporte: {e}"}), 500


@extraccion_sipago_bp.route("/descargar/<nombre_archivo>", methods=["GET"])
def descargar(nombre_archivo):
    ruta = os.path.join(current_app.config["OUTPUT_FOLDER"], nombre_archivo)
    if not os.path.isfile(ruta):
        return jsonify({"ok": False, "mensaje": "Archivo no encontrado."}), 404
    return send_file(ruta, as_attachment=True)