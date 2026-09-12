from flask import Blueprint, request, jsonify, Response, send_file
import sys
import json
import os
import subprocess
from datetime import datetime

vfi_bp = Blueprint('vfi', __name__, url_prefix='/api/vfi')

BASE_BACKEND = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

@vfi_bp.route('/analizar-stream', methods=['GET'])
def analizar_stream():
    """Endpoint SSE para análisis de vigilancia fiscal con progreso en tiempo real."""
    entidad = request.args.get('entidad', 'USME')

    def generar():
        try:
            proc = subprocess.Popen(
                ['python3', 'motor_cruce.py', entidad],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=BASE_BACKEND
            )
            for linea in proc.stderr:
                linea = linea.strip()
                if linea:
                    yield f'event: progress\ndata: "{linea}"\n\n'
            stdout, _ = proc.communicate()
            if stdout.strip():
                resultado = json.loads(stdout.strip())
                yield f'event: done\ndata: {json.dumps(resultado, ensure_ascii=False)}\n\n'
        except Exception as e:
            yield f'event: error\ndata: {{"error": "{str(e)}"}}\n\n'

    return Response(generar(), mimetype='text/event-stream')

@vfi_bp.route('/buscar', methods=['POST', 'GET'])
def buscar_documento():
    """Endpoint para buscar un documento en SIRI y Boletín."""
    doc = request.args.get('doc', '') or (request.json.get('doc', '') if request.is_json else '')
    try:
        proc = subprocess.Popen(
            ['python3', 'buscar_documento.py', doc],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=BASE_BACKEND
        )
        stdout, stderr = proc.communicate()
        if stdout.strip():
            return jsonify(json.loads(stdout.strip())), 200
        return jsonify({"error": "No se pudo procesar la búsqueda"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@vfi_bp.route('/informe-pdf', methods=['POST'])
def informe_pdf():
    """Genera el informe PDF a partir del resultado del análisis (enviado por POST)."""
    try:
        import generar_informe as gi
        resultado = request.get_json(force=True)
        buf = gi.generar_pdf(resultado)
        entidad = (resultado.get('entidad', 'entidad') or 'entidad').replace(' ', '_')
        fecha = datetime.now().strftime('%Y%m%d')
        return send_file(buf, mimetype='application/pdf', as_attachment=True,
                         download_name=f'VigilanciaFiscal_{entidad}_{fecha}.pdf')
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@vfi_bp.route('/informe-excel', methods=['POST'])
def informe_excel():
    """Genera el informe Excel a partir del resultado del análisis (enviado por POST)."""
    try:
        import generar_informe as gi
        resultado = request.get_json(force=True)
        buf = gi.generar_excel(resultado)
        entidad = (resultado.get('entidad', 'entidad') or 'entidad').replace(' ', '_')
        fecha = datetime.now().strftime('%Y%m%d')
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=f'VigilanciaFiscal_{entidad}_{fecha}.xlsx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500
