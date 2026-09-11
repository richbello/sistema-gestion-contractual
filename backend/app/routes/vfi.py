from flask import Blueprint, request, jsonify, Response
import sys
import json
import os
import subprocess
from datetime import datetime

vfi_bp = Blueprint('vfi', __name__, url_prefix='/api/vfi')

@vfi_bp.route('/analizar-stream', methods=['GET'])
def analizar_stream():
    """Endpoint SSE para análisis de vigilancia fiscal con progreso en tiempo real."""
    entidad = request.args.get('entidad', 'USME')
    
    def generar():
        try:
            # Ejecutar motor_cruce.py como subprocess
            proc = subprocess.Popen(
                ['python3', 'motor_cruce.py', entidad],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                cwd=os.path.dirname(os.path.abspath(__file__))
            )
            
            # Leer stderr (logs de progreso)
            for linea in proc.stderr:
                linea = linea.strip()
                if linea:
                    yield f'event: progress\ndata: "{linea}"\n\n'
            
            # Esperar a que termine y leer stdout (JSON resultado)
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
    doc = request.args.get('doc', '') or request.json.get('doc', '')
    
    try:
        proc = subprocess.Popen(
            ['python3', 'buscar_documento.py', doc],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.path.dirname(os.path.abspath(__file__))
        )
        
        stdout, stderr = proc.communicate()
        if stdout.strip():
            resultado = json.loads(stdout.strip())
            return jsonify(resultado), 200
        else:
            return jsonify({"error": "No se pudo procesar la búsqueda"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500
