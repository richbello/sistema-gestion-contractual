from flask import Blueprint, request, jsonify, send_file
from flask_cors import cross_origin
import pandas as pd
from openpyxl import load_workbook
import os

bp = Blueprint('estado_cuenta', __name__)

@bp.route('/procesar', methods=['POST'])
@cross_origin()
def procesar():
    try:
        plantilla_file = request.files.get('plantilla')
        historico_file = request.files.get('historico')
        contrato = request.form.get('contrato', '').strip()

        if not all([plantilla_file, historico_file, contrato]):
            return jsonify({"ok": False, "mensaje": "Faltan archivos o contrato"}), 400

        # Leer histórico
        df_historico = pd.read_excel(historico_file, sheet_name=0)
        pagos = df_historico[df_historico['Referencia'].astype(str).str.contains(contrato, na=False)]
        
        if pagos.empty:
            return jsonify({"ok": False, "mensaje": f"No encontrado: {contrato}"}), 400

        primer = pagos.iloc[0]

        # Cargar plantilla con openpyxl (preserva TODO)
        plantilla_path = '/tmp/plantilla_temp.xlsx'
        plantilla_file.save(plantilla_path)
        wb = load_workbook(plantilla_path)
        ws = wb.active

        # Escribir datos SIN BORRAR FORMATOS
        ws['D5'].value = contrato
        ws['D6'].value = str(primer.get('Nombre', ''))
        ws['D7'].value = float(primer.get('VALOR FINAL DEL CONTRATO', 0)) or 0
        ws['D8'].value = str(primer.get('FECHA INICIAL DE CONTRATO', ''))
        ws['H5'].value = str(primer.get('Proveedor', ''))
        ws['H6'].value = str(primer.get('Nº identificación', ''))
        ws['H7'].value = str(primer.get('Numero RP', ''))
        ws['H8'].value = str(primer.get('FECHA DE TERMINACION FINAL', ''))

        fila = 17
        saldo = float(primer.get('VALOR FINAL DEL CONTRATO', 0)) or 0

        for i, (_, pago) in enumerate(pagos.iterrows(), start=1):
            monto = float(pago.get('Valor Bruto', 0)) or 0
            saldo -= monto

            ws[f'B{fila}'].value = i
            ws[f'C{fila}'].value = str(pago.get('Texto cabecera documento', ''))
            ws[f'D{fila}'].value = monto
            ws[f'E{fila}'].value = saldo
            ws[f'F{fila}'].value = str(pago.get('Doc.compensación', ''))
            ws[f'G{fila}'].value = str(pago.get('Fecha de pago', ''))
            ws[f'H{fila}'].value = str(pago.get('Numero RP', ''))
            ws[f'I{fila}'].value = str(pago.get('CDP Externo', ''))
            ws[f'J{fila}'].value = str(pago.get('CRP Externo', ''))
            fila += 1

        salida_path = f'/tmp/Estado_de_Cuenta_{contrato}.xlsx'
        wb.save(salida_path)

        return send_file(salida_path, as_attachment=True, download_name=f'Estado_de_Cuenta_{contrato}.xlsx')

    except Exception as e:
        return jsonify({"ok": False, "mensaje": str(e)}), 500

