#!/usr/bin/env python3
import sys, csv, json
from datetime import datetime
from collections import defaultdict

SIRI_CSV = "datos_entrada/siri_antecedentes.csv"
BOLETIN_CSV = "datos_entrada/boletin_responsables_fiscales.csv"

def cargar_siri():
    siri = defaultdict(list)
    with open(SIRI_CSV, 'r', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            doc = str(row.get('numero_identificacion', '')).strip()
            if doc:
                siri[doc].append({
                    'nombre': (row.get('primer_nombre', '') + ' ' + row.get('primer_apellido', '')).strip(),
                    'tipo_inhabilidad': row.get('tipo_inhabilidad', ''),
                    'fecha_efectos': row.get('fecha_efectos_juridicos', ''),
                    'duracion': row.get('duracion_anos', ''),
                    'sancion': row.get('sanciones', ''),
                    'fuente': 'SIRI'
                })
    return siri

def cargar_boletin():
    boletin = defaultdict(list)
    with open(BOLETIN_CSV, 'r', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            doc = str(row.get('documento', '')).strip()
            if doc:
                boletin[doc].append({
                    'nombre': row.get('nombre', ''),
                    'entidad_afectada': row.get('Entidad Afectada', ''),
                    'fuente': 'BOLETIN'
                })
    return boletin

def main():
    entidad = sys.argv[1] if len(sys.argv) > 1 else 'USME'
    max_alertas = int(sys.argv[2]) if len(sys.argv) > 2 else 500
    
    siri = cargar_siri()
    boletin = cargar_boletin()
    
    alertas = []
    
    # Cruzar todos los documentos SIRI
    for doc, antecedentes_siri in siri.items():
        for ant in antecedentes_siri:
            alertas.append({
                'nivel': 'ALTA',
                'proveedor_secop': ant['nombre'],
                'documento': doc,
                'sancion': ant['tipo_inhabilidad'],
                'tipo': ant['sancion'],
                'fuente': 'SIRI',
                'valor': 0,
                'entidad': entidad,
                'objeto': 'Antecedente disciplinario',
                'contrato_id': ''
            })
        
        # Buscar en Boletin
        if doc in boletin:
            for ant_b in boletin[doc]:
                alertas.append({
                    'nivel': 'ALTA',
                    'proveedor_secop': ant_b['nombre'],
                    'documento': doc,
                    'sancion': ant_b['entidad_afectada'],
                    'tipo': 'Responsable Fiscal',
                    'fuente': 'BOLETIN',
                    'valor': 0,
                    'entidad': entidad,
                    'objeto': 'Responsabilidad fiscal',
                    'contrato_id': ''
                })
    
    alertas = alertas[:max_alertas]
    
    stats = {
        'CRITICA': len([a for a in alertas if a['nivel'] == 'CRITICA']),
        'ALTA': len([a for a in alertas if a['nivel'] == 'ALTA']),
    }
    
    print(json.dumps({
        "entidad": entidad,
        "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_contratos": len(alertas),
        "total_alertas": len(alertas),
        "criticas": stats['CRITICA'],
        "altas": stats['ALTA'],
        "medias": 0,
        "info": 0,
        "alertas": alertas
    }, ensure_ascii=False))

if __name__ == "__main__":
    main()
