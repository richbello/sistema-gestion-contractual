#!/usr/bin/env python3
import sys, csv, json
from datetime import datetime
from collections import defaultdict

SIRI_CSV = "datos_entrada/siri_antecedentes.csv"
BOLETIN_CSV = "datos_entrada/boletin_responsables_fiscales.csv"

def cargar_siri():
    siri = defaultdict(list)
    try:
        with open(SIRI_CSV, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                doc = str(row.get('numero_identificacion', '')).strip()
                if doc:
                    siri[doc].append({
                        'nombre': (row.get('primer_nombre', '') + ' ' + row.get('primer_apellido', '')).strip(),
                        'tipo_inhabilidad': row.get('tipo_inhabilidad', ''),
                        'sancion': row.get('sanciones', ''),
                        'fuente': 'SIRI'
                    })
    except Exception as e:
        pass
    return siri

def cargar_boletin():
    boletin = defaultdict(list)
    try:
        with open(BOLETIN_CSV, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                doc = str(row.get('documento', '')).strip()
                if doc:
                    boletin[doc].append({
                        'nombre': row.get('nombre', ''),
                        'entidad_afectada': row.get('Entidad Afectada', ''),
                        'fuente': 'BOLETIN'
                    })
    except Exception as e:
        pass
    return boletin

def main():
    entidad = sys.argv[1] if len(sys.argv) > 1 else 'USME'
    max_alertas = int(sys.argv[2]) if len(sys.argv) > 2 else 500
    
    siri = cargar_siri()
    boletin = cargar_boletin()
    alertas = []
    
    # Cruzar SIRI
    for doc, antecedentes in list(siri.items())[:max_alertas]:
        for ant in antecedentes:
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
            
            # Buscar en Boletín
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
    
    resultado = {
        "entidad": entidad,
        "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_contratos": len(alertas),
        "total_alertas": len(alertas),
        "criticas": 0,
        "altas": len(alertas),
        "medias": 0,
        "info": 0,
        "alertas": alertas[:max_alertas]
    }
    
    print(json.dumps(resultado, ensure_ascii=False))

if __name__ == "__main__":
    main()
