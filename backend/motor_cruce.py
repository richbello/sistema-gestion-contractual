#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
MOTOR DE CRUCE CON LÓGICA TEMPORAL — Vigilancia Fiscal IA
============================================================================
Cruza SECOP II × SIRI (Procuraduría) × Boletín (Contraloría) aplicando
la REGLA DE ORO: una alerta es CRÍTICA solo cuando el contrato se firmó
DESPUÉS de que la persona quedó inhabilitada.

Niveles de alerta:
  - CRÍTICA:  contrato firmado DESPUÉS de la inhabilidad vigente
  - ALTA:     inhabilidad vigente, contrato anterior (vigilar renovaciones)
  - MEDIA:    antecedente sin inhabilidad vigente (multa, sanción vencida)
  - INFO:     coincidencia histórica sin vigencia

Uso:
    python3 motor_cruce.py --entidad USME
    python3 motor_cruce.py --entidad IDU --max 3000
    python3 motor_cruce.py --entidad "SECRETARIA DE EDUCACION" --json salida.json
============================================================================
"""

import sys
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
import requests
import csv
import re
import json
import argparse
from datetime import datetime
from pathlib import Path

SECOP_BASE = "https://www.datos.gov.co/resource/jbjy-vk9h.json"
SIRI_CSV = "../datos_entrada/siri_antecedentes.csv"
BOLETIN_CSV = "../datos_entrada/boletin_responsables_fiscales.csv"
OUT_DIR = Path("resultados_cruce")


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", file=sys.stderr)


def norm(v):
    """Normaliza documento: solo dígitos, sin DV ni padding."""
    if v is None:
        return ""
    s = str(v)
    if "-" in s:
        s = s.split("-")[0]
    return re.sub(r"[^0-9]", "", s)


def parse_fecha(fecha_str):
    """Intenta varios formatos de fecha comunes en las fuentes."""
    if not fecha_str:
        return None
    s = str(fecha_str).strip().split("T")[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def fin_inhabilidad(fecha_inicio, anos, meses, dias):
    """Calcula fecha fin de la inhabilidad."""
    if not fecha_inicio:
        return None
    a = int(anos) if str(anos).strip().isdigit() else 0
    m = int(meses) if str(meses).strip().isdigit() else 0
    d = int(dias) if str(dias).strip().isdigit() else 0
    if a == 0 and m == 0 and d == 0:
        return None
    total_dias = a * 365 + m * 30 + d
    from datetime import timedelta
    return fecha_inicio + timedelta(days=total_dias)


def cargar_siri():


def cargar_boletin():



def descargar_secop(entidad, max_reg):
    """Descarga proveedores de una entidad con reintentos."""
    log(f"Descargando SECOP para entidad que contenga: '{entidad}'")
    contratos = []
    page = 500
    offset = 0
    where = f"upper(nombre_entidad) like upper('%{entidad}%')"
    while offset < max_reg:
        lote = None
        for intento in range(1, 4):
            try:
                r = requests.get(SECOP_BASE, params={
                    "$where": where, "$limit": page,
                    "$offset": offset, "$order": ":id"
                }, timeout=120)
                if r.status_code == 200:
                    lote = r.json()
                    break
            except Exception:
                log(f"  reintento {intento}/3 en offset {offset}...")
        if not lote:
            break
        contratos.extend(lote)
        log(f"  {len(contratos)} contratos...")
        offset += page
        if len(lote) < page:
            break
    log(f"Total contratos SECOP: {len(contratos)}")
    return contratos


def procesar(entidad, max_reg):
    siri = cargar_siri()
    boletin = cargar_boletin()

    alertas = []
<<<<<<< HEAD
    hoy = datetime.now()

    for c in contratos:
        doc = norm(c.get("documento_proveedor", ""))
        if not doc:
            continue

        fecha_firma = parse_fecha(c.get("fecha_de_firma", ""))

        # --- Cruce con SIRI (disciplinario/penal) ---
        if doc in siri:
            for s in siri[doc]:
                f_ini = parse_fecha(s.get("fecha_efectos_juridicos", ""))
                f_fin = fin_inhabilidad(
                    f_ini, s.get("duracion_anos", ""),
                    s.get("duracion_mes", ""), s.get("duracion_dias", "")
                )
                tiene_inhab = "INHABILIDAD" in (s.get("sanciones", "") or "").upper() or f_fin is not None
                inhab_vigente = bool(f_fin and f_fin > hoy)

                # REGLA DE ORO
                nivel = "INFO"
                motivo = ""
                if tiene_inhab and inhab_vigente:
                    if fecha_firma and f_ini and fecha_firma >= f_ini:
                        nivel = "CRITICA"
                        motivo = "Contrato firmado DESPUÉS de la inhabilidad"
                    else:
                        nivel = "ALTA"
                        motivo = "Inhabilidad vigente (contrato anterior; vigilar renovación)"
                elif tiene_inhab:
                    nivel = "MEDIA"
                    motivo = "Antecedente con inhabilidad vencida"
                else:
                    nivel = "INFO"
                    motivo = "Antecedente sin inhabilidad (multa/amonestación)"

                nombre_siri = " ".join(filter(None, [
                    s.get("primer_nombre", ""), s.get("segundo_nombre", ""),
                    s.get("primer_apellido", ""), s.get("segundo_apellido", "")
                ])).strip()

                alertas.append({
                    "fuente": "SIRI (Procuraduría)",
                    "nivel": nivel,
                    "motivo": motivo,
                    "documento": doc,
                    "proveedor_secop": c.get("proveedor_adjudicado", ""),
                    "nombre_fuente": nombre_siri,
                    "tipo": s.get("tipo_inhabilidad", ""),
                    "sancion": s.get("sanciones", ""),
                    "fecha_sancion": s.get("fecha_efectos_juridicos", ""),
                    "fin_inhabilidad": f_fin.strftime("%d/%m/%Y") if f_fin else "",
                    "vigente": "SÍ" if inhab_vigente else "NO",
                    "entidad": c.get("nombre_entidad", ""),
                    "valor": c.get("valor_del_contrato", ""),
                    "estado": c.get("estado_contrato", ""),
                    "fecha_firma": c.get("fecha_de_firma", "").split("T")[0] if c.get("fecha_de_firma") else "",
                    "objeto": (c.get("objeto_del_contrato", "") or "")[:200],
                    "autoridad": s.get("autoridad", ""),
                })

        # --- Cruce con Boletín (fiscal) ---
        if doc in boletin:
            for b in boletin[doc]:
                alertas.append({
                    "fuente": "Boletín (Contraloría)",
                    "nivel": "ALTA",
                    "motivo": "Responsable fiscal con contrato público",
                    "documento": doc,
                    "proveedor_secop": c.get("proveedor_adjudicado", ""),
                    "nombre_fuente": b.get("nombre", ""),
                    "tipo": "FISCAL",
                    "sancion": "RESPONSABILIDAD FISCAL",
                    "fecha_sancion": "",
                    "fin_inhabilidad": "",
                    "vigente": "SÍ",
                    "entidad": c.get("nombre_entidad", ""),
                    "valor": c.get("valor_del_contrato", ""),
                    "estado": c.get("estado_contrato", ""),
                    "fecha_firma": c.get("fecha_de_firma", "").split("T")[0] if c.get("fecha_de_firma") else "",
                    "objeto": (c.get("objeto_del_contrato", "") or "")[:200],
                    "autoridad": b.get("Ente que Reporta", ""),
                })

    # Ordenar por gravedad
    orden = {"CRITICA": 0, "ALTA": 1, "MEDIA": 2, "INFO": 3}
    alertas.sort(key=lambda a: orden.get(a["nivel"], 9))

    # Guardar CSV
    csv_path = OUT_DIR / f"alertas_{entidad.replace(' ', '_')}.csv"
    if alertas:
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=alertas[0].keys())
            w.writeheader()
            w.writerows(alertas)

    # Guardar JSON (para el dashboard)
    json_path = OUT_DIR / f"alertas_{entidad.replace(' ', '_')}.json"
    resumen = {
=======
    
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
>>>>>>> 86fc35667d69bce6caa7adbb37bb34c4e5b84556
        "entidad": entidad,
        "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_contratos": len(contratos),
        "registros_siri": len(siri),
        "registros_boletin": len(boletin),
        "total_alertas": len(alertas),
<<<<<<< HEAD
        "criticas": len([a for a in alertas if a["nivel"] == "CRITICA"]),
        "altas": len([a for a in alertas if a["nivel"] == "ALTA"]),
        "medias": len([a for a in alertas if a["nivel"] == "MEDIA"]),
        "info": len([a for a in alertas if a["nivel"] == "INFO"]),
        "alertas": alertas,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(resumen, f, ensure_ascii=False, indent=2)
=======
        "criticas": 0,
        "altas": len(alertas),
        "medias": 0,
        "info": 0,
        "alertas": alertas[:max_alertas]
    }
    
    print(json.dumps(resultado, ensure_ascii=False))
>>>>>>> 86fc35667d69bce6caa7adbb37bb34c4e5b84556

    # Reporte en consola
    print("\n" + "=" * 70, file=sys.stderr)
    print(f"RESULTADO — {entidad}", file=sys.stderr)
    print("=" * 70, file=sys.stderr)
    print(f"  Contratos analizados: {len(contratos)}", file=sys.stderr)
    print(f"  Total alertas:        {len(alertas)}", file=sys.stderr)
    print(f"  [CRIT] CRÍTICAS:          {resumen['criticas']}", file=sys.stderr)
    print(f"  [ALTA] ALTAS:             {resumen['altas']}", file=sys.stderr)
    print(f"  [MEDIA] MEDIAS:            {resumen['medias']}", file=sys.stderr)
    print(f"  [INFO] INFO:              {resumen['info']}", file=sys.stderr)
    print("=" * 70, file=sys.stderr)

    if resumen["criticas"] > 0:
        print("\n[CRIT] ALERTAS CRÍTICAS (contrato posterior a inhabilidad):", file=sys.stderr)
        for a in alertas:
            if a["nivel"] == "CRITICA":
                print(f"  • {a['proveedor_secop']} | {a['sancion']} | firma: {a['fecha_firma']} | sanción: {a['fecha_sancion']}", file=sys.stderr)

    print(f"\n[OK] CSV:  {csv_path}", file=sys.stderr)
    print(f"[OK] JSON: {json_path}", file=sys.stderr)
    return resumen


if __name__ == "__main__":
    import sys as _sys
    args = [a for a in _sys.argv[1:] if not a.startswith("-")]
    entidad = args[0] if len(args) > 0 else "USME"
    maxr = int(args[1]) if len(args) > 1 else 5000
    resumen = procesar(entidad, maxr)
    print(json.dumps(resumen, ensure_ascii=False))
