#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Motor de cruce SECOP II x SIRI x Boletin - Vigilancia Fiscal IA"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")
import csv, re, json
from datetime import datetime, timedelta
from pathlib import Path
import urllib.request
import urllib.parse

BASE = Path(__file__).resolve().parent
SECOP_BASE = "https://www.datos.gov.co/resource/jbjy-vk9h.json"
SIRI_CSV = BASE / "datos_entrada" / "siri_antecedentes.csv"
BOLETIN_CSV = BASE / "datos_entrada" / "boletin_responsables_fiscales.csv"

def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", file=sys.stderr, flush=True)

def norm(v):
    if v is None: return ""
    s = str(v)
    if "-" in s: s = s.split("-")[0]
    return re.sub(r"[^0-9]", "", s)

def parse_fecha(fecha_str):
    if not fecha_str: return None
    s = str(fecha_str).strip().split("T")[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try: return datetime.strptime(s, fmt)
        except ValueError: continue
    return None

def fin_inhabilidad(fi, anos, meses, dias):
    if not fi: return None
    a = int(anos) if str(anos).strip().isdigit() else 0
    m = int(meses) if str(meses).strip().isdigit() else 0
    d = int(dias) if str(dias).strip().isdigit() else 0
    if a==0 and m==0 and d==0: return None
    return fi + timedelta(days=a*365 + m*30 + d)

def cargar_siri():
    siri = {}
    if not SIRI_CSV.exists():
        log(f"ADVERTENCIA: no existe {SIRI_CSV}")
        return siri
    with open(SIRI_CSV, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            doc = norm(row.get("numero_identificacion", ""))
            if doc: siri.setdefault(doc, []).append(row)
    log(f"Procuraduria (SIRI): SIRI indexado: {len(siri)} documentos unicos")
    return siri

def cargar_boletin():
    bol = {}
    if not BOLETIN_CSV.exists():
        log(f"ADVERTENCIA: no existe {BOLETIN_CSV}")
        return bol
    with open(BOLETIN_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            doc = norm(row.get("documento", ""))
            if doc: bol.setdefault(doc, []).append(row)
    log(f"Contraloria (Boletin): Boletin indexado: {len(bol)} documentos unicos")
    return bol

def descargar_secop(entidad, max_reg):
    log(f"SECOP II: Descargando SECOP para entidad que contenga: '{entidad}'")
    contratos = []
    page = 1000
    offset = 0
    where = f"upper(nombre_entidad) like upper('%{entidad}%')"
    while offset < max_reg:
        params = urllib.parse.urlencode({
            "$where": where, "$limit": page, "$offset": offset, "$order": "id_contrato"
        })
        url = f"{SECOP_BASE}?{params}"
        lote = None
        for intento in range(1, 4):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "VFI/1.0"})
                with urllib.request.urlopen(req, timeout=90) as r:
                    lote = json.loads(r.read().decode("utf-8"))
                break
            except Exception as e:
                log(f"  reintento {intento}/3 en offset {offset}...")
        if not lote: break
        contratos.extend(lote)
        log(f"SECOP II: {len(contratos)} contratos...")
        offset += page
        if len(lote) < page: break
    log(f"SECOP II: Total contratos SECOP: {len(contratos)}")
    return contratos

def procesar(entidad, max_reg):
    siri = cargar_siri()
    boletin = cargar_boletin()
    contratos = descargar_secop(entidad, max_reg)
    alertas = []
    hoy = datetime.now()

    for c in contratos:
        doc = norm(c.get("documento_proveedor", ""))
        if not doc: continue
        fecha_firma = parse_fecha(c.get("fecha_de_firma", ""))

        if doc in siri:
            for s in siri[doc]:
                f_ini = parse_fecha(s.get("fecha_efectos_juridicos", ""))
                f_fin = fin_inhabilidad(f_ini, s.get("duracion_anos",""), s.get("duracion_mes",""), s.get("duracion_dias",""))
                tiene_inhab = "INHABILIDAD" in (s.get("sanciones","") or "").upper() or f_fin is not None
                inhab_vigente = bool(f_fin and f_fin > hoy)
                if tiene_inhab and inhab_vigente:
                    if fecha_firma and f_ini and fecha_firma >= f_ini:
                        nivel, motivo = "CRITICA", "Contrato firmado DESPUES de la inhabilidad"
                    else:
                        nivel, motivo = "ALTA", "Inhabilidad vigente (contrato anterior; vigilar renovacion)"
                elif tiene_inhab:
                    nivel, motivo = "MEDIA", "Antecedente con inhabilidad vencida"
                else:
                    nivel, motivo = "INFO", "Antecedente sin inhabilidad (multa/amonestacion)"
                nombre = " ".join(filter(None, [s.get("primer_nombre",""), s.get("segundo_nombre",""), s.get("primer_apellido",""), s.get("segundo_apellido","")])).strip()
                alertas.append({
                    "fuente": "SIRI (Procuraduria)", "nivel": nivel, "motivo": motivo,
                    "documento": doc, "proveedor_secop": c.get("proveedor_adjudicado",""),
                    "nombre_fuente": nombre, "tipo": s.get("tipo_inhabilidad",""),
                    "sancion": s.get("sanciones",""), "fecha_sancion": s.get("fecha_efectos_juridicos",""),
                    "fin_inhabilidad": f_fin.strftime("%d/%m/%Y") if f_fin else "",
                    "vigente": "SI" if inhab_vigente else "NO", "entidad": c.get("nombre_entidad",""),
                    "valor": c.get("valor_del_contrato",""), "estado": c.get("estado_contrato",""),
                    "fecha_firma": c.get("fecha_de_firma","").split("T")[0] if c.get("fecha_de_firma") else "",
                    "objeto": (c.get("objeto_del_contrato","") or "")[:200], "autoridad": s.get("autoridad",""),
                })

        if doc in boletin:
            for b in boletin[doc]:
                alertas.append({
                    "fuente": "Boletin (Contraloria)", "nivel": "ALTA",
                    "motivo": "Responsable fiscal con contrato publico", "documento": doc,
                    "proveedor_secop": c.get("proveedor_adjudicado",""), "nombre_fuente": b.get("nombre",""),
                    "tipo": "FISCAL", "sancion": "RESPONSABILIDAD FISCAL", "fecha_sancion": "",
                    "fin_inhabilidad": "", "vigente": "SI", "entidad": c.get("nombre_entidad",""),
                    "valor": c.get("valor_del_contrato",""), "estado": c.get("estado_contrato",""),
                    "fecha_firma": c.get("fecha_de_firma","").split("T")[0] if c.get("fecha_de_firma") else "",
                    "objeto": (c.get("objeto_del_contrato","") or "")[:200], "autoridad": b.get("Ente que Reporta",""),
                })

    orden = {"CRITICA":0, "ALTA":1, "MEDIA":2, "INFO":3}
    alertas.sort(key=lambda a: orden.get(a["nivel"], 9))
    return {
        "entidad": entidad, "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_contratos": len(contratos), "registros_siri": len(siri), "registros_boletin": len(boletin),
        "total_alertas": len(alertas),
        "criticas": len([a for a in alertas if a["nivel"]=="CRITICA"]),
        "altas": len([a for a in alertas if a["nivel"]=="ALTA"]),
        "medias": len([a for a in alertas if a["nivel"]=="MEDIA"]),
        "info": len([a for a in alertas if a["nivel"]=="INFO"]),
        "alertas": alertas,
    }

if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    entidad = args[0] if len(args) > 0 else "USME"
    maxr = int(args[1]) if len(args) > 1 else 8000
    print(json.dumps(procesar(entidad, maxr), ensure_ascii=False))
