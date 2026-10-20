import sys, csv, re, json
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")

# Misma ruta que motor_cruce.py (backend/datos_entrada), sin depender del cwd.
DATOS = Path(__file__).resolve().parent / "datos_entrada"
SIRI_CSV = DATOS / "siri_antecedentes.csv"
BOLETIN_CSV = DATOS / "boletin_responsables_fiscales.csv"

def norm(v):
    if v is None: return ""
    s = str(v)
    if "-" in s: s = s.split("-")[0]
    return re.sub(r"[^0-9]", "", s)

def dv_nit(base):
    """Digito de verificacion DIAN para un NIT de 9 digitos."""
    pesos = [71, 67, 59, 53, 47, 43, 41, 37, 29, 23, 19, 17, 13, 7, 3]
    s = sum(int(d) * p for d, p in zip(reversed(base), reversed(pesos)))
    r = s % 11
    return r if r < 2 else 11 - r

def claves_boletin(valor):
    """El Boletin trae 'NIT 8020062581' (con DV pegado); SECOP II trae el NIT
    sin DV. Se indexan las dos formas cuando el ultimo digito es un DV valido."""
    doc = norm(valor)
    claves = [doc] if doc else []
    if str(valor).strip().upper().startswith("NIT") and len(doc) == 10 \
            and dv_nit(doc[:9]) == int(doc[9]):
        claves.append(doc[:9])
    return claves

def buscar(doc):
    doc = norm(doc)
    hallazgos = []
    # SIRI
    try:
        with open(SIRI_CSV, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if norm(row.get("numero_identificacion","")) == doc:
                    nombre = " ".join(filter(None, [
                        row.get("primer_nombre",""), row.get("segundo_nombre",""),
                        row.get("primer_apellido",""), row.get("segundo_apellido","")
                    ])).strip()
                    hallazgos.append({
                        "fuente": "SIRI (Procuraduria)",
                        "nombre": nombre,
                        "tipo": row.get("tipo_inhabilidad",""),
                        "sancion": row.get("sanciones",""),
                        "fecha": row.get("fecha_efectos_juridicos",""),
                        "autoridad": row.get("autoridad",""),
                        "entidad": row.get("entidad_sancionado",""),
                    })
    except Exception as e:
        pass
    # Boletin
    try:
        with open(BOLETIN_CSV, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                if doc in claves_boletin(row.get("documento","")) or (len(doc) == 10 and doc[:9] in claves_boletin(row.get("documento",""))):
                    hallazgos.append({
                        "fuente": "Boletin (Contraloria)",
                        "nombre": row.get("nombre",""),
                        "tipo": "FISCAL",
                        "sancion": "RESPONSABILIDAD FISCAL",
                        "fecha": "",
                        "autoridad": row.get("Ente que Reporta",""),
                        "entidad": "",
                    })
    except Exception as e:
        pass
    return {"documento": doc, "total": len(hallazgos), "hallazgos": hallazgos}

if __name__ == "__main__":
    doc = sys.argv[1] if len(sys.argv) > 1 else ""
    print(json.dumps(buscar(doc), ensure_ascii=False))
