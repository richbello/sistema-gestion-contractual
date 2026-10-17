"""
Módulo 10 — Reportes SIVICOF (Contraloría de Bogotá)

Arma el paquete mensual que el área de Presupuesto envía antes del día 9 del
mes siguiente al corte:

  SIVICOF <MES>/
    01_000000005_AAAAMMDD.xlsx   formulario 400 CB-0104 (reservas / OxP)
    CBN-0001_<Mes>_<AAAA>.pdf     certificaciones
    CBN-1093_<Mes>_<AAAA>.pdf
    CBN-1098_<Mes>_<AAAA>.pdf
    CBN-1109_<Mes>_<AAAA>.pdf
    <decreto>.pdf                 solo si hubo modificación presupuestal

Fuentes (BogData / SAP):
  · Reporte de Ejecución de Gastos con corte al último día del mes.
  · Reporte de Anulaciones de CRP.

CB-0104: solo se diligencian C5 y las columnas E, F, G, H, J, K de las filas
11, 12, 19 y 20. Las columnas I, L (porcentajes) y M (saldo) son fórmulas del
formato; aquí solo se recalcula su valor en caché para que el archivo se vea
bien aunque se abra en un visor que no recalcula.

  E  Reservas constituidas   = Ejecución  "Apropiación Vigente // RESERVAS CONSTITUIDAS"
  F  Anulaciones del mes     = Anulaciones CRP: Σ(Liberación total + Liberación parcial)
                               del rubro con Fecha Liberación dentro del mes del corte
  G  Anulaciones acumuladas  = Anulaciones CRP: misma suma para los meses ANTERIORES
                               al del corte (ej. corte julio: F = julio, G = hasta junio).
                               Solo registros con Número de Oficio/Acta de anulación
                               (las correcciones "ERROR EN CARGUE" de ene-mar no tienen).
  H  Reservas definitivas    = Ejecución  "CDP Acumulado // RESERVAS DEFINITIVAS"
  J  Autorización giro mes   = Ejecución  "Giro Mes Presupuestal // AUTORIZACION DE GIRO DEL MES"
  K  Autorización giro acum. = Ejecución  "Giros Acumulados Ppto // AUTORIZACION DE GIRO ACUMULADA"

Controles contra la Ejecución SAP que se devuelven para revisión (no bloquean):
  · F contra el "CDP Mes" (deberían coincidir con signo opuesto).
  · F + G contra "Saldo Apr.Disponible" (= E − H). Una diferencia suele ser un CRP
    anulado cuyo saldo quedó en el CDP ("Saldo p. Comprometer"), que no reduce la
    reserva definitiva.
"""
import calendar
import io
import os
import re
import unicodedata
import zipfile
from datetime import date, datetime

import pandas as pd

# ---------------------------------------------------------------------------
# Configuración de la entidad (para otras alcaldías basta con cambiar esto)
# ---------------------------------------------------------------------------
ENTIDAD_CODIGO = "000000005"                      # código de la entidad en SIVICOF
FONDO = "Fondo de Desarrollo Local de Usme"
RESPONSABLE_NOMBRE = "JORGE ANDRÉS GONZALEZ DÍAZ"
RESPONSABLE_CARGO = "Responsable del área de Presupuesto FDLU"
PROYECTO = "Richard Bello Roncancio – Contratista Apoyo Presupuesto - FDLU"
PROYECTO_INICIALES = "RB"
REVISO = ("Jorge Andrés González Díaz - Profesional universitario 219 -15 "
          "Responsable de Presupuesto - FDLU")
PIE_DIRECCION = [
    "Alcaldía Local de Usme",
    "Carrera 14 A No. 137 B Sur - 32",
    "Plazoleta Fundacional de Usme",
    "(Usme Centro)",
    "Código Postal: 110541",
    "Tel. 7693100 Ext. 306",
    "Información Línea 195",
    "www.usme.gov.co",
]

# Rubros del CB-0104 -> fila del formato
RUBROS = [
    ("O219001", 11),
    ("O219002", 12),
    ("O230689", 19),
    ("O230690", 20),
]

TEXTO_1093_SIN_MODIFICACION = ("no se realizaron modificaciones presupuestales de "
                               "ingresos y gastos durante el mes de {mes} de {anio}")

CERTIFICADOS = {
    "CBN-0001": {
        "titulo": "INFORME CBN – 0001 NO EXISTENCIA DE RECURSOS DE TESORERIA, "
                  "INVERSIONES FINANCIERAS",
        "cuerpo": "Mediante la presente certifico que, en el proceso presupuestal del "
                  "{fondo} no se cuenta con recursos de tesorería para realizar inversiones "
                  "financieras por la naturaleza jurídica y administrativa, al pertenecer al "
                  "nivel central y ser dependencia de la Secretaría Distrital de Gobierno del "
                  "D.C., por consiguiente, no se cuenta con Tesorería propia.",
        "tam": 11,
    },
    "CBN-1093": {
        "titulo": "INFORME CBN – 1093 MODIFICACIONES PRESUPUESTALES DE INGRESOS Y GASTOS",
        "cuerpo": "Mediante la presente certifico que, en el proceso presupuestal del "
                  "{fondo} {texto_1093}.",
        "tam": 12,
    },
    "CBN-1098": {
        "titulo": "INFORME CBN – 1098 SOBRE EVOLUCIÓN DE PATRIMONIOS AUTONOMOS O "
                  "ENCARGOS FIDUCIARIOS",
        "cuerpo": "Mediante la presente certifico que, en el proceso presupuestal del "
                  "{fondo}, no se han suscrito Contratos Públicos de Encargos Fiduciarios.",
        "tam": 12,
    },
    "CBN-1109": {
        "titulo": "INFORME CBN – 1109 EGRESOS DE PATRIMONIOS AUTONOMOS DESAGREGADOS "
                  "POR TERCEROS",
        "cuerpo": "Mediante la presente certifico que, en el proceso presupuestal del "
                  "{fondo}, no se cuenta con recursos de Egresos de Patrimonios Autónomos "
                  "desagregados por terceros.",
        "tam": 12,
    },
}

DIR_PLANTILLAS = os.path.join(os.path.dirname(__file__), "..", "plantillas", "sivicof")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MESES_ABREV = ["Ene", "Feb", "Mar", "Abr", "Mayo", "Jun", "Jul", "Ago", "Sept",
               "Oct", "Nov", "Dic"]


class ErrorSivicof(ValueError):
    """Error de datos de entrada que se muestra tal cual al usuario."""


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------
def _norm(texto):
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", t).strip().upper()


def _num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if not s:
        return 0.0
    # 1.234.567,89 -> 1234567.89 ; 1234567.89 se deja igual
    if "," in s and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _entero(v):
    return int(round(v))


def _leer_excel(origen):
    try:
        return pd.read_excel(origen, header=None, engine="calamine")
    except Exception:
        if hasattr(origen, "seek"):
            origen.seek(0)
        return pd.read_excel(origen, header=None)


def fecha_corte(anio, mes):
    return date(anio, mes, calendar.monthrange(anio, mes)[1])


# ---------------------------------------------------------------------------
# Lectura de la Ejecución de Gastos
# ---------------------------------------------------------------------------
COLUMNAS_EJECUCION = {
    "E": "RESERVAS CONSTITUIDAS",
    "H": "RESERVAS DEFINITIVAS",
    "J": "GIRO DEL MES",
    "K": "GIRO ACUMULADA",
    "cdp_mes": "CDP MES",
    "saldo_disponible": "SALDO APR",
    "saldo_comprometer": "SALDO P. COMPROMETER",
}


def leer_ejecucion(origen):
    df = _leer_excel(origen)
    # La fila de encabezados es la primera que contenga "RESERVAS CONSTITUIDAS"
    fila_enc = None
    for i in range(min(len(df), 15)):
        if any("RESERVAS CONSTITUIDAS" in _norm(v) for v in df.iloc[i].tolist()):
            fila_enc = i
            break
    if fila_enc is None:
        raise ErrorSivicof("El reporte de Ejecución no tiene la columna "
                           "'Apropiación Vigente // RESERVAS CONSTITUIDAS'. "
                           "¿Es el reporte de Ejecución de Gastos de SAP?")

    encabezados = [_norm(v) for v in df.iloc[fila_enc].tolist()]
    cols = {}
    for clave, patron in COLUMNAS_EJECUCION.items():
        idx = next((j for j, h in enumerate(encabezados) if patron in h), None)
        if idx is None:
            raise ErrorSivicof(f"El reporte de Ejecución no tiene la columna '{patron}'.")
        cols[clave] = idx

    datos = {}
    codigos = {c for c, _ in RUBROS}
    for i in range(fila_enc + 1, len(df)):
        celda = str(df.iat[i, 0] if pd.notna(df.iat[i, 0]) else "").strip()
        if not celda:
            continue
        codigo = celda.split()[0].upper()
        if codigo in codigos and codigo not in datos:
            datos[codigo] = {k: _num(df.iat[i, j]) for k, j in cols.items()}

    faltan = [c for c in codigos if c not in datos]
    if faltan:
        raise ErrorSivicof("No se encontraron en la Ejecución los rubros: "
                           + ", ".join(sorted(faltan)))
    return datos


# ---------------------------------------------------------------------------
# Lectura del Reporte de Anulaciones de CRP
# ---------------------------------------------------------------------------
def leer_anulaciones(origen):
    df = _leer_excel(origen)
    fila_enc = None
    for i in range(min(len(df), 15)):
        fila = [_norm(v) for v in df.iloc[i].tolist()]
        if "RUBRO" in fila and any("FECHA LIBERACION" in h for h in fila):
            fila_enc = i
            break
    if fila_enc is None:
        raise ErrorSivicof("El reporte de Anulaciones no tiene las columnas "
                           "'Rubro' y 'Fecha Liberación'.")

    encabezados = [_norm(v) for v in df.iloc[fila_enc].tolist()]

    def col(nombre, obligatoria=True):
        for j, h in enumerate(encabezados):
            if h == nombre or h.startswith(nombre):
                return j
        if obligatoria:
            raise ErrorSivicof(f"El reporte de Anulaciones no tiene la columna '{nombre.title()}'.")
        return None

    c_rubro = col("RUBRO")
    c_total = col("LIBERACION TOTAL")
    c_parcial = col("LIBERACION PARCIAL")
    c_fecha = col("FECHA LIBERACION")
    c_oficio = col("NUMERO OFICIO", obligatoria=False)

    cuerpo = df.iloc[fila_enc + 1:]
    out = pd.DataFrame({
        "rubro": cuerpo[c_rubro].astype(str).str.strip().str.split().str[0].str.upper(),
        "valor": cuerpo[c_total].map(_num) + cuerpo[c_parcial].map(_num),
        "fecha": pd.to_datetime(cuerpo[c_fecha], errors="coerce", dayfirst=True),
        "con_oficio": (cuerpo[c_oficio].notna() & (cuerpo[c_oficio].astype(str).str.strip() != "")
                       if c_oficio is not None else True),
    })
    return out[out["fecha"].notna()]


# ---------------------------------------------------------------------------
# Cálculo del CB-0104
# ---------------------------------------------------------------------------
def calcular(origen_ejecucion, origen_anulaciones, anio, mes):
    ejec = leer_ejecucion(origen_ejecucion)
    anul = leer_anulaciones(origen_anulaciones)
    anul = anul[anul["fecha"].dt.year == anio]
    del_mes = anul[anul["fecha"].dt.month == mes]
    previas = anul[(anul["fecha"].dt.month < mes) & anul["con_oficio"]]

    filas, alertas = [], []
    for codigo, fila in RUBROS:
        e = ejec[codigo]
        F = del_mes.loc[del_mes["rubro"] == codigo, "valor"].sum()
        G = previas.loc[previas["rubro"] == codigo, "valor"].sum()
        f_sap = -e["cdp_mes"]
        g_sap = e["saldo_disponible"] - f_sap

        r = {
            "rubro": codigo, "fila": fila,
            "E": _entero(e["E"]), "F": _entero(F), "G": _entero(G), "H": _entero(e["H"]),
            "J": _entero(e["J"]), "K": _entero(e["K"]),
            "control_f_sap": _entero(f_sap),
            "control_g_sap": _entero(g_sap),
            "control_total_sap": _entero(e["saldo_disponible"]),
            "saldo_comprometer": _entero(e["saldo_comprometer"]),
        }
        r["I"] = r["H"] / r["E"] if r["E"] else 0
        r["L"] = r["K"] / r["H"] if r["H"] else 0
        r["M"] = r["H"] - r["K"]
        filas.append(r)

        if r["F"] != r["control_f_sap"]:
            alertas.append(
                f"{codigo}: anulaciones del mes según CRP ({r['F']:,}) distintas al "
                f"CDP Mes de la Ejecución ({r['control_f_sap']:,}).".replace(",", "."))
        total = r["F"] + r["G"]
        if total != r["control_total_sap"]:
            dif = total - r["control_total_sap"]
            extra = (" Coincide con el Saldo p. Comprometer: CRP anulado cuyo saldo quedó en el CDP"
                     " (no rebaja la reserva definitiva)."
                     if abs(dif) == abs(r["saldo_comprometer"]) and dif else "")
            alertas.append(
                f"{codigo}: anulaciones CRP mes + acumuladas ({total:,}) difieren de Reservas "
                f"constituidas − definitivas en SAP ({r['control_total_sap']:,}) por "
                f"{dif:,}.{extra}".replace(",", "."))
        if r["F"] < 0 or r["G"] < 0:
            alertas.append(f"{codigo}: valor negativo en anulaciones; revise los reportes.")

    return {"filas": filas, "alertas": alertas,
            "fecha_corte": fecha_corte(anio, mes).isoformat()}


# ---------------------------------------------------------------------------
# Excel CB-0104 (se edita el XML de la plantilla para no perder logo ni estilos)
# ---------------------------------------------------------------------------
def _serial_excel(d):
    return (d - date(1899, 12, 30)).days


def _poner_valor(xml, ref, valor):
    patron = re.compile(r'<c r="%s"([^>]*?)(?:/>|>(.*?)</c>)' % ref, re.S)
    m = patron.search(xml)
    if not m:
        raise ErrorSivicof(f"La plantilla CB-0104 no tiene la celda {ref}.")
    attrs = re.sub(r'\s+t="[^"]*"', "", m.group(1))
    interior = m.group(2) or ""
    formula = re.search(r"<f>.*?</f>|<f [^>]*?(?:/>|>.*?</f>)", interior, re.S)
    nuevo = f'<c r="{ref}"{attrs}>{formula.group(0) if formula else ""}<v>{valor}</v></c>'
    return xml[:m.start()] + nuevo + xml[m.end():]


def generar_cb0104(filas, corte):
    ruta = os.path.join(DIR_PLANTILLAS, "CB-0104.xlsx")
    salida = io.BytesIO()
    with zipfile.ZipFile(ruta) as zin, zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            datos = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                xml = datos.decode("utf-8")
                xml = _poner_valor(xml, "C5", _serial_excel(corte))
                for r in filas:
                    n = r["fila"]
                    for col in "EFGHJK":
                        xml = _poner_valor(xml, f"{col}{n}", r[col])
                    xml = _poner_valor(xml, f"I{n}", repr(r["I"]))
                    xml = _poner_valor(xml, f"L{n}", repr(r["L"]))
                    xml = _poner_valor(xml, f"M{n}", r["M"])
                datos = xml.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                xml = datos.decode("utf-8")
                if "fullCalcOnLoad" not in xml:
                    xml = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', xml)
                datos = xml.encode("utf-8")
            zout.writestr(item, datos)
    return salida.getvalue()


def nombre_cb0104(corte):
    return f"01_{ENTIDAD_CODIGO}_{corte:%Y%m%d}.xlsx"


# ---------------------------------------------------------------------------
# Textos de los certificados
# ---------------------------------------------------------------------------
_UNIDADES = ["", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve",
             "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete",
             "dieciocho", "diecinueve", "veinte", "veintiún", "veintidós", "veintitrés",
             "veinticuatro", "veinticinco", "veintiséis", "veintisiete", "veintiocho",
             "veintinueve"]
_DECENAS = {3: "treinta", 4: "cuarenta", 5: "cincuenta", 6: "sesenta", 7: "setenta",
            8: "ochenta", 9: "noventa"}


def _num_letras(n):
    if n < 30:
        return _UNIDADES[n]
    d, u = divmod(n, 10)
    return _DECENAS[d] + ("" if u == 0 else " y " + _UNIDADES[u])


def _anio_letras(anio):
    resto = anio - 2000
    if resto == 0:
        return "dos mil"
    s = _num_letras(resto)
    s = "veintiuno" if s == "veintiún" else re.sub(r"\bun$", "uno", s)
    return "dos mil " + s


def texto_fecha_firma(f):
    """'a los nueve (09) días del mes de octubre de dos mil veintiséis (2026)'"""
    if f.day == 1:
        dia = "al primer (01) día"
    else:
        dia = f"a los {_num_letras(f.day)} ({f.day:02d}) días"
    return f"{dia} del mes de {MESES[f.month - 1]} de {_anio_letras(f.year)} ({f.year})"


def texto_1093(hubo_modificacion, corte, numero_decreto="", fecha_decreto=None):
    if not hubo_modificacion:
        return TEXTO_1093_SIN_MODIFICACION.format(mes=MESES[corte.month - 1], anio=corte.year)
    if not str(numero_decreto).strip():
        raise ErrorSivicof("Indique el número del Decreto Local.")
    if fecha_decreto:
        cuando = f"del {fecha_decreto.day:02d} de {MESES[fecha_decreto.month - 1]} de {fecha_decreto.year}"
    else:
        cuando = f"de {MESES[corte.month - 1]} de {corte.year}"
    return f"se realizó modificación presupuestal mediante Decreto Local No. {str(numero_decreto).strip()} {cuando}"


# ---------------------------------------------------------------------------
# PDF de los certificados (reportlab, mismo diseño del formato Word)
# ---------------------------------------------------------------------------
_FUENTES_OK = False


def _registrar_fuentes():
    global _FUENTES_OK
    if _FUENTES_OK:
        return
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    d = os.path.join(DIR_PLANTILLAS, "fonts")
    pdfmetrics.registerFont(TTFont("Garamond", os.path.join(d, "EBGaramond-Regular.ttf")))
    pdfmetrics.registerFont(TTFont("Garamond-Bold", os.path.join(d, "EBGaramond-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("Garamond-BoldItalic", os.path.join(d, "EBGaramond-BoldItalic.ttf")))
    _FUENTES_OK = True


def generar_certificado_pdf(codigo, fecha_firma, txt_1093=""):
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Paragraph

    _registrar_fuentes()
    cfg = CERTIFICADOS[codigo]
    W, H = letter
    X0, X1 = 85.05, W - 85.05          # márgenes 1701 twips como el Word
    ancho = X1 - X0

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    c.setTitle(f"{codigo} - {RESPONSABLE_CARGO}")
    c.setAuthor(RESPONSABLE_NOMBRE.title())

    def img(nombre, x, top, w, h):
        c.drawImage(os.path.join(DIR_PLANTILLAS, nombre), x, H - top - h, w, h, mask="auto")

    def parrafo(texto, top, estilo, x=X0, w=ancho):
        p = Paragraph(texto, estilo)
        _, alto = p.wrap(w, H)
        p.drawOn(c, x, H - top - alto)
        return top + alto

    st_titulo = ParagraphStyle("t", fontName="Garamond-Bold", fontSize=11, leading=12.6,
                               alignment=TA_CENTER)
    tam = cfg["tam"]
    st_cuerpo = ParagraphStyle("c", fontName="Garamond", fontSize=tam,
                               leading=13.5 if tam == 12 else 14.4, alignment=TA_JUSTIFY)
    st_centro = ParagraphStyle("f", fontName="Garamond", fontSize=12, leading=13.4,
                               alignment=TA_CENTER)
    st_centro_b = ParagraphStyle("fb", parent=st_centro, fontName="Garamond-Bold")

    # Encabezado
    img("logo_bogota_gobierno.jpg", 92.5, 70.9, 167.6, 41.5)
    c.setFont("Garamond", 11)
    c.drawString(95.2, H - 164, ".")
    parrafo("C E R T I F I C A C I O N", 178, st_titulo)

    # Título del informe y cuerpo
    y = parrafo(cfg["titulo"], 216, st_titulo)
    cuerpo = cfg["cuerpo"].format(fondo=FONDO, texto_1093=txt_1093)
    y = parrafo(cuerpo, y + 58, st_cuerpo)
    y = parrafo(f"Dado en Bogotá D.C. {texto_fecha_firma(fecha_firma)}.", y + 38, st_cuerpo
                if tam == 12 else ParagraphStyle("d", parent=st_cuerpo, fontSize=12, leading=13.5))

    # Firma
    y = parrafo(RESPONSABLE_NOMBRE, max(y + 100, 470), st_centro_b)
    y = parrafo(RESPONSABLE_CARGO, y, st_centro)

    # Proyectó / Revisó
    base = y + 58
    c.setFont("Garamond", 9)
    c.drawString(X0, H - base, "Proyectó:")
    c.drawString(155.5, H - base, PROYECTO)
    ancho_p = c.stringWidth(PROYECTO, "Garamond", 9)
    c.setFont("Garamond-BoldItalic", 11)
    c.drawString(155.5 + ancho_p + 5, H - base, PROYECTO_INICIALES)
    c.setFont("Garamond", 9)
    c.drawString(X0, H - base - 11.2, "Revisó y Aprobó:")
    c.drawString(153, H - base - 11.2, REVISO)

    # Pie de página
    c.setFont("Garamond-Bold", 8)
    for i, linea in enumerate(PIE_DIRECCION):
        c.drawString(X0, H - 691 - 9 * i, linea)
    img("codigo_formato.png", 235.5, 694.8, 120, 55.4)
    img("logo_alcaldia_mayor.png", 489, 699.5, 51, 50.7)

    c.showPage()
    c.save()
    return buf.getvalue()


def nombre_certificado(codigo, corte):
    return f"{codigo}_{MESES_ABREV[corte.month - 1]}_{corte.year}.pdf"


# ---------------------------------------------------------------------------
# Paquete completo
# ---------------------------------------------------------------------------
def nombre_carpeta(corte):
    return f"SIVICOF {MESES[corte.month - 1].upper()}"


def armar_paquete(origen_ejecucion, origen_anulaciones, anio, mes, fecha_firma,
                  hubo_modificacion=False, numero_decreto="", fecha_decreto=None,
                  decreto=None):
    """decreto: (nombre_archivo, bytes) o None. Devuelve (zip_bytes, nombre_zip, resultado)."""
    corte = fecha_corte(anio, mes)
    resultado = calcular(origen_ejecucion, origen_anulaciones, anio, mes)
    txt = texto_1093(hubo_modificacion, corte, numero_decreto, fecha_decreto)
    carpeta = nombre_carpeta(corte)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{carpeta}/{nombre_cb0104(corte)}", generar_cb0104(resultado["filas"], corte))
        for codigo in CERTIFICADOS:
            z.writestr(f"{carpeta}/{nombre_certificado(codigo, corte)}",
                       generar_certificado_pdf(codigo, fecha_firma, txt))
        if hubo_modificacion and decreto:
            z.writestr(f"{carpeta}/{decreto[0]}", decreto[1])

    resultado.update({
        "texto_1093": txt,
        "texto_fecha_firma": texto_fecha_firma(fecha_firma),
        "carpeta": carpeta,
    })
    return buf.getvalue(), f"{carpeta}.zip", resultado


def correo_sugerido(corte):
    mes = MESES[corte.month - 1].capitalize()
    return {
        "asunto": f"Reporte SIVICOF  {mes} {corte.year}_FDLU",
        "cuerpo": ("Jefe\n\nBuenas tardes\n\n"
                   f"Remito el informe de SIVICOF del mes de {mes} del año en curso, para que "
                   "lo revise y se surta el trámite correspondiente.\n\nCordialmente"),
    }


def parse_fecha(valor):
    if not valor:
        return None
    if isinstance(valor, (date, datetime)):
        return valor if isinstance(valor, date) else valor.date()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(valor).strip()[:10], fmt).date()
        except ValueError:
            pass
    raise ErrorSivicof(f"Fecha inválida: {valor}")
