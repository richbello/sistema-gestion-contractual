# -*- coding: utf-8 -*-
"""
Normalizacion de texto para cargue en SAP / BOGDATA.

Reglas (Modulo 11 - Validacion de plantillas):
  - Vocales con tilde -> sin tilde (a e i o u u)
  - n con virgulilla -> n
  - Caracteres  * / . , -  -> espacio
  - Espacios multiples colapsados y recorte de extremos
  - Resultado en MAYUSCULAS
"""
import re
import unicodedata

_TRADUCIR_ESPACIO = {ord(c): " " for c in "*/.,-"}


def normalizar_nombre(texto):
    """Sanea un nombre de beneficiario para SAP/BOGDATA."""
    if texto is None:
        return texto
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.translate(_TRADUCIR_ESPACIO)
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()
def quitar_tildes(texto):
    """Quita solo tildes y ñ, conservando mayúsculas/minúsculas,
    puntuación y espacios. Para columnas que no son el nombre."""
    if texto is None:
        return texto
    s = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in s if not unicodedata.combining(c))