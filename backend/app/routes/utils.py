"""
Utilidades compartidas entre los blueprints de la API.
"""
import os
import uuid
from werkzeug.utils import secure_filename


def guardar_archivo(file_storage, carpeta_destino):
    """Guarda un archivo subido con nombre seguro y único, retorna la ruta absoluta."""
    nombre_seguro = secure_filename(file_storage.filename)
    nombre_unico = f"{uuid.uuid4().hex[:8]}_{nombre_seguro}"
    ruta = os.path.join(carpeta_destino, nombre_unico)
    file_storage.save(ruta)
    return ruta


def nombre_salida(carpeta, prefijo, extension="xlsx"):
    return os.path.join(carpeta, f"{prefijo}_{uuid.uuid4().hex[:8]}.{extension}")
