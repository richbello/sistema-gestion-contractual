"""
Rutas que sirven el frontend (archivos estáticos puros, sin plantillas Jinja,
para que el mismo HTML funcione igual aquí que en GitHub Pages).
"""
import os
from flask import Blueprint, send_from_directory

paginas_bp = Blueprint("paginas", __name__)


def _frontend_dir():
    return os.path.join(os.path.dirname(__file__), "..", "..", "..", "frontend")


@paginas_bp.route("/")
def inicio():
    return send_from_directory(_frontend_dir(), "index.html")


@paginas_bp.route("/config.js")
def config_js():
    return send_from_directory(_frontend_dir(), "config.js")
