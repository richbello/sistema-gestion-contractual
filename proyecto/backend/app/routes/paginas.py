"""
Rutas que sirven las páginas del frontend (SPA simple por archivos estáticos).
"""
from flask import Blueprint, render_template

paginas_bp = Blueprint("paginas", __name__)


@paginas_bp.route("/")
def inicio():
    return render_template("index.html")
