"""
Fábrica de la aplicación Flask.
Centraliza configuración, CORS y registro de blueprints (módulos).
"""
import os
from flask import Flask
from flask_cors import CORS


def create_app():
    app = Flask(
        __name__,
        static_folder=os.path.join(os.path.dirname(__file__), "..", "..", "frontend", "static"),
        static_url_path="/static",
    )

    app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "..", "uploads")
    app.config["OUTPUT_FOLDER"] = os.path.join(os.path.dirname(__file__), "..", "outputs")
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

    CORS(app)

    # Registro de módulos (blueprints)
    from .routes.extraccion import extraccion_bp
    from .routes.plantilla_pagos import plantilla_bp
    from .routes.cuadre_pac import cuadre_bp
    from .routes.estado_cuenta import estado_cuenta_bp
    from .routes.cuentas_bancarias import cuentas_bp
    from .routes.validacion_plantilla import validacion_bp
    from .routes.paginas import paginas_bp

    app.register_blueprint(paginas_bp)
    app.register_blueprint(extraccion_bp, url_prefix="/api/extraccion")
    app.register_blueprint(plantilla_bp, url_prefix="/api/plantilla-pagos")
    app.register_blueprint(cuadre_bp, url_prefix="/api/cuadre-pac")
    app.register_blueprint(estado_cuenta_bp, url_prefix="/api/estado-cuenta")
    app.register_blueprint(cuentas_bp, url_prefix="/api/cuentas-bancarias")
    app.register_blueprint(validacion_bp, url_prefix="/api/validacion-plantilla")

    @app.route("/api/salud")
    def salud():
        return {"ok": True, "estado": "activo"}

    return app
