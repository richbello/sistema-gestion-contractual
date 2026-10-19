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
    app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB
    app.config["MAX_FORM_MEMORY_SIZE"] = 50 * 1024 * 1024  # 50 MB: uploads mas grandes van a disco

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)

    # CORS: solo el frontend publicado y desarrollo local.
    # Orígenes adicionales (p. ej. un Codespace) se agregan con CORS_EXTRA="https://a,https://b".
    origenes = [
        "https://richbello.github.io",
        r"http://localhost(:\d+)?",
        r"http://127\.0\.0\.1(:\d+)?",
        r"https://.*\.app\.github\.dev",
    ]
    origenes += [o.strip() for o in os.getenv("CORS_EXTRA", "").split(",") if o.strip()]
    CORS(app, origins=origenes, allow_headers=["Content-Type", "X-API-Key"],
         methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
         expose_headers=["Content-Disposition"])

    from .auth import instalar_autenticacion
    instalar_autenticacion(app)

    # Registro de módulos (blueprints)
    from .routes.extraccion import extraccion_bp
    from .routes.plantilla_pagos import plantilla_bp
    from .routes.cuadre_pac import cuadre_bp
    from .routes.cuadre_pac_sj import cuadre_sj_bp
    from .routes.estado_cuenta import estado_cuenta_bp
    from .routes.cuentas_bancarias import cuentas_bp
    from .routes.validacion_plantilla import validacion_bp
    from .routes.oxp_routes import oxp_bp
    from .routes.vfi import vfi_bp
    from .routes.paginas import paginas_bp
    from .routes.extraccion_sipago import extraccion_sipago_bp
    from .routes.sivicof import sivicof_bp

    app.register_blueprint(paginas_bp)
    app.register_blueprint(extraccion_bp, url_prefix="/api/extraccion")
    app.register_blueprint(plantilla_bp, url_prefix="/api/plantilla-pagos")
    app.register_blueprint(cuadre_bp, url_prefix="/api/cuadre-pac")
    app.register_blueprint(cuadre_sj_bp, url_prefix="/api/cuadre-pac-sj")
    app.register_blueprint(estado_cuenta_bp, url_prefix="/api/estado-cuenta")
    app.register_blueprint(cuentas_bp, url_prefix="/api/cuentas-bancarias")
    app.register_blueprint(validacion_bp, url_prefix="/api/validacion-plantilla")
    app.register_blueprint(oxp_bp, url_prefix="/api/oxp")
    app.register_blueprint(vfi_bp)
    app.register_blueprint(extraccion_sipago_bp, url_prefix="/api/extraccion-sipago")
    app.register_blueprint(sivicof_bp, url_prefix="/api/sivicof")

    @app.route("/api/salud")
    def salud():
        return {"ok": True, "estado": "activo"}

    return app
