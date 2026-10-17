"""
Autenticación por clave de API para todas las rutas /api/*.

- La clave se define en la variable de entorno API_KEY (en Render: Environment).
  Nunca se escribe en el código: el repositorio es público.
- El navegador la envía en el header  X-API-Key  (o en ?k= para los enlaces de
  descarga, que no pueden llevar headers).
- Si API_KEY no está definida, la protección queda APAGADA y se avisa en el log,
  para que un deploy sin la variable no tumbe la aplicación.
- Rutas abiertas: /api/salud y las peticiones OPTIONS (preflight de CORS).
"""
import hmac
import os

from flask import jsonify, request

RUTAS_ABIERTAS = {"/api/salud"}


def _clave_configurada():
    return os.getenv("API_KEY", "").strip()


def instalar_autenticacion(app):
    if not _clave_configurada():
        app.logger.warning("API_KEY no definida: las rutas /api/* están SIN protección.")

    @app.before_request
    def _verificar_clave():
        if request.method == "OPTIONS":
            return None
        ruta = request.path
        if not ruta.startswith("/api/") or ruta in RUTAS_ABIERTAS:
            return None
        esperada = _clave_configurada()
        if not esperada:
            return None
        recibida = (request.headers.get("X-API-Key") or request.args.get("k") or "").strip()
        if recibida and hmac.compare_digest(recibida, esperada):
            return None
        return jsonify({"ok": False, "error": "no_autorizado",
                        "mensaje": "Clave de acceso inválida o ausente."}), 401
