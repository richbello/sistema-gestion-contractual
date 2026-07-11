# -*- coding: utf-8 -*-
"""
Servicio: Registro de cuentas bancarias (CRUD sobre JSON).
Fuente de verdad: backend/app/data/cuentas_bancarias.json
"""
import os
import re
import json
import threading
from datetime import datetime

_DEFAULT_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "..", "data", "cuentas_bancarias.json"
)
_JSON_PATH = os.path.abspath(os.environ.get("CUENTAS_JSON_PATH", _DEFAULT_PATH))

_LOCK = threading.Lock()
_CACHE = None


def _solo_digitos(v):
    if v is None:
        return ""
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\D", "", s)


def _nombre(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip().upper())


def _pad(v, ancho):
    d = _solo_digitos(v)
    return d.zfill(ancho) if d and len(d) < ancho else d


def _cargar():
    global _CACHE
    if _CACHE is not None:
        return _CACHE
    with _LOCK:
        if _CACHE is not None:
            return _CACHE
        if os.path.isfile(_JSON_PATH):
            try:
                with open(_JSON_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                _CACHE = data.get("cuentas", {})
            except Exception:
                _CACHE = {}
        else:
            _CACHE = {}
    return _CACHE


def _persistir():
    with _LOCK:
        os.makedirs(os.path.dirname(_JSON_PATH), exist_ok=True)
        salida = {
            "meta": {
                "generado": datetime.now().isoformat(timespec="seconds"),
                "fuente": "runtime",
                "total": len(_CACHE),
            },
            "cuentas": _CACHE,
        }
        tmp = _JSON_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(salida, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _JSON_PATH)


def buscar_cuenta(cedula):
    cuentas = _cargar()
    ced = _solo_digitos(cedula)
    reg = cuentas.get(ced)
    if reg is None:
        return None
    out = dict(reg)
    out["cedula"] = ced
    return out


def agregar_cuenta(cedula, nombre, tipo_cuenta, codigo_banco,
                   numero_cuenta, sobreescribir=True):
    cuentas = _cargar()
    ced = _solo_digitos(cedula)
    cta = _solo_digitos(numero_cuenta)
    if not ced:
        return {"ok": False, "mensaje": "Cédula/NIT inválido (sin dígitos)."}
    if not cta:
        return {"ok": False, "mensaje": "Número de cuenta inválido."}
    registro = {
        "nombre": _nombre(nombre),
        "tipo_cuenta": _pad(tipo_cuenta, 2) or "02",
        "codigo_banco": _pad(codigo_banco, 3),
        "numero_cuenta": cta,
    }
    existe = ced in cuentas
    if existe and not sobreescribir:
        return {"ok": False,
                "mensaje": f"La cédula {ced} ya existe. Marca sobreescribir para actualizar.",
                "actual": buscar_cuenta(ced)}
    anterior = dict(cuentas[ced]) if existe else None
    cuentas[ced] = registro
    _persistir()
    return {
        "ok": True,
        "mensaje": "Cuenta actualizada." if existe else "Cuenta agregada.",
        "accion": "actualizada" if existe else "agregada",
        "cedula": ced, "anterior": anterior, "registro": registro,
        "total": len(cuentas),
    }


def editar_cuenta(cedula, nombre=None, tipo_cuenta=None, codigo_banco=None,
                  numero_cuenta=None, nueva_cedula=None):
    cuentas = _cargar()
    ced = _solo_digitos(cedula)
    if ced not in cuentas:
        return {"ok": False, "mensaje": f"No existe una cuenta para {ced}."}
    reg = dict(cuentas[ced])
    anterior = dict(reg)
    if nombre is not None and str(nombre).strip() != "":
        reg["nombre"] = _nombre(nombre)
    if tipo_cuenta is not None and str(tipo_cuenta).strip() != "":
        reg["tipo_cuenta"] = _pad(tipo_cuenta, 2)
    if codigo_banco is not None and str(codigo_banco).strip() != "":
        reg["codigo_banco"] = _pad(codigo_banco, 3)
    if numero_cuenta is not None and str(numero_cuenta).strip() != "":
        cta = _solo_digitos(numero_cuenta)
        if not cta:
            return {"ok": False, "mensaje": "Número de cuenta inválido."}
        reg["numero_cuenta"] = cta
    destino = ced
    if nueva_cedula is not None and str(nueva_cedula).strip() != "":
        nced = _solo_digitos(nueva_cedula)
        if not nced:
            return {"ok": False, "mensaje": "La nueva cédula es inválida."}
        if nced != ced and nced in cuentas:
            return {"ok": False, "mensaje": f"Ya existe una cuenta con la cédula {nced}."}
        destino = nced
    if destino != ced:
        del cuentas[ced]
    cuentas[destino] = reg
    _persistir()
    return {"ok": True, "mensaje": "Cuenta modificada.", "cedula": destino,
            "anterior": anterior, "registro": reg, "total": len(cuentas)}


def eliminar_cuenta(cedula):
    cuentas = _cargar()
    ced = _solo_digitos(cedula)
    if ced not in cuentas:
        return {"ok": False, "mensaje": f"No existe una cuenta para {ced}."}
    eliminado = dict(cuentas[ced])
    del cuentas[ced]
    _persistir()
    return {"ok": True, "mensaje": "Cuenta eliminada.", "cedula": ced,
            "eliminado": eliminado, "total": len(cuentas)}


def listar(filtro=None):
    cuentas = _cargar()
    f = _nombre(filtro) if filtro else None
    salida = []
    for ced, r in cuentas.items():
        if f and f not in ced and f not in r.get("nombre", ""):
            continue
        item = dict(r)
        item["cedula"] = ced
        salida.append(item)
    salida.sort(key=lambda x: x.get("nombre", ""))
    return salida


def total():
    return len(_cargar())


def ruta_json():
    return _JSON_PATH


def recargar():
    global _CACHE
    with _LOCK:
        _CACHE = None
    return total()
