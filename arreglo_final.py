#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Arreglo final cesiones masivo: buscar bloque CESION exacto (no instrucciones) + filtrar PAGADA
import os, sys, re

ARCHIVO = "frontend/static/js/estadocuenta-masivo.js"
if not os.path.exists(ARCHIVO):
    print("Corre desde la raiz del proyecto"); sys.exit(1)
src = open(ARCHIVO, encoding="utf-8").read()
import shutil; shutil.copy(ARCHIVO, ARCHIVO+".bak3")

# ARREGLO 1: _buscar_bloque que matchea la celda cuyo texto es EXACTAMENTE "CESION"/"CESION N"
# (no una frase larga que contiene la palabra)
nueva_func_buscar = '''def _buscar_cesion(ws, desde=1):
    for r in range(desde, ws.max_row+1):
        v = ws.cell(row=r, column=2).value
        if not v: continue
        t = _sinac(str(v).strip())
        # aceptar solo "CESION" o "CESION 1", "CESION 2"... (celda corta, no instrucciones)
        if t == "CESION" or (t.startswith("CESION ") and len(t) < 12):
            return r
    return None'''

# insertar la funcion nueva justo despues de _buscar
marca = "def _detectar_cesion_masivo"
src = src.replace(marca, nueva_func_buscar + "\n" + marca, 1)

# ARREGLO 2: reemplazar la llamada _buscar(ws,"CESION"...) por _buscar_cesion(ws)
viejo_fc = '_fc=_buscar(ws,"CESION",col=2,desde=1) or _buscar(ws,chr(67)+chr(69)+chr(83)+chr(73)+chr(211)+chr(78),col=2,desde=1)'
nuevo_fc = "_fc=_buscar_cesion(ws)"
src = src.replace(viejo_fc, nuevo_fc, 1)

# ARREGLO 3: filtrar PAGADA dentro de _detectar_cesion_masivo
viejo_det = '''def _detectar_cesion_masivo(pagos, dc):
    pagos=sorted(pagos, key=lambda p:_fecha(_g(_pn(p),"Fecha de pago")) or datetime.max)'''
nuevo_det = '''def _detectar_cesion_masivo(pagos, dc):
    pagos=[p for p in pagos if _sinac(_g(_pn(p),"Estatus","Estado")) in ("PAGADA","")]
    pagos=sorted(pagos, key=lambda p:_fecha(_g(_pn(p),"Fecha de pago")) or datetime.max)'''
src = src.replace(viejo_det, nuevo_det, 1)

open(ARCHIVO,"w",encoding="utf-8").write(src)
print("Arreglos aplicados:")
print(" 1. _buscar_cesion (bloque real, no instrucciones)")
print(" 2. llamada actualizada")
print(" 3. filtro PAGADA")