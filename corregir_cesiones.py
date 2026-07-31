#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Corrige _detectar_cesion_masivo (ordena por fecha + busca valor final correcto)
import sys, shutil, os, re

ARCHIVO = "frontend/static/js/estadocuenta-masivo.js"

NUEVO = 'def _detectar_cesion_masivo(pagos, dc):\n    pagos=sorted(pagos, key=lambda p:_fecha(_g(_pn(p),"Fecha de pago")) or datetime.max)\n    por_bp={}; orden=[]\n    for p in pagos:\n        pn=_pn(p); bp=_limpiar(_g(pn,"Proveedor"))\n        if not bp: continue\n        if bp not in por_bp: por_bp[bp]=[]; orden.append(bp)\n        por_bp[bp].append(p)\n    if len(orden)<2: return None\n    bpc=orden[0]\n    vi=_num(dc.get("valor_inicial")) if dc else 0.0\n    if not vi:\n        for p in pagos:\n            v=_num(_g(_pn(p),"VALOR FINAL DEL CONTRATO","VALOR INICIAL DEL CONTRATO"))\n            if v>0: vi=v; break\n    pc=por_bp[bpc]; suma=sum(_num(_g(_pn(p),"Valor Bruto")) for p in pc)\n    vces=vi-suma; ces=[]\n    for bp in orden[1:]:\n        pgs=por_bp[bp]; pn0=_pn(pgs[0])\n        ces.append({"bp":bp,"nombre":_txt(_g(pn0,"Nombre")),"pagos":pgs,"valor_cesion":vces})\n    return {"bp_cedente":bpc,"pagos_cedente":pc,"valor_cesion":vces,"cesionarios":ces}'

if not os.path.exists(ARCHIVO):
    print("ERROR: corre desde la raiz del proyecto"); sys.exit(1)
src = open(ARCHIVO, encoding="utf-8").read()
shutil.copy(ARCHIVO, ARCHIVO + ".bak2")
# Reemplazar la funcion _detectar_cesion_masivo completa
patron = re.compile(r'def _detectar_cesion_masivo\(pagos, dc\):.*?return \{"bp_cedente":bpc,"pagos_cedente":pc,"valor_cesion":vces,"cesionarios":ces\}', re.DOTALL)
if not patron.search(src):
    print("ERROR: no encontre la funcion a reemplazar"); sys.exit(1)
src2 = patron.sub(NUEVO, src, count=1)
open(ARCHIVO, "w", encoding="utf-8").write(src2)
print("LISTO. Funcion corregida.")
print("Verifica: node --check", ARCHIVO)