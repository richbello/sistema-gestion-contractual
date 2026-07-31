#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Inserta la logica de cesiones en estadocuenta-masivo.js (autocontenido)
# Uso: python3 insertar_cesiones.py  (desde la raiz del proyecto)

import sys, shutil, os

ARCHIVO = "frontend/static/js/estadocuenta-masivo.js"

B1 = 'def _sinac(s):\n    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).upper()\ndef _buscar(ws, texto, col=3, desde=1, hasta=None):\n    t=texto.strip().upper(); hasta=hasta or ws.max_row\n    for r in range(desde,hasta+1):\n        v=ws.cell(row=r,column=col).value\n        if v and t in str(v).strip().upper(): return r\n    return None\ndef _detectar_cesion_masivo(pagos, dc):\n    por_bp={}; orden=[]\n    for p in pagos:\n        pn=_pn(p); bp=_limpiar(_g(pn,"Proveedor"))\n        if not bp: continue\n        if bp not in por_bp: por_bp[bp]=[]; orden.append(bp)\n        por_bp[bp].append(p)\n    if len(orden)<2: return None\n    bpc=orden[0]\n    vi=_num(dc.get("valor_inicial")) if dc else 0.0\n    if not vi and pagos: vi=_num(_g(_pn(pagos[0]),"VALOR FINAL DEL CONTRATO"))\n    pc=por_bp[bpc]; suma=sum(_num(_g(_pn(p),"Valor Bruto")) for p in pc)\n    vces=vi-suma; ces=[]\n    for bp in orden[1:]:\n        pgs=por_bp[bp]; pn0=_pn(pgs[0])\n        ces.append({"bp":bp,"nombre":_txt(_g(pn0,"Nombre")),"pagos":pgs,"valor_cesion":vces})\n    return {"bp_cedente":bpc,"pagos_cedente":pc,"valor_cesion":vces,"cesionarios":ces}\ndef _llenar_bloque_cesion(ws, c, contrato, dc, fila_titulo):\n    m=_mm(ws)\n    f_cto=_buscar(ws,"CTO Y VIG",col=3,desde=fila_titulo)\n    if not f_cto: return False\n    _wrc(ws,f_cto,4,contrato,m=m); _wrc(ws,f_cto,8,c["bp"],m=m)\n    _wrc(ws,f_cto+1,4,c["nombre"],m=m)\n    _wrc(ws,f_cto+2,4,c["valor_cesion"],money=True,m=m)\n    if c["pagos"]:\n        pn0=_pn(c["pagos"][0]); _wrc(ws,f_cto+2,8,_limpiar(_g(pn0,"Numero RP")),m=m)\n    if dc:\n        fti=_fecha(dc.get("fecha_term_inicial")); ftf=_fecha(dc.get("fecha_term_final"))\n        if fti: _wrc(ws,f_cto+3,4,fti+timedelta(days=1),m=m)\n        if ftf: _wrc(ws,f_cto+3,8,ftf,m=m)\n    f_hdr=_buscar(ws,"PERIODO",col=3,desde=f_cto)\n    if not f_hdr: return False\n    f_pago=f_hdr+1; f_total=_buscar(ws,"TOTAL",col=2,desde=f_pago)\n    slots=(f_total-f_pago) if f_total else 6\n    for i in range(min(len(c["pagos"]),slots)): _dm(ws,f_pago+i)\n    m=_mm(ws); saldo=c["valor_cesion"]; suma=0\n    for i,p in enumerate(c["pagos"][:slots]):\n        pn=_pn(p); mo=_num(_g(pn,"Valor Bruto")); saldo-=mo; suma+=mo; r=f_pago+i\n        fp=_fecha(_g(pn,"Fecha de pago")) or _txt(_g(pn,"Fecha de pago"))\n        _wrc(ws,r,2,i+1,m=m); _wrc(ws,r,3,_txt(_g(pn,"Texto cabecera documento")),m=m)\n        _wrc(ws,r,4,mo,money=True,m=m); _wrc(ws,r,5,saldo,money=True,m=m)\n        _wrc(ws,r,6,_limpiar(_g(pn,"Doc.compensación","Doc compensacion")),m=m)\n        _wrc(ws,r,7,fp,m=m); _wrc(ws,r,8,_limpiar(_g(pn,"Numero RP")),m=m)\n        _wrc(ws,r,9,_limpiar(_g(pn,"CDP Externo")),m=m); _wrc(ws,r,10,_limpiar(_g(pn,"CRP Externo")),m=m)\n    if f_total: _wrc(ws,f_total,4,suma,money=True,m=m); _wrc(ws,f_total,5,saldo,money=True,m=m)\n    return True\n'

B2 = '    # --- CESIONES (deteccion automatica por BP en el historico) ---\n    try:\n        _ces=_detectar_cesion_masivo(pagos, dc)\n        if _ces and _ces["cesionarios"]:\n            _fc=_buscar(ws,"CESION",col=2,desde=1) or _buscar(ws,chr(67)+chr(69)+chr(83)+chr(73)+chr(211)+chr(78),col=2,desde=1)\n            if _fc:\n                _llenar_bloque_cesion(ws, _ces["cesionarios"][0], contrato, dc, _fc)\n    except Exception as _e:\n        pass\n'

if not os.path.exists(ARCHIVO):
    print("ERROR: no encuentro", ARCHIVO)
    print("Corre desde la raiz: /workspaces/sistema-gestion-contractual")
    sys.exit(1)

src = open(ARCHIVO, encoding="utf-8").read()

if "_detectar_cesion_masivo" in src:
    print("YA estaba insertado. No hago nada.")
    sys.exit(0)

shutil.copy(ARCHIVO, ARCHIVO + ".bak")
print("Backup:", ARCHIVO + ".bak")

marca1 = "def llenar(plantilla_b64, pagos_json, contrato, datos_json):"
if marca1 not in src:
    print("ERROR: no encontre def llenar"); sys.exit(1)
src = src.replace(marca1, B1 + "\n" + marca1)

marca2 = "    out=io.BytesIO(); wb.save(out)"
if marca2 not in src:
    print("ERROR: no encontre linea de guardado"); sys.exit(1)
src = src.replace(marca2, B2 + marca2, 1)

open(ARCHIVO, "w", encoding="utf-8").write(src)
print("LISTO. Cesiones insertadas.")
print("Ahora corre: node --check", ARCHIVO)