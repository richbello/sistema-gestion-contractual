# -*- coding: utf-8 -*-
"""
Importador: Excel/CSV de cuentas -> app/data/cuentas_bancarias.json
Uso:  python importar_cuentas.py cuentas.xlsx
"""
import sys, os, re, json, unicodedata
from datetime import datetime


def solo_digitos(v):
    if v is None:
        return ""
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return re.sub(r"\D", "", s)


def limpiar_nombre(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip().upper())


def pad(v, ancho):
    d = solo_digitos(v)
    return d.zfill(ancho) if d and len(d) < ancho else d


def leer_tabular(ruta):
    import pandas as pd
    ext = os.path.splitext(ruta)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(ruta, dtype=str)
    else:
        df = pd.read_csv(ruta, dtype=str, sep=None, engine="python")
    df = df.fillna("")

    def buscar_col(*claves):
        for col in df.columns:
            h = limpiar_nombre(col)
            h = "".join(c for c in unicodedata.normalize("NFD", h)
                        if unicodedata.category(c) != "Mn")
            if all(k in h for k in claves):
                return col
        return None

    col_ced = buscar_col("IDENTIF") or buscar_col("CEDULA") or buscar_col("NIT")
    col_nom = buscar_col("NOMBRE")
    col_tipo = buscar_col("TIPO")
    col_banco = buscar_col("CODIGO", "BANCO") or buscar_col("COD", "BANCO")
    col_cuenta = buscar_col("NUMERO", "CUENTA") or buscar_col("CUENTA", "BANCARIA")

    faltan = [n for n, c in {
        "identificación": col_ced, "nombre": col_nom, "tipo": col_tipo,
        "código banco": col_banco, "número cuenta": col_cuenta,
    }.items() if c is None]
    if faltan:
        raise ValueError("No pude identificar estas columnas: " + ", ".join(faltan)
                         + "\nEncabezados: " + ", ".join(map(str, df.columns)))

    filas = []
    for _, r in df.iterrows():
        filas.append((r[col_ced], r[col_nom], r[col_tipo], r[col_banco], r[col_cuenta]))
    return filas


def construir(filas, ruta_fuente):
    cuentas, duplicados, sospechosas, descartadas = {}, [], [], 0
    for ced_raw, nom_raw, tipo_raw, banco_raw, cuenta_raw in filas:
        ced = solo_digitos(ced_raw)
        cuenta = solo_digitos(cuenta_raw)
        if not ced or not cuenta:
            descartadas += 1
            continue
        nombre = limpiar_nombre(nom_raw)
        tipo = pad(tipo_raw, 2)
        banco = pad(banco_raw, 3)
        motivo = []
        if len(ced) < 5 or len(ced) > 12:
            motivo.append(f"cedula longitud {len(ced)}")
        if len(cuenta) < 5:
            motivo.append(f"cuenta corta ({len(cuenta)})")
        if tipo not in ("01", "02"):
            motivo.append(f"tipo '{tipo}'")
        if motivo:
            sospechosas.append((ced, nombre, ", ".join(motivo)))
        reg = {"nombre": nombre, "tipo_cuenta": tipo,
               "codigo_banco": banco, "numero_cuenta": cuenta}
        if ced in cuentas and cuentas[ced]["numero_cuenta"] != cuenta:
            duplicados.append((ced, nombre, cuentas[ced]["numero_cuenta"], cuenta))
        cuentas[ced] = reg
    salida = {"meta": {"generado": datetime.now().isoformat(timespec="seconds"),
                       "fuente": os.path.basename(ruta_fuente), "total": len(cuentas)},
              "cuentas": cuentas}
    return salida, duplicados, sospechosas, descartadas


def main():
    if len(sys.argv) < 2:
        print("Uso: python importar_cuentas.py <archivo.xlsx|.csv>")
        sys.exit(1)
    ruta = sys.argv[1]
    if not os.path.isfile(ruta):
        print(f"No existe el archivo: {ruta}")
        sys.exit(1)

    filas = leer_tabular(ruta)
    salida, duplicados, sospechosas, descartadas = construir(filas, ruta)

    destino = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "app", "data", "cuentas_bancarias.json")
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    with open(destino, "w", encoding="utf-8") as f:
        json.dump(salida, f, ensure_ascii=False, indent=2)

    print("=" * 60)
    print("  IMPORTACION DE CUENTAS BANCARIAS")
    print("=" * 60)
    print(f"  Filas leidas         : {len(filas)}")
    print(f"  Cuentas guardadas    : {salida['meta']['total']}")
    print(f"  Filas descartadas    : {descartadas}")
    print(f"  Duplicados           : {len(duplicados)}")
    print(f"  Filas sospechosas    : {len(sospechosas)}")
    print(f"  Archivo generado     : {destino}")
    print("-" * 60)
    if duplicados:
        print("  DUPLICADOS (quedo la ultima cuenta):")
        for ced, nom, ant, nue in duplicados[:20]:
            print(f"    {ced}  {nom[:28]:<28} {ant} -> {nue}")
        if len(duplicados) > 20:
            print(f"    ... y {len(duplicados)-20} mas")
        print("-" * 60)
    if sospechosas:
        print("  REVISAR:")
        for ced, nom, mot in sospechosas[:20]:
            print(f"    {ced}  {nom[:28]:<28} -> {mot}")
        if len(sospechosas) > 20:
            print(f"    ... y {len(sospechosas)-20} mas")
        print("-" * 60)
    print("  Muestra:")
    for ced in list(salida["cuentas"])[:3]:
        r = salida["cuentas"][ced]
        print(f"    {ced} | {r['nombre']} | {r['tipo_cuenta']} | "
              f"{r['codigo_banco']} | {r['numero_cuenta']}")
    print("=" * 60)


if __name__ == "__main__":
    main()
