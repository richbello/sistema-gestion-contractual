"""
Procesamiento LOCAL del Módulo 1 (extracción de causaciones).
Evita el túnel de Codespaces (límite 1 MB): lee los archivos desde el disco
del Codespace en vez de subirlos por HTTP.

Uso:
  python procesar_local.py <carpeta_pdfs> <archivo_basegen.xlsx>
Ejemplo:
  python procesar_local.py entrada/pdfs entrada/BASEGEN_JUNIO04D.xlsx
El Excel resultado queda en: salida/RESULTADO_EXTRACCION.xlsx
"""
import os
import sys
import glob
from app.services.extraccion_service import ejecutar_extraccion

def main():
    if len(sys.argv) < 3:
        print("Uso: python procesar_local.py <carpeta_pdfs> <archivo_basegen.xlsx>")
        sys.exit(1)

    carpeta_pdfs = sys.argv[1]
    ruta_base    = sys.argv[2]

    if not os.path.isdir(carpeta_pdfs):
        print(f"ERROR: la carpeta de PDFs no existe: {carpeta_pdfs}")
        sys.exit(1)
    if not os.path.isfile(ruta_base):
        print(f"ERROR: el BASEGEN no existe: {ruta_base}")
        sys.exit(1)

    rutas_pdfs = sorted(glob.glob(os.path.join(carpeta_pdfs, "*.pdf")))
    rutas_pdfs += sorted(glob.glob(os.path.join(carpeta_pdfs, "*.PDF")))
    if not rutas_pdfs:
        print(f"ERROR: no se encontraron PDFs en {carpeta_pdfs}")
        sys.exit(1)

    os.makedirs("salida", exist_ok=True)
    ruta_salida = os.path.join("salida", "RESULTADO_EXTRACCION.xlsx")

    print(f"PDFs encontrados : {len(rutas_pdfs)}")
    print(f"BASEGEN          : {ruta_base}")
    print("Procesando...")

    resultado = ejecutar_extraccion(ruta_base, rutas_pdfs, ruta_salida)

    print("\n===== RESULTADO =====")
    for k, v in resultado.items():
        if k == "errores":
            print(f"errores          : {len(v)}")
            for e in v:
                print(f"   - {e}")
        else:
            print(f"{k:17}: {v}")
    print(f"\nExcel generado   : {os.path.abspath(ruta_salida)}")
    print("Descárgalo: clic derecho sobre salida/RESULTADO_EXTRACCION.xlsx en el explorador -> Download")

if __name__ == "__main__":
    main()
