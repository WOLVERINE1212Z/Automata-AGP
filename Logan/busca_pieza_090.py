import os
import re

RUTA_BASE = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
EXTENSIONES_VALIDAS = {".dwg", ".dxf"}

# Detecta 090 como código aislado o separado por guiones, espacios, guiones bajos, etc.
PATRON_090 = re.compile(r"(?<!\d)090(?!\d)")

def contiene_090_en_texto(texto: str) -> bool:
    return bool(PATRON_090.search(texto))

def revisar_dxf_contenido(ruta_archivo: str) -> bool:
    try:
        with open(ruta_archivo, "r", encoding="utf-8", errors="ignore") as f:
            contenido = f.read()
        return contiene_090_en_texto(contenido)
    except Exception as e:
        print(f"[ERROR] No se pudo leer el DXF: {ruta_archivo}")
        print(f"        Motivo: {e}")
        return False

def buscar_piezas_090(ruta_base: str):
    encontrados = []

    print(f"\nBuscando pieza 090 en:\n{ruta_base}\n")

    for raiz, _, archivos in os.walk(ruta_base):
        for archivo in archivos:
            nombre, ext = os.path.splitext(archivo)
            ext = ext.lower()

            if ext not in EXTENSIONES_VALIDAS:
                continue

            ruta_completa = os.path.join(raiz, archivo)
            ruta_texto = ruta_completa.lower()

            detectado_por = []

            # 1. Buscar 090 en nombre o ruta
            if contiene_090_en_texto(archivo) or contiene_090_en_texto(ruta_completa):
                detectado_por.append("nombre/ruta")

            # 2. Si es DXF, buscar 090 dentro del contenido
            if ext == ".dxf":
                if revisar_dxf_contenido(ruta_completa):
                    detectado_por.append("contenido_dxf")

            if detectado_por:
                encontrados.append({
                    "archivo": archivo,
                    "ruta": ruta_completa,
                    "detectado_por": detectado_por
                })

                print("=" * 80)
                print(f"Archivo       : {archivo}")
                print(f"Ruta          : {ruta_completa}")
                print(f"Detectado por : {', '.join(detectado_por)}")

    print("\n" + "=" * 80)
    print(f"Total encontrados: {len(encontrados)}")

    return encontrados

if __name__ == "__main__":
    resultados = buscar_piezas_090(RUTA_BASE)