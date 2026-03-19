# -*- coding: utf-8 -*-
"""
INTENTO 1 - Script de prueba con ezdxf
Este script intenta leer archivos DWG directamente con ezdxf
para verificar si puede extraer la informacion de las tablas tecnicas.
"""

import os
import sys

# Configurar salida para UTF-8
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Intentar importar ezdxf
try:
    import ezdxf
    print(f"[OK] ezdxf instalado - version {ezdxf.__version__}")
except ImportError:
    print("[ERROR] ezdxf NO esta instalado")
    print("  Ejecutar: pip install ezdxf")
    sys.exit(1)

def probar_dwg(ruta_dwg):
    """Intenta leer un archivo DWG y extraer informacion"""
    print(f"\n{'='*60}")
    print(f"Probando: {ruta_dwg}")
    print('='*60)
    
    try:
        # Intentar abrir el archivo DWG
        doc = ezdxf.readfile(ruta_dwg)
        print(f"[OK] Archivo abierto correctamente")
        print(f"  Version DWG: {doc.dxfversion}")
        
        # Obtener entidades del espacio modelo
        msp = doc.modelspace()
        print(f"\n  Entidades en modelspace: {len(msp)}")
        
        # Contar por tipo
        tipos = {}
        for ent in msp:
            tipo = ent.dxftype()
            tipos[tipo] = tipos.get(tipo, 0) + 1
        
        print("\n  Tipos de entidades encontradas:")
        for tipo, count in sorted(tipos.items(), key=lambda x: -x[1]):
            print(f"    {tipo}: {count}")
        
        # Buscar tablas (TABLE)
        tablas = list(msp.query('TABLE'))
        print(f"\n  Tablas encontradas: {len(tablas)}")
        
        # Buscar bloques con atributos para mayor eficacia
        bloques = list(msp.query('INSERT'))
        print(f"  Bloques (INSERT) encontrados: {len(bloques)}")
        
        # Buscar textos
        textos = list(msp.query('TEXT'))
        mtexts = list(msp.query('MTEXT'))
        print(f"  Textos simples: {len(textos)}")
        print(f"  Textos multilinea: {len(mtexts)}")
        
        # Mostrar algunas entidades de ejemplo
        print("\n  --- Ejemplos de entidades ---")
        
        # Mostrar algunos TEXT
        for i, text in enumerate(texts[:5]):
            if hasattr(text, 'text') and text.text:
                print(f"    TEXT[{i}]: {text.text[:50]}")
        
        # Mostrar algunos MTEXT
        for i, mtext in enumerate(mtexts[:3]):
            if hasattr(mtext, 'text') and mtext.text:
                print(f"    MTEXT[{i}]: {mtext.text[:80]}...")
        
        # Mostrar bloques con atributos
        if bloques:
            print(f"\n  --- Bloques con atributos ---")
            for i, block_ref in enumerate(bloques[:5]):
                if hasattr(block_ref, 'attribs') and block_ref.attribs:
                    print(f"    Bloque[{i}]: {block_ref.dxf.name}")
                    for attr in block_ref.attribs:
                        print(f"      - {attr.dxf.tag}: {attr.dxf.text}")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error al abrir archivo: {type(e).__name__}: {e}")
        return False

def main():
    # Ruta del servidor con planos
    servidor = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
    
    # Verificar si el servidor es accesible
    if not os.path.exists(servidor):
        print(f"[ERROR] No se puede acceder al servidor: {servidor}")
        print("\nIntentando con ruta local de prueba...")
        # Usar una ruta local si el servidor no esta disponible
        servidor = r"."
    
    print(f"Servidor: {servidor}")
    print(f"Existe: {os.path.exists(servidor)}")
    
    # Listar carpetas de vehiculos
    if os.path.exists(servidor):
        try:
            carpetas = [d for d in os.listdir(servidor) if os.path.isdir(os.path.join(servidor, d))]
            print(f"\nCarpetas de vehiculos encontradas: {len(carpetas)}")
            for carpeta in carpetas[:10]:
                print(f"  - {carpeta}")
        except Exception as e:
            print(f"Error listando carpetas: {e}")
    
    # Buscar archivos DWG en el servidor
    print("\nBuscando archivos DWG...")
    archivos_dwg = []
    
    if os.path.exists(servidor):
        for root, dirs, files in os.walk(servidor):
            for file in files:
                if file.lower().endswith('.dwg'):
                    archivos_dwg.append(os.path.join(root, file))
            if len(archivos_dwg) >= 5:  # Limitamos a 5 archivos para la prueba
                break
    
    print(f"Archivos DWG encontrados: {len(archivos_dwg)}")
    
    if archivos_dwg:
        for archivo in archivos_dwg[:5]:
            print(f"  - {archivo}")
        
        # Probar con los primeros 3 archivos
        print("\n" + "="*60)
        print("PROBANDO LECTURA CON EZDXF")
        print("="*60 + "\n")
        
        for archivo in archivos_dwg[:3]:
            probar_dwg(archivo)
    else:
        print("No se encontraron archivos DWG en el servidor")

if __name__ == "__main__":
    main()
