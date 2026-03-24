# -*- coding: utf-8 -*-
"""
INTENTO 2 - Prueba simple de conversion con ODA
"""

import os
import sys
import subprocess
import shutil
import tempfile

# Configurar salida para UTF-8
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Rutas
ODA_EXE = r"C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"
SERVIDOR = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"

# Intentar importar ezdxf
try:
    import ezdxf
    print(f"[OK] ezdxf version {ezdxf.__version__}")
except ImportError:
    print("[ERROR] ezdxf NO instalado")
    sys.exit(1)

def main():
    print("="*60)
    print("INTENTO 2 - Prueba simple ODA + ezdxf")
    print("="*60)
    
    # Un solo archivo de prueba - el primero que encuentre
    archivo_dwg = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS\1514 55 00.dwg"
    
    print(f"\nArchivo DWG: {archivo_dwg}")
    print(f"Existe: {os.path.exists(archivo_dwg)}")
    
    # Crear carpeta temporal
    temp_dir = tempfile.mkdtemp(prefix="dxf_test_")
    print(f"Carpeta temp: {temp_dir}")
    
    # Convertir - ODA usa carpetas como entrada/salida
    carpeta_entrada = os.path.dirname(archivo_dwg)
    nombre_dwg = os.path.basename(archivo_dwg)
    nombre_sin_ext = os.path.splitext(nombre_dwg)[0]
    
    print(f"\n[1] Convirtiendo con ODA...")
    print(f"    Entrada: {carpeta_entrada}")
    print(f"    Salida: {temp_dir}")
    
    # Comando ODA
    cmd = [
        ODA_EXE,
        carpeta_entrada,
        temp_dir,
        "ACAD2018",
        "DXF",
        "0",  # No subdirectorios
        "0"
    ]
    
    print(f"    Comando: {' '.join(cmd)}")
    
    # Ejecutar con timeout más largo (10 minutos)
    try:
        print("    Esperando conversion (puede tardar varios minutos)...")
        resultado = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=600  # 10 minutos
        )
        
        print(f"    Return code: {resultado.returncode}")
        
        if resultado.stdout:
            print(f"    STDOUT: {resultado.stdout[:500]}")
        if resultado.stderr:
            print(f"    STDERR: {resultado.stderr[:500]}")
            
        # Buscar archivo DXF generado
        print(f"\n[2] Buscando archivo DXF generado...")
        
        archivos_dxf = []
        for root, dirs, files in os.walk(temp_dir):
            for f in files:
                if f.lower().endswith('.dxf'):
                    archivos_dxf.append(os.path.join(root, f))
        
        print(f"    Archivos DXF encontrados: {len(archivos_dxf)}")
        for dxf in archivos_dxf:
            print(f"      - {dxf}")
            tamano = os.path.getsize(dxf)
            print(f"        Tamano: {tamano / 1024:.1f} KB")
        
        if archivos_dxf:
            # Leer el primer DXF
            dxffile = archivos_dxf[0]
            print(f"\n[3] Leyendo {os.path.basename(dxffile)} con ezdxf...")
            
            try:
                doc = ezdxf.readfile(dxffile)
                print(f"    Version DXF: {doc.dxfversion}")
                
                msp = doc.modelspace()
                print(f"    Entidades: {len(msp)}")
                
                # Contar entidades
                tipos = {}
                for ent in msp:
                    t = ent.dxftype()
                    tipos[t] = tipos.get(t, 0) + 1
                
                print(f"    Tipos: {tipos}")
                
                # Buscar tablas y bloques
                tablas = list(msp.query('TABLE'))
                bloques = list(msp.query('INSERT'))
                textos = list(msp.query('TEXT'))
                mtexts = list(msp.query('MTEXT'))
                
                print(f"    TABLAS: {len(tablas)}")
                print(f"    BLOQUES: {len(bloques)}")
                print(f"    TEXT: {len(textos)}")
                print(f"    MTEXT: {len(mtexts)}")
                
                # Mostrar ejemplos de textos
                if textos:
                    print(f"\n    TEXT (ejemplos):")
                    for t in textos[:5]:
                        if hasattr(t, 'text') and t.text:
                            print(f"      - {t.text[:80]}")
                
                if mtexts:
                    print(f"\n    MTEXT (ejemplos):")
                    for t in mtexts[:3]:
                        if hasattr(t, 'text') and t.text:
                            print(f"      - {t.text[:100].replace(chr(10), '|')}")
                
                # Bloques con atributos
                print(f"\n    Bloques con atributos:")
                for b in bloques[:5]:
                    if hasattr(b, 'attribs') and b.attribs:
                        print(f"      Bloque: {b.dxf.name}")
                        for a in b.attribs:
                            if hasattr(a.dxf, 'tag') and hasattr(a.dxf, 'text'):
                                print(f"        {a.dxf.tag}: {a.dxf.text}")
                
                print(f"\n[OK] EXITO! La conversion funciono!")
                
            except Exception as e:
                print(f"    ERROR leyendo DXF: {type(e).__name__}: {e}")
        else:
            print(f"    [WARN] No se genero ningun DXF")
            
    except subprocess.TimeoutExpired:
        print("    [ERROR] Timeout - el archivo es muy grande o hay problemas")
    except Exception as e:
        print(f"    ERROR: {type(e).__name__}: {e}")
    finally:
        # Limpiar
        print(f"\nLimpiando...")
        try:
            shutil.rmtree(temp_dir)
            print("    Carpeta temporal eliminada")
        except:
            pass

if __name__ == "__main__":
    main()
