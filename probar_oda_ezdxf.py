# -*- coding: utf-8 -*-
"""
INTENTO 2 - Script de prueba con ODA File Converter + ezdxf
1. Convierte archivos DWG a DXF usando ODAFileConverter
2. Lee los DXF con ezdxf y extrae la informacion
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
    print(f"[OK] ezdxf instalado - version {ezdxf.__version__}")
except ImportError:
    print("[ERROR] ezdxf NO esta instalado")
    print("  Ejecutar: pip install ezdxf")
    sys.exit(1)

def convertir_dwg_a_dxf(archivo_dwg, carpeta_salida):
    """Convierte un archivo DWG a DXF usando ODA File Converter"""
    nombre_archivo = os.path.basename(archivo_dwg)
    carpeta_entrada = os.path.dirname(archivo_dwg)
    
    # ODA requiere rutas sin comillas y parametros específicos
    # ODAFileConverter.exe "input_folder" "output_folder" ACAD2018 DXF 1 0
    cmd = [
        ODA_EXE,
        carpeta_entrada,
        carpeta_salida,
        "ACAD2018",  # Version de AutoCAD
        "DXF",        # Formato de salida
        "1",         # Include subdirectories
        "0"          # Audit before save
    ]
    
    print(f"  Ejecutando ODA: {nombre_archivo} -> DXF...")
    
    try:
        resultado = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 minutos por archivo
        )
        
        if resultado.returncode == 0:
            # Buscar el archivo DXF generado
            nombre_dxf = os.path.splitext(nombre_archivo)[0] + ".dxf"
            ruta_dxf = os.path.join(carpeta_salida, nombre_dxf)
            
            if os.path.exists(ruta_dxf):
                print(f"    [OK] Convertido: {nombre_dxf}")
                return ruta_dxf
            else:
                # ODA puede haber creado subcarpetas
                for root, dirs, files in os.walk(carpeta_salida):
                    for f in files:
                        if f.endswith('.dxf'):
                            return os.path.join(root, f)
                print(f"    [WARN] No se encontro archivo DXF generado")
                return None
        else:
            print(f"    [ERROR] ODA retorno codigo {resultado.returncode}")
            print(f"    stderr: {resultado.stderr[:200] if resultado.stderr else 'N/A'}")
            return None
            
    except subprocess.TimeoutExpired:
        print(f"    [ERROR] Timeout convirtiendo archivo")
        return None
    except Exception as e:
        print(f"    [ERROR] {type(e).__name__}: {e}")
        return None

def leer_dxf(ruta_dxf):
    """Lee un archivo DXF con ezdxf y extrae informacion"""
    print(f"  Leyendo DXF con ezdxf...")
    
    try:
        doc = ezdxf.readfile(ruta_dxf)
        print(f"    [OK] DXF abierto - version: {doc.dxfversion}")
        
        msp = doc.modelspace()
        print(f"    Entidades en modelspace: {len(msp)}")
        
        # Contar tipos de entidades
        tipos = {}
        for ent in msp:
            tipo = ent.dxftype()
            tipos[tipo] = tipos.get(tipo, 0) + 1
        
        print(f"    Tipos de entidades: {dict(tipos)}")
        
        # Buscar tablas, bloques y textos
        tablas = list(msp.query('TABLE'))
        bloques = list(msp.query('INSERT'))
        textos = list(msp.query('TEXT'))
        mtexts = list(msp.query('MTEXT'))
        
        print(f"    TABLAS: {len(tablas)}, BLOQUES: {len(bloques)}, TEXT: {len(textos)}, MTEXT: {len(mtexts)}")
        
        # Extraer datos de bloques con atributos (son los que tienen las tablas técnicas)
        datos_encontrados = []
        
        for block_ref in bloques[:10]:
            if hasattr(block_ref, 'attribs') and block_ref.attribs:
                nombre_bloque = block_ref.dxf.name
                atributos = {}
                for attr in block_ref.attribs:
                    tag = attr.dxf.tag.strip() if hasattr(attr.dxf, 'tag') else ""
                    texto = attr.dxf.text.strip() if hasattr(attr.dxf, 'text') else ""
                    if tag and texto:
                        atributos[tag] = texto
                
                if atributos:
                    datos_encontrados.append({
                        'bloque': nombre_bloque,
                        'atributos': atributos
                    })
        
        # Extraer textos que parecen datos técnicos (buscar patrones como "ACERO -> 30")
        print(f"\n    --- Texts encontrados (ejemplos) ---")
        for i, text in enumerate(texts[:10]):
            if hasattr(text, 'text') and text.text:
                txt = text.text.strip()
                if txt and len(txt) < 100:  # Solo textos cortos
                    print(f"      {txt}")
        
        print(f"\n    --- MTEXT encontrados (ejemplos) ---")
        for i, mtext in enumerate(mtexts[:5]):
            if hasattr(mtext, 'text') and mtext.text:
                txt = mtext.text.strip().replace('\n', ' | ')
                if len(txt) < 200:
                    print(f"      {txt[:150]}...")
        
        if datos_encontrados:
            print(f"\n    --- Bloques con atributos ---")
            for dato in datos_encontrados[:3]:
                print(f"    Bloque: {dato['bloque']}")
                for k, v in dato['atributos'].items():
                    print(f"      {k}: {v}")
        
        return {
            'tipos': tipos,
            'tablas': len(tablas),
            'bloques': len(bloques),
            'textos': len(textos),
            'mtexts': len(mtexts),
            'datos_bloques': datos_encontrados
        }
        
    except Exception as e:
        print(f"    [ERROR] {type(e).__name__}: {e}")
        return None

def main():
    print("="*60)
    print("INTENTO 2 - ODA File Converter + ezdxf")
    print("="*60)
    
    # Verificar ODA
    if not os.path.exists(ODA_EXE):
        print(f"[ERROR] ODA no encontrado en: {ODA_EXE}")
        return
    
    print(f"[OK] ODA encontrado: {ODA_EXE}")
    
    # Crear carpeta temporal para DXF convertidos
    temp_dir = tempfile.mkdtemp(prefix="dwg_to_dxf_")
    print(f"[OK] Carpeta temporal: {temp_dir}")
    
    # Buscar pocos archivos DWG para prueba (solo en una carpeta de vehículo)
    carpeta_prueba = os.path.join(SERVIDOR, "AGP PLANOS TECNICOS")
    
    print(f"\nBuscando archivos DWG en: {carpeta_prueba}")
    archivos_dwg = []
    
    if os.path.exists(carpeta_prueba):
        for file in os.listdir(carpeta_prueba):
            if file.lower().endswith('.dwg'):
                archivos_dwg.append(os.path.join(carpeta_prueba, file))
    
    # Si no hay, buscar en cualquier subcarpeta
    if not archivos_dwg:
        print("Buscando en subcarpetas...")
        for root, dirs, files in os.walk(SERVIDOR):
            for file in files:
                if file.lower().endswith('.dwg'):
                    archivos_dwg.append(os.path.join(root, file))
            if len(archivos_dwg) >= 3:
                break
    
    print(f"Archivos DWG encontrados para prueba: {len(archivos_dwg)}")
    
    if not archivos_dwg:
        print("[ERROR] No se encontraron archivos DWG")
        return
    
    # Probar con los primeros 2 archivos
    for i, archivo_dwg in enumerate(archivos_dwg[:2]):
        print(f"\n{'='*60}")
        print(f"PRUEBA {i+1}: {os.path.basename(archivo_dwg)}")
        print('='*60)
        
        # Convertir a DXF
        print(f"\n[1] Convirtiendo DWG -> DXF...")
        ruta_dxf = convertir_dwg_a_dxf(archivo_dwg, temp_dir)
        
        if ruta_dxf and os.path.exists(ruta_dxf):
            # Leer con ezdxf
            print(f"\n[2] Leyendo DXF con ezdxf...")
            datos = leer_dxf(ruta_dxf)
            
            if datos:
                print(f"\n[OK] Extraccion completada!")
                print(f"  - {datos['tablas']} tablas")
                print(f"  - {datos['bloques']} bloques")
                print(f"  - {datos['textos'] + datos['mtexts']} textos")
        else:
            print(f"\n[ERROR] No se pudo convertir el archivo")
    
    # Limpiar
    print(f"\nLimpiando carpeta temporal...")
    try:
        shutil.rmtree(temp_dir)
        print("[OK] Limpieza completada")
    except:
        pass
    
    print(f"\n{'='*60}")
    print("FIN DE PRUEBA - INTENTO 2")
    print("="*60)

if __name__ == "__main__":
    main()
