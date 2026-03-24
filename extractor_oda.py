"""
CONVERSOR DWG → DXF + EXTRACTOR
================================
Este script:
1. Descarga e instala ODA File Converter
2. Convierte archivos DWG a DXF
3. Extrae las tablas de los archivos

ODA File Converter es una herramienta gratuita de Autodesk
que puede leer casi cualquier versión de DWG.
"""

import os
import sys
import json
import subprocess
import time
import urllib.request
import zipfile
import shutil
from pathlib import Path
from datetime import datetime

# Configuración
ODA_VERSION = "4.3.3"
ODA_URL = f"https://www.opendesign.com/sites/default/files/odaconverter_4.3.3.zip"
INSTALL_DIR = Path("./oda_converter")
EXTRACTED_DIR = INSTALL_DIR / "ODAFileConverter"

def log(msg):
    print(f"  {msg}", flush=True)

def download_oda():
    """Descarga e instala ODA File Converter."""
    if EXTRACTED_DIR.exists():
        log("ODA ya instalado")
        return True
    
    log(f"Descargando ODA File Converter v{ODA_VERSION}...")
    zip_path = INSTALL_DIR / "odaconverter.zip"
    
    try:
        urllib.request.urlretrieve(ODA_URL, zip_path)
        log("Extrayendo...")
        
        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(INSTALL_DIR)
        
        zip_path.unlink()  # Borrar zip
        log("✓ ODA instalado")
        return True
    except Exception as e:
        log(f"Error descargando ODA: {e}")
        return False

def convert_dwg_to_dxf(input_file, output_file, oda_exe):
    """Convierte un archivo DWG a DXF usando ODA."""
    try:
        # ODA usa: ODAFileConverter.exe input output format
        result = subprocess.run([
            str(oda_exe),
            str(input_file),
            str(output_file),
            "DXF"
        ], capture_output=True, timeout=30)
        return output_file.exists()
    except Exception as e:
        return False

def process_with_ezdxf(filepath):
    """Procesa un archivo con ezdxf."""
    try:
        import ezdxf
        doc = ezdxf.readfile(filepath)
        msp = doc.modelspace()
        
        tables = []
        
        # Buscar tablas
        for entity in msp:
            if entity.dxftype() in ('ACAD_TABLE', 'TABLE'):
                tables.append(parse_table(entity))
        
        # Buscar texto
        if not tables:
            tables = extract_text(msp)
        
        return tables
    except Exception as e:
        return [{"error": str(e)}]

def parse_table(entity):
    """Parsea una tabla."""
    try:
        rows = []
        for row in range(entity.dxf.get('num_rows', 0)):
            row_data = []
            for col in range(entity.dxf.get('num_cols', 0)):
                try:
                    cell = entity.cell(row, col)
                    text = cell.dxf.get('text', '') or ''
                    row_data.append(str(text).strip())
                except:
                    row_data.append('')
            if any(row_data):
                rows.append(row_data)
        
        if rows:
            fields = {}
            raw = []
            for row in rows:
                if len(row) >= 2:
                    fields[row[0]] = ' '.join(row[1:])
                    raw.append({'campo': row[0], 'valor': ' '.join(row[1:])})
            
            return {'fields': fields, 'rows': raw, 'num_campos': len(raw)}
    except:
        pass
    return None

def extract_text(msp):
    """Extrae texto agrupado."""
    texts = []
    for entity in msp:
        if entity.dxftype() in ('TEXT', 'MTEXT'):
            try:
                if entity.dxftype() == 'TEXT':
                    content = entity.dxf.get('text', '')
                else:
                    content = entity.plain_mtext()
                texts.append(str(content).strip())
            except:
                continue
    
    if not texts:
        return []
    
    # Agrupar textos en pares campo-valor
    data = []
    for i in range(0, len(texts)-1, 2):
        if i+1 < len(texts):
            data.append({'campo': texts[i], 'valor': texts[i+1]})
    
    if data:
        return [{'fields': {d['campo']: d['valor'] for d in data}, 'rows': data, 'num_campos': len(data)}]
    return []

def main():
    print("="*60)
    print("  EXTRACTOR CON ODA - DWG ANTIGUOS")
    print("="*60)
    
    # Descargar ODA
    if not download_oda():
        print("AAAAAAAAError con ODA. Usando método alternativo...")
        return
    
    oda_exe = EXTRACTED_DIR / "ODAFileConverter.exe"
    if not oda_exe.exists():
        # Buscar en subcarpetas
        for p in INSTALL_DIR.rglob("ODAFileConverter.exe"):
            oda_exe = p
            break
    
    if not oda_exe.exists():
        print("golllll No se encontró ODAFileConverter.exe")
        return
    
    log(f"ODA encontrado: {oda_exe}")
    
    # Procesar vehículos
    base_path = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
    base = Path(base_path)
    
    if not base.exists():
        print(f":0 Ruta no encontrada: {base_path}")
        return
    
    vehicles = []
    
    for vehicle_folder in sorted(base.iterdir()):
        if not vehicle_folder.is_dir() or vehicle_folder.name.startswith('.'):
            continue
        
        print(f"\n📦 Procesando: {vehicle_folder.name}")
        vehicle_data = {
            'vehiculo': vehicle_folder.name,
            'ruta': str(vehicle_folder),
            'archivos': []
        }
        
        # Buscar todos los DWG
        dwg_files = list(vehicle_folder.rglob("*.dwg"))
        log(f"   Encontrados {len(dwg_files)} archivos DWG")
        
        converted_count = 0
        tables_count = 0
        
        for dwg in dwg_files[:10]:  # Procesar solo 10 para probar
            log(f"   Procesando: {dwg.name}")
            
            # Convertir a DXF
            dxf = dwg.with_suffix('.dxf')
            
            if convert_dwg_to_dxf(dwg, dxf, oda_exe):
                converted_count += 1
                
                # Extraer tablas
                tables = process_with_ezdxf(dxf)
                
                vehicle_data['archivos'].append({
                    'archivo': dwg.name,
                    'ruta_relativa': str(dwg.relative_to(vehicle_folder)),
                    'subcarpeta': str(dwg.parent.relative_to(vehicle_folder)),
                    'tablas': tables,
                    'num_tablas': len([t for t in tables if 'error' not in t])
                })
                
                tables_count += len(tables)
                
                # Limpiar DXF temporal
                try:
                    dxf.unlink()
                except:
                    pass
            else:
                vehicle_data['archivos'].append({
                    'archivo': dwg.name,
                    'error': 'No se pudo convertir'
                })
        
        vehicle_data['total_archivos'] = len(vehicle_data['archivos'])
        vehicle_data['fecha_extraccion'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        log(f"   ✓ {converted_count} convertidos, {tables_count} tablas")
        vehicles.append(vehicle_data)
    
    # Guardar JSON
    output = {
        'generado': datetime.now().isoformat(),
        'metodo': 'ODA File Converter + ezdxf',
        'total_vehiculos': len(vehicles),
        'vehiculos': vehicles
    }
    
    output_path = Path("./output/datos_planos.json")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2))
    
    print(f"\n✅ Listo! JSON guardado: {output_path}")
    print(f"   Total vehículos: {len(vehicles)}")

if __name__ == '__main__':
    main()
