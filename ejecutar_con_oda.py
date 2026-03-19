"""
EXTRACTOR DE PLANOS - puta mrda tan fea
"""
import os
import json
from pathlib import Path
from datetime import datetime

# ==================== CONFIGURACION ====================
BASE_PATH = r"C:\output_dxf"
ORIGINAL_PATH = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS"
OUTPUT_JSON = Path("./output/datos_planos.json")
SOLO_VEHICULO = "ACURA"  # Cambiar a "" para procesar todos

def log(msg):
    print(f"  {msg}", flush=True)

def extract_all_text_from_entity(entity):
    """Extrae todo el texto posible de una entidad"""
    texts = []
    
    try:
        if entity.dxftype() == 'TEXT':
            txt = str(entity.dxf.text).strip()
            if txt and len(txt) > 1:
                texts.append(txt)
        elif entity.dxftype() == 'MTEXT':
            if entity.text:
                txt = str(entity.text).strip()
                if txt and len(txt) > 1:
                    texts.append(txt)
        elif entity.dxftype() == 'INSERT':
            # Buscar atributos
            try:
                for attr in entity.attribs():
                    try:
                        txt = str(attr.dxf.text).strip()
                        if txt and len(txt) > 1:
                            texts.append(txt)
                    except:
                        pass
            except:
                pass
    except:
        pass
    
    return texts

def process_dxf_file(dxf_path):
    """Procesa un archivo DXF y extrae TODOS los datos"""
    try:
        import ezdxf
        doc = ezdxf.readfile(str(dxf_path))
        msp = doc.modelspace()
        
        all_data = []
        all_texts = []
        
        # 1. Buscar tablas ACAD_TABLE
        tables_found = 0
        for entity in msp:
            if entity.dxftype() in ('ACAD_TABLE', 'TABLE'):
                try:
                    rows_data = []
                    for row in range(entity.dxf.get('num_rows', 0)):
                        row_data = []
                        for col in range(entity.dxf.get('num_cols', 0)):
                            try:
                                cell = entity.cell(row, col)
                                if cell and cell.text:
                                    text = str(cell.text).strip()
                                    if text and text != ' ':
                                        row_data.append(text)
                            except:
                                pass
                        if row_data:
                            rows_data.append(row_data)
                    
                    if rows_data:
                        tables_found += 1
                        all_data.append({
                            'tipo': 'tabla', 
                            'datos': rows_data, 
                            'filas': len(rows_data),
                            'columnas': len(rows_data[0]) if rows_data else 0
                        })
                        # Agregar textos de la tabla
                        for row in rows_data:
                            all_texts.extend(row)
                except:
                    pass
        
        # 2. Extraer TODOS los textos del archivo
        for entity in msp:
            texts = extract_all_text_from_entity(entity)
            all_texts.extend(texts)
        
        # Agregar textos si hay tablas o si hay suficientes textos
        if tables_found > 0 or len(all_texts) > 0:
            # Filtrar textos relevantes (no solo palabras cortas)
            relevant_texts = []
            for txt in all_texts:
                # Ignorar textos muy cortos o solo símbolos
                if len(txt) > 2:
                    relevant_texts.append(txt)
            
            if relevant_texts:
                all_data.append({
                    'tipo': 'textos',
                    'datos': relevant_texts,
                    'cantidad': len(relevant_texts)
                })
        
        return all_data
        
    except Exception as e:
        return [{'error': str(e)}]

def process_vehicle(vehicle_folder):
    """Procesa un vehículo"""
    name = vehicle_folder.name
    log(f"Vehiculo: {name}")
    
    # Obtener TODOS los archivos DXF
    dxf_files = list(vehicle_folder.rglob('*.dxf'))
    log(f"   Archivos DXF: {len(dxf_files)}")
    
    # Obtener estructura de carpetas - FORMA EXACTA COMO ORIGINAL
    estructura = {}
    for f in dxf_files:
        # Obtener ruta relativa desde la carpeta del vehículo
        rel = f.relative_to(vehicle_folder)
        
        if rel.parent != Path('.'):
            # Construir ruta completa de la carpeta
            carpeta = str(rel.parent).replace('\\', '/')
        else:
            carpeta = 'Raiz'
        
        if carpeta not in estructura:
            estructura[carpeta] = []
        estructura[carpeta].append(rel.name)
    
    data = {
        'vehiculo': name,
        'ruta': str(vehicle_folder),
        'carpetas': len(estructura),
        'total_archivos': len(dxf_files),
        'estructura': estructura,
        'archivos': []
    }
    
    # Procesar cada archivo
    ok = 0
    sin_datos = 0
    
    for f in dxf_files:
        rel = f.relative_to(vehicle_folder)
        
        if rel.parent != Path('.'):
            carpeta = str(rel.parent).replace('\\', '/')
        else:
            carpeta = 'Raiz'
        
        result = {
            'archivo': f.name,
            'ruta_relativa': str(rel).replace('\\', '/'),
            'subcarpeta': carpeta,
            'tipo': 'dxf',
            'tablas': [],
            'num_tablas': 0,
            'procesado': datetime.now().isoformat(),
            'estado': 'ok'
        }
        
        try:
            tables = process_dxf_file(f)
            result['tablas'] = tables
            
            # Contar tablas reales
            num_tablas = len([t for t in tables if t.get('tipo') == 'tabla'])
            result['num_tablas'] = num_tablas
            
            if num_tablas > 0 or len(tables) > 0:
                ok += 1
            else:
                sin_datos += 1
                
        except Exception as e:
            result['tablas'] = [{'error': str(e)}]
            sin_datos += 1
        
        data['archivos'].append(result)
    
    data['archivos_procesados'] = ok
    data['archivos_sin_datos'] = sin_datos
    log(f"   Con datos: {ok}, Sin datos: {sin_datos}")
    
    return data

def main():
    print("\n" + "="*50)
    print("  EXTRACTOR DE PLANOS v5")
    print("  Estructura exacta + mejor extraccion")
    print("="*50)
    
    base = Path(BASE_PATH)
    if not base.exists():
        print(f"ERROR: Ruta no encontrada: {BASE_PATH}")
        return
    
    vehicle_folders = sorted([f for f in base.iterdir() if f.is_dir() and not f.name.startswith('.')])
    
    if SOLO_VEHICULO:
        vehicle_folders = [vf for vf in vehicle_folders if vf.name.upper().startswith(SOLO_VEHICULO.upper())]
    
    log(f"Vehiculos a procesar: {len(vehicle_folders)}")
    
    vehicles = []
    for vf in vehicle_folders:
        vehicles.append(process_vehicle(vf))
        print()
    
    output = {
        'generado': datetime.now().isoformat(),
        'version': '5.0',
        'total_vehiculos': len(vehicles),
        'vehiculos': vehicles
    }
    
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(output, ensure_ascii=False, indent=2))
    
    total = sum(v['total_archivos'] for v in vehicles)
    proc = sum(v['archivos_procesados'] for v in vehicles)
    
    print("\n" + "="*50)
    print(f"  Listo! {len(vehicles)} vehiculos, {total} archivos")
    print(f"  Con datos: {proc}")
    print(f"  Output: {OUTPUT_JSON}")
    print("="*50)

if __name__ == "__main__":
    main()
