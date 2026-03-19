import ezdxf

# Probar con un archivo DWG directamente desde la red
archivo = r"\\192.168.2.37\ingenieria\PRODUCCION\AGP PLANOS TECNICOS\ACURA\ACURA RDX 4D U 2013--398\958 00 00.dwg"

print(f"Intentando leer: {archivo}")

try:
    doc = ezdxf.readfile(archivo)
    print('Version DXF:', doc.dxfversion)
    print('Archivo legible!')
    
    msp = doc.modelspace()
    types = {}
    for e in msp:
        t = e.dxftype()
        types[t] = types.get(t, 0) + 1
    
    print('Total entidades:', sum(types.values()))
    print('Top 10 tipos:')
    for t, c in sorted(types.items(), key=lambda x: -x[1])[:10]:
        print(f'  {t}: {c}')
        
    # Buscar tablas
    tables = [e for e in msp if e.dxftype() in ('ACAD_TABLE', 'TABLE')]
    print(f'\nTablas encontradas: {len(tables)}')
    
    # Buscar textos
    texts = [e for e in msp if e.dxftype() in ('TEXT', 'MTEXT')]
    print(f'Textos encontrados: {len(texts)}')
    
except Exception as e:
    print(f'Error: {e}')
    import traceback
    traceback.print_exc()
