import ezdxf
from pathlib import Path

base = Path(r'C:\output_dxf')
files = list(base.rglob('*.dxf'))
print(f'Total archivos: {len(files)}')

# Probar con más detalle
for f in files[:5]:
    print(f'\n=== {f.name} ===')
    
    try:
        doc = ezdxf.readfile(str(f))
        msp = doc.modelspace()
        
        # Contar entidades
        type_count = {}
        for e in msp:
            t = e.dxftype()
            type_count[t] = type_count.get(t, 0) + 1
        
        print('Top 10 tipos:')
        for t, c in sorted(type_count.items(), key=lambda x: -x[1])[:10]:
            print(f'  {t}: {c}')
        
        # Buscar INSERT (bloques)
        inserts = [e for e in msp if e.dxftype() == 'INSERT']
        print(f'\nBloques INSERT: {len(inserts)}')
        
        # Intentar obtener atributos de los bloques
        for ins in inserts[:3]:
            try:
                name = ins.dxf.name
                print(f'  Bloque: {name}')
                
                # Buscar atributos
                if hasattr(ins, 'attribs'):
                    for attr in ins.attribs():
                        try:
                            txt = str(attr.dxf.text).strip()
                            if txt:
                                print(f'    ATTRIB: {txt}')
                        except:
                            pass
            except Exception as e:
                pass
        
        # Buscar en todas las entidades algo con ACERO, OFFSET, BN
        print('\nBuscando datos técnicos...')
        for e in msp:
            try:
                # Buscar en TEXT
                if e.dxftype() == 'TEXT':
                    txt = str(e.dxf.text).strip().upper()
                    if any(pal in txt for pal in ['ACERO', 'OFFSET', 'BN+', 'BN-']):
                        print(f'  TEXT: {e.dxf.text}')
                # Buscar en MTEXT
                elif e.dxftype() == 'MTEXT' and e.text:
                    txt = str(e.text).strip().upper()
                    if any(pal in txt for pal in ['ACERO', 'OFFSET', 'BN+', 'BN-']):
                        print(f'  MTEXT: {e.text[:100]}')
            except:
                pass

    except Exception as e:
        print(f'Error: {e}')
