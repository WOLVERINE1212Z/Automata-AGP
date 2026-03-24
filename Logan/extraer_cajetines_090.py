import os
import re
import sys
import shutil
import subprocess
from pathlib import Path
import ezdxf


# CONFIGURACION
DWG_FILE = r"\\192.168.2.37\ingenieria\PRODUCCION\MULTICAPA LOGAN MACRO\AUTOMATA\DWG\M1566 006 090.dwg" #ruta del la pieza CAD

ODA_EXE = r"C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"

# Carpeta temporal de trabajo
WORK_DIR = Path(r"C:\temp\pieza_090_extraccion") #carpeta temporal 
ODA_IN = WORK_DIR / "in" #se crea una ruta que se usa como entrada de ODA
ODA_OUT = WORK_DIR / "out" # se crea una salida temporal para ODA

# Version de salida DXF aceptada por ODA. Si falla, prueba ACAD2018 o ACAD2013 PARA LAS VERSIONES
ODA_OUTPUT_VERSION = "ACAD2013"
ODA_RECURSIVE = "0"
ODA_AUDIT = "1"


# FUNCIONES
def limpiar_texto(texto: str) -> str: #Normaliza el texto que sale de DXF antes de analizarlo toma el texto sucio con doble espacio u otras cosas para que sea mas facil de trabajar. str: texto en py ... -> str esto significa que entra y devuelve texto
    if texto is None: # si el str esta vacio lo devuelve vacio
        return ""
    texto = texto.replace("\\P", " ") #representa como un salgto de parrafo o salto de linea, es decir que reemplaza el salto especial por el espacio normal
    texto = re.sub(r"\s+", " ", texto)#se limpian espacio de la variable texto que trajo
    return texto.strip()# strip quita espacios al inicio y al final de la variable texto
                        # el return se usa para entregar un valor para usarlo despues. el print es para mostrar algo en la pantalla.


def es_texto_interesante(texto: str) -> bool: #esta funcion sirve para filtrar los textos del dibujo para quedarse solo con lo valioso.    el bool: es un boleano es decir que resive el texto y devuelve true o false
    if not texto:
        return False

    t = texto.upper() # upper convierte todo en mayusculas...

    palabras_clave = [  #palabras que puedan aparecer en el dibujo, solo son pistas de que el texto puede ser util
        "OFFSET",
        "BN",
        "BN INT",
        "BANDA NEGRA",
        "ACERO",
        "STEEL",
        "VIN",
        "ESPESOR",
        "THK",
        "MM",
        "CÓDIGO",
        "CODIGO",
        "PIEZA",
        "PART",
        "SUNROOF",
        "090",
    ]

    if any(p in t for p in palabras_clave): #esp significa que si la palabra esta dentro de texto puede ser interesante 
        return True

    # También deja pasar textos con números técnicos
    if re.search(r"\d", t):
        return True

    return False #significa que si no encontro ninguna palabra clave o numero se devuelve como false, es decir que el texto no es util


def punto_entidad(ent) -> tuple[float, float]: #esta funcion es para obtener la ubicacion de una entidad de texto. devulve dos valores de x, y   esto ayuda a identificar si estan en el mismo cajetin 
    
    # recive una entidad  que se llama ent y devuelve una tuplan ejemplo:(1520.33, 410.12)
    try: # se maneja excepciones
        if ent.dxftype() == "TEXT": # .dxftype esto devuelve el tipo de entidad del DXF
            p = ent.dxf.insert
            return float(p.x), float(p.y)
        if ent.dxftype() == "MTEXT": # MTEXT ES UN TEXTO MAS COMPLEJO 
            p = ent.dxf.insert
            return float(p.x), float(p.y)
        if ent.dxftype() == "ATTRIB": # SON LOS ATRIVUTOS DE BLOQUE
            p = ent.dxf.insert
            return float(p.x), float(p.y)
    except Exception:
        pass
    return 0.0, 0.0


def extraer_texto_entidad(ent) -> str:  # resive una entidad que es ent y devuelve uno que es str o texto
    tipo = ent.dxftype() # se guarda el tipo de identidad 

    try:
        if tipo == "TEXT":
            return limpiar_texto(ent.dxf.text)

        if tipo == "MTEXT":
            return limpiar_texto(ent.text)

        if tipo == "ATTRIB":
            # Si el atributo tiene MTEXT embebido, plain_mtext() da el contenido sin formato
            if getattr(ent, "has_embedded_mtext_entity", False):
                try:
                    return limpiar_texto(ent.plain_mtext())
                except Exception:
                    pass
            return limpiar_texto(ent.dxf.text)
    except Exception:
        return ""

    return ""


def imprimir_bloque_titulo(texto: str): #Muestra los titulos
    print("\n" + "=" * 100)
    print(texto)
    print("=" * 100)


# ODA: DWG a DXF
def convertir_dwg_a_dxf(dwg_path: str) -> Path:  #entra un dwg_path que es la ruta y la convierte en str
    dwg = Path(dwg_path)

    if not dwg.exists():
        raise FileNotFoundError(f"No existe el archivo DWG: {dwg}")

    if not Path(ODA_EXE).exists(): #verific que ODA exista 
        raise FileNotFoundError(
            f"No encontré ODA en:\n{ODA_EXE}\n\n"
            "Revisa la ruta de instalación y actualiza la variable ODA_EXE."
        )

    # Preparar carpetas temporales
    if WORK_DIR.exists():# si ya existe esta carpeta elimina la carpeta
        shutil.rmtree(WORK_DIR)

    # CREA LAS CARPETAS
    ODA_IN.mkdir(parents=True, exist_ok=True)
    ODA_OUT.mkdir(parents=True, exist_ok=True)

    # Copiamos el DWG a una carpeta simple para que ODA convierta solo ese archivo
    dwg_copia = ODA_IN / dwg.name
    shutil.copy2(dwg, dwg_copia) # copia el archivo original a esa carpeta

    # ODA recibe:
    # source_dir target_dir filter output_version recursive audit
    #ARMA EL COMANDO PARA ODA
    cmd = [
        ODA_EXE,
        str(ODA_IN),
        str(ODA_OUT),
        ODA_OUTPUT_VERSION,
        "DXF",
        ODA_RECURSIVE,
        ODA_AUDIT,
    ]

    imprimir_bloque_titulo("1) CONVIRTIENDO DWG A DXF CON ODA")
    print("Comando:")
    print(" ".join(f'"{c}"' if " " in c else c for c in cmd))

    result = subprocess.run(cmd, capture_output=True, text=True) # subprocess.run: ejecutar programas externos y capture_output guada el text=True hace que la salida salga con el texto estandar y legible

    print("\n[STDOUT]")
    print(result.stdout.strip() or "(sin salida)")
    print("\n[STDERR]")
    print(result.stderr.strip() or "(sin salida)")

    if result.returncode not in (0,):
        raise RuntimeError(f"ODA terminó con código {result.returncode}")

    # Buscar el dxf generado 
    posibles = list(ODA_OUT.glob("*.dxf")) + list(ODA_OUT.rglob("*.dxf"))
    if not posibles:
        raise FileNotFoundError(
            "ODA terminó, pero no encontré el archivo DXF convertido en la carpeta de salida."
        )

    dxf_path = posibles[0]
    print(f"\nDXF generado:\n{dxf_path}")
    return dxf_path


# EXTRACCION DE TEXTO / CAJETINES
def recopilar_entidades_texto(layout, origen: str) -> list[dict]:
    resultados = []

    # TEXT
    for ent in layout.query("TEXT"):
        texto = extraer_texto_entidad(ent)
        if not texto:
            continue

        x, y = punto_entidad(ent)
        resultados.append({
            "origen": origen,
            "tipo_entidad": "TEXT",
            "texto": texto,
            "layer": ent.dxf.layer,
            "x": x,
            "y": y,
        })

    # MTEXT
    for ent in layout.query("MTEXT"):
        texto = extraer_texto_entidad(ent)
        if not texto:
            continue

        x, y = punto_entidad(ent)
        resultados.append({
            "origen": origen,
            "tipo_entidad": "MTEXT",
            "texto": texto,
            "layer": ent.dxf.layer,
            "x": x,
            "y": y,
        })

    # INSERT + ATTRIB
    for ins in layout.query("INSERT"):
        nombre_bloque = ins.dxf.name
        for att in ins.attribs:
            texto = extraer_texto_entidad(att)
            if not texto:
                continue

            x, y = punto_entidad(att)
            resultados.append({
                "origen": origen,
                "tipo_entidad": "ATTRIB",
                "bloque": nombre_bloque,
                "tag": getattr(att.dxf, "tag", ""),
                "texto": texto,
                "layer": att.dxf.layer,
                "x": x,
                "y": y,
            })

    return resultados

#AGRUPA LAS COORDENAS TEXTO QUE TENGA UNA Y ES LA MISMA FILA, Y x → LOS ORDENA
def agrupar_por_cercania(items: list[dict], tolerancia_y: float = 8.0) -> list[list[dict]]:
    """
    Agrupa textos que estén aproximadamente en la misma fila.
    Esto ayuda a reconstruir tablas/cajetines simples.
    """
    if not items:
        return []

    ordenados = sorted(items, key=lambda x: (-x["y"], x["x"]))
    grupos = []

    for item in ordenados:
        agregado = False
        for grupo in grupos:
            y_ref = grupo[0]["y"]
            if abs(item["y"] - y_ref) <= tolerancia_y:
                grupo.append(item)
                agregado = True
                break
        if not agregado:
            grupos.append([item])

    for grupo in grupos:
        grupo.sort(key=lambda x: x["x"])

    return grupos

#SI ENCUENTRA UNA PALABRA CLAVE COMO OFFSET, ASUME QUE EL TEXTO SIGUEINTE EN ESA FILA ES SU VALOR ES UNA APROXIMACION RAZONBLE
def detectar_cajetines(textos: list[dict]) -> list[dict]: #esta funcion resive una lista, que este diccionario tiene elementos de texto
    """
    Heurística inicial:
    - busca textos con palabras típicas de cajetín
    - intenta tomar como valor el texto siguiente en la misma fila
    """
    patrones = [# r"..." raw" y el /b es el limite de la palabra, es decir, encuentra la palabra exacta
        r"\bOFFSET\b", 
        r"\bBN\b",
        r"\bBN INT\b",
        r"\bBANDA NEGRA\b",
        r"\bACERO\b",
        r"\bSTEEL\b",
        r"\bOFFSET PC\b",
    ]
    patron = re.compile("|".join(patrones), re.IGNORECASE) #"|".join(patrones) covierte esta lista, por ejemplo: r"\bOFFSET\b|\bBN\b|\bBN INT\b|\bBANDA NEGRA\b|\bACERO\b|\bSTEEL\b|\bOFFSET PC\b"

    grupos = agrupar_por_cercania(textos) # aqui llama la funcon de agrupar_por_cercania esta funcion agrupa los text que estan aproximadamente en la misma fila del plano por ejemplo: OFFSET en = 12
    cajetines = [] # aqui se guarda los cajetines que se detectan en el dibujo

    for fila in grupos: # este bucle recorre para grupo de textos cada fila es un grupo 
        textos_fila = [x["texto"] for x in fila] # esto saca la fila de cada elemento de esa fila
        fila_str = " | ".join(textos_fila) # se usa todos los textos de la fila en una sola cadena esto para despues depurarlo

        for i, item in enumerate(fila): # este bloque recorre cada texto dentro de la fila "i" es la pósicion de la fila, y item es el diccionario del texto
            if patron.search(item["texto"]):
                valor = ""
                if i + 1 < len(fila): # esto pregunta si esiste un elemento despues  i + 1 es la combinacion para que sea la sigueinte posicion
                    valor = fila[i + 1]["texto"]

                #se agrag un diccionario con la lista de resultados
                cajetines.append({ 
                    "nombre": item["texto"],
                    "valor": valor,
                    "layer": item.get("layer", ""),
                    "x": item.get("x", 0.0),
                    "y": item.get("y", 0.0),
                    "fila": fila_str,
                })

    return cajetines # devuelve la lista completa del cajetin detectado


def analizar_dxf(dxf_path: Path):
    imprimir_bloque_titulo("2) LEYENDO DXF Y BUSCANDO TEXTO / CAJETINES")

    doc = ezdxf.readfile(dxf_path)

    resultados = []

    # Modelspace
    msp = doc.modelspace()
    resultados.extend(recopilar_entidades_texto(msp, "MODELSPACE"))

    # Layouts paperspace
    for layout in doc.layouts:
        if layout.name.lower() == "model":
            continue
        resultados.extend(recopilar_entidades_texto(layout, f"LAYOUT:{layout.name}"))

    print(f"Total textos/atributos encontrados: {len(resultados)}")

    interesantes = [r for r in resultados if es_texto_interesante(r["texto"])]
    print(f"Textos/atributos interesantes: {len(interesantes)}")

    imprimir_bloque_titulo("3) TEXTOS INTERESANTES")
    for r in sorted(interesantes, key=lambda x: (x["origen"], -x["y"], x["x"])):
        bloque = r.get("bloque", "")
        tag = r.get("tag", "")
        extra = f" bloque={bloque} tag={tag}" if bloque or tag else ""
        print(
            f"[{r['origen']}] [{r['tipo_entidad']}] "
            f"layer={r['layer']} x={r['x']:.2f} y={r['y']:.2f}{extra} -> {r['texto']}"
        )

    cajetines = detectar_cajetines(interesantes)

    imprimir_bloque_titulo("4) POSIBLES CAJETINES DETECTADOS")
    if not cajetines:
        print("No pude detectar cajetines con la heurística actual.")
        print("Pero arriba ya quedaron impresos los textos interesantes para ajustar la lógica.")
    else:
        for c in cajetines:
            print(
                f"- nombre={c['nombre']} | valor={c['valor']} | "
                f"layer={c['layer']} | x={c['x']:.2f} | y={c['y']:.2f}"
            )
            print(f"  fila: {c['fila']}")

    return interesantes, cajetines


def main():
    try:
        dxf_path = convertir_dwg_a_dxf(DWG_FILE)
        _, cajetines = analizar_dxf(dxf_path)

        imprimir_bloque_titulo("5) RESUMEN FINAL")
        print(f"Archivo origen: {DWG_FILE}")
        print(f"DXF temporal : {dxf_path}")
        print(f"Total cajetines detectados: {len(cajetines)}")

    except Exception as e:
        print("\n[ERROR]")
        print(str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()