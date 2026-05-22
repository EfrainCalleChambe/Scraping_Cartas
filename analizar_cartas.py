import os
import re
import unicodedata
from collections import Counter, defaultdict

import pandas as pd

# Configuracion
INPUT_DIR = "cartas_montemar_columnas"   # carpeta donde estan los txt generados
OUTPUT_DIR = "analisis_cartas_excel"     # carpeta donde se guardara el excel
OUTPUT_FILE = "analisis_cartas.xlsx"

STOPWORDS = set([
    "y", "de", "la", "que", "el", "en", "a", "los", "se", "del", "las", "por", "un", "para",
    "con", "una", "su", "al", "lo", "como", "mas", "pero", "sus", "le", "ya", "o", "este",
    "entre", "esta", "cuando", "muy", "sin", "sobre", "tambien", "me", "hasta", "hay", "donde",
    "quien", "desde", "todo", "nos", "durante", "todos", "uno", "les", "ni", "parte", "frente",
    "esa", "eso", "ese", "aquello", "aquel", "aquella", "aquellos", "aquellas", "estos", "estas",
])


def extraer_emisor_receptor(texto):
    """Extrae emisor y receptor desde el bloque inicial de Titulo."""
    titulo = re.search(
        r"^T(?:i|í|Ã­)tulo:\s*(.*?)(?=^Total de p(?:a|á|Ã¡)ginas:)",
        texto,
        flags=re.IGNORECASE | re.MULTILINE | re.DOTALL,
    )
    if not titulo:
        return "", ""

    titulo_limpio = re.sub(r"\s+", " ", titulo.group(1)).strip()
    partes = re.search(r"\bLetter from\s+(.+?)\s+to\s+(.+)", titulo_limpio, flags=re.IGNORECASE)
    if not partes:
        return "", ""

    emisor = partes.group(1).strip(" ,")
    receptor = partes.group(2).strip(" ,")
    return emisor, receptor


def limpiar_texto(texto):
    """Elimina lineas de metadatos, encabezados de pagina y separadores."""
    lineas = texto.splitlines()
    nuevas = []
    omitir_titulo = False

    for linea in lineas:
        if re.match(r"^T(?:i|í|Ã­)tulo:", linea):
            omitir_titulo = True
            continue
        if re.match(r"^Total de p(?:a|á|Ã¡)ginas:", linea):
            omitir_titulo = False
            continue
        if omitir_titulo:
            continue
        if re.match(r"^=== P", linea):
            continue
        if re.match(r"^=+$", linea):
            continue
        if re.match(r"^-+$", linea):
            continue
        nuevas.append(linea)

    return "\n".join(nuevas)


def es_caracter_de_palabra(caracter):
    """Acepta letras Unicode, incluidas letras modificadoras como superindices."""
    categoria = unicodedata.category(caracter)
    return categoria.startswith("L") or categoria.startswith("M")


def obtener_palabras(texto):
    """Devuelve palabras en minusculas preservando letras Unicode como qᵉ o ultimamᵗᵉ."""
    texto = texto.lower()
    palabras = []
    actual = []

    for caracter in texto:
        if es_caracter_de_palabra(caracter):
            actual.append(caracter)
        else:
            if actual:
                palabras.append("".join(actual))
                actual = []

    if actual:
        palabras.append("".join(actual))

    return [p for p in palabras if p not in STOPWORDS and len(p) > 1]


def contar_frecuencias(palabras):
    return Counter(palabras)


def obtener_asociaciones(palabras, ventana=5):
    """Cuenta pares de palabras que co-ocurren dentro de una ventana."""
    pares = defaultdict(int)
    for i, palabra in enumerate(palabras):
        for j in range(i + 1, min(i + ventana + 1, len(palabras))):
            par = tuple(sorted([palabra, palabras[j]]))
            pares[par] += 1
    return pares


def procesar_carta(ruta_archivo):
    """Lee un archivo, limpia, cuenta frecuencias y asociaciones."""
    with open(ruta_archivo, "r", encoding="utf-8") as f:
        texto = f.read()

    emisor, receptor = extraer_emisor_receptor(texto)
    texto_limpio = limpiar_texto(texto)
    palabras = obtener_palabras(texto_limpio)
    frecuencias = contar_frecuencias(palabras)
    asociaciones = obtener_asociaciones(palabras, ventana=5)
    return frecuencias, asociaciones, emisor, receptor


def guardar_excel(resultados, output_dir):
    """Guarda un Excel consolidado con dos hojas: Frecuencias y Asociaciones."""
    os.makedirs(output_dir, exist_ok=True)
    salida = os.path.join(output_dir, OUTPUT_FILE)

    filas_frecuencias = []
    filas_asociaciones = []

    for resultado in resultados:
        archivo = resultado["archivo"]
        emisor = resultado["emisor"]
        receptor = resultado["receptor"]

        for palabra, frecuencia in resultado["frecuencias"].items():
            filas_frecuencias.append({
                "Archivo": archivo,
                "Emisor": emisor,
                "Receptor": receptor,
                "Palabra": palabra,
                "Frecuencia": frecuencia,
            })

        for (p1, p2), frecuencia in resultado["asociaciones"].items():
            filas_asociaciones.append({
                "Archivo": archivo,
                "Emisor": emisor,
                "Receptor": receptor,
                "Par de palabras": f"{p1} - {p2}",
                "Co-ocurrencias": frecuencia,
            })

    df_freq = pd.DataFrame(filas_frecuencias)
    df_asoc = pd.DataFrame(filas_asociaciones)

    if not df_freq.empty:
        df_freq = df_freq.sort_values(by=["Archivo", "Frecuencia"], ascending=[True, False])
    if not df_asoc.empty:
        df_asoc = df_asoc.sort_values(by=["Archivo", "Co-ocurrencias"], ascending=[True, False])

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_freq.to_excel(writer, sheet_name="Frecuencias", index=False)
        df_asoc.to_excel(writer, sheet_name="Asociaciones", index=False)

    print(f"Guardado: {salida}")


def main():
    archivos = [f for f in os.listdir(INPUT_DIR) if f.endswith(".txt")]
    if not archivos:
        print(f"No se encontraron archivos .txt en {INPUT_DIR}")
        return

    print(f"Procesando {len(archivos)} cartas...")
    resultados = []

    for archivo in archivos:
        ruta = os.path.join(INPUT_DIR, archivo)
        print(f"Analizando: {archivo}")
        frecuencias, asociaciones, emisor, receptor = procesar_carta(ruta)
        if frecuencias:
            resultados.append({
                "archivo": archivo,
                "emisor": emisor,
                "receptor": receptor,
                "frecuencias": frecuencias,
                "asociaciones": asociaciones,
            })
        else:
            print(f"  No se encontraron palabras en {archivo}")

    if resultados:
        guardar_excel(resultados, OUTPUT_DIR)

    print("Analisis completado.")


if __name__ == "__main__":
    main()
