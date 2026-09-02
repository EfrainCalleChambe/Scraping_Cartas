# -*- coding: utf-8 -*-
"""
MODERNIZAR CARTAS DE MONTEMAR CON API DE OPENAI
VERSIÓN CORREGIDA

Entrada esperada:
    montemar_transcripciones/html_por_pagina/*_page_*.html

Salida:
    montemar_moderno/
    ├── txt_moderno_por_pagina/
    ├── comparacion_original_moderno/
    └── indice_modernizacion.csv

Requisitos:
    pip install openai beautifulsoup4 lxml pandas regex

Antes de ejecutar, configura tu API key:
    CMD:
        set OPENAI_API_KEY=tu_api_key
    PowerShell:
        $env:OPENAI_API_KEY="tu_api_key"

Ejecutar:
    python 01_modernizar_montemar_api_CORREGIDO.py
"""

import os
import re
import time
import html
from pathlib import Path
from typing import List, Dict

import pandas as pd
from bs4 import BeautifulSoup, NavigableString, Tag
from openai import OpenAI

try:
    import regex as regex_unicode
    TIENE_REGEX = True
except ImportError:
    regex_unicode = None
    TIENE_REGEX = False


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_ORIGINAL = Path("montemar_transcripciones") / "html_por_pagina"
CARPETA_SALIDA = Path("montemar_moderno")
CARPETA_TXT_MODERNO = CARPETA_SALIDA / "txt_moderno_por_pagina"
CARPETA_COMPARACION = CARPETA_SALIDA / "comparacion_original_moderno"
ARCHIVO_INDICE = CARPETA_SALIDA / "indice_modernizacion.csv"

# Modelo recomendado para texto. Puedes cambiarlo si deseas.
MODELO = "gpt-5-mini"

# Para pruebas, usa por ejemplo MAX_PAGINAS = 3.
# Para procesar todo, deja None.
MAX_PAGINAS = None

# Si ya existe el TXT moderno, no vuelve a gastar API.
SALTAR_SI_EXISTE = True

# Pausa entre llamadas para evitar saturar la API.
SLEEP_SECONDS = 0.5

# Reintentos si hay error temporal.
MAX_REINTENTOS = 3


# ============================================================
# SUPERÍNDICES
# ============================================================

SUPER_MAP = str.maketrans({
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ", "f": "ᶠ",
    "g": "ᵍ", "h": "ʰ", "i": "ᶦ", "j": "ʲ", "k": "ᵏ", "l": "ˡ",
    "m": "ᵐ", "n": "ⁿ", "o": "ᵒ", "p": "ᵖ", "r": "ʳ", "s": "ˢ",
    "t": "ᵗ", "u": "ᵘ", "v": "ᵛ", "w": "ʷ", "x": "ˣ", "y": "ʸ",
    "z": "ᶻ",
    "A": "ᴬ", "B": "ᴮ", "D": "ᴰ", "E": "ᴱ", "G": "ᴳ", "H": "ᴴ",
    "I": "ᴵ", "J": "ᴶ", "K": "ᴷ", "L": "ᴸ", "M": "ᴹ", "N": "ᴺ",
    "O": "ᴼ", "P": "ᴾ", "R": "ᴿ", "T": "ᵀ", "U": "ᵁ", "V": "ⱽ",
    "W": "ᵂ",
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵",
    "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "+": "⁺", "-": "⁻", "=": "⁼", "(": "⁽", ")": "⁾"
})


def a_superindice(texto: str) -> str:
    return texto.translate(SUPER_MAP)


# ============================================================
# UTILIDADES
# ============================================================

def limpiar_nombre_archivo(s: str, max_len: int = 150) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def extraer_id_carta(nombre_archivo: str) -> str:
    return re.sub(r"_page_\d+\.html$", "", nombre_archivo)


def extraer_num_pagina(nombre_archivo: str) -> int:
    m = re.search(r"_page_(\d+)\.html$", nombre_archivo)
    return int(m.group(1)) if m else 0


def normalizar_espacios(txt: str) -> str:
    txt = html.unescape(str(txt))
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r" *\n *", "\n", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()


def asegurar_carpetas():
    CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)
    CARPETA_TXT_MODERNO.mkdir(parents=True, exist_ok=True)
    CARPETA_COMPARACION.mkdir(parents=True, exist_ok=True)


# ============================================================
# EXTRACCIÓN DE TEXTO DEL HTML ORIGINAL
# ============================================================

def texto_diplomatico(node) -> str:
    if isinstance(node, NavigableString):
        return str(node)

    if not isinstance(node, Tag):
        return ""

    clases = set(node.get("class", []))

    # Superíndice HTML -> Unicode
    if node.name == "span" and "g" in clases and "rend_superior" in clases:
        return a_superindice("".join(texto_diplomatico(c) for c in node.children))

    # Lagunas
    if "gap" in clases:
        reason = ""
        for c in clases:
            if c.startswith("reason_"):
                reason = c.replace("reason_", "")
        return " [GAP_{}] ".format(reason or "gap")

    if node.name == "br":
        return "\n"

    contenido = "".join(texto_diplomatico(c) for c in node.children)

    if node.name in {"p", "div"}:
        return contenido + "\n"

    return contenido


def leer_pagina_html(path: Path) -> Dict:
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "lxml")

    page_tag = soup.select_one("span.page")
    if page_tag is None:
        page_tag = soup.select_one("span.body")
    if page_tag is None:
        page_tag = soup.body or soup

    titulo_tag = soup.find("h1")
    titulo = titulo_tag.get_text(" ", strip=True) if titulo_tag else extraer_id_carta(path.name)

    carta_id = extraer_id_carta(path.name)
    pagina = extraer_num_pagina(path.name)
    texto = normalizar_espacios(texto_diplomatico(page_tag))

    return {
        "carta_id": carta_id,
        "pagina": pagina,
        "titulo": titulo,
        "archivo_html": str(path),
        "texto_original": texto
    }


# ============================================================
# PROMPT Y API
# ============================================================

SYSTEM_PROMPT = (
    "Eres un editor experto en español histórico hispanoamericano del siglo XVIII. "
    "Tu tarea es modernizar transcripciones diplomáticas sin alterar el contenido histórico. "
    "Debes ser conservador: moderniza la escritura, pero no inventes información."
)


def construir_prompt(texto_original: str, carta_id: str, pagina: int) -> str:
    # OJO: no hay variables sueltas fuera de esta función.
    # Por eso aquí evitamos cualquier bloque mal cerrado.
    reglas = """
Convierte la siguiente transcripción diplomática del siglo XVIII a español actual.

Reglas estrictas:
1. No resumas.
2. No agregues información nueva.
3. No elimines nombres propios, lugares, fechas ni cantidades.
4. Conserva el orden de las ideas.
5. Conserva la división en párrafos cuando exista.
6. Moderniza ortografía, acentuación y puntuación.
7. Expande abreviaturas evidentes solo cuando sea seguro, por ejemplo:
   - qᵉ -> que
   - Dⁿ -> Don
   - Sⁿ -> San
   - Franᶜᵒ -> Francisco
   - Nᵗᵒ -> Nuestro
   - testamᵗᵒ -> testamento
8. Si una palabra es dudosa, consérvala y no inventes.
9. Si aparece [GAP_illegible] o una laguna, consérvala como [ilegible].
10. Devuelve solo el texto modernizado, sin comentarios ni explicación.
""".strip()

    return (
        reglas
        + "\n\nIdentificador: " + str(carta_id) + ", página " + str(pagina)
        + "\n\nTRANSCRIPCIÓN ORIGINAL:\n<<<TEXTO_ORIGINAL>>>\n"
        + texto_original
        + "\n<<<FIN_TEXTO_ORIGINAL>>>"
    )


def modernizar_con_api(client: OpenAI, texto_original: str, carta_id: str, pagina: int) -> str:
    prompt = construir_prompt(texto_original, carta_id, pagina)

    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            response = client.responses.create(
                model=MODELO,
                input=[
                    {
                        "role": "system",
                        "content": SYSTEM_PROMPT
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ]
            )

            texto = getattr(response, "output_text", None)
            if texto is None:
                texto = str(response)

            return normalizar_espacios(texto)

        except Exception as e:
            print("  Error intento {}/{}: {}".format(intento, MAX_REINTENTOS, e))
            if intento == MAX_REINTENTOS:
                raise
            time.sleep(2 * intento)

    raise RuntimeError("No se pudo modernizar el texto")


# ============================================================
# COMPARACIÓN SIMPLE
# ============================================================

def tokenizar_basico(texto: str) -> List[str]:
    texto = texto.lower()
    if TIENE_REGEX:
        patron = r"[\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻºª]+"
        return regex_unicode.findall(patron, texto)
    return re.findall(r"[^\W\d_]+", texto, flags=re.UNICODE)


def guardar_comparacion(carta_id: str, pagina: int, original: str, moderno: str):
    nombre = limpiar_nombre_archivo("{}_page_{:03d}_comparacion.txt".format(carta_id, pagina))
    path = CARPETA_COMPARACION / nombre

    contenido = (
        "CARTA: {}\nPÁGINA: {}\n".format(carta_id, pagina)
        + "=" * 80 + "\n"
        + "ORIGINAL DIPLOMÁTICO\n"
        + "=" * 80 + "\n"
        + original
        + "\n\n"
        + "=" * 80 + "\n"
        + "VERSIÓN MODERNIZADA\n"
        + "=" * 80 + "\n"
        + moderno
    )

    path.write_text(contenido, encoding="utf-8")
    return path


# ============================================================
# MAIN
# ============================================================

def main():
    asegurar_carpetas()

    if "OPENAI_API_KEY" not in os.environ:
        raise RuntimeError(
            "No se encontró OPENAI_API_KEY. Configúrala antes de ejecutar.\n"
            "CMD: set OPENAI_API_KEY=tu_api_key\n"
            "PowerShell: $env:OPENAI_API_KEY=\"tu_api_key\""
        )

    if not CARPETA_ORIGINAL.exists():
        raise FileNotFoundError(
            "No existe {}. Primero ejecuta el script de descarga de cartas.".format(CARPETA_ORIGINAL)
        )

    archivos = sorted(CARPETA_ORIGINAL.glob("*_page_*.html"))
    if MAX_PAGINAS is not None:
        archivos = archivos[:MAX_PAGINAS]

    if not archivos:
        raise FileNotFoundError("No se encontraron HTML de páginas en {}".format(CARPETA_ORIGINAL))

    client = OpenAI()
    registros = []

    print("Páginas a modernizar: {}".format(len(archivos)))
    print("Modelo: {}".format(MODELO))

    for i, archivo in enumerate(archivos, start=1):
        data = leer_pagina_html(archivo)
        carta_id = data["carta_id"]
        pagina = data["pagina"]
        original = data["texto_original"]

        base = limpiar_nombre_archivo("{}_page_{:03d}".format(carta_id, pagina))
        salida_txt = CARPETA_TXT_MODERNO / (base + "_moderno.txt")

        print("[{}/{}] {} página {}".format(i, len(archivos), carta_id, pagina))

        if SALTAR_SI_EXISTE and salida_txt.exists():
            moderno = salida_txt.read_text(encoding="utf-8", errors="replace")
            estado = "ya_existia"
            print("  Saltado: ya existe")
        else:
            moderno = modernizar_con_api(client, original, carta_id, pagina)
            salida_txt.write_text(moderno, encoding="utf-8")
            estado = "modernizado"
            time.sleep(SLEEP_SECONDS)

        comparacion_path = guardar_comparacion(carta_id, pagina, original, moderno)

        tokens_original = tokenizar_basico(original)
        tokens_moderno = tokenizar_basico(moderno)

        registros.append({
            "carta_id": carta_id,
            "pagina": pagina,
            "titulo": data["titulo"],
            "archivo_html_original": data["archivo_html"],
            "archivo_txt_moderno": str(salida_txt),
            "archivo_comparacion": str(comparacion_path),
            "estado": estado,
            "tokens_original": len(tokens_original),
            "tokens_moderno": len(tokens_moderno),
            "palabras_unicas_original": len(set(tokens_original)),
            "palabras_unicas_moderno": len(set(tokens_moderno))
        })

        pd.DataFrame(registros).to_csv(ARCHIVO_INDICE, index=False, encoding="utf-8-sig")

    print("\nProceso terminado.")
    print("Salida: {}".format(CARPETA_SALIDA.resolve()))
    print("Índice: {}".format(ARCHIVO_INDICE.resolve()))


if __name__ == "__main__":
    main()
