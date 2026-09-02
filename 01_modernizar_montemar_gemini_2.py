# -*- coding: utf-8 -*-
"""
MODERNIZAR CARTAS DE MONTEMAR CON GEMINI API
VERSIÓN CORREGIDA PARA LOTES Y PÁGINAS PENDIENTES

Entrada:
    montemar_transcripciones/html_por_pagina/

Salida:
    montemar_moderno_gemini/
    ├── txt_moderno_por_pagina/
    ├── html_moderno_por_pagina/
    ├── comparacion_original_moderno/
    ├── indice_modernizacion_gemini.csv
    └── indice_modernizacion_gemini.xlsx

Instalar:
    pip install google-genai beautifulsoup4 lxml pandas openpyxl

Configurar clave:

CMD:
    set GEMINI_API_KEY=TU_API_KEY

PowerShell:
    $env:GEMINI_API_KEY="TU_API_KEY"

Ejecutar:
    python 01_modernizar_montemar_gemini_CORREGIDO_LOTES.py
"""

import os
import re
import time
import html
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup, NavigableString, Tag

from google import genai
from google.genai import types


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_ORIGINAL = Path("montemar_transcripciones")
CARPETA_HTML_ORIGINAL = CARPETA_ORIGINAL / "html_por_pagina"

CARPETA_SALIDA = Path("montemar_moderno_gemini")
CARPETA_TXT_MODERNO = CARPETA_SALIDA / "txt_moderno_por_pagina"
CARPETA_HTML_MODERNO = CARPETA_SALIDA / "html_moderno_por_pagina"
CARPETA_COMPARACION = CARPETA_SALIDA / "comparacion_original_moderno"

INDICE_SALIDA = CARPETA_SALIDA / "indice_modernizacion_gemini.csv"
EXCEL_SALIDA = CARPETA_SALIDA / "indice_modernizacion_gemini.xlsx"

# Puedes probar también: "gemini-2.5-flash-lite"
MODELO = "gemini-2.5-flash-lite"

# Para tu cuota gratuita, usa 20 o menos.
# El script tomará 20 páginas PENDIENTES, no las primeras 20.
MAX_PAGINAS = 20

SALTAR_SI_EXISTE = True

# Si te sale 429 con frecuencia, sube a 20 o 30.
ESPERA_SEGUNDOS = 12

MAX_INTENTOS = 5
TEMPERATURE = 0.2
MAX_CARACTERES_POR_BLOQUE = 12000

# Si aparece cuota agotada, se detiene para no insistir.
DETENER_SI_CUOTA_AGOTADA = True


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

def asegurar_carpetas():
    for carpeta in [
        CARPETA_SALIDA,
        CARPETA_TXT_MODERNO,
        CARPETA_HTML_MODERNO,
        CARPETA_COMPARACION,
    ]:
        carpeta.mkdir(parents=True, exist_ok=True)


def limpiar_nombre_archivo(s: str, max_len: int = 180) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def extraer_num_pagina(nombre_archivo: str) -> int:
    m = re.search(r"_page_(\d+)\.html$", nombre_archivo)
    return int(m.group(1)) if m else 0


def extraer_id_carta(nombre_archivo: str) -> str:
    return re.sub(r"_page_\d+\.html$", "", nombre_archivo)


def extraer_num_carta(carta_id: str):
    m = re.match(r"^(\d+)_", carta_id)
    return int(m.group(1)) if m else None


def normalizar_espacios(txt: str) -> str:
    txt = html.unescape(str(txt))
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r" *\n *", "\n", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()


def nombre_base_salida(path_html: Path) -> str:
    carta_id = extraer_id_carta(path_html.name)
    pagina = extraer_num_pagina(path_html.name)
    return limpiar_nombre_archivo(f"{carta_id}_page_{pagina:03d}")


def ruta_txt_moderno_para(path_html: Path) -> Path:
    return CARPETA_TXT_MODERNO / f"{nombre_base_salida(path_html)}_moderno.txt"


# ============================================================
# EXTRAER TEXTO ORIGINAL
# ============================================================

def texto_diplomatico(node) -> str:
    if isinstance(node, NavigableString):
        return str(node)

    if not isinstance(node, Tag):
        return ""

    clases = set(node.get("class", []))

    if node.name == "span" and "g" in clases and "rend_superior" in clases:
        return a_superindice("".join(texto_diplomatico(c) for c in node.children))

    if "gap" in clases:
        reason = ""
        for c in clases:
            if c.startswith("reason_"):
                reason = c.replace("reason_", "")
        return f" [GAP_{reason or 'gap'}] "

    if node.name == "br":
        return "\n"

    contenido = "".join(texto_diplomatico(c) for c in node.children)

    if node.name in {"p", "div"}:
        return contenido + "\n"

    return contenido


def leer_pagina_original(path_html: Path) -> dict:
    raw = path_html.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "lxml")

    page_tag = soup.select_one("span.page")
    if page_tag is None:
        page_tag = soup.select_one("span.body")
    if page_tag is None:
        page_tag = soup.body or soup

    titulo_tag = soup.find("h1")
    titulo = titulo_tag.get_text(" ", strip=True) if titulo_tag else extraer_id_carta(path_html.name)

    carta_id = extraer_id_carta(path_html.name)
    pagina = extraer_num_pagina(path_html.name)

    texto = normalizar_espacios(texto_diplomatico(page_tag))

    return {
        "carta_id": carta_id,
        "nro_carta": extraer_num_carta(carta_id),
        "titulo": titulo,
        "pagina": pagina,
        "archivo_html": str(path_html),
        "texto_original": texto
    }


# ============================================================
# PROMPT
# ============================================================

def construir_prompt(texto_original: str, carta_id: str, pagina: int) -> str:
    return f"""
Convierte la siguiente transcripción diplomática del siglo XVIII a español actual.

Contexto:
- Es una carta histórica del corpus del Conde de Montemar.
- La transcripción conserva grafías antiguas, abreviaturas, superíndices, tachados y añadidos.
- Tu tarea es crear una versión modernizada en español actual.

Reglas obligatorias:
1. No resumas.
2. No agregues información nueva.
3. No elimines nombres propios, lugares, fechas, cantidades ni referencias familiares.
4. Conserva el orden de las ideas.
5. Conserva la división en párrafos cuando sea posible.
6. Moderniza ortografía, acentuación y puntuación.
7. Expande abreviaturas evidentes cuando sea seguro:
   - qᵉ -> que
   - Dⁿ -> Don
   - Sⁿ -> San
   - Nᵗᵒ -> Nuestro
   - Sᵒʳ -> Señor
   - Hᵒ -> Hermano
   - Franᶜᵒ -> Francisco
8. Moderniza grafías históricas cuando el equivalente sea claro:
   - rresibo -> recibo
   - felisidad -> felicidad
   - pribilejio -> privilegio
   - yndios -> indios
   - thestamento -> testamento
   - sedula -> cédula
9. Si una palabra es dudosa, consérvala sin inventar.
10. Si aparece [GAP_illegible] o una laguna, consérvala como [ilegible].
11. Devuelve únicamente el texto modernizado. No incluyas explicación, títulos ni comentarios.

Identificador de carta: {carta_id}
Página: {pagina}

TRANSCRIPCIÓN ORIGINAL:
<<<TEXTO_ORIGINAL>>>
{texto_original}
<<<FIN_TEXTO_ORIGINAL>>>
""".strip()


def dividir_texto(texto: str, max_chars: int = MAX_CARACTERES_POR_BLOQUE) -> list:
    if len(texto) <= max_chars:
        return [texto]

    parrafos = re.split(r"\n\s*\n", texto)
    bloques = []
    actual = ""

    for p in parrafos:
        p = p.strip()
        if not p:
            continue

        candidato = (actual + "\n\n" + p).strip() if actual else p

        if len(candidato) <= max_chars:
            actual = candidato
        else:
            if actual:
                bloques.append(actual)
            if len(p) <= max_chars:
                actual = p
            else:
                partes = [p[i:i + max_chars] for i in range(0, len(p), max_chars)]
                bloques.extend(partes[:-1])
                actual = partes[-1]

    if actual:
        bloques.append(actual)

    return bloques


# ============================================================
# GEMINI
# ============================================================

def crear_cliente_gemini():
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")

    if not api_key:
        raise RuntimeError(
            "No se encontró GEMINI_API_KEY ni GOOGLE_API_KEY.\n"
            "CMD: set GEMINI_API_KEY=TU_API_KEY\n"
            "PowerShell: $env:GEMINI_API_KEY=\"TU_API_KEY\""
        )

    return genai.Client(api_key=api_key)


def extraer_texto_respuesta(response) -> str:
    if hasattr(response, "text") and response.text:
        return response.text.strip()

    try:
        partes = []
        for cand in response.candidates:
            if cand.content and cand.content.parts:
                for part in cand.content.parts:
                    if getattr(part, "text", None):
                        partes.append(part.text)
        return "\n".join(partes).strip()
    except Exception:
        return ""


def obtener_retry_delay(msg: str) -> int:
    m = re.search(r"retryDelay['\"]?\s*:\s*['\"]?(\d+)s", msg)
    if m:
        return int(m.group(1))

    m = re.search(r"retry in\s+([0-9.]+)s", msg, flags=re.IGNORECASE)
    if m:
        return int(float(m.group(1))) + 5

    return 0


def es_error_cuota(msg: str) -> bool:
    msg_low = msg.lower()
    return (
        "429" in msg_low or
        "resource_exhausted" in msg_low or
        "quota exceeded" in msg_low or
        "free_tier_requests" in msg_low or
        "generaterequestsperday" in msg_low
    )


def modernizar_bloque(client, texto_bloque: str, carta_id: str, pagina: int,
                      num_bloque: int, total_bloques: int) -> str:
    prompt = construir_prompt(texto_bloque, carta_id, pagina)

    if total_bloques > 1:
        prompt += (
            f"\n\nNOTA: Esta es la parte {num_bloque} de {total_bloques} "
            "de la misma página. Moderniza solo esta parte."
        )

    ultimo_error = None

    for intento in range(1, MAX_INTENTOS + 1):
        try:
            response = client.models.generate_content(
                model=MODELO,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=TEMPERATURE,
                    candidate_count=1
                )
            )

            texto = extraer_texto_respuesta(response)

            if not texto:
                raise RuntimeError("Gemini devolvió una respuesta vacía.")

            return texto.strip()

        except Exception as e:
            ultimo_error = e
            msg = str(e)
            print(f"  Error intento {intento}/{MAX_INTENTOS}: {msg}")

            if DETENER_SI_CUOTA_AGOTADA and es_error_cuota(msg):
                retry = obtener_retry_delay(msg)
                print("\n  Se alcanzó una cuota/límite de Gemini.")
                if retry:
                    print(f"  La API sugiere reintentar en aproximadamente {retry} segundos.")
                print("  El script se detendrá para no consumir intentos inútiles.")
                print("  Vuelve a ejecutar más tarde o mañana; continuará con páginas pendientes.\n")
                raise

            espera = ESPERA_SEGUNDOS * intento
            retry = obtener_retry_delay(msg)
            if retry:
                espera = max(espera, retry)

            if intento < MAX_INTENTOS:
                print(f"  Esperando {espera} segundos antes de reintentar...")
                time.sleep(espera)

    raise RuntimeError(f"No se pudo modernizar después de {MAX_INTENTOS} intentos: {ultimo_error}")


def modernizar_con_gemini(client, texto_original: str, carta_id: str, pagina: int) -> str:
    bloques = dividir_texto(texto_original)

    if len(bloques) == 1:
        return modernizar_bloque(client, bloques[0], carta_id, pagina, 1, 1)

    salida = []
    for i, bloque in enumerate(bloques, start=1):
        print(f"  Bloque {i}/{len(bloques)}")
        moderno = modernizar_bloque(client, bloque, carta_id, pagina, i, len(bloques))
        salida.append(moderno)
        time.sleep(ESPERA_SEGUNDOS)

    return "\n\n".join(salida).strip()


# ============================================================
# SALIDAS
# ============================================================

def construir_html_moderno(titulo: str, carta_id: str, pagina: int, texto_moderno: str) -> str:
    parrafos = [p.strip() for p in re.split(r"\n\s*\n", texto_moderno) if p.strip()]
    html_parrafos = "\n".join(f"<p>{html.escape(p)}</p>" for p in parrafos)

    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>{html.escape(titulo)} - Página {pagina} - Modernizado</title>
<style>
body {{
    font-family: Georgia, "Times New Roman", serif;
    font-size: 20px;
    line-height: 1.65;
    margin: 32px;
    max-width: 900px;
}}
.header {{
    font-family: Arial, sans-serif;
    font-size: 14px;
    border-bottom: 1px solid #ddd;
    margin-bottom: 24px;
    padding-bottom: 12px;
}}
p {{
    margin-bottom: 1em;
}}
</style>
</head>
<body>
<div class="header">
<h1>{html.escape(titulo)}</h1>
<div><strong>Carta:</strong> {html.escape(carta_id)}</div>
<div><strong>Página:</strong> {pagina}</div>
<div><strong>Versión:</strong> Español actual modernizado con Gemini</div>
</div>
{html_parrafos}
</body>
</html>
"""


def guardar_comparacion(carta_id: str, pagina: int, titulo: str,
                        original: str, moderno: str, salida_path: Path):
    html_original = "<br>".join(html.escape(original).splitlines())
    html_moderno = "<br>".join(html.escape(moderno).splitlines())

    doc = f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Comparación - {html.escape(carta_id)} - Página {pagina}</title>
<style>
body {{
    font-family: Arial, sans-serif;
    margin: 24px;
}}
.container {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 24px;
}}
.panel {{
    border: 1px solid #ddd;
    padding: 16px;
    border-radius: 8px;
}}
.texto {{
    font-family: Georgia, "Times New Roman", serif;
    font-size: 18px;
    line-height: 1.55;
}}
</style>
</head>
<body>
<h1>Comparación original vs modernizado</h1>
<p><strong>{html.escape(titulo)}</strong></p>
<p><strong>Carta:</strong> {html.escape(carta_id)} | <strong>Página:</strong> {pagina}</p>
<div class="container">
<div class="panel">
<h2>Original diplomático</h2>
<div class="texto">{html_original}</div>
</div>
<div class="panel">
<h2>Español actual</h2>
<div class="texto">{html_moderno}</div>
</div>
</div>
</body>
</html>
"""
    salida_path.write_text(doc, encoding="utf-8")


def cargar_indice_existente() -> list:
    if INDICE_SALIDA.exists():
        try:
            df = pd.read_csv(INDICE_SALIDA, encoding="utf-8-sig")
            return df.to_dict("records")
        except Exception:
            return []
    return []


def guardar_indice(registros: list):
    if not registros:
        return

    df = pd.DataFrame(registros)

    if {"carta_id", "pagina"}.issubset(df.columns):
        df = df.drop_duplicates(subset=["carta_id", "pagina"], keep="last")

    df = df.sort_values(["nro_carta", "carta_id", "pagina"], na_position="last")
    df.to_csv(INDICE_SALIDA, index=False, encoding="utf-8-sig")
    df.to_excel(EXCEL_SALIDA, index=False)


# ============================================================
# SELECCIONAR PENDIENTES
# ============================================================

def obtener_archivos_pendientes() -> tuple[list, int, int]:
    archivos_todos = sorted(CARPETA_HTML_ORIGINAL.glob("*_page_*.html"))

    if not archivos_todos:
        raise FileNotFoundError(
            f"No se encontraron archivos *_page_*.html en {CARPETA_HTML_ORIGINAL}"
        )

    pendientes = []
    ya_modernizados = 0

    for path_html in archivos_todos:
        txt_moderno = ruta_txt_moderno_para(path_html)

        if SALTAR_SI_EXISTE and txt_moderno.exists() and txt_moderno.stat().st_size > 0:
            ya_modernizados += 1
            continue

        pendientes.append(path_html)

    if MAX_PAGINAS is not None:
        pendientes = pendientes[:MAX_PAGINAS]

    return pendientes, len(archivos_todos), ya_modernizados


# ============================================================
# MAIN
# ============================================================

def main():
    asegurar_carpetas()

    if not CARPETA_HTML_ORIGINAL.exists():
        raise FileNotFoundError(
            f"No existe {CARPETA_HTML_ORIGINAL}. "
            "Primero ejecuta el script de descarga de Montemar."
        )

    archivos, total_html, ya_modernizados = obtener_archivos_pendientes()

    print(f"Total de páginas originales: {total_html}")
    print(f"Páginas ya modernizadas: {ya_modernizados}")
    print(f"Páginas pendientes seleccionadas para este lote: {len(archivos)}")
    print(f"MAX_PAGINAS: {MAX_PAGINAS}")
    print(f"Modelo Gemini: {MODELO}")
    print(f"Salida: {CARPETA_SALIDA.resolve()}")

    if not archivos:
        print("\nNo hay páginas pendientes para modernizar.")
        return

    client = crear_cliente_gemini()
    registros = cargar_indice_existente()

    for idx, path_html in enumerate(archivos, start=1):
        pagina_data = leer_pagina_original(path_html)

        carta_id = pagina_data["carta_id"]
        pagina = pagina_data["pagina"]
        titulo = pagina_data["titulo"]
        original = pagina_data["texto_original"]

        base_name = nombre_base_salida(path_html)

        txt_moderno_path = CARPETA_TXT_MODERNO / f"{base_name}_moderno.txt"
        html_moderno_path = CARPETA_HTML_MODERNO / f"{base_name}_moderno.html"
        html_comparacion_path = CARPETA_COMPARACION / f"{base_name}_comparacion.html"

        print(f"\n[{idx}/{len(archivos)}] {carta_id} página {pagina}")

        try:
            if SALTAR_SI_EXISTE and txt_moderno_path.exists() and txt_moderno_path.stat().st_size > 0:
                print("  Ya existe. Saltando.")
                moderno = txt_moderno_path.read_text(encoding="utf-8", errors="replace")
                estado = "saltado_existente"
            else:
                moderno = modernizar_con_gemini(client, original, carta_id, pagina)
                txt_moderno_path.write_text(moderno, encoding="utf-8")
                estado = "modernizado"
                time.sleep(ESPERA_SEGUNDOS)

            html_moderno = construir_html_moderno(titulo, carta_id, pagina, moderno)
            html_moderno_path.write_text(html_moderno, encoding="utf-8")

            guardar_comparacion(
                carta_id=carta_id,
                pagina=pagina,
                titulo=titulo,
                original=original,
                moderno=moderno,
                salida_path=html_comparacion_path
            )

            registros.append({
                "carta_id": carta_id,
                "nro_carta": pagina_data["nro_carta"],
                "titulo": titulo,
                "pagina": pagina,
                "archivo_original_html": pagina_data["archivo_html"],
                "archivo_moderno_txt": str(txt_moderno_path),
                "archivo_moderno_html": str(html_moderno_path),
                "archivo_comparacion_html": str(html_comparacion_path),
                "n_caracteres_original": len(original),
                "n_caracteres_moderno": len(moderno),
                "estado": estado
            })

            guardar_indice(registros)

        except Exception as e:
            print(f"\nERROR en {carta_id} página {pagina}: {e}")
            print("Se guarda el índice con lo procesado hasta ahora y se detiene.")
            guardar_indice(registros)
            break

    print("\nProceso terminado o detenido.")
    print(f"TXT moderno: {CARPETA_TXT_MODERNO.resolve()}")
    print(f"HTML moderno: {CARPETA_HTML_MODERNO.resolve()}")
    print(f"Comparaciones: {CARPETA_COMPARACION.resolve()}")
    print(f"Índice CSV: {INDICE_SALIDA.resolve()}")
    print(f"Índice Excel: {EXCEL_SALIDA.resolve()}")


if __name__ == "__main__":
    main()
