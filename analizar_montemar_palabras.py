# -*- coding: utf-8 -*-
"""
ANÁLISIS DE PALABRAS - CARTAS DE MONTEMAR

Este script trabaja con la carpeta generada por el descargador:
    montemar_transcripciones/

Analiza:
    1. Frecuencia de palabras por carta.
    2. Frecuencia de palabras por página.
    3. Asociaciones / coocurrencias de palabras por carta.
    4. Asociaciones / coocurrencias de palabras por página.
    5. Nube de palabras por carta.
    6. Nube de palabras por página.
    7. Redes de asociaciones por carta y por página.

IMPORTANTE:
    - No moderniza la grafía.
    - No elimina caracteres especiales.
    - Convierte superíndices HTML a caracteres superíndices Unicode:
        q<span class="g rend_superior">e</span>  -> qᵉ
        D<span class="g rend_superior">n</span>  -> Dⁿ
        Fran<span class="g rend_superior">co</span> -> Franᶜᵒ
    - Incluye texto tachado y texto añadido, porque ambos son parte visible
      de la transcripción diplomática.
    - Por defecto NO elimina stopwords. Si quieres quitarlas, cambia:
        USAR_STOPWORDS = True

Requisitos:
    pip install beautifulsoup4 lxml pandas openpyxl matplotlib wordcloud networkx regex

Ejecución:
    python analizar_montemar_palabras.py
"""

import re
import math
import html
from pathlib import Path
from collections import Counter, defaultdict
from itertools import combinations

import pandas as pd
from bs4 import BeautifulSoup, NavigableString, Tag

import matplotlib.pyplot as plt
from matplotlib import font_manager

from wordcloud import WordCloud
import networkx as nx

try:
    import regex as regex_unicode
    TIENE_REGEX = True
except ImportError:
    regex_unicode = None
    TIENE_REGEX = False


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_BASE = Path("montemar_transcripciones")
CARPETA_HTML_PAGINAS = CARPETA_BASE / "html_por_pagina"

CARPETA_SALIDA = Path("analisis_montemar")
CARPETA_FRECUENCIAS = CARPETA_SALIDA / "frecuencias"
CARPETA_ASOCIACIONES = CARPETA_SALIDA / "asociaciones"
CARPETA_NUBES_CARTA = CARPETA_SALIDA / "nubes_palabras" / "por_carta"
CARPETA_NUBES_PAGINA = CARPETA_SALIDA / "nubes_palabras" / "por_pagina"
CARPETA_REDES_CARTA = CARPETA_SALIDA / "redes_asociaciones" / "por_carta"
CARPETA_REDES_PAGINA = CARPETA_SALIDA / "redes_asociaciones" / "por_pagina"

ARCHIVO_EXCEL = CARPETA_SALIDA / "analisis_palabras_montemar.xlsx"

# Frecuencias
MIN_FRECUENCIA_NUBE = 1
MAX_PALABRAS_NUBE = 150

# Asociaciones
VENTANA_ASOCIACION = 4
TOP_PALABRAS_ASOCIACION = 50
TOP_ARISTAS_RED = 35

# Cambia a True si quieres quitar palabras funcionales.
USAR_STOPWORDS = False

# Frecuencia: True agrupa "Lima" y "lima" como "lima".
# No cambia caracteres especiales ni grafía histórica.
PASAR_A_MINUSCULAS = True

# Si quieres limitar pruebas, por ejemplo 2 cartas:
# MAX_CARTAS = 2
MAX_CARTAS = None


STOPWORDS_BASE = {
    "a", "al", "ala", "alas", "alos", "ante", "asi", "con", "contra",
    "de", "del", "dela", "delas", "delos", "desde", "e", "el", "ella",
    "ellos", "en", "entre", "era", "es", "esta", "estas", "este", "esto",
    "estos", "ha", "hai", "hasta", "la", "las", "le", "lo", "los", "mas",
    "me", "mi", "mis", "no", "o", "para", "por", "porque", "q", "qᵉ",
    "que", "se", "si", "sin", "su", "sus", "te", "tu", "tus", "un", "una",
    "y", "yo",
    # Formas frecuentes históricas / abreviadas
    "d", "dn", "dⁿ", "s", "sⁿ", "n", "nᵗᵒ", "r", "p"
}


# ============================================================
# UTILIDADES
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


def limpiar_nombre_archivo(s: str, max_len: int = 110) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def asegurar_carpetas():
    for c in [
        CARPETA_SALIDA,
        CARPETA_FRECUENCIAS,
        CARPETA_ASOCIACIONES,
        CARPETA_NUBES_CARTA,
        CARPETA_NUBES_PAGINA,
        CARPETA_REDES_CARTA,
        CARPETA_REDES_PAGINA,
    ]:
        c.mkdir(parents=True, exist_ok=True)


def extraer_num_pagina(nombre_archivo: str) -> int:
    m = re.search(r"_page_(\d+)\.html$", nombre_archivo)
    return int(m.group(1)) if m else 0


def extraer_id_carta(nombre_archivo: str) -> str:
    return re.sub(r"_page_\d+\.html$", "", nombre_archivo)


def extraer_num_carta(carta_id: str):
    m = re.match(r"^(\d+)_", carta_id)
    return int(m.group(1)) if m else None


def fuente_wordcloud():
    """
    Usa una fuente del sistema para soportar tildes y superíndices.
    No copia ni comparte fuentes.
    """
    candidatos = ["DejaVu Sans", "Arial Unicode MS", "Noto Sans", "Liberation Sans"]
    for f in candidatos:
        try:
            ruta = font_manager.findfont(f, fallback_to_default=False)
            if ruta and Path(ruta).exists():
                return ruta
        except Exception:
            pass
    return None


# ============================================================
# EXTRACCIÓN DE TEXTO DESDE HTML
# ============================================================

def texto_diplomatico(node) -> str:
    """
    Extrae texto del HTML conservando caracteres especiales.
    Convierte etiquetas de superíndice a Unicode superíndice.
    """
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


def leer_texto_pagina_html(html_path: Path) -> dict:
    raw = html_path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "lxml")

    page_tag = soup.select_one("span.page")
    if page_tag is None:
        page_tag = soup.select_one("span.body")
    if page_tag is None:
        page_tag = soup.body or soup

    titulo_tag = soup.find("h1")
    titulo = titulo_tag.get_text(" ", strip=True) if titulo_tag else extraer_id_carta(html_path.name)

    texto = texto_diplomatico(page_tag)
    texto = html.unescape(texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r" *\n *", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto).strip()

    carta_id = extraer_id_carta(html_path.name)
    page_num = extraer_num_pagina(html_path.name)

    return {
        "carta_id": carta_id,
        "nro_carta": extraer_num_carta(carta_id),
        "titulo": titulo,
        "pagina": page_num,
        "archivo_html": str(html_path),
        "texto": texto
    }


# ============================================================
# TOKENIZACIÓN
# ============================================================

def tokenizar(texto: str) -> list:
    """
    Tokeniza conservando letras Unicode, tildes, ñ, superíndices y grafía histórica.
    No moderniza palabras.
    """
    if PASAR_A_MINUSCULAS:
        texto = texto.lower()

    # Quita marcas de corchetes solo si son de GAP, pero conserva GAP como token.
    texto = texto.replace("[GAP_", " GAP_").replace("]", " ")

    if TIENE_REGEX:
        # Letras unicode + marcas + superíndices + ºª.
        patron = r"[\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+(?:[-'’´`~][\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+)*"
        tokens = regex_unicode.findall(patron, texto)
    else:
        # Fallback sin paquete regex.
        patron = r"[^\W\d_ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻºª]+(?:[-'’´`~][^\W\d_]+)*"
        tokens = re.findall(patron, texto, flags=re.UNICODE)

    tokens = [t.strip("-'’´`~") for t in tokens if t.strip("-'’´`~")]

    if USAR_STOPWORDS:
        tokens = [t for t in tokens if t not in STOPWORDS_BASE]

    return tokens


# ============================================================
# FRECUENCIAS
# ============================================================

def frecuencias_tokens(tokens: list) -> pd.DataFrame:
    c = Counter(tokens)
    total = sum(c.values())

    filas = []
    for palabra, frecuencia in c.most_common():
        filas.append({
            "palabra": palabra,
            "frecuencia": frecuencia,
            "porcentaje": (frecuencia / total * 100) if total else 0
        })

    return pd.DataFrame(filas)


# ============================================================
# ASOCIACIONES / COOCURRENCIAS
# ============================================================

def asociaciones_tokens(tokens: list,
                        top_palabras: int = TOP_PALABRAS_ASOCIACION,
                        ventana: int = VENTANA_ASOCIACION) -> pd.DataFrame:
    """
    Calcula coocurrencias dentro de una ventana móvil.
    Ejemplo ventana=4:
        para cada bloque de 4 palabras, se cuentan pares que aparecen juntos.
    """
    if not tokens:
        return pd.DataFrame(columns=["palabra_1", "palabra_2", "coocurrencias"])

    freq = Counter(tokens)
    vocab = {w for w, _ in freq.most_common(top_palabras)}

    pares = Counter()

    for i in range(len(tokens)):
        ventana_tokens = [t for t in tokens[i:i + ventana] if t in vocab]
        ventana_unicos = sorted(set(ventana_tokens))

        for a, b in combinations(ventana_unicos, 2):
            pares[(a, b)] += 1

    filas = [
        {"palabra_1": a, "palabra_2": b, "coocurrencias": n}
        for (a, b), n in pares.most_common()
    ]

    return pd.DataFrame(filas)


# ============================================================
# GRÁFICOS: NUBES Y REDES
# ============================================================

def crear_nube(freq_df: pd.DataFrame, salida_png: Path, titulo: str):
    if freq_df.empty:
        return

    freqs = {
        str(row["palabra"]): int(row["frecuencia"])
        for _, row in freq_df.iterrows()
        if int(row["frecuencia"]) >= MIN_FRECUENCIA_NUBE
    }

    if not freqs:
        return

    font_path = fuente_wordcloud()

    wc = WordCloud(
        width=1800,
        height=1100,
        background_color="white",
        max_words=MAX_PALABRAS_NUBE,
        collocations=False,
        font_path=font_path
    ).generate_from_frequencies(freqs)

    plt.figure(figsize=(16, 10))
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")
    plt.title(titulo, fontsize=18)
    plt.tight_layout()
    plt.savefig(salida_png, dpi=160)
    plt.close()


def crear_red_asociaciones(asoc_df: pd.DataFrame, salida_png: Path, titulo: str):
    if asoc_df.empty:
        return

    df = asoc_df.head(TOP_ARISTAS_RED).copy()

    if df.empty:
        return

    G = nx.Graph()

    for _, row in df.iterrows():
        a = str(row["palabra_1"])
        b = str(row["palabra_2"])
        w = int(row["coocurrencias"])
        if a != b and w > 0:
            G.add_edge(a, b, weight=w)

    if G.number_of_edges() == 0:
        return

    grados = dict(G.degree(weight="weight"))
    pesos = [G[u][v]["weight"] for u, v in G.edges()]
    max_peso = max(pesos) if pesos else 1

    plt.figure(figsize=(14, 10))
    pos = nx.spring_layout(G, seed=123, k=0.7)

    node_sizes = [250 + grados.get(n, 1) * 25 for n in G.nodes()]
    edge_widths = [0.6 + (G[u][v]["weight"] / max_peso) * 4 for u, v in G.edges()]

    nx.draw_networkx_nodes(G, pos, node_size=node_sizes, alpha=0.85)
    nx.draw_networkx_edges(G, pos, width=edge_widths, alpha=0.35)
    nx.draw_networkx_labels(G, pos, font_size=10)

    plt.title(titulo, fontsize=16)
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(salida_png, dpi=160)
    plt.close()


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def main():
    asegurar_carpetas()

    if not CARPETA_HTML_PAGINAS.exists():
        raise FileNotFoundError(
            f"No existe la carpeta {CARPETA_HTML_PAGINAS}. "
            "Primero ejecuta el script de descarga de transcripciones."
        )

    archivos_html = sorted(CARPETA_HTML_PAGINAS.glob("*_page_*.html"))

    if not archivos_html:
        raise FileNotFoundError(
            f"No se encontraron archivos *_page_*.html en {CARPETA_HTML_PAGINAS}"
        )

    print(f"Archivos HTML de páginas encontrados: {len(archivos_html)}")

    paginas = [leer_texto_pagina_html(p) for p in archivos_html]

    df_paginas = pd.DataFrame([
        {
            "carta_id": p["carta_id"],
            "nro_carta": p["nro_carta"],
            "titulo": p["titulo"],
            "pagina": p["pagina"],
            "archivo_html": p["archivo_html"],
            "n_caracteres": len(p["texto"])
        }
        for p in paginas
    ])

    # Limitar cartas si se desea
    if MAX_CARTAS is not None:
        cartas_permitidas = set(
            df_paginas.sort_values(["nro_carta", "carta_id"])["carta_id"]
            .drop_duplicates()
            .head(MAX_CARTAS)
        )
        paginas = [p for p in paginas if p["carta_id"] in cartas_permitidas]
        df_paginas = df_paginas[df_paginas["carta_id"].isin(cartas_permitidas)].copy()

    # Agrupar texto por carta
    textos_cartas = {}
    metadatos_cartas = {}

    for p in paginas:
        cid = p["carta_id"]
        textos_cartas.setdefault(cid, [])
        textos_cartas[cid].append((p["pagina"], p["texto"]))

        metadatos_cartas[cid] = {
            "carta_id": cid,
            "nro_carta": p["nro_carta"],
            "titulo": p["titulo"]
        }

    # ========================================================
    # ANÁLISIS POR PÁGINA
    # ========================================================
    freq_paginas = []
    asoc_paginas = []
    resumen_paginas = []

    print("Analizando por página...")

    for p in paginas:
        cid = p["carta_id"]
        pagina = p["pagina"]
        base_salida = limpiar_nombre_archivo(f"{cid}_page_{pagina:03d}")

        tokens = tokenizar(p["texto"])
        freq_df = frecuencias_tokens(tokens)
        asoc_df = asociaciones_tokens(tokens)

        resumen_paginas.append({
            "carta_id": cid,
            "nro_carta": p["nro_carta"],
            "titulo": p["titulo"],
            "pagina": pagina,
            "total_tokens": len(tokens),
            "total_palabras_unicas": freq_df["palabra"].nunique() if not freq_df.empty else 0
        })

        if not freq_df.empty:
            freq_df.insert(0, "pagina", pagina)
            freq_df.insert(0, "titulo", p["titulo"])
            freq_df.insert(0, "nro_carta", p["nro_carta"])
            freq_df.insert(0, "carta_id", cid)
            freq_paginas.append(freq_df)

            crear_nube(
                freq_df[["palabra", "frecuencia"]],
                CARPETA_NUBES_PAGINA / f"{base_salida}_nube.png",
                f"Nube de palabras - {cid} - página {pagina}"
            )

        if not asoc_df.empty:
            asoc_df.insert(0, "pagina", pagina)
            asoc_df.insert(0, "titulo", p["titulo"])
            asoc_df.insert(0, "nro_carta", p["nro_carta"])
            asoc_df.insert(0, "carta_id", cid)
            asoc_paginas.append(asoc_df)

            crear_red_asociaciones(
                asoc_df[["palabra_1", "palabra_2", "coocurrencias"]],
                CARPETA_REDES_PAGINA / f"{base_salida}_red.png",
                f"Asociaciones - {cid} - página {pagina}"
            )

    df_freq_paginas = pd.concat(freq_paginas, ignore_index=True) if freq_paginas else pd.DataFrame()
    df_asoc_paginas = pd.concat(asoc_paginas, ignore_index=True) if asoc_paginas else pd.DataFrame()
    df_resumen_paginas = pd.DataFrame(resumen_paginas)

    # ========================================================
    # ANÁLISIS POR CARTA
    # ========================================================
    freq_cartas = []
    asoc_cartas = []
    resumen_cartas = []

    print("Analizando por carta...")

    for cid, partes in textos_cartas.items():
        partes_ordenadas = sorted(partes, key=lambda x: x[0])
        texto_carta = "\n".join(t for _, t in partes_ordenadas)
        meta = metadatos_cartas[cid]

        tokens = tokenizar(texto_carta)
        freq_df = frecuencias_tokens(tokens)
        asoc_df = asociaciones_tokens(tokens)

        resumen_cartas.append({
            "carta_id": cid,
            "nro_carta": meta["nro_carta"],
            "titulo": meta["titulo"],
            "total_paginas": len(partes_ordenadas),
            "total_tokens": len(tokens),
            "total_palabras_unicas": freq_df["palabra"].nunique() if not freq_df.empty else 0
        })

        base_salida = limpiar_nombre_archivo(cid)

        if not freq_df.empty:
            freq_df.insert(0, "titulo", meta["titulo"])
            freq_df.insert(0, "nro_carta", meta["nro_carta"])
            freq_df.insert(0, "carta_id", cid)
            freq_cartas.append(freq_df)

            crear_nube(
                freq_df[["palabra", "frecuencia"]],
                CARPETA_NUBES_CARTA / f"{base_salida}_nube.png",
                f"Nube de palabras - {cid}"
            )

        if not asoc_df.empty:
            asoc_df.insert(0, "titulo", meta["titulo"])
            asoc_df.insert(0, "nro_carta", meta["nro_carta"])
            asoc_df.insert(0, "carta_id", cid)
            asoc_cartas.append(asoc_df)

            crear_red_asociaciones(
                asoc_df[["palabra_1", "palabra_2", "coocurrencias"]],
                CARPETA_REDES_CARTA / f"{base_salida}_red.png",
                f"Asociaciones - {cid}"
            )

    df_freq_cartas = pd.concat(freq_cartas, ignore_index=True) if freq_cartas else pd.DataFrame()
    df_asoc_cartas = pd.concat(asoc_cartas, ignore_index=True) if asoc_cartas else pd.DataFrame()
    df_resumen_cartas = pd.DataFrame(resumen_cartas)

    # ========================================================
    # GUARDAR CSV
    # ========================================================
    print("Guardando CSV...")

    df_resumen_cartas.to_csv(CARPETA_SALIDA / "resumen_por_carta.csv", index=False, encoding="utf-8-sig")
    df_resumen_paginas.to_csv(CARPETA_SALIDA / "resumen_por_pagina.csv", index=False, encoding="utf-8-sig")

    df_freq_cartas.to_csv(CARPETA_FRECUENCIAS / "frecuencia_por_carta.csv", index=False, encoding="utf-8-sig")
    df_freq_paginas.to_csv(CARPETA_FRECUENCIAS / "frecuencia_por_pagina.csv", index=False, encoding="utf-8-sig")

    df_asoc_cartas.to_csv(CARPETA_ASOCIACIONES / "asociaciones_por_carta.csv", index=False, encoding="utf-8-sig")
    df_asoc_paginas.to_csv(CARPETA_ASOCIACIONES / "asociaciones_por_pagina.csv", index=False, encoding="utf-8-sig")

    # Versiones top para Excel
    top_freq_carta = (
        df_freq_cartas.groupby("carta_id", group_keys=False)
        .head(100)
        if not df_freq_cartas.empty else df_freq_cartas
    )

    top_freq_pagina = (
        df_freq_paginas.groupby(["carta_id", "pagina"], group_keys=False)
        .head(100)
        if not df_freq_paginas.empty else df_freq_paginas
    )

    top_asoc_carta = (
        df_asoc_cartas.groupby("carta_id", group_keys=False)
        .head(100)
        if not df_asoc_cartas.empty else df_asoc_cartas
    )

    top_asoc_pagina = (
        df_asoc_paginas.groupby(["carta_id", "pagina"], group_keys=False)
        .head(100)
        if not df_asoc_paginas.empty else df_asoc_paginas
    )

    # ========================================================
    # GUARDAR EXCEL
    # ========================================================
    print("Guardando Excel...")

    with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl") as writer:
        df_resumen_cartas.to_excel(writer, sheet_name="resumen_cartas", index=False)
        df_resumen_paginas.to_excel(writer, sheet_name="resumen_paginas", index=False)
        top_freq_carta.to_excel(writer, sheet_name="freq_carta_top100", index=False)
        top_freq_pagina.to_excel(writer, sheet_name="freq_pagina_top100", index=False)
        top_asoc_carta.to_excel(writer, sheet_name="asoc_carta_top100", index=False)
        top_asoc_pagina.to_excel(writer, sheet_name="asoc_pagina_top100", index=False)

    print("\nProceso terminado.")
    print(f"Salida principal: {CARPETA_SALIDA.resolve()}")
    print(f"Excel: {ARCHIVO_EXCEL.resolve()}")
    print(f"Nubes por carta: {CARPETA_NUBES_CARTA.resolve()}")
    print(f"Nubes por página: {CARPETA_NUBES_PAGINA.resolve()}")
    print(f"Redes por carta: {CARPETA_REDES_CARTA.resolve()}")
    print(f"Redes por página: {CARPETA_REDES_PAGINA.resolve()}")


if __name__ == "__main__":
    main()
