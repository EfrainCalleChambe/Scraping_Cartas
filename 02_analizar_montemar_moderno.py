# -*- coding: utf-8 -*-
"""
02_ANALIZAR_MONTEMAR_MODERNO.PY

Repite el análisis de frecuencia, asociaciones, nubes y redes usando la versión
modernizada generada por 01_modernizar_montemar_api.py.

ENTRADA:
    montemar_moderno/txt_moderno_por_pagina/*_moderno.txt

SALIDA:
    analisis_montemar_moderno/
    ├── analisis_palabras_montemar_moderno.xlsx
    ├── resumen_por_carta.csv
    ├── resumen_por_pagina.csv
    ├── frecuencias/
    ├── asociaciones/
    ├── nubes_palabras/
    └── redes_asociaciones/

REQUISITOS:
    pip install pandas openpyxl matplotlib wordcloud networkx regex

EJECUCIÓN:
    python 02_analizar_montemar_moderno.py
"""

import re
import html
from pathlib import Path
from collections import Counter
from itertools import combinations

import pandas as pd
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

CARPETA_MODERNO = Path("montemar_moderno")
CARPETA_TXT_MODERNO = CARPETA_MODERNO / "txt_moderno_por_pagina"

CARPETA_SALIDA = Path("analisis_montemar_moderno")
CARPETA_FRECUENCIAS = CARPETA_SALIDA / "frecuencias"
CARPETA_ASOCIACIONES = CARPETA_SALIDA / "asociaciones"
CARPETA_NUBES_CARTA = CARPETA_SALIDA / "nubes_palabras" / "por_carta"
CARPETA_NUBES_PAGINA = CARPETA_SALIDA / "nubes_palabras" / "por_pagina"
CARPETA_REDES_CARTA = CARPETA_SALIDA / "redes_asociaciones" / "por_carta"
CARPETA_REDES_PAGINA = CARPETA_SALIDA / "redes_asociaciones" / "por_pagina"

ARCHIVO_EXCEL = CARPETA_SALIDA / "analisis_palabras_montemar_moderno.xlsx"

PASAR_A_MINUSCULAS = True
USAR_STOPWORDS = True
MAX_CARTAS = None

MIN_FRECUENCIA_NUBE = 1
MAX_PALABRAS_NUBE = 150
VENTANA_ASOCIACION = 4
TOP_PALABRAS_ASOCIACION = 50
TOP_ARISTAS_RED = 35

STOPWORDS_BASE = {
    "a", "al", "algo", "algunas", "algunos", "ante", "antes", "así", "aunque",
    "como", "con", "contra", "cuando", "de", "del", "desde", "donde", "e", "el",
    "él", "ella", "ellas", "ellos", "en", "entre", "era", "es", "esa", "esas",
    "ese", "eso", "esos", "esta", "está", "estaba", "estaban", "estado", "estas",
    "este", "esto", "estos", "ha", "han", "hasta", "hay", "la", "las", "le", "les",
    "lo", "los", "más", "me", "mi", "mis", "muy", "no", "nos", "o", "para", "pero",
    "por", "porque", "que", "se", "si", "sí", "sin", "sobre", "su", "sus", "te", "tu",
    "tus", "un", "una", "uno", "unos", "y", "ya", "yo"
}


# ============================================================
# UTILIDADES
# ============================================================


def asegurar_carpetas():
    for c in [
        CARPETA_SALIDA, CARPETA_FRECUENCIAS, CARPETA_ASOCIACIONES,
        CARPETA_NUBES_CARTA, CARPETA_NUBES_PAGINA, CARPETA_REDES_CARTA,
        CARPETA_REDES_PAGINA
    ]:
        c.mkdir(parents=True, exist_ok=True)


def limpiar_nombre_archivo(s: str, max_len: int = 120) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def extraer_meta_desde_nombre(nombre: str) -> dict:
    # nombre esperado: <carta_id>_page_001_moderno.txt
    base = nombre.replace("_moderno.txt", "")
    m_page = re.search(r"_page_(\d+)$", base)
    pagina = int(m_page.group(1)) if m_page else 0
    carta_id = re.sub(r"_page_\d+$", "", base)
    m_carta = re.match(r"^(\d+)_", carta_id)
    nro_carta = int(m_carta.group(1)) if m_carta else None
    return {"carta_id": carta_id, "nro_carta": nro_carta, "pagina": pagina}


def fuente_wordcloud():
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
# TOKENIZACIÓN
# ============================================================


def tokenizar(texto: str) -> list:
    if PASAR_A_MINUSCULAS:
        texto = texto.lower()

    if TIENE_REGEX:
        patron = r"[\p{L}\p{M}]+(?:[-'’´`~][\p{L}\p{M}]+)*"
        tokens = regex_unicode.findall(patron, texto)
    else:
        patron = r"[^\W\d_]+(?:[-'’´`~][^\W\d_]+)*"
        tokens = re.findall(patron, texto, flags=re.UNICODE)

    tokens = [t.strip("-'’´`~") for t in tokens if t.strip("-'’´`~")]

    if USAR_STOPWORDS:
        tokens = [t for t in tokens if t not in STOPWORDS_BASE]

    return tokens


# ============================================================
# FRECUENCIAS Y ASOCIACIONES
# ============================================================


def frecuencias_tokens(tokens: list) -> pd.DataFrame:
    c = Counter(tokens)
    total = sum(c.values())
    return pd.DataFrame([
        {"palabra": palabra, "frecuencia": frecuencia,
         "porcentaje": (frecuencia / total * 100) if total else 0}
        for palabra, frecuencia in c.most_common()
    ])


def asociaciones_tokens(tokens: list,
                        top_palabras: int = TOP_PALABRAS_ASOCIACION,
                        ventana: int = VENTANA_ASOCIACION) -> pd.DataFrame:
    if not tokens:
        return pd.DataFrame(columns=["palabra_1", "palabra_2", "coocurrencias"])

    freq = Counter(tokens)
    vocab = {w for w, _ in freq.most_common(top_palabras)}
    pares = Counter()

    for i in range(len(tokens)):
        ventana_tokens = [t for t in tokens[i:i + ventana] if t in vocab]
        for a, b in combinations(sorted(set(ventana_tokens)), 2):
            pares[(a, b)] += 1

    return pd.DataFrame([
        {"palabra_1": a, "palabra_2": b, "coocurrencias": n}
        for (a, b), n in pares.most_common()
    ])


# ============================================================
# GRÁFICOS
# ============================================================


def crear_nube(freq_df: pd.DataFrame, salida_png: Path, titulo: str):
    if freq_df.empty:
        return
    freqs = {str(r["palabra"]): int(r["frecuencia"]) for _, r in freq_df.iterrows()
             if int(r["frecuencia"]) >= MIN_FRECUENCIA_NUBE}
    if not freqs:
        return

    wc = WordCloud(
        width=1800,
        height=1100,
        background_color="white",
        max_words=MAX_PALABRAS_NUBE,
        collocations=False,
        font_path=fuente_wordcloud()
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
    G = nx.Graph()
    for _, row in df.iterrows():
        a, b, w = str(row["palabra_1"]), str(row["palabra_2"]), int(row["coocurrencias"])
        if a != b and w > 0:
            G.add_edge(a, b, weight=w)

    if G.number_of_edges() == 0:
        return

    grados = dict(G.degree(weight="weight"))
    pesos = [G[u][v]["weight"] for u, v in G.edges()]
    max_peso = max(pesos) if pesos else 1

    plt.figure(figsize=(14, 10))
    pos = nx.spring_layout(G, seed=123, k=0.7)
    nx.draw_networkx_nodes(G, pos, node_size=[250 + grados.get(n, 1) * 25 for n in G.nodes()], alpha=0.85)
    nx.draw_networkx_edges(G, pos, width=[0.6 + (G[u][v]["weight"] / max_peso) * 4 for u, v in G.edges()], alpha=0.35)
    nx.draw_networkx_labels(G, pos, font_size=10)
    plt.title(titulo, fontsize=16)
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(salida_png, dpi=160)
    plt.close()


# ============================================================
# MAIN
# ============================================================


def main():
    asegurar_carpetas()

    if not CARPETA_TXT_MODERNO.exists():
        raise FileNotFoundError(f"No existe {CARPETA_TXT_MODERNO}. Primero ejecuta 01_modernizar_montemar_api.py")

    archivos = sorted(CARPETA_TXT_MODERNO.glob("*_moderno.txt"))
    if not archivos:
        raise FileNotFoundError(f"No hay archivos *_moderno.txt en {CARPETA_TXT_MODERNO}")

    paginas = []
    for p in archivos:
        meta = extraer_meta_desde_nombre(p.name)
        texto = p.read_text(encoding="utf-8", errors="replace")
        paginas.append({**meta, "archivo_txt": str(p), "texto": texto})

    if MAX_CARTAS is not None:
        ids = []
        for p in paginas:
            if p["carta_id"] not in ids:
                ids.append(p["carta_id"])
        permitidos = set(ids[:MAX_CARTAS])
        paginas = [p for p in paginas if p["carta_id"] in permitidos]

    print(f"Páginas modernas a analizar: {len(paginas)}")

    textos_cartas = {}
    for p in paginas:
        textos_cartas.setdefault(p["carta_id"], []).append((p["pagina"], p["texto"], p["nro_carta"]))

    # Por página
    freq_paginas, asoc_paginas, resumen_paginas = [], [], []

    for p in paginas:
        tokens = tokenizar(p["texto"])
        freq_df = frecuencias_tokens(tokens)
        asoc_df = asociaciones_tokens(tokens)
        base = limpiar_nombre_archivo(f"{p['carta_id']}_page_{p['pagina']:03d}")

        resumen_paginas.append({
            "carta_id": p["carta_id"], "nro_carta": p["nro_carta"], "pagina": p["pagina"],
            "total_tokens": len(tokens), "total_palabras_unicas": freq_df["palabra"].nunique() if not freq_df.empty else 0
        })

        if not freq_df.empty:
            freq_df.insert(0, "pagina", p["pagina"])
            freq_df.insert(0, "nro_carta", p["nro_carta"])
            freq_df.insert(0, "carta_id", p["carta_id"])
            freq_paginas.append(freq_df)
            crear_nube(freq_df[["palabra", "frecuencia"]], CARPETA_NUBES_PAGINA / f"{base}_nube.png", f"Nube moderna - {p['carta_id']} - página {p['pagina']}")

        if not asoc_df.empty:
            asoc_df.insert(0, "pagina", p["pagina"])
            asoc_df.insert(0, "nro_carta", p["nro_carta"])
            asoc_df.insert(0, "carta_id", p["carta_id"])
            asoc_paginas.append(asoc_df)
            crear_red_asociaciones(asoc_df[["palabra_1", "palabra_2", "coocurrencias"]], CARPETA_REDES_PAGINA / f"{base}_red.png", f"Asociaciones modernas - {p['carta_id']} - página {p['pagina']}")

    # Por carta
    freq_cartas, asoc_cartas, resumen_cartas = [], [], []

    for cid, partes in textos_cartas.items():
        partes = sorted(partes, key=lambda x: x[0])
        texto = "\n".join(t for _, t, _ in partes)
        nro = partes[0][2]
        tokens = tokenizar(texto)
        freq_df = frecuencias_tokens(tokens)
        asoc_df = asociaciones_tokens(tokens)
        base = limpiar_nombre_archivo(cid)

        resumen_cartas.append({
            "carta_id": cid, "nro_carta": nro, "total_paginas": len(partes),
            "total_tokens": len(tokens), "total_palabras_unicas": freq_df["palabra"].nunique() if not freq_df.empty else 0
        })

        if not freq_df.empty:
            freq_df.insert(0, "nro_carta", nro)
            freq_df.insert(0, "carta_id", cid)
            freq_cartas.append(freq_df)
            crear_nube(freq_df[["palabra", "frecuencia"]], CARPETA_NUBES_CARTA / f"{base}_nube.png", f"Nube moderna - {cid}")

        if not asoc_df.empty:
            asoc_df.insert(0, "nro_carta", nro)
            asoc_df.insert(0, "carta_id", cid)
            asoc_cartas.append(asoc_df)
            crear_red_asociaciones(asoc_df[["palabra_1", "palabra_2", "coocurrencias"]], CARPETA_REDES_CARTA / f"{base}_red.png", f"Asociaciones modernas - {cid}")

    df_resumen_cartas = pd.DataFrame(resumen_cartas)
    df_resumen_paginas = pd.DataFrame(resumen_paginas)
    df_freq_cartas = pd.concat(freq_cartas, ignore_index=True) if freq_cartas else pd.DataFrame()
    df_freq_paginas = pd.concat(freq_paginas, ignore_index=True) if freq_paginas else pd.DataFrame()
    df_asoc_cartas = pd.concat(asoc_cartas, ignore_index=True) if asoc_cartas else pd.DataFrame()
    df_asoc_paginas = pd.concat(asoc_paginas, ignore_index=True) if asoc_paginas else pd.DataFrame()

    df_resumen_cartas.to_csv(CARPETA_SALIDA / "resumen_por_carta.csv", index=False, encoding="utf-8-sig")
    df_resumen_paginas.to_csv(CARPETA_SALIDA / "resumen_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_freq_cartas.to_csv(CARPETA_FRECUENCIAS / "frecuencia_por_carta.csv", index=False, encoding="utf-8-sig")
    df_freq_paginas.to_csv(CARPETA_FRECUENCIAS / "frecuencia_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_asoc_cartas.to_csv(CARPETA_ASOCIACIONES / "asociaciones_por_carta.csv", index=False, encoding="utf-8-sig")
    df_asoc_paginas.to_csv(CARPETA_ASOCIACIONES / "asociaciones_por_pagina.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl") as writer:
        df_resumen_cartas.to_excel(writer, sheet_name="resumen_cartas", index=False)
        df_resumen_paginas.to_excel(writer, sheet_name="resumen_paginas", index=False)
        (df_freq_cartas.groupby("carta_id", group_keys=False).head(100) if not df_freq_cartas.empty else df_freq_cartas).to_excel(writer, sheet_name="freq_carta_top100", index=False)
        (df_freq_paginas.groupby(["carta_id", "pagina"], group_keys=False).head(100) if not df_freq_paginas.empty else df_freq_paginas).to_excel(writer, sheet_name="freq_pagina_top100", index=False)
        (df_asoc_cartas.groupby("carta_id", group_keys=False).head(100) if not df_asoc_cartas.empty else df_asoc_cartas).to_excel(writer, sheet_name="asoc_carta_top100", index=False)
        (df_asoc_paginas.groupby(["carta_id", "pagina"], group_keys=False).head(100) if not df_asoc_paginas.empty else df_asoc_paginas).to_excel(writer, sheet_name="asoc_pagina_top100", index=False)

    print("\nProceso terminado.")
    print(f"Salida: {CARPETA_SALIDA.resolve()}")
    print(f"Excel: {ARCHIVO_EXCEL.resolve()}")


if __name__ == "__main__":
    main()
