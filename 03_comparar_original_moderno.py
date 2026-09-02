# -*- coding: utf-8 -*-
"""
03_COMPARAR_ORIGINAL_MODERNO.PY

Compara el corpus diplomático original contra el corpus modernizado por API.

ENTRADAS:
    Original:  montemar_moderno/txt_original_por_pagina/*_original.txt
    Moderno:   montemar_moderno/txt_moderno_por_pagina/*_moderno.txt

SALIDA:
    comparacion_montemar_original_moderno/
    ├── comparacion_original_moderno.xlsx
    ├── csv/
    └── graficos/

COMPARA:
    - tokens por carta/página
    - palabras únicas
    - riqueza léxica
    - hapax
    - diferencias de longitud
    - top palabras original vs moderno
    - palabras que desaparecen o aparecen tras modernización
    - similitud de vocabulario por índice de Jaccard

REQUISITOS:
    pip install pandas openpyxl matplotlib regex

EJECUCIÓN:
    python 03_comparar_original_moderno.py
"""

import re
import html
from pathlib import Path
from collections import Counter

import pandas as pd
import matplotlib.pyplot as plt

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
CARPETA_ORIG = CARPETA_MODERNO / "txt_original_por_pagina"
CARPETA_MOD = CARPETA_MODERNO / "txt_moderno_por_pagina"

CARPETA_SALIDA = Path("comparacion_montemar_original_moderno")
CARPETA_CSV = CARPETA_SALIDA / "csv"
CARPETA_GRAFICOS = CARPETA_SALIDA / "graficos"
ARCHIVO_EXCEL = CARPETA_SALIDA / "comparacion_original_moderno.xlsx"

PASAR_A_MINUSCULAS = True
USAR_STOPWORDS = False
TOP_N = 50

STOPWORDS_BASE = {
    "a", "al", "algo", "algunas", "algunos", "ante", "antes", "así", "aunque",
    "como", "con", "contra", "cuando", "de", "del", "desde", "donde", "e", "el",
    "él", "ella", "ellas", "ellos", "en", "entre", "era", "es", "esa", "esas",
    "ese", "eso", "esos", "esta", "está", "estaba", "estaban", "estado", "estas",
    "este", "esto", "estos", "ha", "han", "hasta", "hay", "la", "las", "le", "les",
    "lo", "los", "más", "me", "mi", "mis", "muy", "no", "nos", "o", "para", "pero",
    "por", "porque", "qᵉ", "que", "se", "si", "sí", "sin", "sobre", "su", "sus", "te",
    "tu", "tus", "un", "una", "uno", "unos", "y", "ya", "yo"
}


# ============================================================
# UTILIDADES
# ============================================================


def asegurar_carpetas():
    for c in [CARPETA_SALIDA, CARPETA_CSV, CARPETA_GRAFICOS]:
        c.mkdir(parents=True, exist_ok=True)


def limpiar_nombre_archivo(s: str, max_len: int = 120) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def meta_desde_nombre_original(nombre: str) -> dict:
    base = nombre.replace("_original.txt", "")
    m_page = re.search(r"_page_(\d+)$", base)
    pagina = int(m_page.group(1)) if m_page else 0
    carta_id = re.sub(r"_page_\d+$", "", base)
    m_carta = re.match(r"^(\d+)_", carta_id)
    nro_carta = int(m_carta.group(1)) if m_carta else None
    return {"carta_id": carta_id, "nro_carta": nro_carta, "pagina": pagina}


def ruta_moderno_para_original(path_original: Path) -> Path:
    return CARPETA_MOD / path_original.name.replace("_original.txt", "_moderno.txt")


# ============================================================
# TOKENIZACIÓN Y MÉTRICAS
# ============================================================


def tokenizar(texto: str) -> list:
    if PASAR_A_MINUSCULAS:
        texto = texto.lower()

    if TIENE_REGEX:
        patron = r"[\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+(?:[-'’´`~][\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+)*"
        tokens = regex_unicode.findall(patron, texto)
    else:
        patron = r"[^\W\d_]+(?:[-'’´`~][^\W\d_]+)*"
        tokens = re.findall(patron, texto, flags=re.UNICODE)

    tokens = [t.strip("-'’´`~") for t in tokens if t.strip("-'’´`~")]

    if USAR_STOPWORDS:
        tokens = [t for t in tokens if t not in STOPWORDS_BASE]

    return tokens


def metricas(tokens: list) -> dict:
    c = Counter(tokens)
    total = len(tokens)
    unicas = len(c)
    hapax = sum(1 for _, n in c.items() if n == 1)
    return {
        "tokens": total,
        "palabras_unicas": unicas,
        "riqueza_lexica": unicas / total if total else 0,
        "hapax": hapax,
        "porcentaje_hapax": hapax / unicas * 100 if unicas else 0
    }


def jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    return len(a & b) / len(a | b) if (a | b) else 0


# ============================================================
# GRÁFICOS
# ============================================================


def grafico_barras(df, x, y, titulo, salida, top=32):
    if df.empty:
        return
    d = df.head(top).copy()
    plt.figure(figsize=(12, max(6, len(d) * 0.35)))
    plt.barh(d[y].astype(str), d[x])
    plt.gca().invert_yaxis()
    plt.title(titulo)
    plt.xlabel(x)
    plt.ylabel(y)
    plt.tight_layout()
    plt.savefig(salida, dpi=160)
    plt.close()


# ============================================================
# MAIN
# ============================================================


def main():
    asegurar_carpetas()

    if not CARPETA_ORIG.exists() or not CARPETA_MOD.exists():
        raise FileNotFoundError("Primero ejecuta 01_modernizar_montemar_api.py")

    archivos_orig = sorted(CARPETA_ORIG.glob("*_original.txt"))
    if not archivos_orig:
        raise FileNotFoundError(f"No hay archivos originales en {CARPETA_ORIG}")

    filas_paginas = []
    frecuencias_original = []
    frecuencias_moderno = []
    vocab_diferencias_pag = []

    textos_carta_orig = {}
    textos_carta_mod = {}
    nro_por_carta = {}

    for p_orig in archivos_orig:
        p_mod = ruta_moderno_para_original(p_orig)
        if not p_mod.exists():
            print(f"ADVERTENCIA: falta moderno para {p_orig.name}")
            continue

        meta = meta_desde_nombre_original(p_orig.name)
        texto_orig = p_orig.read_text(encoding="utf-8", errors="replace")
        texto_mod = p_mod.read_text(encoding="utf-8", errors="replace")

        tok_orig = tokenizar(texto_orig)
        tok_mod = tokenizar(texto_mod)
        m_orig = metricas(tok_orig)
        m_mod = metricas(tok_mod)

        set_orig = set(tok_orig)
        set_mod = set(tok_mod)

        filas_paginas.append({
            **meta,
            "tokens_original": m_orig["tokens"],
            "tokens_moderno": m_mod["tokens"],
            "dif_tokens": m_mod["tokens"] - m_orig["tokens"],
            "pct_dif_tokens": ((m_mod["tokens"] - m_orig["tokens"]) / m_orig["tokens"] * 100) if m_orig["tokens"] else 0,
            "palabras_unicas_original": m_orig["palabras_unicas"],
            "palabras_unicas_moderno": m_mod["palabras_unicas"],
            "dif_palabras_unicas": m_mod["palabras_unicas"] - m_orig["palabras_unicas"],
            "riqueza_original": m_orig["riqueza_lexica"],
            "riqueza_moderna": m_mod["riqueza_lexica"],
            "dif_riqueza": m_mod["riqueza_lexica"] - m_orig["riqueza_lexica"],
            "hapax_original": m_orig["hapax"],
            "hapax_moderno": m_mod["hapax"],
            "jaccard_vocabulario": jaccard(set_orig, set_mod),
            "solo_original_n": len(set_orig - set_mod),
            "solo_moderno_n": len(set_mod - set_orig),
            "archivo_original": str(p_orig),
            "archivo_moderno": str(p_mod),
        })

        for palabra, n in Counter(tok_orig).most_common():
            frecuencias_original.append({**meta, "version": "original", "palabra": palabra, "frecuencia": n})
        for palabra, n in Counter(tok_mod).most_common():
            frecuencias_moderno.append({**meta, "version": "moderno", "palabra": palabra, "frecuencia": n})

        for palabra in sorted(set_orig - set_mod):
            vocab_diferencias_pag.append({**meta, "tipo": "solo_original", "palabra": palabra})
        for palabra in sorted(set_mod - set_orig):
            vocab_diferencias_pag.append({**meta, "tipo": "solo_moderno", "palabra": palabra})

        cid = meta["carta_id"]
        textos_carta_orig.setdefault(cid, []).append((meta["pagina"], texto_orig))
        textos_carta_mod.setdefault(cid, []).append((meta["pagina"], texto_mod))
        nro_por_carta[cid] = meta["nro_carta"]

    df_paginas = pd.DataFrame(filas_paginas)
    df_freq_original = pd.DataFrame(frecuencias_original)
    df_freq_moderno = pd.DataFrame(frecuencias_moderno)
    df_vocab_pag = pd.DataFrame(vocab_diferencias_pag)

    # Carta consolidada
    filas_cartas = []
    freq_cartas = []
    vocab_diferencias_carta = []

    for cid in sorted(textos_carta_orig):
        texto_orig = "\n".join(t for _, t in sorted(textos_carta_orig[cid]))
        texto_mod = "\n".join(t for _, t in sorted(textos_carta_mod.get(cid, [])))
        tok_orig = tokenizar(texto_orig)
        tok_mod = tokenizar(texto_mod)
        m_orig = metricas(tok_orig)
        m_mod = metricas(tok_mod)
        set_orig = set(tok_orig)
        set_mod = set(tok_mod)

        filas_cartas.append({
            "carta_id": cid,
            "nro_carta": nro_por_carta.get(cid),
            "total_paginas": len(textos_carta_orig[cid]),
            "tokens_original": m_orig["tokens"],
            "tokens_moderno": m_mod["tokens"],
            "dif_tokens": m_mod["tokens"] - m_orig["tokens"],
            "pct_dif_tokens": ((m_mod["tokens"] - m_orig["tokens"]) / m_orig["tokens"] * 100) if m_orig["tokens"] else 0,
            "palabras_unicas_original": m_orig["palabras_unicas"],
            "palabras_unicas_moderno": m_mod["palabras_unicas"],
            "dif_palabras_unicas": m_mod["palabras_unicas"] - m_orig["palabras_unicas"],
            "riqueza_original": m_orig["riqueza_lexica"],
            "riqueza_moderna": m_mod["riqueza_lexica"],
            "dif_riqueza": m_mod["riqueza_lexica"] - m_orig["riqueza_lexica"],
            "hapax_original": m_orig["hapax"],
            "hapax_moderno": m_mod["hapax"],
            "jaccard_vocabulario": jaccard(set_orig, set_mod),
            "solo_original_n": len(set_orig - set_mod),
            "solo_moderno_n": len(set_mod - set_orig),
        })

        for palabra, n in Counter(tok_orig).most_common(TOP_N):
            freq_cartas.append({"carta_id": cid, "nro_carta": nro_por_carta.get(cid), "version": "original", "palabra": palabra, "frecuencia": n})
        for palabra, n in Counter(tok_mod).most_common(TOP_N):
            freq_cartas.append({"carta_id": cid, "nro_carta": nro_por_carta.get(cid), "version": "moderno", "palabra": palabra, "frecuencia": n})

        for palabra in sorted(set_orig - set_mod):
            vocab_diferencias_carta.append({"carta_id": cid, "nro_carta": nro_por_carta.get(cid), "tipo": "solo_original", "palabra": palabra})
        for palabra in sorted(set_mod - set_orig):
            vocab_diferencias_carta.append({"carta_id": cid, "nro_carta": nro_por_carta.get(cid), "tipo": "solo_moderno", "palabra": palabra})

    df_cartas = pd.DataFrame(filas_cartas)
    df_freq_cartas = pd.DataFrame(freq_cartas)
    df_vocab_carta = pd.DataFrame(vocab_diferencias_carta)

    # CSV
    df_cartas.to_csv(CARPETA_CSV / "comparacion_por_carta.csv", index=False, encoding="utf-8-sig")
    df_paginas.to_csv(CARPETA_CSV / "comparacion_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_freq_original.to_csv(CARPETA_CSV / "frecuencias_original_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_freq_moderno.to_csv(CARPETA_CSV / "frecuencias_moderno_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_freq_cartas.to_csv(CARPETA_CSV / "top_frecuencias_por_carta_original_moderno.csv", index=False, encoding="utf-8-sig")
    df_vocab_pag.to_csv(CARPETA_CSV / "vocabulario_diferencias_por_pagina.csv", index=False, encoding="utf-8-sig")
    df_vocab_carta.to_csv(CARPETA_CSV / "vocabulario_diferencias_por_carta.csv", index=False, encoding="utf-8-sig")

    # Gráficos
    if not df_cartas.empty:
        d = df_cartas.sort_values("pct_dif_tokens", ascending=False).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        grafico_barras(d, "pct_dif_tokens", "label", "Cambio porcentual de tokens: moderno vs original", CARPETA_GRAFICOS / "pct_dif_tokens_por_carta.png")

        d = df_cartas.sort_values("dif_palabras_unicas", ascending=False).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        grafico_barras(d, "dif_palabras_unicas", "label", "Diferencia de palabras únicas: moderno - original", CARPETA_GRAFICOS / "dif_palabras_unicas_por_carta.png")

        d = df_cartas.sort_values("jaccard_vocabulario", ascending=True).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        grafico_barras(d, "jaccard_vocabulario", "label", "Similitud de vocabulario Jaccard por carta", CARPETA_GRAFICOS / "jaccard_vocabulario_por_carta.png")

    # Excel
    with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl") as writer:
        df_cartas.to_excel(writer, sheet_name="comparacion_cartas", index=False)
        df_paginas.to_excel(writer, sheet_name="comparacion_paginas", index=False)
        df_freq_cartas.to_excel(writer, sheet_name="top_freq_cartas", index=False)
        df_vocab_carta.to_excel(writer, sheet_name="vocab_dif_carta", index=False)
        df_vocab_pag.to_excel(writer, sheet_name="vocab_dif_pagina", index=False)

    print("\nProceso terminado.")
    print(f"Salida: {CARPETA_SALIDA.resolve()}")
    print(f"Excel: {ARCHIVO_EXCEL.resolve()}")


if __name__ == "__main__":
    main()
