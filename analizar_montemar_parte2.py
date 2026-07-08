# -*- coding: utf-8 -*-
"""
SEGUNDA PARTE DE ANÁLISIS - CARTAS DE MONTEMAR

Trabaja con:
    montemar_transcripciones/html_por_pagina/

Exporta en una carpeta nueva:
    analisis_montemar_parte2/

Incluye:
1) Riqueza léxica por carta y página.
2) TF-IDF por carta y página.
3) Entidades desde etiquetas HTML: persName, geogName, placeName, name, date, quote.
4) Marcas diplomáticas: del, add, subst, gap, superíndices.
5) Categorías temáticas personalizables.
6) Gráficos PNG y Excel consolidado.

Requisitos:
    pip install beautifulsoup4 lxml pandas openpyxl matplotlib regex scikit-learn

Ejecución:
    python analizar_montemar_parte2.py
"""

import re
import html
from pathlib import Path
from collections import Counter, defaultdict

import pandas as pd
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup, NavigableString, Tag

try:
    import regex as regex_unicode
    TIENE_REGEX = True
except ImportError:
    regex_unicode = None
    TIENE_REGEX = False

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    TIENE_SKLEARN = True
except ImportError:
    TfidfVectorizer = None
    TIENE_SKLEARN = False


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_BASE = Path("montemar_transcripciones")
CARPETA_HTML_PAGINAS = CARPETA_BASE / "html_por_pagina"

CARPETA_SALIDA = Path("analisis_montemar_parte2")
CARPETA_CSV = CARPETA_SALIDA / "csv"
CARPETA_GRAFICOS = CARPETA_SALIDA / "graficos"
CARPETA_EXTRACTOS = CARPETA_SALIDA / "extractos_html"

ARCHIVO_EXCEL = CARPETA_SALIDA / "analisis_montemar_parte2.xlsx"

MAX_CARTAS = None
PASAR_A_MINUSCULAS = True
USAR_STOPWORDS = False
TOP_TFIDF = 50
TOP_CARTAS_GRAFICO = 32

STOPWORDS_BASE = {
    "a", "al", "ala", "alas", "alos", "ante", "asi", "con", "contra",
    "de", "del", "dela", "delas", "delos", "desde", "e", "el", "ella",
    "ellos", "en", "entre", "era", "es", "esta", "estas", "este", "esto",
    "estos", "ha", "hai", "hasta", "la", "las", "le", "lo", "los", "mas",
    "me", "mi", "mis", "no", "o", "para", "por", "porque", "q", "qᵉ",
    "que", "se", "si", "sin", "su", "sus", "te", "tu", "tus", "un", "una",
    "y", "yo", "d", "dn", "dⁿ", "s", "sⁿ", "n", "nᵗᵒ", "r", "p"
}

CATEGORIAS_TEMATICAS = {
    "familia_parentesco": [
        "hermano", "hermana", "hermanos", "padre", "madre", "tia", "tía",
        "abuelo", "abuela", "hijo", "hija", "hijos", "familia", "casa",
        "pariente", "sobrino", "sobrina", "primo", "prima", "amado",
        "amante", "obediencia"
    ],
    "administracion_cargos": [
        "corregimiento", "corregidor", "oficio", "oficios", "escribano",
        "consulado", "sedula", "cédula", "pribilejio", "privilegio",
        "derechos", "nombramiento", "theniente", "teniente", "audiencia",
        "provincia", "navios", "navío", "navieros"
    ],
    "religion": [
        "bautismo", "oleo", "óleo", "curato", "yndios", "indios",
        "capellania", "capellanía", "capellanias", "capellanías",
        "missa", "misa", "missas", "misas", "presbitero", "presbítero",
        "cofradia", "cofradía", "cura", "padre", "sacramento", "iglesia"
    ],
    "economia_patrimonio": [
        "mayorasgo", "mayorazgo", "renta", "rrenta", "principal",
        "pension", "pensión", "derechos", "intereses", "utilidad",
        "utilidades", "dinero", "finca", "cassas", "casas", "censos",
        "vinculo", "vínculo", "patronato", "bienes", "soldada"
    ],
    "lugares_territorio": [
        "lima", "canta", "guanta", "sebilla", "sevilla", "america",
        "américa", "apongo", "mar", "norte", "pueblo", "provincia",
        "sercado", "cercado", "santa", "ana"
    ],
    "escritura_documentos": [
        "carta", "cartas", "letras", "papeles", "partida", "partidas",
        "testamento", "thestamento", "testamᵗᵒ", "certificasion",
        "sertificasion", "informe", "escritura", "clausula", "cláusula",
        "libro", "sedula", "cédula"
    ]
}

SUPER_MAP = str.maketrans({
    "a": "ᵃ", "b": "ᵇ", "c": "ᶜ", "d": "ᵈ", "e": "ᵉ", "f": "ᶠ",
    "g": "ᵍ", "h": "ʰ", "i": "ᶦ", "j": "ʲ", "k": "ᵏ", "l": "ˡ",
    "m": "ᵐ", "n": "ⁿ", "o": "ᵒ", "p": "ᵖ", "r": "ʳ", "s": "ˢ",
    "t": "ᵗ", "u": "ᵘ", "v": "ᵛ", "w": "ʷ", "x": "ˣ", "y": "ʸ",
    "z": "ᶻ", "A": "ᴬ", "B": "ᴮ", "D": "ᴰ", "E": "ᴱ", "G": "ᴳ",
    "H": "ᴴ", "I": "ᴵ", "J": "ᴶ", "K": "ᴷ", "L": "ᴸ", "M": "ᴹ",
    "N": "ᴺ", "O": "ᴼ", "P": "ᴾ", "R": "ᴿ", "T": "ᵀ", "U": "ᵁ",
    "V": "ⱽ", "W": "ᵂ", "0": "⁰", "1": "¹", "2": "²", "3": "³",
    "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹"
})


def a_superindice(texto: str) -> str:
    return texto.translate(SUPER_MAP)


def asegurar_carpetas():
    for c in [CARPETA_SALIDA, CARPETA_CSV, CARPETA_GRAFICOS, CARPETA_EXTRACTOS]:
        c.mkdir(parents=True, exist_ok=True)


def limpiar_nombre_archivo(s: str, max_len: int = 120) -> str:
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


def tokenizar(texto: str) -> list:
    if PASAR_A_MINUSCULAS:
        texto = texto.lower()

    texto = texto.replace("[GAP_", " GAP_").replace("]", " ")

    if TIENE_REGEX:
        patron = r"[\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+(?:[-'’´`~][\p{L}\p{M}ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻᴬᴮᴰᴱᴳᴴᴵᴶᴷᴸᴹᴺᴼᴾᴿᵀᵁⱽᵂºª]+)*"
        tokens = regex_unicode.findall(patron, texto)
    else:
        patron = r"[^\W\d_ᵃᵇᶜᵈᵉᶠᵍʰᶦʲᵏˡᵐⁿᵒᵖʳˢᵗᵘᵛʷˣʸᶻºª]+(?:[-'’´`~][^\W\d_]+)*"
        tokens = re.findall(patron, texto, flags=re.UNICODE)

    tokens = [t.strip("-'’´`~") for t in tokens if t.strip("-'’´`~")]
    if USAR_STOPWORDS:
        tokens = [t for t in tokens if t not in STOPWORDS_BASE]
    return tokens


def leer_pagina_html(path: Path) -> dict:
    raw = path.read_text(encoding="utf-8", errors="replace")
    soup = BeautifulSoup(raw, "lxml")

    page_tag = soup.select_one("span.page") or soup.select_one("span.body") or soup.body or soup
    titulo_tag = soup.find("h1")
    titulo = titulo_tag.get_text(" ", strip=True) if titulo_tag else extraer_id_carta(path.name)

    carta_id = extraer_id_carta(path.name)
    pagina = extraer_num_pagina(path.name)
    texto = normalizar_espacios(texto_diplomatico(page_tag))

    return {
        "carta_id": carta_id,
        "nro_carta": extraer_num_carta(carta_id),
        "titulo": titulo,
        "pagina": pagina,
        "archivo_html": str(path),
        "texto": texto,
        "soup": soup,
        "page_tag": page_tag
    }


def calcular_metricas_lexicas(tokens: list) -> dict:
    total = len(tokens)
    conteo = Counter(tokens)
    unicas = len(conteo)
    hapax = sum(1 for _, n in conteo.items() if n == 1)
    return {
        "total_tokens": total,
        "palabras_unicas": unicas,
        "riqueza_lexica": (unicas / total) if total else 0,
        "hapax": hapax,
        "porcentaje_hapax": (hapax / unicas * 100) if unicas else 0
    }


def identidad_tokenizer(texto: str):
    return tokenizar(texto)


def calcular_tfidf(documentos: list, metadatos: list, nivel: str) -> pd.DataFrame:
    if not TIENE_SKLEARN:
        print("ADVERTENCIA: scikit-learn no está instalado. Se omite TF-IDF.")
        return pd.DataFrame()
    if len(documentos) < 2:
        return pd.DataFrame()

    vectorizer = TfidfVectorizer(
        tokenizer=identidad_tokenizer,
        preprocessor=None,
        token_pattern=None,
        lowercase=False,
        min_df=1
    )
    matriz = vectorizer.fit_transform(documentos)
    palabras = vectorizer.get_feature_names_out()

    filas = []
    for i, meta in enumerate(metadatos):
        fila = matriz.getrow(i)
        datos = sorted(zip(fila.indices, fila.data), key=lambda x: x[1], reverse=True)[:TOP_TFIDF]
        for idx, score in datos:
            out = dict(meta)
            out["nivel"] = nivel
            out["palabra"] = palabras[idx]
            out["tfidf"] = float(score)
            filas.append(out)
    return pd.DataFrame(filas)


def extraer_entidades(pagina: dict) -> list:
    page_tag = pagina["page_tag"]
    clases_entidad = ["persName", "geogName", "placeName", "name", "date", "quote"]
    filas = []

    for clase in clases_entidad:
        for tag in page_tag.select(f".{clase}"):
            texto = normalizar_espacios(texto_diplomatico(tag))
            if texto:
                filas.append({
                    "carta_id": pagina["carta_id"],
                    "nro_carta": pagina["nro_carta"],
                    "titulo": pagina["titulo"],
                    "pagina": pagina["pagina"],
                    "tipo_entidad": clase,
                    "texto": texto
                })
    return filas


def clases_place(tag: Tag) -> str:
    if tag is None:
        return ""
    clases = set(tag.get("class", []))
    return ",".join(sorted([c for c in clases if c.startswith("place_")]))


def extraer_marcas_diplomaticas(pagina: dict) -> list:
    page_tag = pagina["page_tag"]
    filas = []

    base = {
        "carta_id": pagina["carta_id"],
        "nro_carta": pagina["nro_carta"],
        "titulo": pagina["titulo"],
        "pagina": pagina["pagina"]
    }

    for tag in page_tag.select(".del"):
        texto = normalizar_espacios(texto_diplomatico(tag))
        filas.append({**base, "tipo_marca": "del", "subtipo": "", "texto": texto,
                      "texto_del": texto, "texto_add": "", "posicion_add": ""})

    for tag in page_tag.select(".add"):
        texto = normalizar_espacios(texto_diplomatico(tag))
        filas.append({**base, "tipo_marca": "add", "subtipo": clases_place(tag), "texto": texto,
                      "texto_del": "", "texto_add": texto, "posicion_add": clases_place(tag)})

    for tag in page_tag.select(".subst"):
        del_tag = tag.select_one(".del")
        add_tag = tag.select_one(".add")
        texto_del = normalizar_espacios(texto_diplomatico(del_tag)) if del_tag else ""
        texto_add = normalizar_espacios(texto_diplomatico(add_tag)) if add_tag else ""
        filas.append({**base, "tipo_marca": "subst", "subtipo": clases_place(add_tag),
                      "texto": f"{texto_del} -> {texto_add}".strip(),
                      "texto_del": texto_del, "texto_add": texto_add, "posicion_add": clases_place(add_tag)})

    for tag in page_tag.select(".gap"):
        clases = set(tag.get("class", []))
        reason = ""
        for c in clases:
            if c.startswith("reason_"):
                reason = c.replace("reason_", "")
        filas.append({**base, "tipo_marca": "gap", "subtipo": reason,
                      "texto": normalizar_espacios(texto_diplomatico(tag)),
                      "texto_del": "", "texto_add": "", "posicion_add": ""})

    for tag in page_tag.select(".g.rend_superior"):
        filas.append({**base, "tipo_marca": "superindice", "subtipo": "rend_superior",
                      "texto": normalizar_espacios(texto_diplomatico(tag)),
                      "texto_del": "", "texto_add": "", "posicion_add": ""})
    return filas


def analizar_categorias(tokens: list, meta: dict) -> list:
    conteo = Counter(tokens)
    total_tokens = len(tokens)
    filas = []
    for categoria, palabras in CATEGORIAS_TEMATICAS.items():
        palabras_norm = [p.lower() if PASAR_A_MINUSCULAS else p for p in palabras]
        frecuencia = sum(conteo.get(p, 0) for p in palabras_norm)
        palabras_encontradas = sorted([p for p in palabras_norm if conteo.get(p, 0) > 0])
        filas.append({
            **meta,
            "categoria": categoria,
            "frecuencia_categoria": frecuencia,
            "porcentaje_tokens": (frecuencia / total_tokens * 100) if total_tokens else 0,
            "palabras_encontradas": ", ".join(palabras_encontradas)
        })
    return filas


def guardar_barra_horizontal(df, x_col, y_col, titulo, salida, top=TOP_CARTAS_GRAFICO):
    if df.empty:
        return
    d = df.copy().head(top)
    plt.figure(figsize=(12, max(6, len(d) * 0.35)))
    plt.barh(d[y_col].astype(str), d[x_col])
    plt.gca().invert_yaxis()
    plt.title(titulo)
    plt.xlabel(x_col)
    plt.ylabel(y_col)
    plt.tight_layout()
    plt.savefig(salida, dpi=160)
    plt.close()


def graficar_categorias(df_cat_carta: pd.DataFrame):
    if df_cat_carta.empty:
        return

    total_cat = (
        df_cat_carta.groupby("categoria", as_index=False)["frecuencia_categoria"]
        .sum()
        .sort_values("frecuencia_categoria", ascending=False)
    )
    plt.figure(figsize=(12, 7))
    plt.barh(total_cat["categoria"], total_cat["frecuencia_categoria"])
    plt.gca().invert_yaxis()
    plt.title("Frecuencia total por categoría temática")
    plt.xlabel("Frecuencia")
    plt.ylabel("Categoría")
    plt.tight_layout()
    plt.savefig(CARPETA_GRAFICOS / "categorias_tematicas_total.png", dpi=160)
    plt.close()

    pivot = df_cat_carta.pivot_table(
        index="carta_id",
        columns="categoria",
        values="frecuencia_categoria",
        aggfunc="sum",
        fill_value=0
    )
    if pivot.empty:
        return
    pivot = pivot.head(TOP_CARTAS_GRAFICO)
    ax = pivot.plot(kind="bar", stacked=True, figsize=(16, 8))
    ax.set_title("Categorías temáticas por carta")
    ax.set_xlabel("Carta")
    ax.set_ylabel("Frecuencia")
    plt.xticks(rotation=75, ha="right")
    plt.tight_layout()
    plt.savefig(CARPETA_GRAFICOS / "categorias_tematicas_por_carta.png", dpi=160)
    plt.close()


def main():
    asegurar_carpetas()

    if not CARPETA_HTML_PAGINAS.exists():
        raise FileNotFoundError(
            f"No existe {CARPETA_HTML_PAGINAS}. Primero ejecuta el script de descarga."
        )

    archivos = sorted(CARPETA_HTML_PAGINAS.glob("*_page_*.html"))
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos *_page_*.html en {CARPETA_HTML_PAGINAS}")

    print(f"Páginas HTML encontradas: {len(archivos)}")
    paginas = [leer_pagina_html(p) for p in archivos]

    if MAX_CARTAS is not None:
        ids = []
        for p in paginas:
            if p["carta_id"] not in ids:
                ids.append(p["carta_id"])
        ids_permitidos = set(ids[:MAX_CARTAS])
        paginas = [p for p in paginas if p["carta_id"] in ids_permitidos]

    textos_paginas = []
    textos_cartas_dict = defaultdict(list)
    meta_cartas = {}

    for p in paginas:
        tokens = tokenizar(p["texto"])
        textos_paginas.append({**{k: p[k] for k in ["carta_id", "nro_carta", "titulo", "pagina"]},
                               "texto": p["texto"], "tokens": tokens})
        textos_cartas_dict[p["carta_id"]].append((p["pagina"], p["texto"]))
        meta_cartas[p["carta_id"]] = {k: p[k] for k in ["carta_id", "nro_carta", "titulo"]}

    textos_cartas = []
    for cid, partes in textos_cartas_dict.items():
        partes = sorted(partes, key=lambda x: x[0])
        texto = "\n".join(t for _, t in partes)
        meta = meta_cartas[cid]
        textos_cartas.append({**meta, "texto": texto, "tokens": tokenizar(texto), "total_paginas": len(partes)})

    print("Calculando riqueza léxica...")
    df_lex_pag = pd.DataFrame([
        {"carta_id": p["carta_id"], "nro_carta": p["nro_carta"], "titulo": p["titulo"],
         "pagina": p["pagina"], **calcular_metricas_lexicas(p["tokens"])}
        for p in textos_paginas
    ])
    df_lex_cartas = pd.DataFrame([
        {"carta_id": c["carta_id"], "nro_carta": c["nro_carta"], "titulo": c["titulo"],
         "total_paginas": c["total_paginas"], **calcular_metricas_lexicas(c["tokens"])}
        for c in textos_cartas
    ])

    print("Calculando TF-IDF...")
    df_tfidf_cartas = calcular_tfidf(
        [c["texto"] for c in textos_cartas],
        [{"carta_id": c["carta_id"], "nro_carta": c["nro_carta"], "titulo": c["titulo"]} for c in textos_cartas],
        "carta"
    )
    df_tfidf_paginas = calcular_tfidf(
        [p["texto"] for p in textos_paginas],
        [{"carta_id": p["carta_id"], "nro_carta": p["nro_carta"], "titulo": p["titulo"], "pagina": p["pagina"]} for p in textos_paginas],
        "pagina"
    )

    print("Extrayendo entidades HTML...")
    df_entidades = pd.DataFrame([fila for p in paginas for fila in extraer_entidades(p)])
    if not df_entidades.empty:
        df_entidades_resumen_carta = (
            df_entidades.groupby(["carta_id", "nro_carta", "titulo", "tipo_entidad", "texto"], as_index=False)
            .size().rename(columns={"size": "frecuencia"})
            .sort_values(["carta_id", "tipo_entidad", "frecuencia"], ascending=[True, True, False])
        )
        df_entidades_resumen_pagina = (
            df_entidades.groupby(["carta_id", "nro_carta", "titulo", "pagina", "tipo_entidad", "texto"], as_index=False)
            .size().rename(columns={"size": "frecuencia"})
            .sort_values(["carta_id", "pagina", "tipo_entidad", "frecuencia"], ascending=[True, True, True, False])
        )
        df_entidades_totales_carta = (
            df_entidades.groupby(["carta_id", "nro_carta", "titulo"], as_index=False)
            .size().rename(columns={"size": "total_entidades"})
            .sort_values("total_entidades", ascending=False)
        )
    else:
        df_entidades_resumen_carta = pd.DataFrame()
        df_entidades_resumen_pagina = pd.DataFrame()
        df_entidades_totales_carta = pd.DataFrame()

    print("Extrayendo marcas diplomáticas...")
    df_marcas = pd.DataFrame([fila for p in paginas for fila in extraer_marcas_diplomaticas(p)])
    if not df_marcas.empty:
        df_marcas_resumen_carta = (
            df_marcas.groupby(["carta_id", "nro_carta", "titulo", "tipo_marca"], as_index=False)
            .size().rename(columns={"size": "frecuencia"})
            .sort_values(["carta_id", "tipo_marca"])
        )
        df_marcas_resumen_pagina = (
            df_marcas.groupby(["carta_id", "nro_carta", "titulo", "pagina", "tipo_marca"], as_index=False)
            .size().rename(columns={"size": "frecuencia"})
            .sort_values(["carta_id", "pagina", "tipo_marca"])
        )
        df_marcas_totales_carta = (
            df_marcas.groupby(["carta_id", "nro_carta", "titulo"], as_index=False)
            .size().rename(columns={"size": "total_marcas"})
            .sort_values("total_marcas", ascending=False)
        )
    else:
        df_marcas_resumen_carta = pd.DataFrame()
        df_marcas_resumen_pagina = pd.DataFrame()
        df_marcas_totales_carta = pd.DataFrame()

    print("Calculando categorías temáticas...")
    df_cat_cartas = pd.DataFrame([
        fila
        for c in textos_cartas
        for fila in analizar_categorias(c["tokens"], {"carta_id": c["carta_id"], "nro_carta": c["nro_carta"], "titulo": c["titulo"]})
    ])
    df_cat_paginas = pd.DataFrame([
        fila
        for p in textos_paginas
        for fila in analizar_categorias(p["tokens"], {"carta_id": p["carta_id"], "nro_carta": p["nro_carta"], "titulo": p["titulo"], "pagina": p["pagina"]})
    ])

    print("Guardando extractos...")
    if not df_marcas.empty:
        for tipo in ["subst", "del", "add", "gap"]:
            df_tipo = df_marcas[df_marcas["tipo_marca"] == tipo].copy()
            if not df_tipo.empty:
                df_tipo.to_csv(CARPETA_EXTRACTOS / f"extractos_{tipo}.csv", index=False, encoding="utf-8-sig")
    if not df_entidades.empty:
        for tipo in ["persName", "geogName", "placeName", "name", "date", "quote"]:
            df_tipo = df_entidades[df_entidades["tipo_entidad"] == tipo].copy()
            if not df_tipo.empty:
                df_tipo.to_csv(CARPETA_EXTRACTOS / f"extractos_{tipo}.csv", index=False, encoding="utf-8-sig")

    print("Guardando CSV...")
    salidas_csv = {
        "riqueza_lexica_por_carta.csv": df_lex_cartas,
        "riqueza_lexica_por_pagina.csv": df_lex_pag,
        "tfidf_por_carta.csv": df_tfidf_cartas,
        "tfidf_por_pagina.csv": df_tfidf_paginas,
        "entidades_detalle.csv": df_entidades,
        "entidades_resumen_por_carta.csv": df_entidades_resumen_carta,
        "entidades_resumen_por_pagina.csv": df_entidades_resumen_pagina,
        "marcas_diplomaticas_detalle.csv": df_marcas,
        "marcas_diplomaticas_resumen_por_carta.csv": df_marcas_resumen_carta,
        "marcas_diplomaticas_resumen_por_pagina.csv": df_marcas_resumen_pagina,
        "categorias_tematicas_por_carta.csv": df_cat_cartas,
        "categorias_tematicas_por_pagina.csv": df_cat_paginas,
    }
    for nombre, df in salidas_csv.items():
        df.to_csv(CARPETA_CSV / nombre, index=False, encoding="utf-8-sig")

    print("Generando gráficos...")
    if not df_lex_cartas.empty:
        d = df_lex_cartas.sort_values("riqueza_lexica", ascending=False).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        guardar_barra_horizontal(d, "riqueza_lexica", "label", "Riqueza léxica por carta", CARPETA_GRAFICOS / "riqueza_lexica_por_carta.png")

        d2 = df_lex_cartas.sort_values("hapax", ascending=False).copy()
        d2["label"] = d2["nro_carta"].astype(str).str.zfill(2) + " - " + d2["carta_id"].astype(str).str.slice(0, 35)
        guardar_barra_horizontal(d2, "hapax", "label", "Hapax por carta", CARPETA_GRAFICOS / "hapax_por_carta.png")

    if not df_marcas_totales_carta.empty:
        d = df_marcas_totales_carta.sort_values("total_marcas", ascending=False).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        guardar_barra_horizontal(d, "total_marcas", "label", "Total de marcas diplomáticas por carta", CARPETA_GRAFICOS / "marcas_diplomaticas_por_carta.png")

    if not df_entidades_totales_carta.empty:
        d = df_entidades_totales_carta.sort_values("total_entidades", ascending=False).copy()
        d["label"] = d["nro_carta"].astype(str).str.zfill(2) + " - " + d["carta_id"].astype(str).str.slice(0, 35)
        guardar_barra_horizontal(d, "total_entidades", "label", "Total de entidades HTML por carta", CARPETA_GRAFICOS / "entidades_por_carta.png")

    graficar_categorias(df_cat_cartas)

    print("Guardando Excel...")
    with pd.ExcelWriter(ARCHIVO_EXCEL, engine="openpyxl") as writer:
        df_lex_cartas.to_excel(writer, sheet_name="lexica_cartas", index=False)
        df_lex_pag.to_excel(writer, sheet_name="lexica_paginas", index=False)
        df_tfidf_cartas.to_excel(writer, sheet_name="tfidf_cartas", index=False)
        df_tfidf_paginas.to_excel(writer, sheet_name="tfidf_paginas", index=False)
        df_entidades_resumen_carta.to_excel(writer, sheet_name="entidades_carta", index=False)
        df_entidades_resumen_pagina.to_excel(writer, sheet_name="entidades_pagina", index=False)
        df_entidades.to_excel(writer, sheet_name="entidades_detalle", index=False)
        df_marcas_resumen_carta.to_excel(writer, sheet_name="marcas_carta", index=False)
        df_marcas_resumen_pagina.to_excel(writer, sheet_name="marcas_pagina", index=False)
        df_marcas.to_excel(writer, sheet_name="marcas_detalle", index=False)
        df_cat_cartas.to_excel(writer, sheet_name="categorias_carta", index=False)
        df_cat_paginas.to_excel(writer, sheet_name="categorias_pagina", index=False)

    print("\nProceso terminado.")
    print(f"Carpeta de salida: {CARPETA_SALIDA.resolve()}")
    print(f"Excel: {ARCHIVO_EXCEL.resolve()}")
    print(f"CSV: {CARPETA_CSV.resolve()}")
    print(f"Gráficos: {CARPETA_GRAFICOS.resolve()}")
    print(f"Extractos: {CARPETA_EXTRACTOS.resolve()}")


if __name__ == "__main__":
    main()
