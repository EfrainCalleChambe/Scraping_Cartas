# -*- coding: utf-8 -*-
"""
DESCARGAR TRANSCRIPCIONES DIPLOMÁTICAS - CARTAS DE MONTEMAR
VERSIÓN FINAL: HTML LIMPIO + SUSTITUCIONES CORREGIDAS

Sitio:
    https://montemar.library.illinois.edu/Home/TheLetters

Esta versión corrige el problema de espacios exagerados:
    - NO usa white-space: pre-wrap.
    - El navegador colapsa espacios como en la web original.
    - No aplica CSS original al HTML exportado para evitar superposición.
    - Mantiene etiquetas diplomáticas:
        del, add, subst, rend_superior, gap, quote, etc.
    - Muestra la carta en una columna limpia.
    - Permite ajustar el ancho de lectura con ANCHO_TEXTO_PX.

Requisitos:
    pip install requests beautifulsoup4 pandas lxml openpyxl

Ejecución:
    python descargar_montemar_transcripciones_final.py
"""

import re
import time
import html
from pathlib import Path
from urllib.parse import urljoin, urlparse, parse_qs

import requests
import pandas as pd
from bs4 import BeautifulSoup, NavigableString, Tag


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_URL = "https://montemar.library.illinois.edu"
LIST_URL = "https://montemar.library.illinois.edu/Home/TheLetters"

OUT_DIR = Path("montemar_transcripciones")
OUT_HTML = OUT_DIR / "html_por_pagina"
OUT_TXT = OUT_DIR / "txt_por_pagina"
OUT_CSS = OUT_DIR / "css_original"

SLEEP_SECONDS = 0.7

# Ajusta este valor si quieres líneas más largas o más cortas.
# 760 se parece más a una columna de lectura.
# 900 o 1000 da líneas más largas.
ANCHO_TEXTO_PX = 860

# El CSS original se descarga como respaldo, pero NO se aplica,
# porque genera superposición en algunos casos.
DESCARGAR_CSS_ORIGINAL = True

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (compatible; MontemarTranscriptionDownloader/1.0; "
        "+research-use)"
    )
}


# ============================================================
# CSS LIMPIO
# ============================================================

def construir_css_limpio() -> str:
    return f"""
html,
body {{
    background: #ffffff;
    color: #111111;
}}

body {{
    font-family: Georgia, "Times New Roman", serif;
    font-size: 25px;
    line-height: 1.65;
    margin: 24px;
}}

.export-wrapper {{
    max-width: {ANCHO_TEXTO_PX}px;
}}

.export-header {{
    font-family: Arial, sans-serif;
    font-size: 14px;
    line-height: 1.4;
    margin-bottom: 24px;
    padding-bottom: 12px;
    border-bottom: 1px solid #ddd;
}}

.export-header h1 {{
    font-family: Arial, sans-serif;
    font-size: 20px;
    line-height: 1.3;
    margin: 0 0 8px 0;
}}

.export-header a {{
    color: #0645ad;
}}

.body {{
    display: block;
    max-width: {ANCHO_TEXTO_PX}px;
}}

.page {{
    display: block;
    max-width: {ANCHO_TEXTO_PX}px;

    /*
       MUY IMPORTANTE:
       white-space: normal evita los espacios enormes causados
       por los saltos e indentaciones internas del HTML.
    */
    white-space: normal;

    overflow-wrap: normal;
    word-break: normal;
}}

.export-page-block {{
    display: block;
    margin-top: 28px;
    padding-top: 14px;
    border-top: 1px solid #ddd;
}}

.export-page-title {{
    font-family: Arial, sans-serif;
    font-size: 18px;
    margin: 0 0 12px 0;
}}

/* No queremos columnas forzadas */
.columns2,
.columns1,
.current-page,
.hidden-page {{
    display: block !important;
    column-count: 1 !important;
    columns: auto !important;
    visibility: visible !important;
}}

/*
   En el HTML original, cada párrafo es un span.
   Se deja inline para que fluya como en la web.
*/
.p,
.continuation,
.opener,
.closer,
.dateline,
.quote {{
    display: inline;
}}

/* Separación suave entre párrafos. */
.p::after,
.continuation::after {{
    content: " ";
}}

/* Notas, firmas y posiciones */
.note.place_right {{
    display: inline;
    margin-right: 0.4em;
}}

.note.place_bottom_left {{
    display: block;
    margin-top: 1em;
}}

.signed.rend_right,
.rend_right {{
    display: block;
    text-align: right;
}}

.rend_left {{
    text-align: left;
}}

/* Tachados */
.del {{
    text-decoration: line-through;
    color: #111111;
}}

/* Añadidos */
.add {{
    color: #b00000;
}}

/* Superíndices */
.g.rend_superior {{
    vertical-align: super;
    font-size: 60%;
    line-height: 0;
}}

/* Añadido interlineal */
.add.place_interlinear {{
    color: #b00000;
    vertical-align: super;
    font-size: 70%;
    line-height: 0;
}}

/*
   Sustituciones corregidas:
   "caso que" queda en la línea normal y ocupa su espacio.
   "cuando" se dibuja encima sin empujar el texto siguiente.
*/
.subst {{
    display: inline-block;
    position: relative;
    padding-top: 0.65em;
    margin-left: 0.05em;
    margin-right: 0.05em;
    vertical-align: baseline;
    line-height: 1;
}}

.subst > .del {{
    display: inline;
    text-decoration: line-through;
    color: #111111;
    white-space: nowrap;
}}

.subst > .add.place_above {{
    position: absolute;
    left: 0;
    top: -0.05em;
    color: #b00000;
    font-size: 70%;
    line-height: 1;
    white-space: nowrap;
    text-decoration: none;
}}

/* Otros añadidos por ubicación */
.add.place_above {{
    color: #b00000;
    vertical-align: super;
    font-size: 70%;
    line-height: 0;
}}

.add.place_below {{
    color: #b00000;
    vertical-align: sub;
    font-size: 70%;
    line-height: 0;
}}

/* Lagunas */
.gap {{
    display: inline-block;
    min-width: 3ch;
    border-bottom: 1px dotted #444;
}}

/* Subrayado punteado */
.hi.rend_dashed_underline {{
    border-bottom: 1px dashed #111;
}}

/* Oculta el visor de imagen si apareciera accidentalmente */
#openseadragon,
.openseadragon-container,
.openseadragon-canvas,
canvas {{
    display: none !important;
}}
"""


# ============================================================
# AUXILIARES
# ============================================================

def limpiar_nombre_archivo(s: str, max_len: int = 90) -> str:
    s = html.unescape(str(s))
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    s = re.sub(r"[\u0000-\u001f]+", "", s)
    return s[:max_len].strip(" ._")


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    r = session.get(url, headers=HEADERS, timeout=60)
    r.raise_for_status()
    r.encoding = r.apparent_encoding or "utf-8"
    return BeautifulSoup(r.text, "lxml")


def normalizar_espacios(txt: str) -> str:
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r" *\n *", "\n", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()


def texto_hijos_con_marcas(node: Tag) -> str:
    return "".join(extraer_texto_con_marcas(c) for c in node.children)


def obtener_lugar_add(add_tag: Tag) -> str:
    if add_tag is None:
        return ""

    add_classes = set(add_tag.get("class", []))

    lugares = [
        "place_above",
        "place_interlinear",
        "place_below",
        "place_margin",
        "place_left",
        "place_right",
        "place_top",
        "place_bottom",
        "place_bottom_left",
        "place_bottom_right",
    ]

    encontrados = [c for c in lugares if c in add_classes]
    return " {" + ",".join(encontrados) + "}" if encontrados else ""


# ============================================================
# CSS ORIGINAL SOLO COMO RESPALDO
# ============================================================

def corregir_urls_css(css_text: str, css_url: str) -> str:
    def repl(match):
        raw = match.group(1).strip().strip('"').strip("'")
        if raw.startswith(("data:", "http://", "https://", "#")):
            return f"url({raw})"
        return f"url({urljoin(css_url, raw)})"

    return re.sub(r"url\(([^)]+)\)", repl, css_text)


def obtener_css_original(session: requests.Session, soup: BeautifulSoup, page_url: str) -> str:
    css_total = []
    vistos = set()

    OUT_CSS.mkdir(parents=True, exist_ok=True)

    if not DESCARGAR_CSS_ORIGINAL:
        return ""

    for link in soup.find_all("link", href=True):
        rel_val = link.get("rel", [])
        rel = rel_val.lower() if isinstance(rel_val, str) else " ".join(rel_val).lower()
        href = link["href"]

        if "stylesheet" not in rel and not href.lower().endswith(".css"):
            continue

        css_url = urljoin(page_url, href)

        if css_url in vistos:
            continue

        vistos.add(css_url)

        try:
            r = session.get(css_url, headers=HEADERS, timeout=60)
            r.raise_for_status()
            r.encoding = r.apparent_encoding or "utf-8"

            css_text = corregir_urls_css(r.text, css_url)

            safe_css_name = limpiar_nombre_archivo(
                urlparse(css_url).path.strip("/").replace("/", "_"),
                max_len=120
            )
            if not safe_css_name.endswith(".css"):
                safe_css_name += ".css"

            local_css_path = OUT_CSS / safe_css_name
            local_css_path.write_text(css_text, encoding="utf-8")

            css_total.append(f"\n/* CSS original: {css_url} */\n")
            css_total.append(css_text)

        except Exception as e:
            print(f"  ADVERTENCIA: no se pudo descargar CSS {css_url}: {e}")

    return "\n".join(css_total)


# ============================================================
# DETECCIÓN DE CARTAS
# ============================================================

def obtener_enlaces_cartas(session: requests.Session):
    soup = get_soup(session, LIST_URL)

    cartas = []
    vistos = set()

    for a in soup.find_all("a", href=True):
        href = a["href"]

        if "/Home/Letter/" not in href:
            continue

        url = urljoin(BASE_URL, href)

        if url in vistos:
            continue

        vistos.add(url)

        titulo = " ".join(a.get_text(" ", strip=True).split())

        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        f_id = qs.get("f", [""])[0]

        if not f_id:
            f_id = parsed.path.rstrip("/").split("/")[-1]

        cartas.append({
            "nro_carta": len(cartas) + 1,
            "titulo": titulo,
            "url": url,
            "f_id": f_id
        })

    return cartas


# ============================================================
# HTML A TXT
# ============================================================

def extraer_texto_con_marcas(node) -> str:
    if isinstance(node, NavigableString):
        return str(node)

    if not isinstance(node, Tag):
        return ""

    classes = set(node.get("class", []))

    if "subst" in classes:
        del_tag = node.find(class_="del")
        add_tag = node.find(class_="add")

        texto_del = texto_hijos_con_marcas(del_tag).strip() if del_tag else ""
        texto_add = texto_hijos_con_marcas(add_tag).strip() if add_tag else ""
        lugar_add = obtener_lugar_add(add_tag) if add_tag else ""

        if texto_del and texto_add:
            return f"[SUST: ~~{texto_del}~~ → {texto_add}{lugar_add}]"

        if texto_del:
            return f"[SUST_DEL: ~~{texto_del}~~]"

        if texto_add:
            return f"[SUST_ADD: {texto_add}{lugar_add}]"

        return f"[SUST: {texto_hijos_con_marcas(node).strip()}]"

    if "del" in classes:
        return f"[DEL: {texto_hijos_con_marcas(node).strip()}]"

    if "add" in classes:
        contenido = texto_hijos_con_marcas(node).strip()
        lugar_add = obtener_lugar_add(node)
        return f"[ADD: {contenido}{lugar_add}]"

    if node.name == "span" and "g" in classes and "rend_superior" in classes:
        return f"^{{{texto_hijos_con_marcas(node)}}}"

    if "gap" in classes:
        reason = ""
        for c in classes:
            if c.startswith("reason_"):
                reason = c.replace("reason_", "")

        style = node.get("style", "")
        width = ""
        m = re.search(r"width\s*:?\s*([^;]+)", style)
        if m:
            width = f",width={m.group(1).strip()}"

        return f"[GAP:{reason or 'gap'}{width}]"

    if node.name == "br":
        return "\n"

    if node.name in {"p", "div"}:
        return texto_hijos_con_marcas(node) + "\n"

    return texto_hijos_con_marcas(node)


# ============================================================
# EXTRACCIÓN DE PÁGINAS
# ============================================================

def extraer_paginas_carta(session: requests.Session, carta: dict):
    soup = get_soup(session, carta["url"])
    obtener_css_original(session, soup, carta["url"])

    titulo_h1 = soup.find(["h1", "h2"])
    titulo_pagina = (
        " ".join(titulo_h1.get_text(" ", strip=True).split())
        if titulo_h1 else carta["titulo"]
    )

    body = soup.select_one("span.body")

    if body is None:
        print(f"ADVERTENCIA: no se encontró span.body en {carta['url']}")
        return []

    paginas = []
    page_tags = body.select('span.page[id^="page-"]')

    for page_tag in page_tags:
        page_id = page_tag.get("id", "")
        m = re.search(r"page-(\d+)", page_id)
        page_num = int(m.group(1)) if m else len(paginas) + 1

        inner_html = page_tag.decode_contents()

        txt_marcas = extraer_texto_con_marcas(page_tag)
        txt_marcas = normalizar_espacios(txt_marcas)

        paginas.append({
            "nro_carta": carta["nro_carta"],
            "f_id": carta["f_id"],
            "titulo": carta["titulo"],
            "titulo_pagina": titulo_pagina,
            "url": carta["url"],
            "page_id": page_id,
            "page_num": page_num,
            "html": inner_html,
            "txt": txt_marcas
        })

    return paginas


# ============================================================
# ARMADO HTML
# ============================================================

def construir_html_pagina(titulo: str, url: str, page_id: str, page_num: int,
                         html_transcripcion: str) -> str:
    css = construir_css_limpio()

    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>{html.escape(titulo)} - Page {page_num}</title>
<style>
{css}
</style>
</head>
<body>
<div class="export-wrapper">

<div class="export-header">
<h1>{html.escape(titulo)}</h1>
<div><strong>Fuente:</strong> <a href="{html.escape(url)}">{html.escape(url)}</a></div>
<div><strong>Page:</strong> {page_num}</div>
</div>

<span class="body">
<span id="{html.escape(page_id)}" class="page columns2 current-page">
{html_transcripcion}
</span>
</span>

</div>
</body>
</html>
"""


def construir_html_carta_completa(titulo: str, url: str, html_paginas: list) -> str:
    css = construir_css_limpio()

    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>{html.escape(titulo)}</title>
<style>
{css}
</style>
</head>
<body>
<div class="export-wrapper">

<div class="export-header">
<h1>{html.escape(titulo)}</h1>
<div><strong>Fuente:</strong> <a href="{html.escape(url)}">{html.escape(url)}</a></div>
</div>

{chr(10).join(html_paginas)}

</div>
</body>
</html>
"""


# ============================================================
# GUARDADO
# ============================================================

def guardar_carta(carta: dict, paginas: list):
    safe_id = limpiar_nombre_archivo(carta["f_id"] or f"carta_{carta['nro_carta']:02d}", 60)
    safe_title = limpiar_nombre_archivo(carta["titulo"], 70)

    base_name = f"{carta['nro_carta']:02d}_{safe_id}_{safe_title}" if safe_title else f"{carta['nro_carta']:02d}_{safe_id}"

    registros = []
    html_completo = []
    txt_completo = []

    for p in paginas:
        page_num = p["page_num"]

        html_name = f"{base_name}_page_{page_num:03d}.html"
        txt_name = f"{base_name}_page_{page_num:03d}.txt"

        html_path = OUT_HTML / html_name
        txt_path = OUT_TXT / txt_name

        html_doc = construir_html_pagina(
            titulo=p["titulo"],
            url=p["url"],
            page_id=p["page_id"],
            page_num=page_num,
            html_transcripcion=p["html"]
        )

        html_path.write_text(html_doc, encoding="utf-8")
        txt_path.write_text(p["txt"], encoding="utf-8")

        html_completo.append(
            f"<div class='export-page-block'>\n"
            f"<div class='export-page-title'>Page {page_num}</div>\n"
            f"<span id='{html.escape(p['page_id'])}' class='page columns2'>\n"
            f"{p['html']}\n"
            f"</span>\n"
            f"</div>"
        )

        txt_completo.append(f"=== PAGE {page_num} ===\n{p['txt']}")

        registros.append({
            "nro_carta": p["nro_carta"],
            "f_id": p["f_id"],
            "titulo": p["titulo"],
            "url": p["url"],
            "page_num": page_num,
            "page_id": p["page_id"],
            "archivo_html": str(html_path),
            "archivo_txt": str(txt_path)
        })

    carta_html_path = OUT_DIR / f"{base_name}_COMPLETA.html"
    carta_txt_path = OUT_DIR / f"{base_name}_COMPLETA.txt"

    carta_html_doc = construir_html_carta_completa(
        titulo=carta["titulo"],
        url=carta["url"],
        html_paginas=html_completo
    )

    carta_html_path.write_text(carta_html_doc, encoding="utf-8")
    carta_txt_path.write_text("\n\n".join(txt_completo), encoding="utf-8")

    return registros


# ============================================================
# MAIN
# ============================================================

def main():
    OUT_HTML.mkdir(parents=True, exist_ok=True)
    OUT_TXT.mkdir(parents=True, exist_ok=True)
    OUT_CSS.mkdir(parents=True, exist_ok=True)

    session = requests.Session()

    print("Leyendo lista de cartas...")
    cartas = obtener_enlaces_cartas(session)
    print(f"Cartas encontradas: {len(cartas)}")

    todos_registros = []

    for carta in cartas:
        print(f"Descargando carta {carta['nro_carta']:02d}: {carta['titulo']}")

        try:
            paginas = extraer_paginas_carta(session, carta)
            registros = guardar_carta(carta, paginas)
            todos_registros.extend(registros)
            print(f"  páginas extraídas: {len(paginas)}")

        except Exception as e:
            print(f"  ERROR en {carta['url']}: {e}")

        time.sleep(SLEEP_SECONDS)

    df = pd.DataFrame(todos_registros)

    csv_path = OUT_DIR / "indice_transcripciones_montemar.csv"
    xlsx_path = OUT_DIR / "indice_transcripciones_montemar.xlsx"

    if not df.empty:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        df.to_excel(xlsx_path, index=False)

    print("\nProceso terminado.")
    print(f"Carpeta de salida: {OUT_DIR.resolve()}")

    if not df.empty:
        print(f"Índice CSV: {csv_path.resolve()}")
        print(f"Índice Excel: {xlsx_path.resolve()}")
    else:
        print("No se generó índice porque no se extrajeron páginas.")


if __name__ == "__main__":
    main()
