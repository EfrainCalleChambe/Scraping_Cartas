import requests
from bs4 import BeautifulSoup
import re
import os
import time
from math import ceil

BASE_URL = "https://montemar.library.illinois.edu"
OUTPUT_DIR = "cartas_montemar_columnas"

def get_soup(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return BeautifulSoup(response.text, 'html.parser')
    except Exception as e:
        print(f"Error: {e}")
        return None

def convert_supers(tag):
    sup_map = {
        'a':'ᵃ','b':'ᵇ','c':'ᶜ','d':'ᵈ','e':'ᵉ','f':'ᶠ','g':'ᵍ','h':'ʰ','i':'ⁱ','j':'ʲ',
        'k':'ᵏ','l':'ˡ','m':'ᵐ','n':'ⁿ','o':'ᵒ','p':'ᵖ','q':'ᑫ','r':'ʳ','s':'ˢ','t':'ᵗ',
        'u':'ᵘ','v':'ᵛ','w':'ʷ','x':'ˣ','y':'ʸ','z':'ᶻ',
        '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹'
    }
    text = tag.get_text()
    return ''.join(sup_map.get(ch, ch) for ch in text)

def html_to_plain_text(html):
    soup = BeautifulSoup(html, 'html.parser')
    for sup in soup.find_all('span', class_=re.compile(r'rend_superior')):
        sup.replace_with(convert_supers(sup))
    for br in soup.find_all('br'):
        br.replace_with('\n')
    text = soup.get_text()
    return text

def extract_pages(soup):
    page_spans = soup.find_all('span', class_=re.compile(r'\bpage\b'))
    if not page_spans:
        body = soup.find('body')
        if body:
            return [{'number': 1, 'text': html_to_plain_text(str(body))}]
        return []
    pages = []
    for idx, span in enumerate(page_spans, 1):
        inner_html = ''.join(str(child) for child in span.contents)
        text = html_to_plain_text(inner_html)
        pages.append({'number': idx, 'text': text})
    return pages

def split_into_two_columns(text):
    """Divide el texto en dos columnas (por mitad de líneas) y las une con tabulador."""
    lines = text.splitlines()
    if not lines:
        return ""
    mid = ceil(len(lines) / 2)
    left = lines[:mid]
    right = lines[mid:]
    # Igualar longitudes
    max_len = max(len(left), len(right))
    left += [''] * (max_len - len(left))
    right += [''] * (max_len - len(right))
    # Unir cada par con tabulador
    col_lines = [f"{l}\t{r}" for l, r in zip(left, right)]
    return "\n".join(col_lines)

def extract_letter_title(soup):
    title_tag = soup.find('span', class_='title')
    return html_to_plain_text(str(title_tag)) if title_tag else "Sin título"

def sanitize_filename(name, max_len=100):
    name = re.sub(r'[\\/*?:"<>|]', '', name)
    name = re.sub(r'[\r\n\t]+', ' ', name)
    name = name.strip()
    if len(name) > max_len:
        name = name[:max_len].rstrip()
    return name

def save_letter(title, pages, index):
    safe_title = sanitize_filename(title)
    filename = f"{index:03d}_{safe_title}.txt"
    filepath = os.path.join(OUTPUT_DIR, filename)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"Título: {title}\n")
        f.write(f"Total de páginas: {len(pages)}\n")
        f.write("=" * 80 + "\n\n")
        for page in pages:
            f.write(f"=== PÁGINA {page['number']} (dos columnas simuladas por líneas) ===\n")
            two_col_text = split_into_two_columns(page['text'])
            f.write(two_col_text)
            f.write("\n\n" + "-" * 80 + "\n\n")
    print(f"Guardado: {filepath}")

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print("Obteniendo página principal...")
    main_soup = get_soup(f"{BASE_URL}/Home/TheLetters")
    if not main_soup:
        return

    # Extraer enlaces de cartas
    letters = []
    for a in main_soup.find_all('a', href=re.compile(r'/Home/Letter/')):
        href = a.get('href')
        if href and 'pages=' in href:
            letters.append({
                'url': BASE_URL + href,
                'title': a.get_text(strip=True)
            })

    print(f"Encontradas {len(letters)} cartas.")
    for i, letter in enumerate(letters, 1):
        print(f"Procesando {i}/{len(letters)}: {letter['title'][:60]}...")
        letter_soup = get_soup(letter['url'])
        if not letter_soup:
            continue
        title = extract_letter_title(letter_soup)
        pages = extract_pages(letter_soup)
        if pages:
            save_letter(title, pages, i)
        else:
            print(f"  Sin contenido en {letter['url']}")
        time.sleep(1)

    print("Extracción completada.")

if __name__ == "__main__":
    main()