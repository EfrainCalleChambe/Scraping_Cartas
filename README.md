# Scraping de Cartas Montemar

Proyecto en Python para descargar cartas del sitio Montemar Library y generar un analisis de frecuencias y co-ocurrencias de palabras en archivos Excel.

## Estructura

- `extraer_cartas.py`: descarga las cartas desde `https://montemar.library.illinois.edu` y guarda archivos `.txt` en `cartas_montemar_columnas/`.
- `analizar_cartas.py`: lee los `.txt` generados, extrae emisor/receptor, calcula frecuencias y asociaciones de palabras, y guarda resultados en `analisis_cartas_excel/`.
- `requirements.txt`: dependencias base del proyecto.

## Requisitos

- Python 3.10 o superior.
- Dependencias de Python:
  - `pandas`
  - `openpyxl`
  - `requests`
  - `beautifulsoup4`

## Instalacion

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install requests beautifulsoup4
```

## Uso

Primero descarga las cartas:

```powershell
python extraer_cartas.py
```

Luego genera el analisis en Excel:

```powershell
python analizar_cartas.py
```

El archivo consolidado se guarda en:

```text
analisis_cartas_excel/analisis_cartas.xlsx
```

## Notas

Las carpetas de salida y archivos comprimidos se consideran generados localmente, por lo que estan excluidos del control de versiones mediante `.gitignore`.
