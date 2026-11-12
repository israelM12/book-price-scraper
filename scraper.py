"""
Scraper de precios de productos -> Excel.

Extrae el catalogo completo de https://books.toscrape.com (una tienda
publica hecha para practicar scraping), limpia los datos con pandas y los
exporta a un archivo Excel con formato y una grafica de precios.

Es un ejemplo del tipo de trabajo que piden muchos clientes: "sacame los
precios de la competencia y pasalos a Excel". El sitio objetivo se puede
cambiar; la estructura del codigo sirve para cualquier catalogo.

Uso:
    python scraper.py                 # extrae todas las paginas
    python scraper.py --pages 5       # solo las primeras 5 paginas
    python scraper.py --output mis_precios.xlsx

Autor: Israel Martinez
"""

import argparse
import sys
import time
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
import pandas as pd


# --- Configuracion -----------------------------------------------------------

BASE_URL = "https://books.toscrape.com/"
# Mapa para convertir el rating en palabra ("Three") a numero (3).
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
# Un User-Agent de navegador real evita que algunos sitios bloqueen el bot.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}
REQUEST_TIMEOUT = 15      # segundos antes de rendirse con una pagina
PAUSE_BETWEEN_PAGES = 0.5  # cortesia: no saturar el servidor


# --- Descarga ----------------------------------------------------------------

def fetch_page(session, url, retries=3):
    """Descarga una pagina con reintentos. Devuelve el HTML o None."""
    for intento in range(1, retries + 1):
        try:
            resp = session.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            return resp.text
        except requests.RequestException as e:
            print(f"  ! Error en {url} (intento {intento}/{retries}): {e}")
            time.sleep(1.5 * intento)  # espera creciente antes de reintentar
    return None


# --- Parseo ------------------------------------------------------------------

def parse_books(html, page_url):
    """Extrae la lista de libros de una pagina de catalogo."""
    soup = BeautifulSoup(html, "html.parser")
    libros = []

    for card in soup.select("article.product_pod"):
        # Titulo: esta en el atributo 'title' del enlace (el texto se corta).
        link = card.select_one("h3 a")
        titulo = link["title"].strip() if link else ""

        # Precio: viene como "£51.77"; lo guardamos crudo y tambien limpio.
        precio_raw = card.select_one(".price_color")
        precio_texto = precio_raw.get_text(strip=True) if precio_raw else ""

        # Rating: esta como clase CSS, p. ej. "star-rating Three".
        estrellas_tag = card.select_one("p.star-rating")
        rating = 0
        if estrellas_tag:
            for cls in estrellas_tag.get("class", []):
                if cls in RATING_MAP:
                    rating = RATING_MAP[cls]
                    break

        # Disponibilidad.
        disp = card.select_one(".availability")
        disponibilidad = disp.get_text(strip=True) if disp else ""

        # URL del detalle del producto (por si se quiere profundizar luego).
        detalle = urljoin(page_url, link["href"]) if link else ""

        libros.append(
            {
                "titulo": titulo,
                "precio_texto": precio_texto,
                "rating": rating,
                "disponibilidad": disponibilidad,
                "url": detalle,
            }
        )

    return libros


def find_next_page(html, page_url):
    """Devuelve la URL de la siguiente pagina, o None si es la ultima."""
    soup = BeautifulSoup(html, "html.parser")
    next_link = soup.select_one("li.next a")
    if next_link:
        return urljoin(page_url, next_link["href"])
    return None


# --- Scraping completo -------------------------------------------------------

def scrape_catalog(start_url=BASE_URL, max_pages=None):
    """Recorre el catalogo pagina por pagina y junta todos los libros."""
    session = requests.Session()
    session.headers.update(HEADERS)

    todos = []
    url = start_url
    pagina = 1

    while url:
        print(f"-> Pagina {pagina}: {url}")
        html = fetch_page(session, url)
        if html is None:
            print("  ! No se pudo descargar; se detiene aqui.")
            break

        libros = parse_books(html, url)
        todos.extend(libros)
        print(f"   {len(libros)} productos encontrados (total: {len(todos)})")

        if max_pages and pagina >= max_pages:
            break

        url = find_next_page(html, url)
        pagina += 1
        time.sleep(PAUSE_BETWEEN_PAGES)

    return todos


# --- Limpieza ----------------------------------------------------------------

def build_dataframe(libros):
    """Convierte la lista de dicts en un DataFrame limpio."""
    df = pd.DataFrame(libros)
    if df.empty:
        return df

    # Precio: quitar el simbolo de moneda y pasar a numero decimal.
    df["precio"] = (
        df["precio_texto"]
        .str.replace(r"[^\d.]", "", regex=True)  # deja solo digitos y punto
        .astype(float)
    )

    # Normalizar disponibilidad a un booleano claro.
    df["en_stock"] = df["disponibilidad"].str.contains("stock", case=False)

    # Ordenar columnas de forma legible y quitar el precio crudo.
    df = df[["titulo", "precio", "rating", "en_stock", "url"]]
    df = df.sort_values("precio", ascending=False).reset_index(drop=True)
    return df


# --- Export a Excel ----------------------------------------------------------

def export_excel(df, path):
    """Guarda el DataFrame en Excel con formato y una grafica de precios."""
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, PatternFill, Alignment

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Productos")

        wb = writer.book
        ws = writer.sheets["Productos"]

        # Encabezados en negrita con fondo de color.
        encabezado_fill = PatternFill("solid", fgColor="2F5496")
        encabezado_font = Font(color="FFFFFF", bold=True)
        for col in range(1, df.shape[1] + 1):
            celda = ws.cell(row=1, column=col)
            celda.fill = encabezado_fill
            celda.font = encabezado_font
            celda.alignment = Alignment(horizontal="center")

        # Ancho de columnas automatico (aproximado).
        anchos = {"A": 45, "B": 12, "C": 10, "D": 12, "E": 50}
        for col, ancho in anchos.items():
            ws.column_dimensions[col].width = ancho

        # Formato de moneda en la columna de precio (B).
        for row in range(2, df.shape[0] + 2):
            ws.cell(row=row, column=2).number_format = '"£"#,##0.00'

        # --- Hoja de resumen con estadisticas ---
        resumen = wb.create_sheet("Resumen")
        stats = [
            ("Total de productos", len(df)),
            ("Precio promedio", round(df["precio"].mean(), 2)),
            ("Precio minimo", round(df["precio"].min(), 2)),
            ("Precio maximo", round(df["precio"].max(), 2)),
            ("En stock", int(df["en_stock"].sum())),
            ("Rating promedio", round(df["rating"].mean(), 2)),
        ]
        resumen["A1"] = "Metrica"
        resumen["B1"] = "Valor"
        resumen["A1"].font = resumen["B1"].font = Font(bold=True)
        for i, (nombre, valor) in enumerate(stats, start=2):
            resumen.cell(row=i, column=1, value=nombre)
            resumen.cell(row=i, column=2, value=valor)
        resumen.column_dimensions["A"].width = 22
        resumen.column_dimensions["B"].width = 14

        # --- Grafica: top 10 productos mas caros ---
        top = df.head(10)
        grafica_ws = wb.create_sheet("Grafica")
        grafica_ws["A1"] = "Producto"
        grafica_ws["B1"] = "Precio"
        for i, (_, fila) in enumerate(top.iterrows(), start=2):
            # Titulo recortado para que quepa en la etiqueta.
            grafica_ws.cell(row=i, column=1, value=fila["titulo"][:30])
            grafica_ws.cell(row=i, column=2, value=fila["precio"])

        chart = BarChart()
        chart.title = "Top 10 productos mas caros"
        chart.type = "bar"
        chart.y_axis.title = "Precio (£)"
        datos = Reference(grafica_ws, min_col=2, min_row=1,
                          max_row=1 + len(top))
        categorias = Reference(grafica_ws, min_col=1, min_row=2,
                               max_row=1 + len(top))
        chart.add_data(datos, titles_from_data=True)
        chart.set_categories(categorias)
        chart.height = 10
        chart.width = 22
        grafica_ws.add_chart(chart, "D2")

    print(f"-> Excel generado: {path}")


# --- Main --------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Scraper de precios -> Excel con formato."
    )
    parser.add_argument("--pages", type=int, default=None,
                        help="Numero maximo de paginas a extraer (default: todas).")
    parser.add_argument("--output", default="precios.xlsx",
                        help="Nombre del archivo Excel de salida.")
    args = parser.parse_args()

    print("=== Scraper de precios ===\n")
    libros = scrape_catalog(max_pages=args.pages)

    if not libros:
        print("No se extrajo ningun producto. Revisa tu conexion o el sitio.")
        sys.exit(1)

    df = build_dataframe(libros)
    print(f"\nTotal extraido: {len(df)} productos")
    print(df.head(5).to_string(index=False))

    export_excel(df, args.output)
    print("\nListo. Abre el Excel para ver los datos, el resumen y la grafica.")


if __name__ == "__main__":
    main()
