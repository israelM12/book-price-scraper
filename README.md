# 🛒 Scraper de precios → Excel

Herramienta en **Python** que extrae el catálogo de productos de una tienda en línea, limpia los datos y los exporta a un **archivo Excel con formato, estadísticas y una gráfica**.

Es justo el tipo de tarea que piden muchos negocios: *"sácame los precios de la competencia y pásalos a una hoja de Excel ordenada"*. Este proyecto automatiza todo ese proceso de principio a fin.

> Como sitio de demostración se usa [books.toscrape.com](https://books.toscrape.com), una tienda pública creada específicamente para practicar web scraping. La misma estructura de código se adapta a cualquier catálogo.

---

## ✨ Qué hace

1. **Scraping con paginación** — recorre todas las páginas del catálogo automáticamente.
2. **Extracción de datos** — título, precio, calificación (rating), disponibilidad y URL de cada producto.
3. **Limpieza con pandas** — convierte el precio a número, normaliza el stock y ordena los resultados.
4. **Exportación a Excel** — genera un `.xlsx` con tres hojas:
   - **Productos**: la tabla completa con formato y moneda.
   - **Resumen**: métricas clave (precio promedio, mínimo, máximo, rating promedio...).
   - **Gráfica**: top 10 de productos más caros.

---

## 🚀 Uso

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar (extrae todo el catálogo)
python scraper.py

# Opciones
python scraper.py --pages 5                 # solo las primeras 5 páginas
python scraper.py --output mis_precios.xlsx # nombre de salida personalizado
```

---

## 📊 Ejemplo de salida

En la carpeta [`sample_output/`](sample_output/) hay un Excel de muestra generado por el script, con las tres hojas (productos, resumen y gráfica).

| titulo | precio | rating | en_stock |
|--------|-------:|:------:|:--------:|
| Our Band Could Be Your Life | £56.17 | 3 | ✅ |
| The Boys in the Boat | £52.92 | 3 | ✅ |
| Starving Hearts | £50.81 | 2 | ✅ |

---

## 🧰 Tecnologías

- **Python 3**
- **requests** — descarga de páginas con reintentos y cabeceras de navegador
- **BeautifulSoup** — parseo del HTML
- **pandas** — limpieza y transformación de datos
- **openpyxl** — generación del Excel con formato y gráficas

---

## 🧩 Estructura del código

El script está dividido en funciones claras y reutilizables:

- `fetch_page()` — descarga con reintentos ante errores de red.
- `parse_books()` — extrae los datos de cada producto de una página.
- `find_next_page()` — detecta la siguiente página del catálogo.
- `scrape_catalog()` — orquesta el recorrido completo.
- `build_dataframe()` — limpia y estructura los datos.
- `export_excel()` — crea el Excel con formato, resumen y gráfica.

---

## 📝 Notas

- El scraper incluye una pausa cortés entre páginas para no saturar el servidor.
- Hace scraping únicamente de un sitio público de práctica. Antes de usarlo con un sitio real, revisa siempre los *Términos de Servicio* y el archivo `robots.txt` del sitio.
