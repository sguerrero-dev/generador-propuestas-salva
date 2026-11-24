# Generador de propuestas solares · Salva Energy

Herramienta web que toma el Excel de preventa (plantilla de análisis económico) y genera en el navegador:

- **Presentación** "Análisis conceptual" (PDF o PowerPoint 16:9, 14 diapositivas con gráficas)
- **Cotización** (PDF A4, 3 páginas)

El Excel se procesa localmente en el navegador; no se sube a ningún servidor.

## Uso
1. Abre la página y arrastra el `.xlsx` de preventa.
2. Completa los datos del cliente (nombre, NIT, número de cotización, fecha…).
3. Descarga la presentación y la cotización.

## Requisitos de la plantilla
Lee posiciones fijas de las hojas `SIGO`, `Capital propio`, `Liquidación`, `Resumen`, `Datos2`, `Cálculos`, `Capex completo`, `A O&M`, `Prestamo 5 años` y `Producción`. No renombres hojas ni insertes/borres filas o columnas en ellas. Guarda el Excel después de recalcular.

## Archivos
- `index.html` — interfaz
- `extract.js` — lectura del Excel (mapa de celdas)
- `docs.js` — diapositivas y cotización
- `lib/` — SheetJS, Chart.js, html2canvas, jsPDF, PptxGenJS
- `img/` — logos y diapositivas fijas
