# Plan de Embarque Walmart

Sube el CSV que descargas del portal de proveedores de Walmart y la app
genera el Excel completo, calcado de tu archivo `envios_WM.xlsx`:

- **Hoja con el CSV crudo** (nombre = nombre del archivo) — respaldo tal
  cual llega del portal.
- **Detalle Envíos** — hoja maestra.
- **Resumen por Tienda** — piezas totales, estilos y EANs por tienda.
- **Lineal x Tienda** — una fila por CAJA física (agrupada por tienda +
  estilo + color), con hasta N posiciones (EAN/Talla/Color/Cantidad) y la
  columna "EAN / Cantidad" ya combinada, más "Número de Cajas x tienda"
  (ej. "2 de 3").
- **Una hoja por cada PO** — mismo formato que Lineal x Tienda, filtrada a
  ese pedido, con la columna extra "Núm. EANs x Caja".

También genera los **QR para CEDIS**: un QR por caja (mismo agrupado que
"Lineal x Tienda"), con el mismo contenido que la columna "EAN / Cantidad"
del Excel.

## Validado contra tu archivo real

Esta versión ya se ajustó comparando directo contra tu `envios_WM.xlsx`:
mismos nombres de hoja y columna, mismo agrupado por caja, mismo separador
de línea en "EAN / Cantidad", encabezado azul `#1F4E78` con Calibri 11
negrita (igual que tu archivo). Si al probarla con otro CSV real algo no
cuadra, dime exactamente qué ajustar.

## Desplegar (mismos pasos que la app de etiquetas)

1. Sube estos 3 archivos (`app.py`, `requirements.txt`, `LEEME.md`) a un
   repositorio nuevo en GitHub (puede ser el mismo tipo de proceso que ya
   hiciste: crear repo → Add file → Upload files → Commit).
2. En share.streamlit.io → Create app → Deploy a public app from GitHub.
3. Repository: `tu-usuario/nombre-del-repo`, Branch: `main`,
   Main file path: `app.py` → Deploy.

## Notas

- Sin base de datos: cada CSV se procesa en memoria y se descarta al cerrar
  la pestaña.
- El CEDIS se extrae tomando los primeros dígitos del campo "Cadena"
  (ej. "7464 NVA WAL-MART DE..." → CEDIS 7464). Si algún cliente/CEDIS no
  sigue ese patrón, avísame para ajustar la lógica.
