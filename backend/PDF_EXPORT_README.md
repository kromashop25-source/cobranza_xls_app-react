# PDF Export Quick Guide

1. Levanta la API (por ejemplo `python run_app.py` o el ejecutable CobranzaApp).
2. Llama a `GET /pdf/debug-blocks?file_path=RUTA&hoja=HOJA` para confirmar los bloques por vendedor.
3. Ejecuta `POST /pdf/export` con un JSON como `{ "file_path": "E:\\RUTA\\COBRANZA.xlsx", "hoja_base": "OFICINA (VES)", "carpeta_salida": null }`. Los PDFs quedan en `PDFS` junto al Excel.

Notas rapidas: las hojas SUR y NORTE se exportan completas; el log vive en `pdf_export.log`.
