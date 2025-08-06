# pdf-excel-word-python
# 🧾 Plantilla Proforma Excel con openpyxl (Python) 🚀

Este script permite **llenar una plantilla Excel con merges complejos**, tabla de items variable y totales, usando Python y [openpyxl](https://openpyxl.readthedocs.io/en/stable/).

## 👨‍💻 Características

- Soporta plantillas .xlsx con celdas combinadas en cabecera, tabla y leyendas.
- Inserta cualquier cantidad de items y mueve la fila de TOTAL dinámicamente.
- Respeta todos los merges de la plantilla, solo ajusta lo necesario.
- Copia formatos (estilo, bordes, alineación, etc.) de filas originales a las nuevas.
- **100% libre de errores “MergedCell is read-only”** o “los merges se dañan”.

## 📂 Estructura esperada

- `template.xlsx` → Tu plantilla base, con celdas combinadas.
- `datos.json` → El JSON de datos, que incluye la cabecera y los items.
- `proforma_final.xlsx` → El resultado generado.

## 📝 Uso

1. **Prepara tu plantilla** (`template.xlsx`) con los merges que quieras.  
   ⚠️ IMPORTANTE: Las filas de la tabla de items (A14:G?) y el TOTAL deben tener su propio formato y merges, pero **no deben tener merges verticales que bajen desde la tabla hacia las leyendas**.

2. Pon tu JSON de datos como `datos.json`.

3. Ejecuta el script:
   ```bash
   python excelservice.py
