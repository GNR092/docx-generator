# DOCX/PDF Generator Skill

Generador de archivos `.docx` y `.pdf` desde texto plano o Markdown, sin dependencias externas (excepto reportlab para PDF).

## Que hace

- Convierte contenido `.md` o `.txt` a Word OpenXML (`.docx`) o PDF.
- Soporta formato inline: negrita, cursiva, codigo inline y enlaces.
- Soporta bloques: encabezados, listas, citas, bloques de codigo y tablas Markdown.
- Genera componentes DOCX necesarios (`document.xml`, `styles.xml`, `numbering.xml`, `docProps`).

## Requisitos

- Python 3.9+ (recomendado 3.10+).
- Para DOCX: No requiere librerias externas (usa solo libreria estandar).
- Para PDF: Requiere `pip install reportlab` (sin LibreOffice, 100% Python puro).

  ```bash
  pip install reportlab
  ```

  Si no esta instalada, el script lo indicara antes de abortar.

## Script principal

- `scripts/docx_skill.py` (unificado para DOCX y PDF)

## Uso rapido

```bash
python scripts/docx_skill.py --output reporte.docx --title "Reporte tecnico" --input reporte.md
python scripts/docx_skill.py --output reporte.pdf --title "Reporte tecnico" --input reporte.md
```

## Opciones CLI

- `--output` (requerido): ruta del archivo de salida (`.docx` o `.pdf`).
- `--format`: formato de salida (`docx` o `pdf`). Por defecto se infiere de la extension.
- `--title`: titulo del documento (se agrega como encabezado `#`).
- `--input`: archivo fuente (`.md` o `.txt`).
- `--line`: linea inline (se puede repetir varias veces).
- `--author`: autor para metadatos del documento.
- `--subject`: asunto para metadatos del documento.
- `--keywords`: palabras clave (texto libre, por ejemplo separadas por comas).

## Ejemplos

### 1) DOCX desde Markdown

```bash
python scripts/docx_skill.py --output .docs/incidente.docx --title "Incidente" --input incidente.md
```

### 2) PDF desde Markdown

```bash
python scripts/docx_skill.py --output reporte.pdf --title "Reporte" --input reporte.md
```

### 3) Solo lineas inline

```bash
python scripts/docx_skill.py --output .docs/nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
```

### 4) Con metadatos

```bash
python scripts/docx_skill.py \
  --output .docs/reporte.pdf \
  --title "Reporte semanal" \
  --input reporte.md \
  --author "Tu Nombre" \
  --subject "Estado del proyecto" \
  --keywords "reporte,estado,proyecto"
```

## Markdown soportado

### Inline

- `**negrita**`
- `*cursiva*`
- `` `codigo` ``
- `[texto](https://ejemplo.com)` y `mailto:` (solo en DOCX)

### Bloques

- Encabezados: `#` a `######`
- Listas no ordenadas: `-` o `*`
- Listas ordenadas: `1.`, `2.`, etc.
- Citas: `>`
- Bloques de codigo: triple acento grave
- Tablas Markdown con separador `|---|---|`

## Como funciona internamente

### DOCX
1. Lee lineas desde `--input` y/o `--line`.
2. Parsea Markdown inline y de bloques usando el modulo `parsers.py`.
3. Genera XML OpenXML para:
   - `word/document.xml`
   - `word/styles.xml`
   - `word/numbering.xml`
   - `docProps/core.xml` y `docProps/app.xml`
   - relaciones (`_rels/.rels` y `word/_rels/document.xml.rels`)
4. Empaqueta todo en un `.docx` (zip con estructura OpenXML).

### PDF
1. Lee lineas desde `--input` y/o `--line`.
2. Usa el mismo parser de `parsers.py` para obtener bloques.
3. Genera elementos ReportLab (Paragraph, Table, ListFlowable, etc.).
4. Construye el PDF con `SimpleDocTemplate`.

## Notas

- Si no pasas contenido, genera un documento minimo con titulo/fecha por defecto.
- El parser esta orientado a Markdown practico (no pretende cubrir el 100% de CommonMark).
- La carpeta `.docs/` es ideal para salidas locales de prueba.
- Para PDF, el script verifica que Python y reportlab esten instalados antes de proceder.