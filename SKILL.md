---
name: docx-generator
description: Genera documentos .docx y .pdf desde Markdown o texto plano sin dependencias externas (excepto reportlab para PDF).
---

# Skill: DOCX/PDF Generator

Genera documentos `.docx` o `.pdf` desde texto plano o Markdown usando un script local.

## Script

- `scripts/docx_skill.py`

## Uso rapido

```bash
python scripts/docx_skill.py --output reporte.docx --title "Reporte tecnico" --input reporte.md
python scripts/docx_skill.py --output reporte.pdf --title "Reporte tecnico" --input reporte.md
```

## Opciones

- `--output` (requerido): ruta destino del archivo de salida (`.docx` o `.pdf`)
- `--format`: formato de salida (`docx` o `pdf`). Por defecto se infiere de la extension del archivo.
- `--title`: titulo inicial del documento
- `--input`: archivo fuente (`.md` o `.txt`)
- `--author`: autor en metadatos del documento
- `--lang`: idioma de ortografia del documento (por defecto `es-ES`; usa `en-US` para ingles)
- `--subject`: asunto en metadatos del documento
- `--keywords`: palabras clave (separadas por comas)
- `--line`: linea inline; se puede repetir varias veces

## Ejemplos

```bash
python scripts/docx_skill.py --output nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
python scripts/docx_skill.py --output incidente.docx --title "Incidente" --input incidente.md
python scripts/docx_skill.py --output reporte.pdf --title "Reporte" --input reporte.md
```

## Notas

- Soporta Markdown basico: negrita `**texto**`, cursiva `*texto*`, y codigo inline `` `codigo` ``.
- Soporta enlaces Markdown `[texto](https://url)` con hipervinculo real en Word (solo DOCX).
- Soporta bloques: encabezados `#`, listas (`-`, `*`, `1.`) con numeracion nativa, citas (`>`) y bloques de codigo con triple acento grave.
- Soporta tablas Markdown con separador (`|---|---|`) convertidas a tablas nativas.
- Convierte `--title` en encabezado principal automaticamente.
- Para generar PDF se requiere `pip install reportlab` (solo Python, sin LibreOffice). Si no esta instalada, el script lo indicara antes de abortar.

  ```bash
  pip install reportlab
  ```
- Genera `styles.xml`, `numbering.xml`, y metadatos `docProps` (core/app) sin dependencias externas.
- Ideal para reportes tecnicos rapidos y evidencia de troubleshooting.