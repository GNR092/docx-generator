---
name: docx-generator
description: Genera documentos .docx desde Markdown o texto plano sin dependencias externas.
---

# Skill: DOCX Generator

Genera documentos `.docx` simples desde texto plano o Markdown usando un script local.

## Script

- `scripts/docx_skill.py`

## Uso rapido

```bash
python scripts/docx_skill.py --output reporte.docx --title "Reporte tecnico" --input reporte.md
```

## Opciones

- `--output` (requerido): ruta destino del `.docx`
- `--title`: titulo inicial del documento
- `--input`: archivo fuente (`.md` o `.txt`)
- `--author`: autor en metadatos del documento
- `--subject`: asunto en metadatos del documento
- `--keywords`: palabras clave (separadas por comas)
- `--line`: linea inline; se puede repetir varias veces

## Ejemplos

```bash
python scripts/docx_skill.py --output nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
python scripts/docx_skill.py --output incidente.docx --title "Incidente" --input incidente.md
```

## Notas

- Soporta Markdown basico: negrita `**texto**`, cursiva `*texto*`, y codigo inline `` `codigo` ``.
- Soporta enlaces Markdown `[texto](https://url)` con hipervinculo real en Word.
- Soporta bloques: encabezados `#`, listas (`-`, `*`, `1.`) con numeracion Word nativa, citas (`>`) y bloques de codigo con triple acento grave.
- Soporta tablas Markdown con separador (`|---|---|`) convertidas a tablas DOCX.
- Convierte `--title` en encabezado principal automaticamente.
- Genera `styles.xml`, `numbering.xml`, y metadatos `docProps` (core/app) sin dependencias externas.
- Ideal para reportes tecnicos rapidos y evidencia de troubleshooting.
