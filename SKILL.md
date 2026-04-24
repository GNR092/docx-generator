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
- `--line`: linea inline; se puede repetir varias veces

## Ejemplos

```bash
python scripts/docx_skill.py --output nota.docx --title "Nota" --line "Linea 1" --line "Linea 2"
python scripts/docx_skill.py --output incidente.docx --title "Incidente" --input incidente.md
```

## Notas

- Soporta Markdown basico: negrita `**texto**`, cursiva `*texto*`, y codigo inline `` `codigo` ``.
- Soporta bloques: encabezados `#`, listas (`-`, `*`, `1.`), citas (`>`) y bloques de codigo con triple acento grave.
- Convierte `--title` en encabezado principal automaticamente.
- El documento sigue siendo Word OpenXML minimo, sin dependencias externas.
- Ideal para reportes tecnicos rapidos y evidencia de troubleshooting.
