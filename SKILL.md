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

- El documento es formato Word OpenXML minimo (sin estilos avanzados).
- Ideal para reportes tecnicos rapidos y evidencia de troubleshooting.
