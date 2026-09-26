# Uniform Measurement Auto-Filler v6

## Fix in v6
- Fixed the `name 'clean' is not defined` startup error.
- Header matching is tolerant of Excel line breaks, non-breaking spaces, repeated whitespace, underscores, hyphens, and punctuation.
- Tamil and English measurement headers are matched by name rather than column position.
- Boys/Girls detection uses the Gender column first, then the sheet name.
- Supports one or many school template workbooks with the same pattern workbook.
- Preserves the uploaded template filename and downloads it as `<original name> FILLED.xlsx`.

## Run
```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```
Then open http://127.0.0.1:5000


## Schema-flexible measurement fix
- Pattern sheets can contain different measurement schemas for different uniform types.
- Bilingual Tamil/English pattern headers are recognized correctly.
- Template validation no longer assumes every Girls file uses the skirt/waist schema.
- The app accepts additional measurement columns such as leg circumference, leg height, coat height, coat chest, and coat shoulder when present.
- Existing legacy positional pattern behavior is retained as a fallback, so previously working Boys/legacy templates continue to receive values.
- Tested against C.M.S. Seven Wells Girls 6-8, Girls 1-5, and Boys 1-8 templates.
