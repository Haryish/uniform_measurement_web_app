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
