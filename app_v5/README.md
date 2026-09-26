# Uniform Measurement Auto-Filler Web App

Local Flask application for filling school uniform measurement Excel templates from a reusable pattern workbook.

## New in this version
- Process multiple template `.xlsx` files in one submission.
- **Next file** workflow: after processing, click **Next file, keep same patterns** and upload another template without re-uploading the pattern workbook.
- Strict template header validation. If required header cells are missing or do not match the expected school template, the app returns **INVALID FILE** and does not generate a filled workbook for that file.
- Generated download name preserves the template filename and adds ` FILLED` before `.xlsx`.

## Run
```bash
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Template validation
The current validator expects the uploaded school workbook to contain the standard student-table columns in this order:
`S.No`, `Student's Name`, `Gender`, `EMIS Number`, `Class`, `Section`, followed by the appropriate Tamil measurement headings used by the Boys/Girls exports.
