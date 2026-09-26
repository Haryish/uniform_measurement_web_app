
# Uniform Measurement Auto-Filler

A local Flask web app for filling school uniform measurement Excel templates from a reusable pattern workbook.

## What it does

1. Upload a pattern workbook.
2. Upload the school's Excel template/export file.
3. Detect student rows, gender, class and measurement columns.
4. Choose an exact class/gender pattern where available.
5. Assign patterns round-robin or randomly.
6. Fill blank measurement cells while preserving the uploaded template.
7. Automatically normalize odd measurements to the next even number.
8. Reject/skip invalid 0 or 3-digit values.
9. Download the completed Excel file.

It also has a separate "Generate additional patterns" option. This should be used only with approved/appropriate base patterns because generated values are synthetic variants, not real student measurements.

## Run on Windows

Install Python 3.10+.

Open Command Prompt in this folder:

    python -m venv .venv
    .venv\Scripts\activate
    pip install -r requirements.txt
    python app.py

Then open:

    http://127.0.0.1:5000

## Input pattern format

Each pattern sheet should contain a heading named `Pattern` (or the bilingual Tamil/English equivalent) followed by numeric measurement columns.

Sheet names such as:

- Boys - Class 1
- Boys - Class 2
- Girls - Class 1
- Girls - Class 5-7

are used to identify gender and class.

## Template format

The supplied school templates work with the app because they have:

- S.No
- Student's Name
- Gender
- EMIS Number
- Class
- Section

followed by measurement columns.

The app preserves the workbook and writes values into those measurement columns.

## Important

For official student records, use actual/approved measurements or an approved pattern allocation. The generator is intended for creating pattern sets, not fabricating real-world student measurements.
