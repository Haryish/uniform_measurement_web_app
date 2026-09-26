from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path
from io import BytesIO
import random
import re
import tempfile
import os
import uuid
import shutil

app = Flask(__name__)
app.secret_key = "uniform-measurement-local-app"

UPLOAD_DIR = Path(tempfile.gettempdir()) / "uniform_measurement_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Expected school template headers. The app intentionally validates these before processing.
COMMON_HEADERS = ["s.no", "student's name", "gender", "emis number", "class", "section"]
BOYS_MEASUREMENT_HEADERS = ["தோள் பட்டை", "உயரம்", "மார்பு சுற்றளவு", "கை நீளம்", "கை சுற்றளவு", "இடுப்பு சுற்றளவு", "கால் உயரம்", "தொடை சுற்றளவு"]
GIRLS_MEASUREMENT_HEADERS = ["தோள் பட்டை", "உயரம்", "மார்பு சுற்றளவு", "கை நீளம்", "கை சுற்றளவு", "பாவாடை உயரம்", "இடுப்பு சுற்றளவு"]


def clean(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def job_dir():
    jid = session.get("job_id")
    if not jid:
        jid = uuid.uuid4().hex
        session["job_id"] = jid
    p = UPLOAD_DIR / jid
    p.mkdir(parents=True, exist_ok=True)
    return p


def parse_group_from_sheet(title):
    t = clean(title)
    gender = "Boys" if "boy" in t or "male" in t else "Girls" if "girl" in t or "female" in t else None
    m = re.search(r"class\s*([0-9]+)(?:\s*-\s*([0-9]+))?", t)
    if not m:
        m = re.search(r"([0-9]+)\s*-\s*([0-9]+)", t)
    if m:
        c1 = int(m.group(1)); c2 = int(m.group(2) or c1)
        classes = list(range(c1, c2 + 1))
    else:
        classes = []
    return gender, classes


def find_pattern_header(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [clean(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if any(v == "pattern" or v.startswith("pattern") or "வடிவம்" in v for v in vals):
            return r
    return None


def read_patterns(fileobj):
    wb = load_workbook(fileobj, data_only=True)
    patterns = []
    for ws in wb.worksheets:
        header_row = find_pattern_header(ws)
        if not header_row:
            continue
        headers = [clean(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
        pcol = next((i + 1 for i, h in enumerate(headers) if h == "pattern" or h.startswith("pattern") or "வடிவம்" in h), 1)
        gender, classes = parse_group_from_sheet(ws.title)
        if not gender:
            sample = str(ws.cell(header_row + 1, pcol).value or "")
            gender = "Girls" if sample.upper().startswith("G") else "Boys"
        if not classes:
            for rr in range(header_row + 1, min(ws.max_row, header_row + 4) + 1):
                s = str(ws.cell(rr, pcol).value or "")
                nums = re.findall(r"\d+", s)
                if nums:
                    classes = [int(nums[0])]
                    break
        for r in range(header_row + 1, ws.max_row + 1):
            pid = ws.cell(r, pcol).value
            if pid in (None, ""):
                continue
            vals = []
            for c in range(pcol + 1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    v = int(v)
                    if v <= 0 or v >= 100 or v % 2:
                        v = v + 1 if 0 < v < 99 and v % 2 else None
                    if v is not None:
                        vals.append(v)
            if vals:
                patterns.append({"id": str(pid), "gender": gender, "classes": classes, "values": vals})
    return patterns


def even_value(v):
    if v is None or v == "":
        return None
    try:
        x = int(float(v))
    except Exception:
        return None
    if x <= 0 or x >= 100:
        return None
    return x if x % 2 == 0 else x + 1


def generate_patterns(base_patterns, variants=3, seed=42):
    rng = random.Random(seed)
    out = []
    seen = set()
    for p in base_patterns:
        vals = [even_value(v) for v in p["values"]]
        out.append(dict(p, values=vals))
        seen.add((p["gender"], tuple(p["classes"]), tuple(vals)))
    for p in base_patterns:
        for k in range(variants):
            vals = []
            for v in p["values"]:
                delta = rng.choice([-4, -2, 0, 2, 4])
                nv = max(10, min(98, v + delta))
                if nv % 2:
                    nv += 1
                vals.append(nv)
            key = (p["gender"], tuple(p["classes"]), tuple(vals))
            if key in seen:
                continue
            seen.add(key)
            out.append({"id": f'{p["id"]}-V{k + 1}', "gender": p["gender"], "classes": p["classes"], "values": vals})
    return out


def patterns_to_workbook(patterns):
    wb = Workbook()
    wb.remove(wb.active)
    groups = {}
    for p in patterns:
        for cls in p["classes"]:
            groups.setdefault((p["gender"], cls), []).append(p)
    for (gender, cls), plist in sorted(groups.items(), key=lambda x: (x[0][0], x[0][1])):
        ws = wb.create_sheet(f"{gender} - Class {cls}")
        maxn = max(len(p["values"]) for p in plist)
        headers = ["Pattern"] + [f"Measurement {i}" for i in range(1, maxn + 1)]
        ws.append(headers)
        for p in plist:
            ws.append([p["id"]] + p["values"])
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[chr(64 + col) if col <= 26 else "A"].width = 16
    return wb


def normalize_header(v):
    return clean(v).replace("\n", " ")


def detect_sheet_gender(ws, header_row):
    """Infer gender from the student rows, falling back to the sheet title."""
    gender_values = []
    for r in range(header_row + 1, ws.max_row + 1):
        v = clean(ws.cell(r, 3).value)
        if v:
            gender_values.append(v)
    if any(v in ("female", "girl", "girls", "பெண்", "மாணவி", "மாணவிகள்") or "female" in v or "girl" in v for v in gender_values):
        return "girls"
    if any(v in ("male", "boy", "boys", "ஆண்", "மாணவன்", "மாணவர்கள்") or "male" in v or "boy" in v for v in gender_values):
        return "boys"
    title = clean(ws.title)
    if "girl" in title or "female" in title or "பெண்" in title:
        return "girls"
    if "boy" in title or "male" in title or "ஆண்" in title:
        return "boys"
    return None


def validate_template_workbook(wb):
    """Validate the common student headers and infer boys/girls from the sheet data, not the sheet name."""
    errors = []
    valid_sheets = 0
    for ws in wb.worksheets:
        header_row = None
        for r in range(1, min(ws.max_row, 20) + 1):
            vals = [normalize_header(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
            if vals[:6] == COMMON_HEADERS:
                header_row = r
                break
        if not header_row:
            errors.append(f"Sheet '{ws.title}': required header row not found or does not match.")
            continue

        headers = [normalize_header(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)]
        gender_hint = detect_sheet_gender(ws, header_row)
        if not gender_hint:
            errors.append(f"Sheet '{ws.title}': could not determine whether this is a boys or girls template from the Gender column.")
            continue

        expected_measurements = GIRLS_MEASUREMENT_HEADERS if gender_hint == "girls" else BOYS_MEASUREMENT_HEADERS
        actual_measurements = headers[6:6 + len(expected_measurements)]
        if actual_measurements != expected_measurements:
            errors.append(
                f"Sheet '{ws.title}': measurement headers do not match the expected {gender_hint} template. "
                f"Expected: {', '.join(expected_measurements)}"
            )
            continue
        valid_sheets += 1

    if not wb.worksheets:
        errors.append("Workbook contains no worksheets.")
    if errors:
        return False, errors
    if valid_sheets == 0:
        return False, ["No valid student measurement sheet was found."]
    return True, []

def find_template_header(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [clean(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if vals[:6] == COMMON_HEADERS:
            return r
    return None


def template_info(ws):
    hr = find_template_header(ws)
    if not hr:
        raise ValueError(f"Could not find the valid student table header in sheet '{ws.title}'.")
    headers = [clean(ws.cell(hr, c).value) for c in range(1, ws.max_column + 1)]
    gender_hint = detect_sheet_gender(ws, hr)
    if gender_hint == "girls":
        expected = GIRLS_MEASUREMENT_HEADERS
    elif gender_hint == "boys":
        expected = BOYS_MEASUREMENT_HEADERS
    else:
        raise ValueError(f"Could not determine gender/template type in sheet '{ws.title}'.")
    if headers[6:6 + len(expected)] != expected:
        raise ValueError(f"Invalid measurement headers in sheet '{ws.title}'.")
    measurement_cols = list(range(7, 7 + len(expected)))
    return hr, measurement_cols, 5, 3, 2, gender_hint

def choose_patterns(all_patterns, gender, cls, allow_fallback=True):
    exact = [p for p in all_patterns if p["gender"] == gender and cls in p["classes"]]
    if exact:
        return exact, False
    if allow_fallback:
        same = [p for p in all_patterns if p["gender"] == gender]
        if same:
            def dist(p):
                return min(abs(cls - c) for c in p["classes"])
            d = min(dist(p) for p in same)
            return [p for p in same if dist(p) == d], True
    return [], False


def fill_template(template_file, patterns, mode="round_robin", seed=42, allow_fallback=True, overwrite=False):
    wb = load_workbook(template_file)
    ok, errors = validate_template_workbook(wb)
    if not ok:
        raise ValueError("INVALID FILE: " + " ".join(errors))
    rng = random.Random(seed)
    report = []
    for ws in wb.worksheets:
        hr, mcols, class_col, gender_col, name_col, sheet_gender = template_info(ws)
        counters = {}
        for r in range(hr + 1, ws.max_row + 1):
            name = ws.cell(r, name_col).value
            if not name:
                continue
            raw_gender = str(ws.cell(r, gender_col).value or "")
            gender = "Girls" if "female" in raw_gender.lower() or "girl" in raw_gender.lower() else "Boys"
            try:
                cls = int(str(ws.cell(r, class_col).value).strip())
            except Exception:
                report.append(("WARNING", ws.title, f"Row {r}: invalid class"))
                continue
            plist, fallback = choose_patterns(patterns, gender, cls, allow_fallback)
            if not plist:
                report.append(("WARNING", ws.title, f"Row {r} ({name}): no pattern for {gender} class {cls}"))
                continue
            key = (gender, cls)
            idx = counters.get(key, 0)
            p = rng.choice(plist) if mode == "random" else plist[idx % len(plist)]
            counters[key] = idx + 1
            vals = p["values"]
            if len(vals) < len(mcols):
                report.append(("WARNING", ws.title, f"Row {r} ({name}): pattern {p['id']} has {len(vals)} values but template expects {len(mcols)}"))
            for i, c in enumerate(mcols):
                if i >= len(vals):
                    break
                if overwrite or ws.cell(r, c).value in (None, ""):
                    ws.cell(r, c).value = even_value(vals[i])
            report.append(("FILLED", ws.title, f"Row {r}: {name} <- {p['id']}" + (" (fallback)" if fallback else "")))
    return wb, report


def save_pattern_to_job(pattern_file):
    p = job_dir() / "patterns.xlsx"
    pattern_file.save(p)
    return p


def load_job_patterns():
    p = job_dir() / "patterns.xlsx"
    if not p.exists():
        return None
    with p.open("rb") as f:
        patterns = read_patterns(f)
    if not patterns:
        raise ValueError("No pattern sheets were detected. A sheet must contain a 'Pattern' heading and numeric measurement columns.")
    return patterns


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        pattern_file = request.files.get("pattern_file")
        template_files = request.files.getlist("template_files")
        if not template_files or all(not f.filename for f in template_files):
            single = request.files.get("template_file")
            if single and single.filename:
                template_files = [single]
        mode = request.form.get("mode", "round_robin")
        seed = int(request.form.get("seed", "42") or 42)
        allow_fallback = request.form.get("fallback") == "on"
        overwrite = request.form.get("overwrite") == "on"
        if pattern_file and pattern_file.filename:
            try:
                patterns = read_patterns(pattern_file)
                if not patterns:
                    raise ValueError("No pattern sheets were detected. A sheet must contain a 'Pattern' heading and numeric measurement columns.")
                save_pattern_to_job(pattern_file)
            except Exception as e:
                flash(str(e))
                return redirect(url_for("index"))
        else:
            try:
                patterns = load_job_patterns()
                if not patterns:
                    flash("Please upload the pattern workbook first.")
                    return redirect(url_for("index"))
            except Exception as e:
                flash(str(e))
                return redirect(url_for("index"))
        if not template_files:
            flash("Please select at least one school template workbook.")
            return redirect(url_for("index"))
        results = []
        for template_file in template_files:
            try:
                # Validate before doing any processing, so an invalid workbook is never silently accepted.
                raw = template_file.read()
                template_file.stream.seek(0)
                wb_check = load_workbook(BytesIO(raw), data_only=True)
                ok, errors = validate_template_workbook(wb_check)
                if not ok:
                    raise ValueError("INVALID FILE: " + " ".join(errors))
                wb, report = fill_template(BytesIO(raw), patterns, mode, seed, allow_fallback, overwrite)
                out_path = job_dir() / f"{uuid.uuid4().hex}.xlsx"
                wb.save(out_path)
                original_name = secure_filename(template_file.filename or "template.xlsx")
                template_path = Path(original_name)
                filled_name = f"{template_path.stem} FILLED{template_path.suffix or '.xlsx'}"
                results.append({
                    "name": filled_name,
                    "path": str(out_path),
                    "filled": sum(1 for x in report if x[0] == "FILLED"),
                    "warnings": sum(1 for x in report if x[0] == "WARNING"),
                    "report": report[-30:],
                })
            except Exception as e:
                results.append({"name": template_file.filename, "error": str(e)})
        session["last_results"] = results
        return render_template("result.html", results=results, pattern_saved=True)
    return render_template("index.html", pattern_exists=(job_dir() / "patterns.xlsx").exists())


@app.route("/next", methods=["GET"])
def next_file():
    try:
        patterns = load_job_patterns()
        if not patterns:
            flash("Pattern workbook is not available. Please start a new session and upload it.")
    except Exception as e:
        flash(str(e))
    return render_template("next.html")


@app.route("/download/<path:filename>")
def download_result(filename):
    path = job_dir() / filename
    if not path.exists():
        flash("Generated file is no longer available.")
        return redirect(url_for("index"))
    name = request.args.get("name") or path.name
    return send_file(path, as_attachment=True, download_name=name)


@app.route("/generate-patterns", methods=["POST"])
def generate():
    pattern_file = request.files.get("pattern_file")
    variants = int(request.form.get("variants", "3") or 3)
    seed = int(request.form.get("seed", "42") or 42)
    if not pattern_file:
        flash("Upload an existing pattern workbook first.")
        return redirect(url_for("index"))
    try:
        base = read_patterns(pattern_file)
        if not base:
            raise ValueError("No patterns found.")
        generated = generate_patterns(base, variants, seed)
        wb = patterns_to_workbook(generated)
        out = BytesIO(); wb.save(out); out.seek(0)
        return send_file(out, as_attachment=True, download_name="generated_patterns.xlsx")
    except Exception as e:
        flash(str(e)); return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
