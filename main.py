#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SETCODE.py — Local file cleaner / column exporter

Supports TXT, CSV, LOG, XLS, and XLSX files.
- Upload a file in your browser
- Preview detected columns/rows
- Choose columns manually and their output order
- Optionally trim whitespace, remove blank rows, and deduplicate rows
- Download a plain-text export using | as the separator

This is a general-purpose data formatter. It does not scan message text for
verification codes or automatically pair phone numbers with codes.
"""

from __future__ import annotations

import csv
import io
import os
import re
import tempfile
from pathlib import Path
from typing import Any

from flask import Flask, request, render_template_string, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("SETCODE_SECRET", "change-this-local-secret")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB
ALLOWED_EXTENSIONS = {".txt", ".csv", ".log", ".xls", ".xlsx"}

PAGE = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SETCODE — File Cleaner</title>
<style>
:root { color-scheme: dark; }
* { box-sizing: border-box; }
body { margin:0; background:#0b1020; color:#eef2ff; font:16px/1.5 system-ui,Arial,sans-serif; }
.wrap { max-width:1000px; margin:0 auto; padding:24px 16px 48px; }
header { padding:24px; border:1px solid #29334f; background:#121a30; border-radius:18px; }
h1 { margin:0 0 6px; font-size:clamp(28px,5vw,42px); }
.sub { color:#aeb9d6; margin:0; }
.card { margin-top:18px; padding:22px; border:1px solid #29334f; background:#121a30; border-radius:16px; }
label { display:block; font-weight:650; margin:12px 0 6px; }
input[type=file], input[type=text], select { width:100%; padding:12px; border:1px solid #3a4666; border-radius:10px; background:#0b1020; color:#eef2ff; }
button, .btn { display:inline-block; padding:12px 18px; border:0; border-radius:10px; background:#7c9cff; color:#071127; font-weight:800; cursor:pointer; text-decoration:none; margin-top:14px; }
button:hover, .btn:hover { filter:brightness(1.08); }
.note { color:#aeb9d6; font-size:14px; }
.alert { padding:12px 14px; border-radius:10px; background:#48252d; margin:12px 0; }
.grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:14px; }
.checks { display:flex; gap:18px; flex-wrap:wrap; margin-top:14px; }
.checks label { display:flex; gap:8px; align-items:center; font-weight:500; margin:0; }
.tablewrap { overflow:auto; border:1px solid #29334f; border-radius:10px; margin-top:12px; max-height:380px; }
table { border-collapse:collapse; width:100%; min-width:560px; font-size:13px; }
th,td { border-bottom:1px solid #29334f; padding:9px; text-align:left; max-width:260px; overflow-wrap:anywhere; }
th { position:sticky; top:0; background:#1a2540; }
.small { font-size:13px; color:#aeb9d6; }
</style>
</head>
<body><main class="wrap">
<header>
<h1>SETCODE</h1>
<p class="sub">Upload a file, preview its columns, clean the rows, and export selected fields to TXT.</p>
</header>
{% with messages = get_flashed_messages() %}
{% if messages %}{% for message in messages %}<div class="alert">{{ message }}</div>{% endfor %}{% endif %}
{% endwith %}

{% if not data %}
<section class="card">
<h2>1. Upload file</h2>
<form method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
<label for="file">Choose TXT / CSV / LOG / XLS / XLSX</label>
<input id="file" type="file" name="file" accept=".txt,.csv,.log,.xls,.xlsx" required>
<p class="note">Files are processed locally by this app on your computer. Maximum file size: 25 MB.</p>
<button type="submit">Upload and Preview</button>
</form>
</section>
{% else %}
<section class="card">
<h2>2. Preview</h2>
<p class="note">File: {{ filename }} · Rows detected: {{ row_count }}</p>
<div class="tablewrap"><table>
<thead><tr>{% for col in columns %}<th>{{ col }}</th>{% endfor %}</tr></thead>
<tbody>{% for row in preview %}<tr>{% for cell in row %}<td>{{ cell }}</td>{% endfor %}</tr>{% endfor %}</tbody>
</table></div>
</section>
<section class="card">
<h2>3. Choose fields to export</h2>
<form method="post" action="{{ url_for('export') }}">
<input type="hidden" name="token" value="{{ token }}">
<div class="grid">
<div>
<label for="first">Output field 1 (required)</label>
<select id="first" name="first" required>
{% for col in columns %}<option value="{{ loop.index0 }}">{{ col }}</option>{% endfor %}
</select>
</div>
<div>
<label for="second">Output field 2 (optional)</label>
<select id="second" name="second">
<option value="">Do not include a second field</option>
{% for col in columns %}<option value="{{ loop.index0 }}">{{ col }}</option>{% endfor %}
</select>
</div>
<div>
<label for="third">Output field 3 (optional)</label>
<select id="third" name="third">
<option value="">Do not include a third field</option>
{% for col in columns %}<option value="{{ loop.index0 }}">{{ col }}</option>{% endfor %}
</select>
</div>
<div>
<label for="separator">Output separator</label>
<select id="separator" name="separator">
<option value="|">Pipe |</option><option value=",">Comma ,</option><option value="tab">Tab</option>
</select>
</div>
</div>
<div class="checks">
<label><input type="checkbox" name="trim" checked> Trim spaces</label>
<label><input type="checkbox" name="skip_blank" checked> Skip blank output rows</label>
<label><input type="checkbox" name="dedupe"> Remove duplicate output rows</label>
</div>
<button type="submit">Create TXT file</button>
<a class="btn" href="{{ url_for('reset') }}">Start over</a>
</form>
<p class="small">Choose fields manually. This tool does not search message text for OTPs or automatically link codes to phone numbers.</p>
</section>
{% endif %}
</main></body></html>
"""

# A single active upload is held in memory; no uploaded file is saved to disk.
STORE: dict[str, dict[str, Any]] = {}


def allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    # Avoid Excel values such as 123.0 when they represent whole numbers.
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_delimited(raw: bytes, ext: str) -> tuple[list[str], list[list[str]]]:
    text = raw.decode("utf-8-sig", errors="replace")
    lines = text.splitlines()
    if not lines:
        return ["Column 1"], []
    sample = "\n".join(lines[:20])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        rows = list(csv.reader(io.StringIO(text), dialect))
    except csv.Error:
        # TXT/LOG fallback: one non-empty line per row; split common delimiters.
        rows = []
        for line in lines:
            if not line.strip():
                rows.append([line])
                continue
            if "|" in line:
                rows.append(line.split("|"))
            elif "\t" in line:
                rows.append(line.split("\t"))
            elif "," in line:
                rows.append(next(csv.reader([line])))
            else:
                rows.append([line])
    width = max((len(row) for row in rows), default=1)
    if width == 0:
        width = 1
    headers = [f"Column {i + 1}" for i in range(width)]
    # If the first row looks like a header, use it as column labels.
    if rows and any(str(v).strip() for v in rows[0]):
        first = [str(v).strip() for v in rows[0]]
        if len(first) == width and len(set(first)) == len(first) and all(first):
            headers = first
            rows = rows[1:]
    normalized = []
    for row in rows:
        row = [cell_text(v) for v in row]
        normalized.append(row + [""] * (width - len(row)))
    return headers, normalized


def read_excel(raw: bytes, ext: str) -> tuple[list[str], list[list[str]]]:
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("Excel support requires pandas and openpyxl/xlrd. Install requirements.txt.") from exc
    engine = "openpyxl" if ext == ".xlsx" else "xlrd"
    try:
        frame = pd.read_excel(io.BytesIO(raw), header=None, dtype=object, engine=engine)
    except Exception as exc:
        raise RuntimeError(f"Could not read this Excel file: {exc}") from exc
    if frame.empty:
        return ["Column 1"], []
    values = frame.where(frame.notna(), "").values.tolist()
    width = max(len(r) for r in values)
    first = [cell_text(v).strip() for v in values[0]]
    headers = [f"Column {i + 1}" for i in range(width)]
    if first and all(first) and len(set(first)) == len(first):
        headers = first
        values = values[1:]
    rows = []
    for row in values:
        normalized = [cell_text(v) for v in row]
        rows.append(normalized + [""] * (width - len(normalized)))
    return headers, rows


@app.get("/")
def index():
    return render_template_string(PAGE, data=None)


@app.post("/upload")
def upload():
    file = request.files.get("file")
    if not file or not file.filename:
        flash("Please choose a file first.")
        return redirect(url_for("index"))
    filename = secure_filename(file.filename)
    ext = Path(filename).suffix.lower()
    if not allowed(filename):
        flash("Unsupported file type. Use TXT, CSV, LOG, XLS, or XLSX.")
        return redirect(url_for("index"))
    raw = file.read()
    if not raw:
        flash("The selected file is empty.")
        return redirect(url_for("index"))
    try:
        if ext in {".xls", ".xlsx"}:
            columns, rows = read_excel(raw, ext)
        else:
            columns, rows = read_delimited(raw, ext)
    except Exception as exc:
        flash(str(exc))
        return redirect(url_for("index"))
    token = os.urandom(16).hex()
    STORE.clear()
    STORE[token] = {"filename": filename, "columns": columns, "rows": rows}
    return render_template_string(
        PAGE, data=True, filename=filename, columns=columns,
        preview=rows[:50], row_count=len(rows), token=token
    )


@app.post("/export")
def export():
    token = request.form.get("token", "")
    item = STORE.get(token)
    if not item:
        flash("Upload session expired. Please upload the file again.")
        return redirect(url_for("index"))
    columns, rows = item["columns"], item["rows"]
    chosen = []
    for name in ("first", "second", "third"):
        val = request.form.get(name, "")
        if val != "":
            try:
                idx = int(val)
                if idx < 0 or idx >= len(columns):
                    raise ValueError
                chosen.append(idx)
            except ValueError:
                flash("Invalid column selection.")
                return redirect(url_for("index"))
    if not chosen:
        flash("Select at least one output field.")
        return redirect(url_for("index"))
    sep_value = request.form.get("separator", "|")
    sep = "\t" if sep_value == "tab" else sep_value
    trim = request.form.get("trim") == "on"
    skip_blank = request.form.get("skip_blank") == "on"
    dedupe = request.form.get("dedupe") == "on"
    output_lines = []
    seen = set()
    for row in rows:
        fields = [row[i] if i < len(row) else "" for i in chosen]
        if trim:
            fields = [field.strip() for field in fields]
        if skip_blank and not any(fields):
            continue
        line = sep.join(fields)
        if dedupe:
            if line in seen:
                continue
            seen.add(line)
        output_lines.append(line)
    payload = ("\n".join(output_lines) + ("\n" if output_lines else "")).encode("utf-8")
    output = io.BytesIO(payload)
    output.seek(0)
    base = Path(item["filename"]).stem or "cleaned"
    return send_file(output, mimetype="text/plain; charset=utf-8",
                     as_attachment=True, download_name=f"{base}_cleaned.txt")


@app.get("/reset")
def reset():
    STORE.clear()
    return redirect(url_for("index"))


@app.errorhandler(413)
def too_large(_error):
    return "File too large. Maximum upload size is 25 MB.", 413


if __name__ == "__main__":
    print("\nSETCODE File Cleaner")
    print("Open this address in your browser: http://127.0.0.1:5000")
    print("Stop the server with Ctrl+C.\n")
    app.run(host="127.0.0.1", port=26283, debug=False)
