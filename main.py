"""
============================================================
 Flask Script: NUMBER|CODE Extractor — Clean TXT Download
============================================================
 Supported files : .txt, .csv, .xlsx (Excel) — sab
 Data format     : NUMBER|CODE  (har line me, ek ke niche ek)
 Output (UI)     : Series-wise tables (5/6/7/8 digit...)
 Download        : SIRF NUMBER|CODE wali plain .txt file
                   (koi header / faltu data nahi)
                   Order: chhote digit series pehle, har series
                   me numbers chhote → bade
 Port            : 26283

 Chalane ke liye:
   pip install flask openpyxl
   python app.py
   browser me: http://127.0.0.1:26283
============================================================
"""

import io
import re
import uuid
import csv
from collections import defaultdict

from flask import Flask, request, render_template_string, send_file, abort

# Excel . (optional dependency)
try:
    from openpyxl import load_workbook
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

app = Flask(__name__)
results_store = {}


# ---------------- Code extraction ----------------
def extract_code(raw):
    """Field se OTP code nikalta hai (NUMBER|CODE format)."""
    if not raw:
        return ""
    raw = raw.strip()

    # 1) Agar already pure digits hain -> direct use karo
    if raw.isdigit():
        return raw

    # 2) Old format support (Arabic message): "الرمز الخاص بك هو ... #880421"
    m = re.search(r"هو[\u200F\u200E\s]*(\d+)", raw)
    if m:
        return m.group(1)

    m = re.search(r"#(\d+)", raw)
    if m:
        return m.group(1)

    # 3) Fallback: 4-8 digit ka group (typical OTP length)
    groups = re.findall(r"\d+", raw)
    for g in groups:
        if 4 <= len(g) <= 8:
            return g
    return groups[-1] if groups else ""


# ---------------- File parsing ----------------
def parse_text(text):
    """TXT/CSV content parse kare — NUMBER|CODE (ya old CSV format)."""
    rows = []
    lines = text.splitlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Primary: pipe-separated NUMBER|CODE
        parts = line.split("|")
        if len(parts) >= 2:
            num_raw, code_raw = parts[0], parts[1]
        else:
            # Fallback: purana CSV format (date, operator, number, name, msg, ...)
            fields = next(csv.reader([line]), [])
            if len(fields) >= 5:
                num_raw, code_raw = fields[2], fields[4]
            elif len(fields) == 2:
                num_raw, code_raw = fields
            else:
                continue

        number = re.sub(r"\D", "", num_raw)      # sirf digits
        code = extract_code(code_raw)
        if number and code and code.isdigit():
            rows.append((number, code))

    return rows


def parse_excel(raw):
    """Excel (.xlsx) file parse kare — pehle do columns = NUMBER, CODE."""
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    rows = []

    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = [str(v) for v in row if v is not None]
            if len(vals) < 2:
                continue
            number = re.sub(r"\D", "", vals[0])
            code = extract_code(vals[1])
            if number and code and code.isdigit():
                rows.append((number, code))

    return rows


def group_by_series(rows):
    """
    Codes ko digit-count ke hisab se series me group karta hai.
    Returns: list of (length, [(number, code), ...]) — har series me
             numbers ascending (chhote → bade) sorted.
    """
    series = defaultdict(list)
    for number, code in rows:
        series[len(code)].append((number, code))

    ordered = []
    for length in sorted(series.keys()):
        items = sorted(series[length], key=lambda x: int(x[0]))
        ordered.append((length, items))
    return ordered


# ---------------- HTML Template (Tailwind) ----------------
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NUMBER|CODE Series Extractor</title>
<script src="https://unpkg.com/@tailwindcss/browser@4"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style> body { font-family: 'Inter', system-ui, sans-serif; } </style>
</head>
<body class="bg-gradient-to-br from-slate-100 via-blue-50 to-indigo-100 min-h-screen flex items-start justify-center p-6">

  <div class="w-full max-w-4xl bg-white rounded-2xl shadow-xl ring-1 ring-slate-200 overflow-hidden">

    <!-- Header -->
    <div class="bg-gradient-to-r from-indigo-600 to-blue-600 px-6 py-5 text-white">
      <h1 class="text-2xl font-extrabold tracking-tight">🔢 NUMBER|CODE Series Extractor</h1>
      <p class="text-indigo-100 text-sm mt-1">
        TXT / CSV / Excel — koi bhi file upload karein. Download me SIRF NUMBER|CODE milega (koi faltu data nahi).
      </p>
    </div>

    <!-- Info / Note (upar) -->
    <div class="px-6 pt-5">
      <div class="flex items-start gap-3 bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
        <span class="text-xl">💡</span>
        <p class="text-sm text-amber-800">
          <b>Note:</b> कुछ भी change / update करना हो तो बताएं — मैं script को आपके हिसाब से adjust कर दूंगा!
          (Jaise output format, grouping rule, ya koi aur file type)
        </p>
      </div>
    </div>

    <!-- Upload card (sabse upar) -->
    <div class="px-6 py-5">
      <form method="post" enctype="multipart/form-data" class="flex flex-col sm:flex-row gap-3 items-stretch sm:items-center">
        <label class="flex-1">
          <input type="file" name="file" accept=".csv,.txt,.xlsx"
                 required
                 class="w-full text-sm text-slate-600 file:mr-4 file:py-2.5 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer">
        </label>
        <button type="submit"
                class="inline-flex items-center justify-center gap-2 px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold rounded-lg shadow transition">
          <img src="https://unpkg.com/lucide-static@1.44.0/icons/upload.svg" class="w-4 h-4" alt="">
          Upload &amp; Extract
        </button>
      </form>

      {% if error %}
        <p class="mt-3 text-sm text-red-600 font-medium">{{ error }}</p>
      {% endif %}
      {% if filename %}
        <p class="mt-3 text-xs text-slate-500">
          📄 Uploaded: <span class="font-semibold text-slate-700">{{ filename }}</span>
          • Total: <span class="font-semibold text-slate-700">{{ total }}</span> records
          • Series: {% for length, _ in series %}<span class="font-semibold text-indigo-600">{{ length }}-digit</span>{% if not loop.last %}, {% endif %}{% endfor %}
        </p>
      {% endif %}
    </div>

    <!-- Results (neeche) — series-wise -->
    {% if series %}
    <div class="px-6 pb-6 space-y-6">

      <!-- Download (sabse neeche button) — plain TXT -->
      <div class="flex items-center justify-between">
        <h2 class="text-lg font-bold text-slate-800">Results — Series-wise</h2>
        <a href="/download/{{ token }}"
           class="inline-flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded-lg shadow transition whitespace-nowrap">
          <img src="https://unpkg.com/lucide-static@1.44.0/icons/download.svg" class="w-4 h-4" alt="">
          Download NUMBER|CODE (.txt)
        </a>
      </div>

      <p class="text-xs text-slate-500 -mt-4">
        📄 Download file me sirf <span class="font-mono font-semibold text-indigo-600">NUMBER|CODE</span> hoga — ek ke niche ek, koi header / faltu data nahi.
      </p>

      {% for length, items in series %}
      <!-- Ek series block -->
      <div class="border border-slate-200 rounded-xl overflow-hidden">
        <div class="bg-slate-100 px-4 py-2.5 flex items-center justify-between">
          <h3 class="text-sm font-bold text-slate-700">
            {{ length }}-Digit Series
            <span class="ml-2 text-xs font-semibold text-indigo-600 bg-indigo-100 px-2 py-0.5 rounded-full">{{ items|length }} records</span>
          </h3>
          <span class="text-xs text-slate-400">Numbers: chhote → bade</span>
        </div>
        <div class="max-h-56 overflow-auto">
          <table class="w-full text-sm">
            <thead class="sticky top-0 bg-white text-slate-500 text-xs">
              <tr class="text-left border-b border-slate-100">
                <th class="px-4 py-2 font-semibold">#</th>
                <th class="px-4 py-2 font-semibold">Number</th>
                <th class="px-4 py-2 font-semibold">Code ({{ length }} digit)</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-100">
              {% for number, code in items %}
              <tr class="hover:bg-indigo-50/60">
                <td class="px-4 py-2 text-slate-400">{{ loop.index }}</td>
                <td class="px-4 py-2 font-mono font-semibold text-slate-800">{{ number }}</td>
                <td class="px-4 py-2">
                  <span class="inline-block bg-indigo-100 text-indigo-800 font-mono font-bold px-3 py-1 rounded-md">{{ code }}</span>
                </td>
              </tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      </div>
      {% endfor %}

    </div>
    {% endif %}

    <div class="px-6 py-4 bg-slate-50 border-t border-slate-100 text-center text-xs text-slate-500">
      Flask script 🐍 • TXT / CSV / Excel → NUMBER|CODE → Clean TXT Download
    </div>
  </div>
</body>
</html>
"""


# ---------------- Routes ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        f = request.files.get("file")
        if f is None or f.filename == "":
            return render_template_string(HTML, error="Pehle file select karein!")

        raw = f.read()
        ext = f.filename.lower().rsplit(".", 1)[-1] if "." in f.filename else ""

        # File type ke hisab se parse karo
        try:
            if ext == "xlsx":
                if not HAS_EXCEL:
                    return render_template_string(
                        HTML,
                        error="Excel file ke liye 'pip install openpyxl' karein.",
                        filename=f.filename,
                    )
                rows = parse_excel(raw)
            else:
                text = None
                for enc in ("utf-8-sig", "utf-8", "latin-1"):
                    try:
                        text = raw.decode(enc)
                        break
                    except UnicodeDecodeError:
                        continue
                rows = parse_text(text or raw.decode("utf-8", errors="ignore"))
        except Exception as e:
            return render_template_string(
                HTML,
                error=f"File parse karte waqt error: {e}",
                filename=f.filename,
            )

        if not rows:
            return render_template_string(
                HTML,
                error="File me koi valid NUMBER|CODE nahi mila. Format: 249110142363|880421",
                filename=f.filename,
            )

        series = group_by_series(rows)
        token = uuid.uuid4().hex
        results_store[token] = series

        return render_template_string(
            HTML,
            series=series,
            total=len(rows),
            filename=f.filename,
            token=token,
        )

    return render_template_string(HTML)


@app.route("/download/<token>")
def download(token):
    """
    Clean TXT download — sirf NUMBER|CODE lines.
    Koi header nahi, koi faltu data nahi, ek ke niche ek.
    """
    if token not in results_store:
        abort(404)

    series = results_store.pop(token)

    buf = io.StringIO()
    for length, items in series:
        for number, code in items:
            buf.write(f"{number}|{code}\n")

    data = io.BytesIO(buf.getvalue().encode("utf-8"))
    return send_file(
        data,
        as_attachment=True,
        download_name="numbers_codes.txt",
        mimetype="text/plain",
    )


if __name__ == "__main__":
    # ---- Port: 26283 ----
    app.run(debug=True, host="0.0.0.0", port=26283)
