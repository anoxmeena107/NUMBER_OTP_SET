"""
============================================================
 Flask Script: NUMBER|CODE Extractor — ANOX + WhatsApp
============================================================
 - ANOX logo (header me)
 - WhatsApp floating button -> wa.me/916375727004
 Modes:
   1) Number + Code  : NUMBER|CODE series-wise (5/6/7/8 digit)
   2) Sirf Number     : koi bhi file -> sirf numbers
 Supported files : .txt, .csv, .xlsx, .xls
 Download        : clean .txt (sirf requested data)
 Port            : 26283

 Chalane ke liye:
   pip install -r requirements.txt
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

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl  # noqa: F401
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import xlrd  # noqa: F401
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

app = Flask(__name__)
results_store = {}


# ---------------- Code / Number extraction ----------------
def extract_code(raw):
    """Field se OTP code nikalta hai. Na mile to '' return karta hai."""
    if raw is None:
        return ""
    raw = str(raw).strip()

    if raw.isdigit():
        return raw

    # Arabic message: "الرمز الخاص بك هو ... #880421"
    m = re.search(r"هو[\u200F\u200E\s]*(\d+)", raw)
    if m:
        return m.group(1)

    m = re.search(r"#(\d+)", raw)
    if m:
        return m.group(1)

    groups = re.findall(r"\d+", raw)
    for g in groups:
        if 4 <= len(g) <= 8:
            return g
    return groups[-1] if groups else ""


def clean_number(raw):
    """Number field se sirf digits nikalta hai."""
    if raw is None:
        return ""
    return re.sub(r"\D", "", str(raw))


def add_pair(rows, number, code):
    """Code na bhi ho to bhi number save hota hai (code '') — dono modes ke liye."""
    if number:
        rows.append((number, code))


# ---------------- File parsing ----------------
def parse_text(text):
    """TXT/CSV parse — NUMBER|CODE ya purana CSV format. Code na ho to ''."""
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        parts = line.split("|")
        if len(parts) >= 2:
            num_raw, code_raw = parts[0], parts[1]
        else:
            fields = next(csv.reader([line]), [])
            if len(fields) >= 5:
                num_raw, code_raw = fields[2], fields[4]
            elif len(fields) == 2:
                num_raw, code_raw = fields
            else:
                continue

        add_pair(rows, clean_number(num_raw), extract_code(code_raw))

    return rows


def parse_with_pandas(raw, ext):
    """pandas se CSV / Excel parse. Pehle 2 columns (ya NUMBER/CODE header)."""
    rows = []
    buf = io.BytesIO(raw)
    try:
        if ext == "csv":
            df = pd.read_csv(buf, dtype=str, keep_default_na=False)
        elif ext == "xlsx":
            df = pd.read_excel(buf, engine="openpyxl", dtype=str, keep_default_na=False)
        elif ext == "xls":
            df = pd.read_excel(buf, engine="xlrd", dtype=str, keep_default_na=False)
        else:
            return rows
    except Exception:
        return rows

    cols = list(df.columns)
    num_col = code_col = None
    for c in cols:
        cl = str(c).strip().lower()
        if cl in ("number", "num", "mobile", "phone", "no"):
            num_col = c
        elif cl in ("code", "otp", "otp code", "password", "pin"):
            code_col = c

    if num_col is None and cols:
        num_col = cols[0]
    if code_col is None and len(cols) >= 2:
        code_col = cols[1]

    for _, row in df.iterrows():
        nv = row.get(num_col) if num_col else None
        cv = row.get(code_col) if code_col else None
        add_pair(rows, clean_number(nv), extract_code(cv))

    return rows


def parse_excel_openpyxl(raw):
    """Fallback: .xlsx via openpyxl (agar pandas na ho)."""
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    rows = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            vals = [str(v) if v is not None else "" for v in row]
            if len(vals) < 2:
                continue
            add_pair(rows, clean_number(vals[0]), extract_code(vals[1]))
    return rows


def group_by_series(rows):
    """Code mode: digit-count ke hisab se series, har series me numbers ascending."""
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
<title>ANOX | Number|Code Extractor</title>
<script src="https://unpkg.com/@tailwindcss/browser@4"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style> body { font-family: 'Inter', system-ui, sans-serif; } </style>
</head>
<body class="bg-gradient-to-br from-slate-100 via-blue-50 to-indigo-100 min-h-screen flex items-start justify-center p-6">

  <div class="w-full max-w-4xl bg-white rounded-2xl shadow-xl ring-1 ring-slate-200 overflow-hidden">

    <!-- Header with ANOX logo -->
    <div class="bg-gradient-to-r from-indigo-600 to-blue-600 px-6 py-5 text-white flex items-center justify-between flex-wrap gap-3">
      <div class="flex items-center gap-3">
        <!-- ANOX logo mark -->
        <div class="w-11 h-11 rounded-xl bg-white/15 backdrop-blur border border-white/30 flex items-center justify-center shadow-lg">
          <span class="text-white font-extrabold text-2xl leading-none">A</span>
        </div>
        <div>
          <h1 class="text-2xl font-extrabold tracking-tight leading-none">ANOX</h1>
          <p class="text-indigo-100 text-xs mt-1 uppercase tracking-widest font-medium">Number | Code Extractor</p>
        </div>
      </div>
      <p class="text-indigo-100 text-xs max-w-xs hidden sm:block">
        TXT / CSV / Excel (.xlsx · .xls) — koi bhi file. Do modes: Number+Code ya Sirf Number.
      </p>
    </div>

    <!-- Info / Note (upar) -->
    <div class="px-6 pt-5">
      <div class="flex items-start gap-3 bg-amber-50 border border-amber-200 rounded-xl px-4 py-3">
        <span class="text-xl">💡</span>
        <p class="text-sm text-amber-800">
          <b>Note:</b> कुछ भी change / update करना हो तो बताएं — मैं script को आपके हिसाब से adjust कर दूंगा!
        </p>
      </div>
    </div>

    <!-- Upload card (sabse upar) -->
    <div class="px-6 py-5">
      <form method="post" enctype="multipart/form-data" class="space-y-4">
        <div class="flex flex-col sm:flex-row gap-3 items-stretch sm:items-center">
          <label class="flex-1">
            <input type="file" name="file" accept=".csv,.txt,.xlsx,.xls" required
                   class="w-full text-sm text-slate-600 file:mr-4 file:py-2.5 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer">
          </label>
          <button type="submit"
                  class="inline-flex items-center justify-center gap-2 px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold rounded-lg shadow transition">
            <img src="https://unpkg.com/lucide-static@1.44.0/icons/upload.svg" class="w-4 h-4" alt="">
            Upload &amp; Extract
          </button>
        </div>

        <!-- Mode selection -->
        <div class="flex flex-wrap items-center gap-x-6 gap-y-2 bg-slate-50 border border-slate-200 rounded-xl px-4 py-3">
          <span class="text-sm font-semibold text-slate-600">Mode:</span>
          <label class="inline-flex items-center gap-2 cursor-pointer">
            <input type="radio" name="mode" value="code" checked class="accent-indigo-600 w-4 h-4">
            <span class="text-sm text-slate-700">Number + Code (NUMBER|CODE)</span>
          </label>
          <label class="inline-flex items-center gap-2 cursor-pointer">
            <input type="radio" name="mode" value="numbers" class="accent-indigo-600 w-4 h-4">
            <span class="text-sm text-slate-700">Sirf Number (Only Numbers)</span>
          </label>
        </div>
      </form>

      {% if error %}
        <p class="mt-3 text-sm text-red-600 font-medium">{{ error }}</p>
      {% endif %}
      {% if filename %}
        <p class="mt-3 text-xs text-slate-500">
          📄 Uploaded: <span class="font-semibold text-slate-700">{{ filename }}</span>
          • Total: <span class="font-semibold text-slate-700">{{ total }}</span> records
          {% if kind == 'code' %}
          • Series: {% for length, _ in series %}<span class="font-semibold text-indigo-600">{{ length }}-digit</span>{% if not loop.last %}, {% endif %}{% endfor %}
          {% endif %}
        </p>
      {% endif %}
    </div>

    <!-- Results — Code mode (series-wise) -->
    {% if kind == 'code' and series %}
    <div class="px-6 pb-6 space-y-6">
      <div class="flex items-center justify-between">
        <h2 class="text-lg font-bold text-slate-800">Results — Series-wise</h2>
        <a href="/download/{{ token }}"
           class="inline-flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded-lg shadow transition whitespace-nowrap">
          <img src="https://unpkg.com/lucide-static@1.44.0/icons/download.svg" class="w-4 h-4" alt="">
          Download NUMBER|CODE (.txt)
        </a>
      </div>
      <p class="text-xs text-slate-500 -mt-4">
        📄 Download me sirf <span class="font-mono font-semibold text-indigo-600">NUMBER|CODE</span> — ek ke niche ek, koi header nahi.
      </p>

      {% for length, items in series %}
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

    <!-- Results — Sirf Number mode -->
    {% if kind == 'numbers' and numbers %}
    <div class="px-6 pb-6">
      <div class="flex items-center justify-between">
        <h2 class="text-lg font-bold text-slate-800">Results — Sirf Numbers ({{ numbers|length }})</h2>
        <a href="/download/{{ token }}"
           class="inline-flex items-center gap-2 px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold rounded-lg shadow transition whitespace-nowrap">
          <img src="https://unpkg.com/lucide-static@1.44.0/icons/download.svg" class="w-4 h-4" alt="">
          Download Numbers (.txt)
        </a>
      </div>
      <p class="text-xs text-slate-500 -mt-4 mb-3">
        📄 Download me sirf numbers — ek ke niche ek, koi faltu data nahi.
      </p>
      <div class="max-h-96 overflow-auto rounded-xl border border-slate-200">
        <table class="w-full text-sm">
          <thead class="sticky top-0 bg-slate-100 text-slate-700">
            <tr class="text-left">
              <th class="px-4 py-3 font-semibold">#</th>
              <th class="px-4 py-3 font-semibold">Number</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100">
            {% for number in numbers %}
            <tr class="hover:bg-indigo-50/60">
              <td class="px-4 py-2.5 text-slate-400">{{ loop.index }}</td>
              <td class="px-4 py-2.5 font-mono font-semibold text-slate-800">{{ number }}</td>
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
    </div>
    {% endif %}

    <div class="px-6 py-4 bg-slate-50 border-t border-slate-100 text-center text-xs text-slate-500">
      ANOX 🐍 • TXT / CSV / Excel (xlsx·xls) → Number+Code / Sirf Number → Download
    </div>
  </div>

  <!-- WhatsApp floating button — click par +91 63757 27004 par chat khul ta hai -->
  <a href="https://wa.me/916375727004?text=Namaste!%20ANOX%20Number%20Code%20Extractor%20se%20baat%20karni%20hai."
     target="_blank" rel="noopener noreferrer"
     aria-label="WhatsApp: +91 63757 27004"
     class="fixed bottom-6 right-6 z-50 group flex items-center gap-2">
    <!-- Hover label (desktop) -->
    <span class="hidden md:inline-block bg-white text-slate-700 text-xs font-semibold px-3 py-2 rounded-full shadow-lg ring-1 ring-slate-200 opacity-0 group-hover:opacity-100 transition-opacity duration-200 translate-x-2 group-hover:translate-x-0">
      +91 63757 27004
    </span>
    <!-- WhatsApp circular button -->
    <span class="w-14 h-14 rounded-full bg-[#25D366] hover:bg-[#1faf54] shadow-xl shadow-green-500/40 flex items-center justify-center transition-transform duration-200 hover:scale-110">
      <svg viewBox="0 0 24 24" fill="white" class="w-7 h-7" aria-hidden="true">
        <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413Z"/>
      </svg>
    </span>
  </a>

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

        mode = request.form.get("mode", "code")
        raw = f.read()
        ext = f.filename.lower().rsplit(".", 1)[-1] if "." in f.filename else ""

        try:
            if ext == "txt":
                text = None
                for enc in ("utf-8-sig", "utf-8", "latin-1"):
                    try:
                        text = raw.decode(enc)
                        break
                    except UnicodeDecodeError:
                        continue
                rows = parse_text(text or raw.decode("utf-8", errors="ignore"))

            elif ext in ("csv", "xlsx", "xls"):
                rows = []
                if ext == "xlsx" and not HAS_OPENPYXL:
                    return render_template_string(HTML, error="Excel .xlsx ke liye 'pip install openpyxl' karein.", filename=f.filename)
                if ext == "xls" and not HAS_XLRD:
                    return render_template_string(HTML, error="Excel .xls ke liye 'pip install xlrd' karein.", filename=f.filename)

                if HAS_PANDAS:
                    rows = parse_with_pandas(raw, ext)
                if not rows and ext == "xlsx" and HAS_OPENPYXL:
                    rows = parse_excel_openpyxl(raw)
                if not rows and ext == "csv":
                    rows = parse_text(raw.decode("utf-8", errors="ignore"))
            else:
                return render_template_string(HTML, error="Sirf .txt / .csv / .xlsx / .xls file upload karein.", filename=f.filename)

        except Exception as e:
            return render_template_string(HTML, error=f"File parse karte waqt error: {e}", filename=f.filename)

        # Mode ke hisab se result taiyar karo
        if mode == "numbers":
            numbers = sorted([n for n, _ in rows if n], key=int)
            if not numbers:
                return render_template_string(HTML, error="File me koi number nahi mila.", filename=f.filename)
            token = uuid.uuid4().hex
            results_store[token] = {"kind": "numbers", "numbers": numbers}
            return render_template_string(HTML, kind="numbers", numbers=numbers,
                                          total=len(numbers), filename=f.filename, token=token)

        # code mode
        valid = [(n, c) for n, c in rows if c]
        if not valid:
            return render_template_string(HTML, error="File me koi valid NUMBER|CODE nahi mila (5/6/7/8 digit code nahi mila). Sirf Number mode try karein.", filename=f.filename)

        series = group_by_series(valid)
        token = uuid.uuid4().hex
        results_store[token] = {"kind": "code", "series": series}
        return render_template_string(HTML, kind="code", series=series,
                                      total=len(valid), filename=f.filename, token=token)

    return render_template_string(HTML)


@app.route("/download/<token>")
def download(token):
    """Clean TXT download — mode ke according: sirf numbers ya NUMBER|CODE."""
    if token not in results_store:
        abort(404)

    data = results_store.pop(token)
    buf = io.StringIO()

    if data["kind"] == "numbers":
        for n in data["numbers"]:
            buf.write(f"{n}\n")
        fname, mime = "numbers_only.txt", "text/plain"
    else:
        for _, items in data["series"]:
            for number, code in items:
                buf.write(f"{number}|{code}\n")
        fname, mime = "numbers_codes.txt", "text/plain"

    out = io.BytesIO(buf.getvalue().encode("utf-8"))
    return send_file(out, as_attachment=True, download_name=fname, mimetype=mime)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=26283)
