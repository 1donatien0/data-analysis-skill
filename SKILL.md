---
name: data-analysis
version: 1.0.0
author: Donatien Byakombe
year: 2026
license: MIT
compatibility: claude-code cowork
allowed-tools:
  - Read
  - Write
  - Edit
  - Bash
  - Glob
description: |
  Analyze data files and produce an interactive HTML report with charts, statistics, and plain-language insights.
  Use whenever the user wants to explore, understand, or get insights from data files.
  Trigger for any request involving CSV, Excel (.xlsx), JSON, or TSV files - even casual phrasing like
  "what's in this file?", "can you look at my data?", "analyze my spreadsheet", "find trends",
  "what does this data say?", "make a chart from this", "summarize my dataset", "which products sell best?",
  "is there anything unusual in here?", or "help me understand these numbers".
  Also trigger when the user pastes a table or mentions numbers they want interpreted.
  When in doubt, use this skill - data questions without explicit tool context almost always benefit from it.
---

# Data Analysis Skill

## Goal

Turn raw data into clear, actionable insights with an interactive HTML report - charts, statistics, and a plain-language summary that anyone can read, even without a technical background.

## When you receive a data request

1. **Find the data** - look in the current workspace folder for CSV, XLSX, JSON, or TSV files. If the user named a file, locate it. If they didn't specify, list what's available and use the most relevant one (or ask if there are several candidates).
2. **Understand the question** - what does the user actually want to know? "Analyze this" is broad; infer intent from context (e.g. sales file → they want trends and top performers; HR file → headcount, turnover; finance file → totals, anomalies).
3. **Run the analysis** - use the Python script below to extract stats, correlations, and chart data.
4. **Generate the HTML report** - use the template below to build a self-contained, browser-ready file.
5. **Save and present** - save the report to the workspace folder and share the file with the user.

---

## Step 1 - Analyze the data with Python

Run this script (adapt the file path if needed). It loads CSV, TSV, Excel, and JSON without hardcoded column names.

```python
import csv, json, re, sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

def install_packages():
    import subprocess
    cmd = [sys.executable, "-m", "pip", "install", "pandas", "openpyxl", "xlrd", "-q"]
    try:
        subprocess.run(cmd + ["--break-system-packages"], check=True)
    except Exception:
        subprocess.run(cmd, check=True)

try:
    import pandas as pd
    import numpy as np
except ImportError:
    install_packages()
    import pandas as pd
    import numpy as np

def fail(message):
    print(json.dumps({"error": message}, ensure_ascii=False))
    sys.exit(1)

def is_missing(value):
    if value is None or isinstance(value, (list, dict, tuple, set)):
        return value is None
    try:
        return bool(pd.isna(value))
    except Exception:
        return False

def clean(value):
    if isinstance(value, dict):
        return {str(k): clean(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [clean(v) for v in value]
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        number = float(value)
        return None if not np.isfinite(number) else number
    if isinstance(value, (np.bool_, bool)):
        return bool(value)
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if is_missing(value):
        return None
    return value

def round_or_none(value, digits=4):
    number = clean(value)
    if isinstance(number, bool) or not isinstance(number, (int, float)):
        return None
    return round(float(number), digits)

def dedupe_columns(df):
    seen = {}
    names = []
    for col in df.columns:
        base = str(col).strip() or "colonne"
        count = seen.get(base, 0) + 1
        seen[base] = count
        names.append(base if count == 1 else f"{base}_{count}")
    out = df.copy()
    out.columns = names
    return out

def decimal_char(sample, sep):
    if sep != ";":
        return "."
    comma_decimals = len(re.findall(r"\d,\d", sample))
    dot_decimals = len(re.findall(r"\d\.\d", sample))
    return "," if comma_decimals > dot_decimals else "."

def detect_sep(sample):
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except Exception:
        counts = {sep: sample.count(sep) for sep in [",", ";", "\t", "|"]}
        best = max(counts, key=counts.get)
        return best if counts[best] else ","

def read_delimited(file_path):
    last_error = "format non reconnu"
    for encoding in ["utf-8-sig", "utf-8", "cp1252", "latin-1"]:
        try:
            sample = file_path.read_text(encoding=encoding)[:8192]
        except Exception as exc:
            last_error = str(exc)
            continue
        if sample.strip() == "":
            fail(f"Le fichier est vide : {file_path.name}")
        sep = detect_sep(sample)
        decimal = decimal_char(sample, sep)
        try:
            df = pd.read_csv(file_path, sep=sep, encoding=encoding, decimal=decimal)
        except Exception as exc:
            last_error = str(exc)
            continue
        if df.shape[1] == 1:
            for alt in [",", ";", "\t", "|"]:
                if alt == sep or sample.count(alt) == 0:
                    continue
                alt_decimal = decimal_char(sample, alt)
                try:
                    alt_df = pd.read_csv(file_path, sep=alt, encoding=encoding, decimal=alt_decimal)
                except Exception:
                    continue
                if alt_df.shape[1] > df.shape[1]:
                    df, sep, decimal = alt_df, alt, alt_decimal
        return df, {
            "encoding": encoding, "separator": sep, "decimal": decimal,
            "sheets": [], "sheet_used": None,
        }
    fail(f"Impossible de lire le fichier texte ({last_error}).")

def read_excel(file_path):
    engine = "openpyxl" if file_path.suffix.lower() == ".xlsx" else "xlrd"
    try:
        book = pd.ExcelFile(file_path, engine=engine)
    except Exception as exc:
        if "xlrd" in str(exc) or "openpyxl" in str(exc):
            install_packages()
            try:
                book = pd.ExcelFile(file_path, engine=engine)
            except Exception as retry_exc:
                fail(f"Impossible de lire le classeur Excel ({retry_exc}).")
        else:
            fail(f"Impossible de lire le classeur Excel ({exc}).")
    sheets = [str(name) for name in book.sheet_names]
    if not sheets:
        fail("Le classeur Excel ne contient aucune feuille.")
    return book.parse(sheets[0]), {
        "encoding": None, "separator": None, "sheets": sheets, "sheet_used": sheets[0],
    }

def read_json(file_path):
    try:
        text = file_path.read_text(encoding="utf-8-sig")
    except Exception as exc:
        fail(f"JSON illisible ({exc}).")
    if text.strip() == "":
        fail(f"Le fichier est vide : {file_path.name}")
    meta = {"encoding": "utf-8", "separator": None, "sheets": [], "sheet_used": None}
    try:
        raw = json.loads(text)
    except json.JSONDecodeError:
        try:
            return pd.read_json(file_path, lines=True), meta
        except Exception as exc:
            fail(f"JSON illisible ({exc}).")
    if isinstance(raw, list):
        if len(raw) == 0:
            fail("Le JSON ne contient aucune ligne.")
        return pd.json_normalize(raw), meta
    if isinstance(raw, dict):
        list_values = [v for v in raw.values() if isinstance(v, list) and v and isinstance(v[0], dict)]
        if len(raw) == 1 and len(list_values) == 1:
            return pd.json_normalize(list_values[0]), meta
        try:
            return pd.DataFrame(raw), meta
        except Exception as exc:
            fail(f"Le JSON n'est pas un tableau ({exc}).")
    fail("Le JSON doit être un tableau d'objets, un objet de colonnes, ou du JSON Lines.")

def read_table(file_path):
    if not file_path.exists() or not file_path.is_file():
        fail(f"Fichier introuvable : {file_path}")
    suffix = file_path.suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        return read_excel(file_path)
    if suffix == ".json":
        return read_json(file_path)
    return read_delimited(file_path)

def is_text_series(series):
    return (
        pd.api.types.is_object_dtype(series)
        or pd.api.types.is_string_dtype(series)
        or isinstance(series.dtype, pd.CategoricalDtype)
    )

def looks_like_dates(series):
    sample = series.dropna().head(30)
    if len(sample) < 6:
        return False
    as_text = sample.astype(str).str.strip()
    if as_text.str.fullmatch(r"\d{1,4}").mean() >= 0.8:
        return False
    parsed = pd.to_datetime(sample, errors="coerce")
    return float(parsed.notna().mean()) >= 0.8

def finite_series(series):
    numbers = pd.to_numeric(series, errors="coerce")
    return numbers.replace([np.inf, -np.inf], np.nan).dropna()

def monthly(indexed):
    for rule in ("ME", "M"):
        try:
            grouped = indexed.resample(rule)
            try:
                summed = grouped.sum(min_count=1)
            except TypeError:
                summed = grouped.sum()
            averaged = grouped.mean()
            summed = summed.dropna()
            if len(summed) == 0:
                return None, None
            return summed, averaged.reindex(summed.index)
        except Exception:
            continue
    return None, None

file_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data.csv")
df, load_meta = read_table(file_path)
df = dedupe_columns(df)
if df.shape[1] == 0:
    fail("Aucune colonne détectée.")

source_rows = int(len(df))
sampled = False
sample_size = None
if source_rows > 100_000:
    sample_size = 10_000
    df = df.sample(n=sample_size, random_state=42).sort_index()
    sampled = True
if len(df) == 0:
    fail("Le fichier ne contient aucune ligne de données.")

parsed_dates = {}
date_cols = []
for col in df.columns:
    series = df[col]
    if pd.api.types.is_datetime64_any_dtype(series):
        parsed_dates[col] = series
        date_cols.append(col)
    elif is_text_series(series) and looks_like_dates(series):
        parsed_dates[col] = pd.to_datetime(series, errors="coerce")
        date_cols.append(col)

numeric_cols = [
    col for col in df.select_dtypes(include=[np.number]).columns
    if not pd.api.types.is_bool_dtype(df[col])
]

info = {
    "file": file_path.name,
    "rows": int(len(df)),
    "source_rows": source_rows,
    "sampled": sampled,
    "sample_size": sample_size,
    "columns": int(len(df.columns)),
    "column_names": [str(col) for col in df.columns],
    "dtypes": {str(col): str(df[col].dtype) for col in df.columns},
    "missing": {
        str(col): int(df[col].isna().sum())
        for col in df.columns
        if int(df[col].isna().sum()) > 0
    },
    "sheets": load_meta.get("sheets") or [],
    "sheet_used": load_meta.get("sheet_used"),
    "encoding": load_meta.get("encoding"),
    "separator": load_meta.get("separator"),
    "decimal": load_meta.get("decimal"),
    "date_columns": [str(col) for col in date_cols],
}

stats = {}
for col in numeric_cols:
    series = finite_series(df[col])
    if len(series) == 0:
        continue
    distribution = {}
    unique = int(series.nunique())
    if 1 < unique <= 50 and unique < len(series):
        counts = series.value_counts().head(10)
        distribution = {str(clean(key)): int(value) for key, value in counts.items()}
    values = series.to_numpy(dtype=float)
    if len(values) >= 2 and float(np.max(values)) > float(np.min(values)):
        hist_counts, hist_edges = np.histogram(values, bins=10)
    else:
        only = float(values[0]) if len(values) else 0.0
        hist_counts, hist_edges = np.array([len(values)]), np.array([only, only])
    stats[str(col)] = {
        "count": int(len(series)),
        "min": round_or_none(series.min()),
        "max": round_or_none(series.max()),
        "mean": round_or_none(series.mean()),
        "median": round_or_none(series.median()),
        "std": round_or_none(series.std()) if len(series) > 1 else None,
        "sum": round_or_none(series.sum()),
        "distribution": distribution,
        "histogram": [int(v) for v in hist_counts.tolist()],
        "histogram_edges": [round_or_none(v) for v in hist_edges.tolist()],
    }

cat_cols = [col for col in df.columns if col not in date_cols and is_text_series(df[col])]
cat_stats = {}
for col in cat_cols[:6]:
    present = df[col].dropna()
    if len(present) == 0:
        continue
    top = present.astype(str).value_counts().head(15)
    mode = present.astype(str).mode()
    cat_stats[str(col)] = {
        "unique": int(present.nunique()),
        "top_values": {str(key): int(value) for key, value in top.items()},
        "most_common": str(mode.iloc[0]) if len(mode) else "",
    }

time_series = {}
if date_cols and numeric_cols:
    date_col = date_cols[0]
    for num_col in numeric_cols[:3]:
        frame = pd.DataFrame({
            "date": parsed_dates[date_col],
            "value": pd.to_numeric(df[num_col], errors="coerce"),
        }).dropna()
        frame = frame[np.isfinite(frame["value"].to_numpy(dtype=float))]
        if len(frame) < 6:
            continue
        indexed = frame.sort_values("date").set_index("date")["value"]
        summed, averaged = monthly(indexed)
        if summed is None:
            continue
        time_series[str(num_col)] = {
            "dates": [str(stamp.date()) for stamp in summed.index],
            "values": [round_or_none(v) for v in summed.tolist()],
            "means": [round_or_none(v) for v in averaged.tolist()] if averaged is not None else [],
            "date_col": str(date_col),
            "aggregation": "sum",
        }

anomalies = {}
for col in numeric_cols[:5]:
    series = finite_series(df[col])
    if len(series) < 4:
        continue
    q1, q3 = float(series.quantile(0.25)), float(series.quantile(0.75))
    iqr = q3 - q1
    if not np.isfinite(iqr) or iqr == 0:
        continue
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    outliers = series[(series < lower) | (series > upper)]
    if len(outliers) == 0:
        continue
    anomalies[str(col)] = {
        "count": int(len(outliers)),
        "pct": round(len(outliers) / len(series) * 100, 1),
        "lower": round_or_none(lower),
        "upper": round_or_none(upper),
        "examples": [round_or_none(v) for v in outliers.head(5).tolist()],
    }

correlations = []
if len(numeric_cols) >= 2:
    matrix = df[numeric_cols[:20]].apply(pd.to_numeric, errors="coerce").corr()
    names = list(matrix.columns)
    for i, left in enumerate(names):
        for right in names[i + 1:]:
            score = clean(matrix.loc[left, right])
            if isinstance(score, bool) or not isinstance(score, (int, float)):
                continue
            correlations.append({"a": str(left), "b": str(right), "r": round(float(score), 4)})
    correlations.sort(key=lambda item: abs(item["r"]), reverse=True)
    correlations = correlations[:10]

preview = df.head(5).copy()
for col in preview.columns:
    if pd.api.types.is_datetime64_any_dtype(preview[col]):
        preview[col] = preview[col].dt.strftime("%Y-%m-%d")
sample_rows = json.loads(preview.to_json(orient="records", default_handler=str))

result = clean({
    "info": info,
    "stats": stats,
    "cat_stats": cat_stats,
    "time_series": time_series,
    "anomalies": anomalies,
    "correlations": correlations,
    "sample_rows": sample_rows,
})
print(json.dumps(result, ensure_ascii=False, allow_nan=False))
```

`info.sampled` est vrai quand le fichier dépasse 100 000 lignes : le script n'analyse alors que 10 000 lignes tirées au hasard, et `info.source_rows` garde le total. `time_series.*.values` est une somme mensuelle ; `means` est la moyenne mensuelle (à utiliser pour un prix, un taux, un âge ou un pourcentage). `correlations` contient les paires Pearson les plus fortes. `info.sheets` liste les autres feuilles Excel. En cas d'échec, le script affiche un JSON `{"error": "..."}` au lieu d'une trace.

Save the output JSON to a temp file (e.g., `analysis_result.json`) so you can embed it in the HTML report.

---

## Step 2 - Write the HTML report

Build a **single self-contained HTML file** that embeds:
- The analysis data as an inline JSON variable
- Chart.js from CDN for interactive charts
- Plain CSS (no framework needed)

If the analysis JSON contains the character sequence that closes a script element, split that sequence before pasting it into the page (write it as `<\/script>`). Otherwise the browser ends the script early and the charts never render.

### Report structure

Every report must include these sections (skip sections with no data):

1. **Header** - dataset name, number of rows × columns, date generated
2. **Plain-language summary** - 3–5 bullet points written in everyday language. Answer the user's actual question. Examples: "Sales peaked in March (+42% vs January)", "Product A accounts for 63% of revenue", "12 rows have unusually high values in the 'price' column - worth checking."
3. **Key metrics** - prominent KPI cards (totals, averages, maximums) for the 3–4 most meaningful numeric columns
4. **Charts** (pick the most informative for the data):
   - **Bar chart** - for comparing categories (e.g. sales by product, count by region)
   - **Line chart** - for time series data
   - **Histogram** - for distributions (e.g. age distribution, price range)
5. **Correlations** - only when `correlations` is non-empty. State the strongest pairs in plain language ("le chiffre d'affaires suit presque exactement la quantité, r = 0,92"). Skip weak pairs (|r| < 0.7) unless the user asked about a relationship.
6. **Data quality** - missing values, detected anomalies/outliers. If `info.sampled` is true, repeat in the header that the figures come from 10,000 rows drawn from `info.source_rows`.
7. **Sample data table** - first 5 rows, clean and readable

### HTML template skeleton

```html
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Analyse - [DATASET_NAME]</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f8f9fa; color: #212529; }
    .container { max-width: 1100px; margin: 0 auto; padding: 2rem 1.5rem; }
    h1 { font-size: 1.8rem; font-weight: 700; margin-bottom: .25rem; }
    .subtitle { color: #6c757d; font-size: .95rem; margin-bottom: 2rem; }
    h2 { font-size: 1.2rem; font-weight: 600; margin: 2rem 0 1rem; color: #343a40; border-bottom: 2px solid #dee2e6; padding-bottom: .4rem; }
    .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 1rem; }
    .kpi { background: white; border-radius: 10px; padding: 1.25rem 1.5rem; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    .kpi .label { font-size: .8rem; text-transform: uppercase; letter-spacing: .05em; color: #6c757d; margin-bottom: .3rem; }
    .kpi .value { font-size: 1.7rem; font-weight: 700; color: #0d6efd; }
    .kpi .sub { font-size: .8rem; color: #adb5bd; margin-top: .2rem; }
    .insights { background: white; border-radius: 10px; padding: 1.25rem 1.5rem; box-shadow: 0 1px 4px rgba(0,0,0,.08); margin-bottom: 1rem; }
    .insights ul { padding-left: 1.2rem; }
    .insights li { margin-bottom: .5rem; line-height: 1.5; }
    .chart-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 440px), 1fr)); gap: 1.5rem; }
    .chart-box { background: white; border-radius: 10px; padding: 1.25rem; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    .chart-box canvas { max-height: 280px; }
    .table-wrap { overflow-x: auto; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    table { width: 100%; border-collapse: collapse; font-size: .88rem; background: white; }
    th { background: #f1f3f5; padding: .6rem .9rem; text-align: left; font-weight: 600; font-size: .78rem; text-transform: uppercase; letter-spacing: .05em; color: #495057; }
    td { padding: .55rem .9rem; border-top: 1px solid #f1f3f5; }
    tr:hover td { background: #f8f9fa; }
    .badge { display: inline-block; padding: .2rem .55rem; border-radius: 20px; font-size: .75rem; font-weight: 600; }
    .badge-warn { background: #fff3cd; color: #856404; }
    .badge-ok { background: #d1e7dd; color: #0f5132; }
    .quality-row { display: flex; flex-wrap: wrap; gap: .75rem; margin-top: .5rem; }
    .quality-item { background: white; border-radius: 8px; padding: .75rem 1rem; box-shadow: 0 1px 4px rgba(0,0,0,.08); font-size: .88rem; }
    footer { text-align: center; color: #adb5bd; font-size: .8rem; margin-top: 3rem; padding-top: 1rem; border-top: 1px solid #dee2e6; }
  </style>
</head>
<body>
<div class="container">
  <h1>📊 [DATASET_NAME]</h1>
  <p class="subtitle">[ROWS] lignes · [COLS] colonnes · Rapport généré le [DATE]</p>

  <h2>Résumé</h2>
  <div class="insights">
    <ul>
      <!-- INSERT 3-5 plain-language insights here -->
    </ul>
  </div>

  <h2>Indicateurs clés</h2>
  <div class="kpi-grid">
    <!-- INSERT KPI cards here -->
  </div>

  <h2>Visualisations</h2>
  <div class="chart-grid">
    <!-- INSERT chart boxes here -->
  </div>

  <h2>Qualité des données</h2>
  <div class="quality-row">
    <!-- INSERT missing values and anomaly badges here -->
  </div>

  <h2>Aperçu des données</h2>
  <div class="table-wrap">
  <table>
    <!-- INSERT table headers and first 5 rows here -->
  </table>
  </div>

  <footer>Généré par Claude · skill data-analysis · Concepteur Donatien Byakombe 2026</footer>
</div>

<script>
// Paste the analysis JSON here.
const data = /* PASTE JSON HERE */;

// Example bar chart:
// new Chart(document.getElementById('myChart'), {
//   type: 'bar',
//   data: { labels: Object.keys(data.cat_stats['column'].top_values),
//           datasets: [{ label: 'Count', data: Object.values(data.cat_stats['column'].top_values),
//                        backgroundColor: '#0d6efd' }] },
//   options: { responsive: true, plugins: { legend: { display: false } } }
// });
</script>
</body>
</html>
```

### Chart color palette

Use these consistent colors:
- Primary bars/lines: `#0d6efd`
- Secondary: `#6610f2`, `#fd7e14`, `#20c997`, `#dc3545`
- Background fills: same colors at 20% opacity (add `33` to hex)

---

## Step 3 - Write insights in plain language

The most important part of the report is the **plain-language summary**. Write it as if explaining to a colleague who doesn't know statistics:

✅ Good: "Ventes moyennes de 4 200 € par mois. Pic en novembre (+68 %), probablement lié aux fêtes."
❌ Bad: "The mean of the 'revenue' column is 4200 with a standard deviation of 890."

Rules:
- Lead with the most surprising or actionable finding
- Include magnitudes and percentages, not just "high" or "low"
- If there are anomalies, flag them clearly: "⚠️ 8 lignes ont des valeurs inhabituellement élevées dans 'prix' - vérification recommandée"
- Adapt the language to the domain (sales → "chiffre d'affaires", HR → "effectifs", etc.)

---

## Output file naming

Save the report as `[original_filename]_analyse.html` in the same folder as the source data file (or in the workspace folder if that's where the user works).

---

## Edge cases and what NOT to do

- **Empty or malformed file** - tell the user clearly what's wrong and what to fix; don't guess or fabricate data
- **Very large file (>100k rows)** - the script keeps 10,000 random rows and sets `info.sampled`. Note this prominently in the report header, with `info.source_rows`
- **Single column** - run univariate analysis only (histogram + stats); don't invent correlations
- **All text, no numbers** - focus on word frequencies, top categories, and missing values; skip numeric KPIs
- **Multiple sheets (Excel)** - the script analyzes `info.sheet_used` (the first sheet) and lists every name in `info.sheets`; mention the sheets that were not analyzed
- **Unreadable file** - the script prints `{"error": "..."}` and exits. Show that message to the user; don't fabricate a report
- **JSON shape** - a list of objects, one object of columns, or JSON Lines. Anything else is an error, not a dataset to guess through
- **User asks a specific question** (e.g. "which region sells the most?") - answer that question prominently in the insights section first, then add the general report below
- **No data files in the folder** - list what IS there and ask the user to point to the right file; don't fabricate
- **Don't over-engineer** - if the file has 3 rows and 2 columns, a simple table with 2 observations is the right output, not a full dashboard

---

## Process (two-pass)

1. **Analyze** - run the Python script, capture stats and chart data
2. **Draft insights** - write 3–5 plain-language findings; ask: "Does this actually answer the user's question?"
3. **Audit** - scan for anything vague or jargon-heavy: replace "high standard deviation" with "values vary a lot (from X to Y)"; replace "outlier detected" with "⚠️ 8 rows have unusually high values in 'price' - worth checking"
4. **Build report** - embed insights + charts into the HTML template; verify charts render correctly
5. **Save and present** - use the correct filename convention and share the file

---

## Full worked example

**User prompt:** "j'ai un CSV avec mes ventes, dis-moi quels produits marchent le mieux"

**File:** `ventes.csv` with columns: date, produit, region, quantite, ca

**Step 1 - Python analysis output (abbreviated):**
```json
{
  "cat_stats": {
    "produit": {
      "top_values": {"Laptop Pro": 18, "Souris": 12, "Écran 27\"": 8}
    }
  },
  "stats": {
    "ca": {"sum": 576911, "mean": 13736, "max": 67548}
  }
}
```

**Step 2 - Draft insights:**
> "Le CA total est de 576 911 €. Le Laptop Pro est en tête. Il y a une forte progression en fin d'année."

**Step 3 - Audit pass:** too vague. Rewrite:
> "Laptop Pro génère 400 092 € de CA (69% du total), grâce à son prix unitaire de 1 299 €. Les ventes doublent de juillet à novembre. Novembre est le mois record avec 95 388 €."

**Step 4 - Final report includes:**
- KPI card: CA total 576 911 € · Meilleur produit Laptop Pro · Mois record Novembre
- Bar chart: CA par produit (Laptop Pro dominant)
- Line chart: évolution mensuelle (montée nette en Q4)
- Plain-language summary as above

**Output file:** `ventes_analyse.html`

---

## Version history

- **1.0.0** - Initial release. Python analysis for CSV, TSV, Excel, and JSON (encoding, sampling, IQR anomalies, monthly sum and mean, Pearson correlations), self-contained HTML template with Chart.js, two-pass insight process, edge case guidance.
