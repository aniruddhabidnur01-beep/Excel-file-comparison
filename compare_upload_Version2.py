#!/usr/bin/env python3
"""
compare_upload.py

Simple Flask app to upload two Excel files, compare them sheet-by-sheet, and download a differences Excel file.

Usage:
    pip install flask pandas openpyxl numpy
    python compare_upload.py
Then open http://127.0.0.1:5000/ in your browser.

Comparison behavior:
- Compares all sheets present in either workbook by default.
- Aligns rows and columns (fills missing entries with NaN).
- Reports cell-level differences with sheet, 0-based row_index, 1-based excel_row, column, left_value, right_value.
- Options: ignore-case (for text), tolerance (numeric).
"""
from flask import Flask, request, send_file, render_template_string, flash
import pandas as pd
import numpy as np
import io
import tempfile
from typing import Any, List, Dict, Tuple

app = Flask(__name__)
# Change secret key for production
app.secret_key = "change-me-for-production"

HTML_FORM = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Excel Compare</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; }
    label { display:block; margin: .5rem 0; }
    input[type="file"] { display:block; }
    .info { margin-top: 1rem; color: #444; }
    .error { color: red; }
  </style>
</head>
<body>
  <h1>Upload two Excel files to compare</h1>
  <form method="post" enctype="multipart/form-data">
    <label>Left file: <input type="file" name="left" required></label>
    <label>Right file: <input type="file" name="right" required></label>
    <label>Sheets (comma-separated, optional): <input type="text" name="sheets" placeholder="e.g. Sheet1,Sheet2"></label>
    <label>Ignore case: <input type="checkbox" name="ignore_case"></label>
    <label>Numeric tolerance: <input type="text" name="tolerance" value="0.0"></label>
    <br>
    <input type="submit" value="Compare and Download Differences">
  </form>
  {% with messages = get_flashed_messages() %}
    {% if messages %}
      <div class="error">
      {% for m in messages %}
        <div>{{ m }}</div>
      {% endfor %}
      </div>
    {% endif %}
  {% endwith %}
  <div class="info">
    <p>Output: an Excel workbook containing a <strong>differences</strong> sheet and a <strong>summary</strong> sheet.</p>
    <p>For large files, consider running a CLI comparison or adding file size limits.</p>
  </div>
</body>
</html>
"""

def is_number(x: Any) -> bool:
    try:
        # handle numpy scalars and Python numbers
        if isinstance(x, (np.floating, np.integer, np.number)):
            return np.isfinite(x)
        return isinstance(x, (int, float))
    except Exception:
        try:
            float(x)
            return True
        except Exception:
            return False

def compare_values(a: Any, b: Any, tol: float = 0.0, ignore_case: bool = False) -> bool:
    # Treat None and NaN as equal
    if a is None and b is None:
        return True
    if pd.isna(a) and pd.isna(b):
        return True

    # Numeric comparison when both are numeric
    if is_number(a) and is_number(b):
        try:
            return abs(float(a) - float(b)) <= tol
        except Exception:
            pass

    # Text comparison (fall back to string)
    try:
        sa = "" if pd.isna(a) else str(a)
        sb = "" if pd.isna(b) else str(b)
        if ignore_case:
            return sa.strip().lower() == sb.strip().lower()
        return sa.strip() == sb.strip()
    except Exception:
        return a == b

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    df = df.reset_index(drop=True)
    return df

def compare_sheets(df_left: pd.DataFrame, df_right: pd.DataFrame, sheet_name: str, tol: float, ignore_case: bool) -> Tuple[List[Dict], Dict]:
    diffs: List[Dict] = []
    summary = {
        "sheet": sheet_name,
        "left_rows": len(df_left),
        "right_rows": len(df_right),
        "left_cols": len(df_left.columns),
        "right_cols": len(df_right.columns),
        "diff_count": 0
    }

    df_left = normalize_dataframe(df_left)
    df_right = normalize_dataframe(df_right)

    all_columns = list(sorted(set(df_left.columns).union(set(df_right.columns)), key=lambda x: str(x)))
    max_rows = max(len(df_left), len(df_right))

    # Reindex to same shape
    df_left = df_left.reindex(columns=all_columns, fill_value=np.nan)
    df_right = df_right.reindex(columns=all_columns, fill_value=np.nan)
    df_left = df_left.reindex(range(max_rows), fill_value=np.nan)
    df_right = df_right.reindex(range(max_rows), fill_value=np.nan)

    for r in range(max_rows):
        for c in all_columns:
            a = df_left.at[r, c] if c in df_left.columns else np.nan
            b = df_right.at[r, c] if c in df_right.columns else np.nan
            if not compare_values(a, b, tol=tol, ignore_case=ignore_case):
                diffs.append({
                    "sheet": sheet_name,
                    "row_index": r,
                    "excel_row": r + 1,
                    "column": c,
                    "left_value": a,
                    "right_value": b
                })

    summary["diff_count"] = len(diffs)
    return diffs, summary

@app.route("/", methods=["GET", "POST"])
def upload_and_compare():
    if request.method == "POST":
        left_file = request.files.get("left")
        right_file = request.files.get("right")
        sheets_input = (request.form.get("sheets") or "").strip()
        ignore_case = bool(request.form.get("ignore_case"))
        tol_input = (request.form.get("tolerance") or "0.0").strip()

        try:
            tolerance = float(tol_input) if tol_input != "" else 0.0
        except ValueError:
            flash("Invalid tolerance value; using 0.0")
            tolerance = 0.0

        if not left_file or not right_file:
            flash("Both files are required.")
            return render_template_string(HTML_FORM)

        # Save uploaded files to temporary files and read with pandas
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp_left, tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp_right:
            left_file.save(tmp_left.name)
            right_file.save(tmp_right.name)
            try:
                left_book = pd.read_excel(tmp_left.name, sheet_name=None)
                right_book = pd.read_excel(tmp_right.name, sheet_name=None)
            except Exception as e:
                flash(f"Failed to read uploaded Excel files: {e}")
                return render_template_string(HTML_FORM)

            left_sheets = set(left_book.keys())
            right_sheets = set(right_book.keys())

            if sheets_input:
                requested = [s.strip() for s in sheets_input.split(",") if s.strip()]
                sheets_to_compare = [s for s in requested if s in left_sheets or s in right_sheets]
                if not sheets_to_compare:
                    flash("No requested sheets found in either file.")
                    return render_template_string(HTML_FORM)
            else:
                sheets_to_compare = sorted(list(left_sheets.union(right_sheets)))

            all_diffs: List[Dict] = []
            summaries: List[Dict] = []

            for sheet in sheets_to_compare:
                if sheet not in left_sheets:
                    summaries.append({"sheet": sheet, "note": "only_in_right"})
                    continue
                if sheet not in right_sheets:
                    summaries.append({"sheet": sheet, "note": "only_in_left"})
                    continue

                df_left = left_book[sheet]
                df_right = right_book[sheet]
                diffs, summary = compare_sheets(df_left, df_right, sheet, tol=tolerance, ignore_case=ignore_case)
                all_diffs.extend(diffs)
                summaries.append(summary)

            # Prepare an in-memory Excel workbook with results
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                if all_diffs:
                    diffs_df = pd.DataFrame(all_diffs)
                    diffs_df.to_excel(writer, sheet_name="differences", index=False)
                else:
                    pd.DataFrame([{"note": "no differences found"}]).to_excel(writer, sheet_name="differences", index=False)

                summary_df = pd.DataFrame(summaries)
                summary_df.to_excel(writer, sheet_name="summary", index=False)

            output.seek(0)
            return send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="differences.xlsx"
            )

    return render_template_string(HTML_FORM)

if __name__ == "__main__":
    # For development only. Use a WSGI server for production.
    app.run(debug=True, host="127.0.0.1", port=5000)