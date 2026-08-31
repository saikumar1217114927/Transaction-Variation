"""
AMC vs Our-Data Comparator
--------------------------
Upload two files (AMC export + our own data), each expected to contain
columns roughly like: client, code, date, amount.

The app auto-detects those columns even if the headers are named a bit
differently (e.g. "Client Name", "Txn Date", "Amt"), normalizes the date
to yyyy-mm-dd, and finds every row in the AMC file whose
(client, code, date, amount) combination does NOT appear in our file.

Result is returned as a downloadable .xlsx file.

Run:
    pip install flask pandas openpyxl
    python app.py
Then open http://127.0.0.1:5000 in your browser.
"""

from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import pandas as pd
import io
import os
import re

app = Flask(__name__)
app.secret_key = "amc-compare-secret"

# Candidate header names -> standard column name
COLUMN_ALIASES = {
    "client": ["client", "clientname", "client name", "customer", "customername", "name"],
    "code": ["code", "clientcode", "client code", "accountcode", "account code", "acccode", "id"],
    "date": ["date", "txndate", "txn date", "transactiondate", "transaction date", "valuedate", "value date"],
    "amount": ["amount", "amt", "value", "txnamount", "txn amount", "transactionamount"],
}


def normalize_header(h):
    return re.sub(r"[^a-z0-9]", "", str(h).strip().lower())


def detect_columns(df):
    """Return a dict mapping standard name -> actual column name in df (or None if not found).
    code/date are required for matching; client and amount are optional (client is shown for
    reference only and is not used to match records)."""
    normalized_to_actual = {normalize_header(c): c for c in df.columns}
    mapping = {}
    for std_name, aliases in COLUMN_ALIASES.items():
        found = None
        for alias in aliases:
            key = normalize_header(alias)
            if key in normalized_to_actual:
                found = normalized_to_actual[key]
                break
        if found is None and std_name in ("code", "date"):
            raise ValueError(
                f"Could not find a '{std_name}' column. "
                f"Columns present: {list(df.columns)}"
            )
        mapping[std_name] = found
    return mapping


def read_any(file_storage):
    filename = file_storage.filename or ""
    data = file_storage.read()
    buf = io.BytesIO(data)
    if filename.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(buf)
    else:
        return pd.read_csv(buf)


def load_and_standardize(file_storage):
    df = read_any(file_storage)
    mapping = detect_columns(df)

    if mapping["amount"] is not None:
        # Match on actual value, decimals dropped (round then treat as whole number)
        amount = pd.to_numeric(df[mapping["amount"]], errors="coerce").round(0)
    else:
        amount = pd.Series([None] * len(df))

    if mapping["client"] is not None:
        client = df[mapping["client"]].astype(str).str.strip()
    else:
        client = pd.Series([""] * len(df))

    out = pd.DataFrame({
        "client": client,
        "code": df[mapping["code"]].astype(str).str.strip(),
        "date": pd.to_datetime(df[mapping["date"]], errors="coerce").dt.strftime("%Y-%m-%d"),
        "amount": amount,
    })
    out["has_amount"] = mapping["amount"] is not None
    return out


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/sample_template", methods=["GET"])
def sample_template():
    """Provide a sample .xlsx showing the expected columns/format."""
    sample = pd.DataFrame({
        "client": ["Acme Corp", "Beta Industries", "Acme Corp"],
        "code": ["CL001", "CL002", "CL001"],
        "date": ["2026-01-15", "2026-02-01", "2026-02-20"],
        "amount": [15000.00, 7250.50, 3000.00],
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Sample")
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="sample_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/compare", methods=["POST"])
def compare():
    amc_file = request.files.get("amc_file")
    our_file = request.files.get("our_file")

    if not amc_file or not our_file or amc_file.filename == "" or our_file.filename == "":
        flash("Please upload both files.")
        return redirect(url_for("index"))

    try:
        amc_df = load_and_standardize(amc_file)
        our_df = load_and_standardize(our_file)
    except Exception as e:
        flash(f"Error reading files: {e}")
        return redirect(url_for("index"))

    # Match on code + date (+ amount, if both files have it). Client is NOT used for
    # matching — only shown in the output for reference.
    both_have_amount = amc_df["has_amount"].iloc[0] and our_df["has_amount"].iloc[0]
    key_cols = ["code", "date", "amount"] if both_have_amount else ["code", "date"]

    amc_compare_df = amc_df.drop(columns=["has_amount"])
    our_compare_df = our_df.drop(columns=["has_amount"])
    if not both_have_amount:
        amc_compare_df = amc_compare_df.drop(columns=["amount"])
        our_compare_df = our_compare_df.drop(columns=["amount"])

    merged = amc_compare_df.merge(
        our_compare_df.drop_duplicates(subset=key_cols)[key_cols],
        on=key_cols,
        how="left",
        indicator=True,
    )
    missing = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        missing.to_excel(writer, index=False, sheet_name="Missing in Our File")
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="missing_in_our_file.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    # Local dev only. On Render, gunicorn (see Procfile) runs the app instead.
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
