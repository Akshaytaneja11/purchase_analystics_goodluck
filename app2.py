"""
Goodluck Engineering — Purchase Analytics & AI Insights Tool
=============================================================

A Streamlit dashboard that reads the monthly "Purchase MRN Report" Excel
export, visualizes spend trends, and uses the Groq API to generate
AI-written insights and answer free-form questions about the data.

HOW TO RUN
----------
1. Install dependencies:      pip install -r requirements.txt
2. Add your Groq API key (pick ONE option):
     a) Easiest: paste it into the sidebar text box when the app is running.
     b) Recommended: create a file .streamlit/secrets.toml with:
            GROQ_API_KEY = "gsk_your_key_here"
     c) Or set an environment variable:  export GROQ_API_KEY=gsk_your_key_here
3. Launch:                    streamlit run app.py
4. Upload the monthly purchase Excel file in the sidebar.

DATA FORMAT THIS APP EXPECTS
-----------------------------
The same layout as "Purchase MRN Report" exports:
  Row 1: month headers (dates) over each Quantity/Amount pair, "Grand Total"
         over the last pair.
  Row 2: column labels — Buy-from Vendor Name, Print Description, No_,
         Unit of Measure Code, then Quantity/Amount pairs per month.
  Row 3+: one line per (vendor, item) combination.

The parser below does NOT hardcode month names or the number of months —
it reads however many month columns are present, so the same file works
whether next month's export has 1 month or 12.
"""

import io
import os
import builtins
import re
import json
import warnings
from datetime import datetime, date

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------------
# Groq client is optional at import time — the app still loads if the
# package or key isn't set up yet, so the dashboard works standalone.
# ------------------------------------------------------------------
try:
    from groq import Groq
except ImportError:
    Groq = None


# =====================================================================
# PAGE CONFIG & THEME
# =====================================================================
st.set_page_config(
    page_title="Goodluck Purchase Analytics",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY = "#2C7DA0"      # steel blue
PRIMARY_DARK = "#014F86"
ACCENT = "#F4A261"       # warm accent for highlights
DANGER = "#E63946"
SUCCESS = "#2A9D8F"
BG_CARD = "#0F2438"
PALETTE = ["#2C7DA0", "#61A5C2", "#89C2D9", "#A9D6E5", "#014F86",
           "#F4A261", "#2A9D8F", "#E9C46A"]


def format_inr(value: float) -> str:
    """Formats a rupee amount in Indian Lakh/Crore style for compact display
    in metric cards, e.g. 26638210 -> '₹2.66 Cr', 93960680 -> '₹93.96 L'."""
    if value is None:
        return "₹0"
    sign = "-" if value < 0 else ""
    v = abs(value)
    if v >= 1_00_00_000:
        return f"{sign}₹{v/1_00_00_000:,.2f} Cr"
    elif v >= 1_00_000:
        return f"{sign}₹{v/1_00_000:,.2f} L"
    elif v >= 1_000:
        return f"{sign}₹{v/1_000:,.1f} K"
    return f"{sign}₹{v:,.0f}"

st.markdown(f"""
<style>
    .main {{ background-color: #0B1622; }}
    .block-container {{ padding-top: 1.5rem; }}

    h1, h2, h3 {{ color: #E8F1F8; }}

    div[data-testid="stMetric"] {{
        background: linear-gradient(135deg, {BG_CARD} 0%, #16324a 100%);
        border: 1px solid #1E3A52;
        border-radius: 12px;
        padding: 14px 16px 10px 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    }}
    div[data-testid="stMetricLabel"] {{
        color: #9FC5DB;
        white-space: normal;
        overflow: visible;
    }}
    div[data-testid="stMetricValue"] {{
        color: #FFFFFF;
        font-size: 1.5rem;
        white-space: normal;
        overflow: visible;
        text-overflow: unset;
        word-break: break-word;
    }}
    div[data-testid="stMetricDelta"] {{ white-space: normal; }}

    .insight-card {{
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 10px;
        font-size: 0.95rem;
        line-height: 1.4;
        border-left: 5px solid;
    }}
    .insight-trend {{ background: rgba(44,125,160,0.12); border-color: {PRIMARY}; }}
    .insight-anomaly {{ background: rgba(230,57,70,0.10); border-color: {DANGER}; }}
    .insight-reco {{ background: rgba(42,157,143,0.10); border-color: {SUCCESS}; }}

    section[data-testid="stSidebar"] {{
        background-color: #0F1E2E;
    }}
</style>
""", unsafe_allow_html=True)


# =====================================================================
# DATA PARSING
# =====================================================================
# Header labels (lower-cased, exact match) used to locate the identity columns.
# Columns are found by LABEL, not position, so extra columns (e.g. "Direct Unit
# Cost") or a different column order in future exports won't break the parser.
FIXED_COL_LABELS = {
    "vendor": ["buy-from vendor name", "vendor name", "vendor", "supplier", "supplier name"],
    "description": ["print description", "description", "item description", "item name"],
    "item_code": ["no_", "no.", "no", "item no", "item no.", "item no_", "item code", "item_code", "code"],
    "uom": ["unit of measure code", "unit of measure", "uom"],
}


def _norm(val) -> str:
    return "" if (not isinstance(val, str) and pd.isna(val)) else str(val).strip().lower()


def _as_period_date(val):
    """Return a Timestamp if `val` is a date-like header (real date or text such
    as 'Jul 2026'); None for blanks, 'Grand Total', or anything else."""
    if val is None or (not isinstance(val, str) and pd.isna(val)):
        return None
    if isinstance(val, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(val)
    if isinstance(val, str):
        s = val.strip()
        if not s or "total" in s.lower():
            return None
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            ts = pd.to_datetime(s, errors="coerce")
        return None if pd.isna(ts) else ts
    return None


@st.cache_data(show_spinner=False)
def parse_purchase_file(file_bytes: bytes, file_name: str):
    """
    Parses the wide MRN-style Excel export into a tidy long-format DataFrame:
    one row per (vendor, item, month) with quantity & amount.

    Layout-tolerant:
      * finds the header row by looking for 'Quantity' / 'Amount' labels
      * finds vendor / description / item code / UOM columns by label
      * ignores any extra columns (e.g. 'Direct Unit Cost')
      * reads however many month pairs exist; an optional 'Grand Total' pair
        and/or a bottom totals row are used only to validate the result
    """
    raw = pd.read_excel(io.BytesIO(file_bytes), header=None)
    if raw.shape[0] < 3 or raw.shape[1] < 4:
        raise ValueError("The sheet is too small to be a purchase report.")

    # --- 1. locate the header row (the one containing Quantity + Amount labels) ---
    header_row = None
    for i in range(min(15, len(raw))):
        labels_i = [_norm(v) for v in raw.iloc[i]]
        if "quantity" in labels_i and "amount" in labels_i:
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            "Couldn't find a header row containing 'Quantity' and 'Amount' columns "
            "in the first 15 rows of the sheet."
        )
    labels = [_norm(v) for v in raw.iloc[header_row]]
    measure_idx = [i for i, l in enumerate(labels) if l in ("quantity", "amount")]
    first_measure = min(measure_idx)

    # --- 2. locate identity columns by label ---
    fixed_cols = {}
    for idx, label in enumerate(labels):
        if label in ("quantity", "amount"):
            continue
        for key, names in FIXED_COL_LABELS.items():
            if key not in fixed_cols and label in names:
                fixed_cols[key] = idx
                break
    for idx, label in enumerate(labels):  # looser fallbacks
        if label in ("quantity", "amount"):
            continue
        if "vendor" not in fixed_cols and "vendor" in label:
            fixed_cols["vendor"] = idx
        elif "description" not in fixed_cols and "description" in label:
            fixed_cols["description"] = idx
        elif "uom" not in fixed_cols and "unit of measure" in label:
            fixed_cols["uom"] = idx
    for key, default in [("vendor", 0), ("description", 1), ("item_code", 2), ("uom", 3)]:
        if key not in fixed_cols and default < first_measure and default not in fixed_cols.values():
            fixed_cols[key] = default  # positional last resort

    if "vendor" not in fixed_cols:
        raise ValueError(f"Couldn't find a vendor column. Header labels seen: {labels}")
    if "description" not in fixed_cols and "item_code" not in fixed_cols:
        raise ValueError(f"Couldn't find an item description or item code column. Header labels seen: {labels}")

    # --- 3. find the month-header row (nearest non-empty row above the header row) ---
    period_row = None
    for i in range(header_row - 1, -1, -1):
        if raw.iloc[i].notna().any():
            period_row = i
            break

    # forward-fill period headers across the measure columns (handles merged cells)
    header_vals, last = {}, None
    for idx in range(first_measure, raw.shape[1]):
        v = raw.iat[period_row, idx] if period_row is not None else None
        blank = v is None or (not isinstance(v, str) and pd.isna(v)) or (isinstance(v, str) and not v.strip())
        if not blank:
            last = v
        header_vals[idx] = last

    # --- 4. pair Quantity/Amount columns into periods ---
    periods, unlabeled, j = [], 0, 0
    while j < len(measure_idx) - 1:
        qi, ai = measure_idx[j], measure_idx[j + 1]
        if labels[qi] == "quantity" and labels[ai] == "amount":
            hv = header_vals.get(qi)
            d = _as_period_date(hv)
            if d is not None:
                periods.append({"qty_idx": qi, "amt_idx": ai, "label": d.strftime("%b %Y"),
                                "order": d, "is_total": False})
            elif isinstance(hv, str) and "total" in hv.lower():
                periods.append({"qty_idx": qi, "amt_idx": ai, "label": "Grand Total",
                                "order": None, "is_total": True})
            else:  # no usable date header — keep the data, label it generically
                unlabeled += 1
                periods.append({"qty_idx": qi, "amt_idx": ai, "label": f"Period {unlabeled}",
                                "order": pd.Timestamp(2000, 1, 1) + pd.DateOffset(months=unlabeled - 1),
                                "is_total": False})
            j += 2
        else:
            j += 1

    real_periods = [p for p in periods if not p["is_total"]]
    total_period = next((p for p in periods if p["is_total"]), None)
    if not real_periods:
        raise ValueError("Found the header row but no Quantity/Amount month columns under it.")

    # --- 5. melt into long format ---
    data = raw.iloc[header_row + 1:].reset_index(drop=True)
    records, totals_row = [], None
    for _, r in data.iterrows():
        vendor_raw = r[fixed_cols["vendor"]]
        vendor = "" if pd.isna(vendor_raw) else str(vendor_raw).strip()
        if not vendor or vendor.lower() == "nan":
            # a row with numbers but no vendor is the sheet's own totals row
            if any(pd.notna(pd.to_numeric(r[p["amt_idx"]], errors="coerce")) for p in real_periods):
                totals_row = r
            continue

        desc_raw = r[fixed_cols["description"]] if "description" in fixed_cols else None
        code_raw = r[fixed_cols["item_code"]] if "item_code" in fixed_cols else None
        description = str(desc_raw).strip() if pd.notna(desc_raw) else None
        item_code = str(code_raw).strip() if pd.notna(code_raw) else "N/A"
        if not description:
            description = item_code  # fall back to item code if no description
        uom_raw = r[fixed_cols["uom"]] if "uom" in fixed_cols else None
        uom = str(uom_raw).strip() if pd.notna(uom_raw) else "NOS"

        for p in real_periods:
            amt = pd.to_numeric(r[p["amt_idx"]], errors="coerce")
            if pd.isna(amt):
                continue
            qty = pd.to_numeric(r[p["qty_idx"]], errors="coerce")
            records.append({
                "vendor": vendor,
                "description": description,
                "item_code": item_code,
                "uom": uom,
                "period": p["label"],
                "period_order": p["order"],
                "quantity": float(qty) if pd.notna(qty) else 0.0,
                "amount": float(amt),
            })

    df_long = pd.DataFrame(records)
    if df_long.empty:
        raise ValueError(
            "Header row and month columns were found, but no purchase line items had amounts. "
            f"Columns matched: {fixed_cols}; periods: {[p['label'] for p in periods]}."
        )
    df_long = df_long.sort_values("period_order")

    # --- 6. validate against the file's own totals (Grand Total column and/or bottom totals row) ---
    notes = []
    if total_period is not None:
        file_total = pd.to_numeric(data[total_period["amt_idx"]], errors="coerce").sum()
        computed_total = df_long["amount"].sum()
        if abs(file_total - computed_total) > max(1.0, 0.001 * file_total):
            notes.append(
                f"computed total ({computed_total:,.0f}) differs from the file's Grand Total "
                f"column ({file_total:,.0f})"
            )
    if totals_row is not None:
        for p in real_periods:
            file_val = pd.to_numeric(totals_row[p["amt_idx"]], errors="coerce")
            computed = df_long.loc[df_long["period"] == p["label"], "amount"].sum()
            if pd.notna(file_val) and abs(file_val - computed) > max(1.0, 0.001 * abs(file_val)):
                notes.append(f"{p['label']}: computed {computed:,.0f} vs file's total row {file_val:,.0f}")
    validation_note = (
        "Note: " + "; ".join(notes) + ". Check for merged cells or a non-standard layout."
        if notes else None
    )

    month_order = [p["label"] for p in sorted(real_periods, key=lambda p: p["order"])]
    return df_long, month_order, validation_note


# =====================================================================
# AGGREGATION HELPERS
# =====================================================================
def monthly_summary(df: pd.DataFrame, month_order: list) -> pd.DataFrame:
    g = df.groupby("period", as_index=False).agg(
        total_spend=("amount", "sum"),
        total_qty=("quantity", "sum"),
        line_items=("amount", "count"),
        vendors=("vendor", "nunique"),
    )
    g["period"] = pd.Categorical(g["period"], categories=month_order, ordered=True)
    return g.sort_values("period")


def vendor_summary(df: pd.DataFrame, top_n=None) -> pd.DataFrame:
    g = df.groupby("vendor", as_index=False).agg(
        total_spend=("amount", "sum"),
        line_items=("amount", "count"),
    ).sort_values("total_spend", ascending=False)
    g["share_pct"] = 100 * g["total_spend"] / g["total_spend"].sum()
    return g.head(top_n) if top_n else g


def item_summary(df: pd.DataFrame, top_n=None) -> pd.DataFrame:
    g = df.groupby("description", as_index=False).agg(
        total_spend=("amount", "sum"),
        total_qty=("quantity", "sum"),
        vendor_count=("vendor", "nunique"),
    ).sort_values("total_spend", ascending=False)
    return g.head(top_n) if top_n else g


def detect_anomalies(df: pd.DataFrame, month_order: list) -> list:
    """
    Simple, explainable anomaly detection done in plain pandas —
    computed here (not by the LLM) so the numbers are trustworthy.
    The AI's job is to explain these, not invent them.
    """
    anomalies = []

    # 1. Vendor month-over-month spend spikes
    piv = df.pivot_table(index="vendor", columns="period", values="amount", aggfunc="sum").fillna(0)
    piv = piv.reindex(columns=[m for m in month_order if m in piv.columns])
    if piv.shape[1] >= 2:
        for vendor in piv.index:
            series = piv.loc[vendor]
            for i in range(1, len(series)):
                prev, curr = series.iloc[i - 1], series.iloc[i]
                if prev > 0 and curr > prev * 3 and (curr - prev) > 0.02 * df["amount"].sum():
                    anomalies.append(
                        f"{vendor}: spend jumped from {prev:,.0f} to {curr:,.0f} "
                        f"({series.index[i-1]} → {series.index[i]}), a {(curr/prev - 1)*100:.0f}% increase."
                    )

    # 2. Item unit-price volatility (same item, very different price per unit across purchases)
    unit_df = df[df["quantity"] > 0].copy()
    unit_df["unit_price"] = unit_df["amount"] / unit_df["quantity"]
    for desc, grp in unit_df.groupby("description"):
        if len(grp) >= 3:
            cv = grp["unit_price"].std() / grp["unit_price"].mean() if grp["unit_price"].mean() else 0
            if cv and cv > 0.6:
                anomalies.append(
                    f"'{desc}': unit price varies widely across purchases "
                    f"(from {grp['unit_price'].min():,.1f} to {grp['unit_price'].max():,.1f} per unit)."
                )

    return anomalies[:10]  # cap so the prompt stays small


def price_trend_analysis(df: pd.DataFrame, month_order: list, min_purchases=2) -> pd.DataFrame:
    """
    For each item *within a single unit of measure*, tracks average unit
    price per month and flags items whose price has risen from their
    first to most recent purchase in this data.

    Grouping by (description, uom) together — not just description — is
    deliberate: the same item can appear under more than one UOM in the
    source data (e.g. "PVC PIPE 1/2\"" logged as both ROLL and MTR), and
    averaging across mismatched units produces false price-spike signals.
    Cross-UOM cases are reported separately by uom_inconsistencies().
    """
    unit_df = df[df["quantity"] > 0].copy()
    unit_df["unit_price"] = unit_df["amount"] / unit_df["quantity"]

    # Exclude rows with no real item description (the parser falls back to
    # the raw item/ledger code when the description is blank). Internal
    # codes like "CWIP-P&M/GLE/000052" are project/asset ledger codes that
    # get reused across many unrelated items — treating them as one
    # comparable "item" produces fake price gaps between totally
    # different products.
    unit_df = unit_df[unit_df["description"] != unit_df["item_code"]]
    unit_df = unit_df[~unit_df["description"].apply(is_generic_charge)]

    # Exclude items with ANY unit-of-measure inconsistency, even within
    # a single matching UOM subset — a mislabeled row can still land in
    # the "right" bucket by coincidence, so once an item's units are
    # known to be unreliable, treat none of its price data as trustworthy.
    flagged_items = set(uom_inconsistencies(df)["item"]) if not uom_inconsistencies(df).empty else set()
    unit_df = unit_df[~unit_df["description"].isin(flagged_items)]

    pivot = unit_df.pivot_table(index=["description", "uom"], columns="period",
                                 values="unit_price", aggfunc="mean")
    pivot = pivot.reindex(columns=[m for m in month_order if m in pivot.columns])

    rows = []
    for (desc, uom), series in pivot.iterrows():
        valid = series.dropna()
        if len(valid) < min_purchases:
            continue
        first_price, last_price = valid.iloc[0], valid.iloc[-1]
        if first_price <= 0:
            continue
        pct_change = (last_price - first_price) / first_price * 100
        rows.append({
            "item": desc,
            "uom": uom,
            "first_period": valid.index[0],
            "first_price": round(first_price, 2),
            "latest_period": valid.index[-1],
            "latest_price": round(last_price, 2),
            "pct_change": round(pct_change, 1),
            "months_tracked": len(valid),
        })

    result = pd.DataFrame(rows)
    return result.sort_values("pct_change", ascending=False) if not result.empty else result


GENERIC_CHARGE_KEYWORDS = [
    "service charge", "handling charge", "freight", "cartage",
    "labour charge", "labor charge", "transport", "loading charge",
    "unloading charge", "miscellaneous", "misc.", "other charge", "round off",
]


def is_generic_charge(description: str) -> bool:
    """True if the description is a generic charge/fee label rather than
    a specific, comparable product — e.g. 'SERVICE CHARGE' billed by four
    unrelated vendors for four unrelated services isn't the same 'item'
    just because the label matches."""
    d = str(description).lower()
    return any(kw in d for kw in GENERIC_CHARGE_KEYWORDS)


def uom_inconsistencies(df: pd.DataFrame) -> pd.DataFrame:
    """
    Flags items recorded under more than one unit of measure across
    purchases (e.g. the same item logged as both 'ROLL' and 'MTR').
    Any price comparison for these items should be treated with caution
    until the units are reconciled — this is a data-quality flag, not
    a price or savings finding.
    """
    rows = []
    for desc, grp in df.groupby("description"):
        uoms = sorted(grp["uom"].dropna().unique())
        if len(uoms) > 1:
            rows.append({
                "item": desc,
                "units_used": ", ".join(uoms),
                "unit_count": len(uoms),
                "purchase_count": len(grp),
                "total_spend": round(grp["amount"].sum(), 2),
            })
    result = pd.DataFrame(rows)
    return result.sort_values("total_spend", ascending=False) if not result.empty else result


def consolidation_opportunities(df: pd.DataFrame, min_vendors=2) -> pd.DataFrame:
    """
    Finds items bought from more than one vendor at different average
    unit prices — the gap between what was actually paid and what it
    would have cost at the cheapest vendor's price is the potential
    saving from consolidating that item to one supplier.

    Grouped by (description, uom) so a vendor selling by ROLL is never
    compared against one selling by MTR for the "same" item. Items with
    no real description (only an internal ledger/asset code) are
    excluded — see price_trend_analysis() for why.
    """
    unit_df = df[df["quantity"] > 0].copy()
    unit_df["unit_price"] = unit_df["amount"] / unit_df["quantity"]

    unit_df = unit_df[unit_df["description"] != unit_df["item_code"]]
    unit_df = unit_df[~unit_df["description"].apply(is_generic_charge)]

    flagged_items = set(uom_inconsistencies(df)["item"]) if not uom_inconsistencies(df).empty else set()
    unit_df = unit_df[~unit_df["description"].isin(flagged_items)]

    rows = []
    for (desc, uom), grp in unit_df.groupby(["description", "uom"]):
        vendor_prices = grp.groupby("vendor").agg(
            avg_price=("unit_price", "mean"),
            total_qty=("quantity", "sum"),
            total_spend=("amount", "sum"),
        )
        if vendor_prices.shape[0] < min_vendors:
            continue
        cheapest_vendor = vendor_prices["avg_price"].idxmin()
        cheapest_price = vendor_prices["avg_price"].min()
        total_qty = vendor_prices["total_qty"].sum()
        total_spend = vendor_prices["total_spend"].sum()
        potential_savings = total_spend - (cheapest_price * total_qty)
        if potential_savings > 1:  # ignore rounding-level noise
            rows.append({
                "item": desc,
                "uom": uom,
                "vendor_count": int(vendor_prices.shape[0]),
                "cheapest_vendor": cheapest_vendor,
                "cheapest_avg_price": round(cheapest_price, 2),
                "current_total_spend": round(total_spend, 2),
                "potential_savings": round(potential_savings, 2),
            })

    result = pd.DataFrame(rows)
    return result.sort_values("potential_savings", ascending=False) if not result.empty else result


def build_data_context(df: pd.DataFrame, month_order: list) -> dict:
    """Compact, numeric summary handed to the LLM — never the raw rows."""
    m = monthly_summary(df, month_order)
    v = vendor_summary(df, top_n=10)
    it = item_summary(df, top_n=10)
    mom = []
    for i in range(1, len(m)):
        prev, curr = m.iloc[i - 1], m.iloc[i]
        if prev["total_spend"] > 0:
            pct = (curr["total_spend"] - prev["total_spend"]) / prev["total_spend"] * 100
            mom.append({"from": str(prev["period"]), "to": str(curr["period"]), "change_pct": round(pct, 1)})

    return {
        "date_range": month_order,
        "total_spend": round(df["amount"].sum(), 2),
        "total_line_items": int(len(df)),
        "vendor_count": int(df["vendor"].nunique()),
        "item_count": int(df["description"].nunique()),
        "monthly_totals": [
            {"period": str(r.period), "spend": round(r.total_spend, 2), "vendors": int(r.vendors)}
            for r in m.itertuples()
        ],
        "month_over_month_pct_change": mom,
        "top_10_vendors_by_spend": [
            {"vendor": r.vendor, "spend": round(r.total_spend, 2), "share_pct": round(r.share_pct, 1)}
            for r in v.itertuples()
        ],
        "top_10_items_by_spend": [
            {"item": r.description, "spend": round(r.total_spend, 2), "qty": round(r.total_qty, 1)}
            for r in it.itertuples()
        ],
        "detected_anomalies": detect_anomalies(df, month_order),
        "items_with_rising_unit_price": [
            {"item": r.item, "uom": r.uom, "from_price": r.first_price, "to_price": r.latest_price,
             "pct_change": r.pct_change}
            for r in price_trend_analysis(df, month_order).head(5).itertuples()
        ] if not price_trend_analysis(df, month_order).empty else [],
        "top_consolidation_opportunities": [
            {"item": r.item, "uom": r.uom, "vendor_count": r.vendor_count, "cheapest_vendor": r.cheapest_vendor,
             "potential_savings": r.potential_savings}
            for r in consolidation_opportunities(df).head(5).itertuples()
        ] if not consolidation_opportunities(df).empty else [],
        "items_with_inconsistent_units_of_measure": [
            {"item": r.item, "units_used": r.units_used, "total_spend": r.total_spend}
            for r in uom_inconsistencies(df).head(10).itertuples()
        ] if not uom_inconsistencies(df).empty else [],
        "note_for_ai": (
            "Price trend and consolidation figures are already computed within a single "
            "unit of measure per item, so they are safe to quote. Items listed under "
            "'items_with_inconsistent_units_of_measure' were excluded from those price "
            "comparisons because their unit of measure is inconsistent in the source data — "
            "flag these as a data-quality issue to fix, not as a price or savings finding."
        ),
    }


# =====================================================================
# GROQ CLIENT
# =====================================================================
def get_groq_client():
    key = st.session_state.get("groq_api_key") or os.environ.get("GROQ_API_KEY")
    if not key:
        try:
            key = st.secrets.get("GROQ_API_KEY")
        except Exception:
            key = None
    if not key or Groq is None:
        return None
    return Groq(api_key=key)


def call_groq(client, system_prompt: str, user_prompt: str, model: str, json_mode=False, temperature=0.3, max_tokens=1500):
    kwargs = dict(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    resp = client.chat.completions.create(**kwargs)
    return resp.choices[0].message.content or ""


@st.cache_data(show_spinner=False)
def generate_ai_insights(context_json: str, model: str, _client_key: str) -> dict:
    """Cached on the data context + model, so re-rendering doesn't re-call the API."""
    client = get_groq_client()
    system_prompt = (
        "You are an experienced procurement/finance analyst. You will be given a JSON "
        "summary of a company's purchase data (monthly totals, top vendors, top items, "
        "and pre-computed anomalies). Write insights ONLY from these numbers — never "
        "invent figures that aren't present. "
        "Respond with STRICT JSON only, no markdown fences, in this exact shape: "
        '{"trends": [string, ...], "anomalies": [string, ...], "recommendations": [string, ...]}. '
        "Each list should have 3-5 short, concrete, leadership-readable bullet points. "
        "For 'anomalies', explain the pre-computed anomalies given to you in plain business "
        "language (why it might matter), rather than restating raw numbers."
    )
    raw = call_groq(client, system_prompt, context_json, model, json_mode=True, temperature=0.2)
    raw = re.sub(r"^```json|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    return json.loads(raw)


# ---------------------------------------------------------------------
# CHAT ENGINE
# Step 1 (deterministic, in Python): find the items / vendors / months the
#        question is about and compute exact breakdowns from the full data.
# Step 2 (Groq): explain those numbers in plain business language.
# Only if step 1 finds nothing does the app ask Groq to write a pandas
# expression as a fallback for analytical questions.
# ---------------------------------------------------------------------
_STOPWORDS = set("""
a an the is are was were be been am do does did has have had i we you our us me my your
what which who whom whose when where why how much many more most less least highest lowest
biggest largest top total spend spent spending purchase purchased purchases bought buy buying
paid pay cost costs price prices rate rates value amount amounts vendor vendors supplier
suppliers wise by per of on in for to from with and or at as it its this that these those
item items product products material materials show give tell list find get compare
comparison breakdown details detail please any all each between about than then there their
across during month months monthly year quarter
""".split())

_MONTH_WORDS = {
    "jan": "Jan", "january": "Jan", "feb": "Feb", "february": "Feb", "mar": "Mar", "march": "Mar",
    "apr": "Apr", "april": "Apr", "may": "May", "jun": "Jun", "june": "Jun", "jul": "Jul",
    "july": "Jul", "aug": "Aug", "august": "Aug", "sep": "Sep", "sept": "Sep", "september": "Sep",
    "oct": "Oct", "october": "Oct", "nov": "Nov", "november": "Nov", "dec": "Dec", "december": "Dec",
}


def _stem(tok: str) -> str:
    return tok[:-1] if len(tok) > 3 and tok.endswith("s") else tok


def _name_tokens(text) -> set:
    return {_stem(t) for t in re.findall(r"[a-z0-9]+", str(text).lower())}


def _match_names(q_tokens: list, names) -> list:
    """Names (item descriptions or vendor names) that contain the most of the
    question's keywords. Needs at least 2 matching keywords (or 1 if the
    question only has 1) and at least half of them; returns the best tier."""
    qset = set(q_tokens)
    if not qset:
        return []
    need = min(2, len(qset))
    scored = []
    for name in names:
        nt = _name_tokens(name)
        hit = len(qset & nt)
        if hit >= need and hit / len(qset) >= 0.5:
            scored.append((hit / len(qset), -len(nt), name))
    if not scored:
        return []
    best = max(s[0] for s in scored)
    return [n for s, _, n in sorted(scored, reverse=True) if s == best]


def _clean(v):
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return None
    if isinstance(v, (np.floating, float)):
        return round(float(v), 2)
    if isinstance(v, (np.integer,)):
        return int(v)
    return v


def _records(frame: pd.DataFrame) -> list:
    return [{k: _clean(v) for k, v in row.items()} for row in frame.to_dict("records")]


def build_lookup(question: str, df: pd.DataFrame, month_order: list) -> dict:
    """Exact, pre-computed answers for the items / vendors / months named in
    the question. Everything here comes straight from the uploaded data."""
    raw_toks = re.findall(r"[a-z0-9]+", question.lower())
    abbrs = {_MONTH_WORDS[t] for t in raw_toks if t in _MONTH_WORDS}
    months = [m for m in month_order if m.split()[0] in abbrs]
    kw = [_stem(t) for t in raw_toks if t not in _STOPWORDS and t not in _MONTH_WORDS]

    scope_df = df[df["period"].isin(months)] if months else df
    out = {"scope": ", ".join(months) if months else "all months in the file"}

    # --- items (by description words, or by an item code quoted in the question) ---
    item_names = _match_names(kw, df["description"].unique())
    q_lower = question.lower()
    for code in df["item_code"].unique():
        if len(code) >= 5 and code.lower() in q_lower:
            item_names += list(df.loc[df["item_code"] == code, "description"].unique())
    item_names = list(dict.fromkeys(item_names))
    more_items = 0
    if len(item_names) > 8:
        ranked = df[df["description"].isin(item_names)].groupby("description")["amount"].sum()
        keep = list(ranked.sort_values(ascending=False).index[:8])
        more_items, item_names = len(item_names) - 8, keep

    if item_names:
        sub = scope_df[scope_df["description"].isin(item_names)]
        block = {"descriptions_matched": item_names, "other_matches_not_shown": more_items}
        if sub.empty:
            block["note"] = f"These items exist in the file but have no purchases in: {out['scope']}."
        else:
            bv = sub.groupby(["vendor", "uom"], as_index=False).agg(
                spend=("amount", "sum"), qty=("quantity", "sum"), purchases=("amount", "count"))
            bv["avg_unit_price"] = np.where(bv["qty"] > 0, bv["spend"] / bv["qty"].where(bv["qty"] > 0, 1), np.nan)
            bv = bv.sort_values("spend", ascending=False)
            bm = sub.groupby("period", as_index=False).agg(spend=("amount", "sum"), qty=("quantity", "sum"))
            bm["period"] = pd.Categorical(bm["period"], categories=month_order, ordered=True)
            bm = bm.sort_values("period").astype({"period": str})
            block.update({
                "total_spend": _clean(sub["amount"].sum()),
                "total_qty": _clean(sub["quantity"].sum()),
                "vendor_count": int(sub["vendor"].nunique()),
                "by_vendor": _records(bv.head(15)),
                "by_month": _records(bm),
            })
        out["matched_items"] = block

    # --- vendors (by name words) ---
    vendor_names = _match_names(kw, df["vendor"].unique())
    if vendor_names:
        ranked_v = df[df["vendor"].isin(vendor_names)].groupby("vendor")["amount"].sum().sort_values(ascending=False)
        vblocks = []
        for v in list(ranked_v.index[:3]):
            sv = scope_df[scope_df["vendor"] == v]
            if sv.empty:
                vblocks.append({"vendor": v, "note": f"No purchases from this vendor in: {out['scope']}."})
                continue
            vm = sv.groupby("period", as_index=False)["amount"].sum().rename(columns={"amount": "spend"})
            vm["period"] = pd.Categorical(vm["period"], categories=month_order, ordered=True)
            vm = vm.sort_values("period").astype({"period": str})
            vi = (sv.groupby(["description", "uom"], as_index=False)
                    .agg(spend=("amount", "sum"), qty=("quantity", "sum"))
                    .sort_values("spend", ascending=False).head(10))
            vblocks.append({"vendor": v, "total_spend": _clean(sv["amount"].sum()),
                            "by_month": _records(vm), "top_items": _records(vi)})
        out["matched_vendors"] = vblocks

    # --- month views (when months are named but nothing more specific matched) ---
    if months and "matched_items" not in out and "matched_vendors" not in out:
        views = []
        for m in months[:3]:
            md = df[df["period"] == m]
            tv = md.groupby("vendor", as_index=False)["amount"].sum().rename(columns={"amount": "spend"}) \
                   .sort_values("spend", ascending=False).head(10)
            ti = md.groupby("description", as_index=False)["amount"].sum().rename(columns={"amount": "spend"}) \
                   .sort_values("spend", ascending=False).head(10)
            views.append({"period": m, "total_spend": _clean(md["amount"].sum()),
                          "top_vendors": _records(tv), "top_items": _records(ti)})
        out["month_view"] = views

    if kw and not any(k in out for k in ("matched_items", "matched_vendors")):
        out["note"] = ("No item description or vendor name in the uploaded file matched these "
                       f"keywords: {kw}.")
    return out


_SAFE_BUILTINS = {n: getattr(builtins, n) for n in (
    "len", "round", "sum", "min", "max", "abs", "sorted", "str", "int", "float",
    "list", "dict", "set", "range", "enumerate", "zip", "any", "all")}


def _clean_code(text: str) -> str:
    text = (text or "").strip()
    text = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", text).strip()
    return text.strip("`").strip()


def _run_pandas_expr(code: str, df: pd.DataFrame):
    """Evaluate one pandas expression against df.
    Returns (result_text_or_None, error_text_or_None)."""
    forbidden = ["import", "__", "open(", "exec", "eval(", "os.", "sys.", ";", "lambda"]
    if not code or code.upper() == "NONE":
        return None, "no expression was produced"
    if any(tok in code.lower() for tok in forbidden):
        return None, "expression blocked by the safety filter"
    try:
        result = eval(code, {"__builtins__": _SAFE_BUILTINS}, {"df": df, "pd": pd, "np": np})  # noqa: S307
    except Exception as e:
        return None, f"{type(e).__name__}: {e}"
    if isinstance(result, (pd.DataFrame, pd.Series)):
        return (None, "expression returned no rows") if len(result) == 0 else (result.head(25).to_string(), None)
    return str(result), None


def answer_chat_question(question: str, df: pd.DataFrame, context: dict, model: str, history: list):
    """Returns (answer_text, debug_info)."""
    client = get_groq_client()
    if client is None:
        return "⚠️ No Groq API key configured yet — add one in the sidebar to enable AI answers.", {}

    month_order = context.get("date_range", [])
    lookup = build_lookup(question, df, month_order)
    has_lookup = any(k in lookup for k in ("matched_items", "matched_vendors", "month_view"))
    debug = {"months_in_scope": lookup["scope"], "lookup_found_data": has_lookup,
             "items_matched": lookup.get("matched_items", {}).get("descriptions_matched", []),
             "vendors_matched": [v["vendor"] for v in lookup.get("matched_vendors", [])]}

    # Fallback for analytical questions that name no specific item/vendor/month:
    # ask Groq for a single pandas expression, run it, and report the result.
    computed_result = None
    if not has_lookup:
        code_system = (
            "You convert a business question into a SINGLE python expression (no assignments, no "
            "imports) that computes the answer from the pandas DataFrame `df` (and `pd`, `np`). "
            "Columns: vendor (str), description (str), item_code (str), uom (str), period (str, "
            "e.g. 'Jul 2026'), period_order (datetime), quantity (float), amount (float, rupees). "
            "Item names and vendor names are messy strings: for any text lookup use "
            "df['description'].str.contains('keyword', case=False, na=False) — never ==. "
            "Respond with ONLY the expression, no explanation, no markdown. If it cannot be "
            "answered from this data, respond with: NONE"
        )
        try:
            code = _clean_code(call_groq(client, code_system, question, model, temperature=0.0, max_tokens=1024))
            computed_result, err = _run_pandas_expr(code, df)
            debug.update({"generated_expression": code, "expression_error": err})
        except Exception as e:
            debug["expression_error"] = f"Groq call failed: {e}"

    final_system = (
        "You are a procurement data assistant for the leadership team of an Indian manufacturing "
        "company. Answer the question directly and briefly using ONLY the data provided. "
        "'lookup' holds exact figures computed from the uploaded file for the items, vendors or "
        "months in the question — treat it as authoritative and quote its numbers as given. "
        "'computed_result' (if present) is the output of a query run on the data. "
        "'data_summary_context' is a high-level summary. Write rupee amounts in Indian style "
        "(₹X.XX lakh or ₹X.XX crore for large values, plain ₹ with commas for small ones). "
        "For vendor-wise questions, list each vendor with its amount, largest first. "
        "If the lookup says nothing in the file matched, say plainly that it isn't in this file — "
        "never invent figures. 'scope' says which months the lookup covers; mention it if it is "
        "not all months."
    )
    recent = [{"role": m["role"], "content": m["content"]} for m in history[-7:-1]]
    payload = {"question": question, "recent_conversation": recent, "lookup": lookup,
               "computed_result": computed_result, "data_summary_context": context}
    try:
        answer = call_groq(client, final_system, json.dumps(payload, default=str), model,
                           temperature=0.2, max_tokens=1500)
    except Exception as e:
        return f"⚠️ Groq API error: {e}", debug
    return answer, {**debug, "lookup": lookup}


# =====================================================================
# SIDEBAR
# =====================================================================
with st.sidebar:
    st.markdown("## ⚙️ Goodluck Purchase Analytics")
    st.caption("Upload the monthly MRN purchase report to begin.")

    uploaded_file = st.file_uploader("Purchase report (.xlsx)", type=["xlsx"])

    st.markdown("---")
    st.markdown("#### AI Configuration (Groq)")
    default_key = os.environ.get("GROQ_API_KEY", "")
    api_key_input = st.text_input(
        "Groq API Key", value=st.session_state.get("groq_api_key", default_key),
        type="password", help="Get a key at console.groq.com/keys. Stored only for this session."
    )
    st.session_state["groq_api_key"] = api_key_input

    model_choice = st.selectbox(
        "Groq model",
        ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "openai/gpt-oss-120b"],
        index=0,
        help="70b-versatile is the recommended general-purpose model. 8b-instant is faster/cheaper.",
    )

    if Groq is None:
        st.warning("`groq` package not installed. Run: pip install groq")
    elif not api_key_input:
        st.info("Paste your Groq API key above to enable AI Insights & Chat.")
    else:
        st.success("Groq configured ✓")


# =====================================================================
# MAIN
# =====================================================================
st.title("Purchase Analytics & AI Insights")

if uploaded_file is None:
    st.info("👈 Upload a purchase report Excel file from the sidebar to get started.")
    st.stop()

try:
    df, month_order, validation_note = parse_purchase_file(uploaded_file.getvalue(), uploaded_file.name)
except Exception as e:
    st.error(f"Couldn't read this file: {e}")
    st.stop()

if validation_note:
    st.warning(validation_note)

tab_dashboard, tab_insights, tab_savings, tab_chat = st.tabs(
    ["📊 Dashboard", "🧠 AI Insights", "💰 Price & Savings", "💬 AI Chat"]
)

# ---------------------------------------------------------------
# TAB 1: DASHBOARD
# ---------------------------------------------------------------
with tab_dashboard:
    m = monthly_summary(df, month_order)
    total_spend = df["amount"].sum()
    latest_spend = m.iloc[-1]["total_spend"] if len(m) else 0
    prev_spend = m.iloc[-2]["total_spend"] if len(m) > 1 else None
    mom_delta = f"{(latest_spend/prev_spend - 1)*100:+.1f}% vs prior month" if prev_spend else None

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Spend", format_inr(total_spend))
    c2.metric(f"Latest Month ({m.iloc[-1]['period']})" if len(m) else "Latest Month",
              format_inr(latest_spend), mom_delta)
    c3.metric("Vendors", f"{df['vendor'].nunique():,}")
    c4.metric("Distinct Items", f"{df['description'].nunique():,}")
    c5.metric("Line Items", f"{len(df):,}")

    st.markdown("###  ")
    col_a, col_b = st.columns([3, 2])

    with col_a:
        fig = px.bar(m, x="period", y="total_spend", text_auto=".2s",
                     title="Monthly Purchase Spend", color_discrete_sequence=[PRIMARY])
        fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                           font_color="#E8F1F8", xaxis_title="", yaxis_title="Spend (₹)")
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        v10 = vendor_summary(df, top_n=8)
        others = total_spend - v10["total_spend"].sum()
        pie_labels = list(v10["vendor"]) + (["Others"] if others > 0 else [])
        pie_values = list(v10["total_spend"]) + ([others] if others > 0 else [])
        fig2 = go.Figure(data=[go.Pie(labels=pie_labels, values=pie_values, hole=0.45,
                                       marker=dict(colors=PALETTE))])
        fig2.update_layout(title="Spend Share by Vendor (Top 8)",
                            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            font_color="#E8F1F8", legend=dict(font=dict(size=10)))
        st.plotly_chart(fig2, use_container_width=True)

    col_c, col_d = st.columns(2)
    with col_c:
        v_top = vendor_summary(df, top_n=10).sort_values("total_spend")
        fig3 = px.bar(v_top, x="total_spend", y="vendor", orientation="h",
                      title="Top 10 Vendors by Spend", color_discrete_sequence=[PRIMARY_DARK])
        fig3.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            font_color="#E8F1F8", xaxis_title="Spend (₹)", yaxis_title="")
        st.plotly_chart(fig3, use_container_width=True)

    with col_d:
        it_top = item_summary(df, top_n=10).sort_values("total_spend")
        fig4 = px.bar(it_top, x="total_spend", y="description", orientation="h",
                      title="Top 10 Items / Materials by Spend", color_discrete_sequence=[ACCENT])
        fig4.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            font_color="#E8F1F8", xaxis_title="Spend (₹)", yaxis_title="")
        st.plotly_chart(fig4, use_container_width=True)

    with st.expander("🔍 Explore raw data"):
        vendor_filter = st.multiselect("Filter by vendor", sorted(df["vendor"].unique()))
        period_filter = st.multiselect("Filter by month", month_order)
        filtered = df.copy()
        if vendor_filter:
            filtered = filtered[filtered["vendor"].isin(vendor_filter)]
        if period_filter:
            filtered = filtered[filtered["period"].isin(period_filter)]
        st.dataframe(
            filtered[["vendor", "description", "item_code", "uom", "period", "quantity", "amount"]],
            use_container_width=True, height=350,
        )

# ---------------------------------------------------------------
# TAB 2: AI INSIGHTS
# ---------------------------------------------------------------
with tab_insights:
    st.markdown("Generates trends, anomalies, and recommendations from the current data using Groq.")
    context = build_data_context(df, month_order)

    if st.button("🔄 Generate / Refresh Insights", type="primary"):
        client = get_groq_client()
        if client is None:
            st.error("Add a Groq API key in the sidebar first.")
        else:
            with st.spinner("Analyzing purchase data..."):
                try:
                    insights = generate_ai_insights(
                        json.dumps(context, default=str), model_choice,
                        st.session_state.get("groq_api_key", "")[:6],  # cache key fragment, not the whole secret
                    )
                    st.session_state["last_insights"] = insights
                except Exception as e:
                    st.error(f"Couldn't generate insights: {e}")

    insights = st.session_state.get("last_insights")
    if insights:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📈 Trends")
            for t in insights.get("trends", []):
                st.markdown(f'<div class="insight-card insight-trend">{t}</div>', unsafe_allow_html=True)
            st.markdown("#### 🚨 Anomalies")
            for a in insights.get("anomalies", []):
                st.markdown(f'<div class="insight-card insight-anomaly">{a}</div>', unsafe_allow_html=True)
        with col2:
            st.markdown("#### ✅ Recommendations")
            for r in insights.get("recommendations", []):
                st.markdown(f'<div class="insight-card insight-reco">{r}</div>', unsafe_allow_html=True)
    else:
        st.caption("Click the button above to generate the first set of insights.")

# ---------------------------------------------------------------
# TAB 3: PRICE TREND & CONSOLIDATION SAVINGS
# ---------------------------------------------------------------
with tab_savings:
    st.markdown("### ⚠️ Unit of Measure Inconsistencies")
    st.caption("Items logged under more than one unit of measure across purchases (e.g. the same "
               "item recorded as both ROLL and MTR). These are excluded from the price trend and "
               "consolidation analysis below until reconciled, since comparing prices across "
               "different units produces false price-swing signals.")

    uom_issues = uom_inconsistencies(df)
    if uom_issues.empty:
        st.success("No unit-of-measure inconsistencies detected — all items use a consistent unit.")
    else:
        st.dataframe(uom_issues, use_container_width=True, height=min(300, 45 + 35 * len(uom_issues)))

    st.markdown("---")
    st.markdown("### 📈 Price Trend Watch")
    st.caption("Items whose average unit price rose between their first and most recent "
               "purchase in this data — catches price creep even when total spend looks flat. "
               "Compared strictly within the same unit of measure (see flags above).")

    trend_df = price_trend_analysis(df, month_order)
    if trend_df.empty:
        st.info("Not enough repeat purchases of the same item yet to track price trends "
                 "(need at least 2 months of data for the same item, same unit).")
    else:
        trend_df["item_label"] = trend_df["item"] + " (" + trend_df["uom"] + ")"
        rising = trend_df[trend_df["pct_change"] > 5].head(15)
        if rising.empty:
            st.success("No significant price increases detected across repeat purchases.")
        else:
            fig5 = px.bar(rising.sort_values("pct_change"), x="pct_change", y="item_label",
                          orientation="h", text_auto=".1f",
                          title="Items with Rising Unit Price (%)", color_discrete_sequence=[DANGER])
            fig5.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                                font_color="#E8F1F8", xaxis_title="% change in unit price", yaxis_title="")
            st.plotly_chart(fig5, use_container_width=True)
        with st.expander("See full price trend table (all items, rising or falling)"):
            st.dataframe(trend_df.drop(columns="item_label"), use_container_width=True, height=300)

    st.markdown("---")
    st.markdown("### 💰 Consolidation Opportunities")
    st.caption("Items bought from more than one vendor at different average prices — "
               "the gap versus the cheapest vendor's price is the potential saving "
               "from consolidating that item to a single supplier. Compared strictly "
               "within the same unit of measure.")

    consol_df = consolidation_opportunities(df)
    if consol_df.empty:
        st.info("No clear multi-vendor consolidation opportunities found in this data.")
    else:
        consol_df["item_label"] = consol_df["item"] + " (" + consol_df["uom"] + ")"
        total_potential = consol_df["potential_savings"].sum()
        st.metric("Total Estimated Potential Savings", format_inr(total_potential))

        top_savings = consol_df.head(10)
        fig6 = px.bar(top_savings.sort_values("potential_savings"), x="potential_savings", y="item_label",
                      orientation="h", text_auto=".2s",
                      title="Top Potential Savings by Item", color_discrete_sequence=[SUCCESS])
        fig6.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                            font_color="#E8F1F8", xaxis_title="Potential Savings (₹)", yaxis_title="")
        st.plotly_chart(fig6, use_container_width=True)
        with st.expander("See full consolidation opportunities table"):
            st.dataframe(consol_df.drop(columns="item_label"), use_container_width=True, height=300)

# ---------------------------------------------------------------
# TAB 4: AI CHAT
# ---------------------------------------------------------------
with tab_chat:
    st.markdown("Ask anything about this data — spend by vendor, month, item, trends, or comparisons.")

    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    def _show_debug(debug):
        if debug:
            with st.expander("🔍 How this answer was found"):
                st.caption(f"Months covered: {debug.get('months_in_scope', '—')}")
                if debug.get("items_matched"):
                    st.write("**Items matched:**", ", ".join(debug["items_matched"]))
                if debug.get("vendors_matched"):
                    st.write("**Vendors matched:**", ", ".join(debug["vendors_matched"]))
                if debug.get("generated_expression") is not None:
                    st.code(debug["generated_expression"] or "(none)", language="python")
                if debug.get("expression_error"):
                    st.warning(debug["expression_error"])
                if debug.get("lookup"):
                    st.json(debug["lookup"], expanded=False)

    for msg in st.session_state["chat_history"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])
            if msg["role"] == "assistant":
                _show_debug(msg.get("debug"))

    question = st.chat_input("e.g. Which vendor did we spend the most with in July?")
    if question:
        st.session_state["chat_history"].append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)

        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                context = build_data_context(df, month_order)
                answer, debug = answer_chat_question(
                    question, df, context, model_choice, st.session_state["chat_history"]
                )
                st.markdown(answer)
                _show_debug(debug)
        st.session_state["chat_history"].append({"role": "assistant", "content": answer, "debug": debug})

    if st.session_state["chat_history"]:
        if st.button("Clear chat"):
            st.session_state["chat_history"] = []
            st.rerun()