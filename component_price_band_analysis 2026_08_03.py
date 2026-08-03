"""
Price Band / ASP Benchmarking Analysis  (Dual-Benchmark Version)
=================================================================
Identifies pricing inconsistencies and revenue-uplift opportunities
by comparing each transaction against two complementary benchmarks:

  Market benchmark  – ASP percentiles across ALL customers.
  Peer benchmark    – ASP percentiles among customers with the same
                      GPO, Account Type, and Revenue Tier (cascading
                      to broader groups when the peer pool is too
                      small).

Peer-group cascade (most specific first):
  1. GPO + Account Type + Revenue Tier  (if >= MIN_PEER_GROUP customers)
  2. GPO + Account Type                 (fallback)
  3. GPO only                           (fallback)
  4. Account Type only                  (fallback)
  5. Market                             (final fallback)

Usage:  python <script>.py [path_to_file.xlsx]
        If no path is given a file-picker dialog opens.
"""

import sys, os, subprocess, warnings, datetime
from pathlib import Path

# -- Dependency bootstrap ------------------------------------------------
_PKGS = {"pandas": "pandas", "openpyxl": "openpyxl", "xlsxwriter": "xlsxwriter"}
def _ensure():
    miss = [p for i, p in _PKGS.items() if not __import__(i, fromlist=["_"]) is None and False]
    miss = []
    for i, p in _PKGS.items():
        try: __import__(i)
        except ImportError: miss.append(p)
    if not miss: return
    print(f"Installing {', '.join(miss)} ...")
    try: subprocess.check_call([sys.executable, "-m", "pip", "install", "-q"] + miss)
    except subprocess.CalledProcessError:
        sys.exit(f"pip install failed. Run:  pip install {' '.join(miss)}")
_ensure()

import pandas as pd
import numpy as np
from difflib import get_close_matches

# ═══════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════
NUM_TIERS       = 3       # Revenue tiers (Tier 1 = largest)
MIN_PEER_GROUP  = 10       # Min distinct customers in a peer group;
                          # groups below this cascade to the next level.

# Source column names for the two new dimensions
GPO_COL       = "GPO / Buying Grp 1"
ACCT_TYPE_COL = "Account Type"


# ═══════════════════════════════════════════════════════════════════
# COLUMN NAME RESOLUTION – handles common variations
# ═══════════════════════════════════════════════════════════════════

# Canonical name -> list of known alternative spellings / abbreviations.
# All comparisons are done after normalisation (lowercase, whitespace-
# collapsed, underscores replaced with spaces), so only genuinely
# different *words* need to be listed here.
COLUMN_ALIASES: dict[str, list[str]] = {
    "Business Partner": [
            "business partner name",
            "bp", "customer", "customer name",
        ],
        "Account Type": [
            "acct type", "accttype", "account category",
        ],
        "GPO / Buying Grp 1": [
            "gpo", "buying group", "buying grp",
        ],
        "Product Group": [
            "prod group", "prod grp", "product grp",
        ],
        "Market Segment": [
            "mkt segment", "market seg",
        ],
        "Material Group": [
            "mat group", "matl group", "material grp", "mat grp",
        ],
        "Pricing Component Type": [
            "product type", "US Pricing Component Type", "component type", "comp type",
        ],
        "Pricing Component": [
            "L5 Product Brand", "US Pricing Component", "price component",
        ],
        "Selling UOM Unit Qty": [
            "Selling UOM Invoice Qty", "selling uom unit quantity", "selling uom qty",
            "uom unit qty", "uom unit quantity", "uom qty",
            "unit qty", "unit quantity", "units", "quantity",
        ],
        "Discount %": [
            "discount pct", "discount percent", "disc %", "disc pct",
            "discount rate",
        ],
        "Invoice ASP": [
            "inv asp", "asp", "avg selling price", "average selling price",
        ],
        "Invoice Total": [
            "inv total", "invoice amount", "total invoice",
        ],
        "Standard COGS": [
            "std cogs", "cogs", "cost of goods sold", "standard cost",
        ],
        "Std Margin": [
            "standard margin", "standard margin $", "margin",
        ],
        "Margin %": [
            "margin pct", "margin %", "margin percent", "margin rate",
        ],
}


def _normalise(name: str) -> str:
    """Lowercase, strip, replace underscores/hyphens with spaces, collapse
    consecutive spaces and strip trailing/leading whitespace."""
    s = str(name).strip().lower()
    s = s.replace("_", " ").replace("-", " ")
    while "  " in s:
        s = s.replace("  ", " ")
    return s


def _build_alias_lookup() -> dict[str, str]:
    """Return a dict mapping every normalised alias to its canonical name.
    The canonical name's own normalised form is also included so that
    simple case / whitespace differences are resolved automatically."""
    lookup: dict[str, str] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        lookup[_normalise(canonical)] = canonical
        for alias in aliases:
            norm = _normalise(alias)
            if norm not in lookup:           # first-registered wins
                lookup[norm] = canonical
    return lookup

_ALIAS_LOOKUP = _build_alias_lookup()


def _resolve_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename DataFrame columns that are not exact canonical matches but
    can be identified via normalisation or known aliases.

    Returns the DataFrame with renamed columns and prints each mapping
    applied so the user can verify correctness.
    """
    canonical_set = set(COLUMN_ALIASES.keys())
    rename_map: dict[str, str] = {}
    already_taken: set[str] = set(df.columns) & canonical_set  # already correct

    for col in df.columns:
        if col in canonical_set:
            continue                         # exact match, no rename needed
        norm = _normalise(col)
        candidate = _ALIAS_LOOKUP.get(norm)
        if candidate and candidate not in already_taken:
            rename_map[col] = candidate
            already_taken.add(candidate)

    if rename_map:
        print("  Column-name variations resolved:")
        for old, new in rename_map.items():
            print(f"    '{old}'  →  '{new}'")
        df = df.rename(columns=rename_map)

    return df


def _suggest_matches(missing: list[str], available: list[str]) -> str:
    """For each missing column, suggest close matches from the available
    columns to help the user fix their input file."""
    lines: list[str] = []
    avail_lower = [c.lower() for c in available]
    for col in missing:
        close = get_close_matches(col.lower(), avail_lower, n=3, cutoff=0.5)
        if close:
            suggestions = [available[avail_lower.index(c)] for c in close]
            lines.append(f"  • '{col}'  – did you mean: {suggestions}?")
        else:
            lines.append(f"  • '{col}'  – no close match found")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════
# EXCLUSION FILTERS – add values to exclude matching rows
# ═══════════════════════════════════════════════════════════════════
EXCLUDE_BUSINESS_PARTNER  = []
EXCLUDE_PRICING_COMPONENT = ["ROSA BRAIN - SERVICE"]
EXCLUDE_COMPONENT_TYPE    = []
EXCLUDE_MATERIAL_GROUP    = []
EXCLUDE_PRODUCT_GROUP     = []
EXCLUDE_GPO               = []
EXCLUDE_ACCOUNT_TYPE      = []

# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════
def _pct(s, q):
    """Percentile that handles empty / all-NaN series."""
    v = s.dropna()
    return np.percentile(v, q) if len(v) else np.nan

def _safe_div(num, denom):
    """Element-wise division replacing zero denominators with NaN."""
    return num / denom.replace(0, np.nan)

# Below-market flag sets used in multiple roll-ups
_BELOW_MED = {"Below Median", "Below P25", "Below P10", "Zero ASP"}
_BELOW_P25 = {"Below P25", "Below P10", "Zero ASP"}
_BELOW_PEER_MED = {"Below Peer Median", "Below Peer P25", "Zero ASP"}
_BELOW_PEER_P25 = {"Below Peer P25", "Zero ASP"}


def _pick_file() -> str:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        p = filedialog.askopenfilename(
            title="Select Customer Component Data File",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        root.destroy()
        if not p: print("No file selected."); sys.exit(0)
        return p
    except ImportError:
        sys.exit("tkinter unavailable – pass file path as argument.")


# ═══════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════
REQUIRED_COLS = [
    "Business Partner", ACCT_TYPE_COL, GPO_COL,
    "Product Group", "Market Segment", "Material Group",
    "Pricing Component Type", "Pricing Component",
    "Selling UOM Unit Qty", "Discount %",
    "Invoice ASP", "Invoice Total",
    "Standard COGS", "Std Margin", "Margin %",
]

def load_data(path: str) -> pd.DataFrame:
    print(f"Loading  {path} ...")
    df = pd.read_excel(path, sheet_name="Data Table")

    # --- Strip stray whitespace from column headers -----------------
    df.columns = [str(c).strip() for c in df.columns]

    # --- Resolve common column-name variations ----------------------
    df = _resolve_columns(df)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        hints = _suggest_matches(missing, list(df.columns))
        sys.exit(
            f"ERROR – {len(missing)} required column(s) not found after "
            f"alias resolution:\n{hints}\n\n"
            f"Columns present in file:\n  {list(df.columns)}\n\n"
            f"TIP: Add new aliases to the COLUMN_ALIASES dict at the top "
            f"of this script to handle your file's naming convention."
        )
    print(f"  {len(df):,} rows loaded")

    filters = {
        "Business Partner":       EXCLUDE_BUSINESS_PARTNER,
        "Pricing Component":      EXCLUDE_PRICING_COMPONENT,
        "Pricing Component Type": EXCLUDE_COMPONENT_TYPE,
        "Material Group":         EXCLUDE_MATERIAL_GROUP,
        "Product Group":          EXCLUDE_PRODUCT_GROUP,
        GPO_COL:                  EXCLUDE_GPO,
        ACCT_TYPE_COL:            EXCLUDE_ACCOUNT_TYPE,
    }
    for col, vals in filters.items():
        if vals and col in df.columns:
            pre = len(df)
            df = df[~df[col].isin(vals)]
            n = pre - len(df)
            if n: print(f"  Excluded {n:,} rows  ({col} in {vals})")

    print(f"  {len(df):,} rows after filters  |  "
          f"{df['Business Partner'].nunique()} customers  |  "
          f"{df['Pricing Component'].nunique()} components")
    return df


# ═══════════════════════════════════════════════════════════════════
# REVENUE TIER ASSIGNMENT
# ═══════════════════════════════════════════════════════════════════
def assign_revenue_tiers(df: pd.DataFrame) -> pd.DataFrame:
    print(f"Assigning {NUM_TIERS} revenue tiers ...")
    cr = (df.groupby("Business Partner")["Invoice Total"]
          .sum().reset_index(name="_CustRev"))
    labels = [f"Tier {i}" for i in range(1, NUM_TIERS + 1)]
    cr["Revenue_Tier"] = pd.qcut(
        cr["_CustRev"].rank(method="first"), NUM_TIERS, labels=labels[::-1])
    for t in labels:
        m = cr["Revenue_Tier"] == t
        print(f"  {t}: {m.sum():,} customers, "
              f"${cr.loc[m, '_CustRev'].sum():,.0f}")
    return df.merge(cr[["Business Partner", "_CustRev", "Revenue_Tier"]],
                    on="Business Partner", how="left")


# ═══════════════════════════════════════════════════════════════════
# 1. GLOBAL (MARKET) COMPONENT BANDS
# ═══════════════════════════════════════════════════════════════════
def build_market_bands(df: pd.DataFrame) -> pd.DataFrame:
    print("Building market (global) price bands ...")
    g = df.groupby("Pricing Component")
    b = g.agg(
        Component_Type    = ("Pricing Component Type", "first"),
        Material_Group    = ("Material Group", "first"),
        Product_Group     = ("Product Group", "first"),
        Record_Lines      = ("Invoice ASP", "size"),
        Customers         = ("Business Partner", "nunique"),
        Total_Revenue     = ("Invoice Total", "sum"),
        Total_Units       = ("Selling UOM Unit Qty", "sum"),
        Mean_ASP          = ("Invoice ASP", "mean"),
        P10_ASP           = ("Invoice ASP", lambda s: _pct(s, 10)),
        P25_ASP           = ("Invoice ASP", lambda s: _pct(s, 25)),
        Median_ASP        = ("Invoice ASP", lambda s: _pct(s, 50)),
        P75_ASP           = ("Invoice ASP", lambda s: _pct(s, 75)),
        P90_ASP           = ("Invoice ASP", lambda s: _pct(s, 90)),
        Min_ASP           = ("Invoice ASP", "min"),
        Max_ASP           = ("Invoice ASP", "max"),
        Std_ASP           = ("Invoice ASP", "std"),
        Mean_Discount     = ("Discount %", "mean"),
        Median_Discount   = ("Discount %", "median"),
        Mean_Margin_Pct   = ("Margin %", "mean"),
        Median_Margin_Pct = ("Margin %", "median"),
    ).reset_index()
    b["CoV"]              = _safe_div(b["Std_ASP"], b["Mean_ASP"])
    b["ASP_Range"]        = b["Max_ASP"] - b["Min_ASP"]
    b["ASP_Range_vs_Med"] = _safe_div(b["ASP_Range"], b["Median_ASP"])
    return b.sort_values("Total_Revenue", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 2. PEER BANDS  (cascading: GPO+AcctType → GPO → AcctType → Mkt)
# ═══════════════════════════════════════════════════════════════════
def _band_at_level(df, group_cols, prefix):
    """Compute P25 / Median ASP and customer count per group."""
    g = df.groupby(["Pricing Component"] + group_cols)
    out = g.agg(
        _custs  = ("Business Partner", "nunique"),
        _p25    = ("Invoice ASP", lambda s: _pct(s, 25)),
        _median = ("Invoice ASP", lambda s: _pct(s, 50)),
    ).reset_index()
    out.rename(columns={
        "_custs":  f"{prefix}_Custs",
        "_p25":    f"{prefix}_P25",
        "_median": f"{prefix}_Median",
    }, inplace=True)
    out[f"{prefix}_Valid"] = out[f"{prefix}_Custs"] >= MIN_PEER_GROUP
    return out


def build_peer_bands(df):
    """Return four DataFrames for the cascade:
    L1 (GPO+AcctType+Tier), L2 (GPO+AcctType), L3 (GPO), L4 (AcctType)."""
    print("Building peer bands (GPO x Account Type x Revenue Tier cascade) ...")
    l1 = _band_at_level(df, [GPO_COL, ACCT_TYPE_COL, "Revenue_Tier"], "L1")
    l2 = _band_at_level(df, [GPO_COL, ACCT_TYPE_COL],                 "L2")
    l3 = _band_at_level(df, [GPO_COL],                                "L3")
    l4 = _band_at_level(df, [ACCT_TYPE_COL],                           "L4")
    for lbl, frame in [("L1 (GPO+Type+Tier)", l1), ("L2 (GPO+Type)", l2),
                        ("L3 (GPO only)", l3), ("L4 (Type only)", l4)]:
        v = frame[f"{lbl[:2]}_Valid"].sum()
        print(f"  {lbl}: {len(frame):,} groups, {v:,} valid")
    return l1, l2, l3, l4


# ═══════════════════════════════════════════════════════════════════
# 3. FLAG RECORDS  (market + peer benchmarks)
# ═══════════════════════════════════════════════════════════════════
def flag_records(df, mkt_bands, peer_l1, peer_l2, peer_l3, peer_l4):
    print("Flagging below-benchmark records ...")

    # --- Market benchmarks ---
    mkt = mkt_bands.set_index("Pricing Component")[
        ["P10_ASP", "P25_ASP", "Median_ASP", "P75_ASP", "Mean_ASP",
         "Customers", "Record_Lines"]]
    out = df.merge(mkt, left_on="Pricing Component",
                   right_index=True, how="left", suffixes=("", "_mkt"))

    out["Price_Band"] = "At/Above Median"
    out.loc[out["Invoice ASP"] < out["Median_ASP"], "Price_Band"] = "Below Median"
    out.loc[out["Invoice ASP"] < out["P25_ASP"],    "Price_Band"] = "Below P25"
    out.loc[out["Invoice ASP"] < out["P10_ASP"],    "Price_Band"] = "Below P10"
    out.loc[out["Invoice ASP"] == 0,                "Price_Band"] = "Zero ASP"

    out["ASP_vs_Mkt_Med_Pct"] = _safe_div(out["Invoice ASP"] - out["Median_ASP"], out["Median_ASP"])
    out["ASP_vs_Mkt_P25_Pct"] = _safe_div(out["Invoice ASP"] - out["P25_ASP"],    out["P25_ASP"])

    out["Mkt_Uplift_Med"] = 0.0
    m = (out["Invoice ASP"] < out["Median_ASP"]) & (out["Median_ASP"] > 0)
    out.loc[m, "Mkt_Uplift_Med"] = (out.loc[m, "Median_ASP"] - out.loc[m, "Invoice ASP"]) * out.loc[m, "Selling UOM Unit Qty"]

    out["Mkt_Uplift_P25"] = 0.0
    m2 = (out["Invoice ASP"] < out["P25_ASP"]) & (out["P25_ASP"] > 0)
    out.loc[m2, "Mkt_Uplift_P25"] = (out.loc[m2, "P25_ASP"] - out.loc[m2, "Invoice ASP"]) * out.loc[m2, "Selling UOM Unit Qty"]

    # --- Peer benchmarks (cascade: L1 → L2 → L3 → L4 → Market) ---
    out = out.merge(peer_l1, on=["Pricing Component", GPO_COL, ACCT_TYPE_COL, "Revenue_Tier"], how="left")
    out = out.merge(peer_l2, on=["Pricing Component", GPO_COL, ACCT_TYPE_COL], how="left")
    out = out.merge(peer_l3, on=["Pricing Component", GPO_COL], how="left")
    out = out.merge(peer_l4, on=["Pricing Component", ACCT_TYPE_COL], how="left")

    # Pick the best valid level: L1 → L2 → L3 → L4 → Market
    out["Peer_P25"]    = np.nan
    out["Peer_Median"] = np.nan
    out["Peer_Custs"]  = np.nan
    out["Peer_Level"]  = "Market"

    # L4 (least specific fallback — Account Type only)
    v4 = out["L4_Valid"].fillna(False).astype(bool)
    out.loc[v4, "Peer_P25"]    = out.loc[v4, "L4_P25"]
    out.loc[v4, "Peer_Median"] = out.loc[v4, "L4_Median"]
    out.loc[v4, "Peer_Custs"]  = out.loc[v4, "L4_Custs"]
    out.loc[v4, "Peer_Level"]  = "Account Type"

    # L3 overrides L4 (GPO only)
    v3 = out["L3_Valid"].fillna(False).astype(bool)
    out.loc[v3, "Peer_P25"]    = out.loc[v3, "L3_P25"]
    out.loc[v3, "Peer_Median"] = out.loc[v3, "L3_Median"]
    out.loc[v3, "Peer_Custs"]  = out.loc[v3, "L3_Custs"]
    out.loc[v3, "Peer_Level"]  = "GPO"

    # L2 overrides L3 (GPO + Account Type)
    v2 = out["L2_Valid"].fillna(False).astype(bool)
    out.loc[v2, "Peer_P25"]    = out.loc[v2, "L2_P25"]
    out.loc[v2, "Peer_Median"] = out.loc[v2, "L2_Median"]
    out.loc[v2, "Peer_Custs"]  = out.loc[v2, "L2_Custs"]
    out.loc[v2, "Peer_Level"]  = "GPO + Account Type"

    # L1 overrides all (GPO + Account Type + Revenue Tier — most specific)
    v1 = out["L1_Valid"].fillna(False).astype(bool)
    out.loc[v1, "Peer_P25"]    = out.loc[v1, "L1_P25"]
    out.loc[v1, "Peer_Median"] = out.loc[v1, "L1_Median"]
    out.loc[v1, "Peer_Custs"]  = out.loc[v1, "L1_Custs"]
    out.loc[v1, "Peer_Level"]  = "GPO + Account Type + Tier"

    # Market fallback for anything still NaN
    fb = out["Peer_Median"].isna()
    out.loc[fb, "Peer_Median"] = out.loc[fb, "Median_ASP"]
    out.loc[fb, "Peer_P25"]    = out.loc[fb, "P25_ASP"]

    # Peer price band
    out["Peer_Band"] = "At/Above Peer Median"
    out.loc[out["Invoice ASP"] < out["Peer_Median"], "Peer_Band"] = "Below Peer Median"
    out.loc[out["Invoice ASP"] < out["Peer_P25"],    "Peer_Band"] = "Below Peer P25"
    out.loc[out["Invoice ASP"] == 0,                 "Peer_Band"] = "Zero ASP"

    out["ASP_vs_Peer_Med_Pct"] = _safe_div(out["Invoice ASP"] - out["Peer_Median"], out["Peer_Median"])
    out["ASP_vs_Peer_P25_Pct"] = _safe_div(out["Invoice ASP"] - out["Peer_P25"],    out["Peer_P25"])

    out["Peer_Uplift_Med"] = 0.0
    pm = (out["Invoice ASP"] < out["Peer_Median"]) & (out["Peer_Median"] > 0)
    out.loc[pm, "Peer_Uplift_Med"] = (out.loc[pm, "Peer_Median"] - out.loc[pm, "Invoice ASP"]) * out.loc[pm, "Selling UOM Unit Qty"]

    out["Peer_Uplift_P25"] = 0.0
    pp = (out["Invoice ASP"] < out["Peer_P25"]) & (out["Peer_P25"] > 0)
    out.loc[pp, "Peer_Uplift_P25"] = (out.loc[pp, "Peer_P25"] - out.loc[pp, "Invoice ASP"]) * out.loc[pp, "Selling UOM Unit Qty"]

    # Drop intermediate L1/L2/L3/L4 columns
    drop = [c for c in out.columns if c.startswith(("L1_", "L2_", "L3_", "L4_"))]
    out.drop(columns=drop, inplace=True)

    return out


# ═══════════════════════════════════════════════════════════════════
# 4-5. ROLL-UP FUNCTIONS
# ═══════════════════════════════════════════════════════════════════
def customer_benchmark(fl):
    print("Building customer benchmark ...")
    c = fl.groupby("Business Partner").agg(
        Account_Type          = (ACCT_TYPE_COL, "first"),
        GPO                   = (GPO_COL, "first"),
        Revenue_Tier          = ("Revenue_Tier", "first"),
        Product_Groups        = ("Product Group", "nunique"),
        Total_Lines           = ("Invoice ASP", "size"),
        Total_Revenue         = ("Invoice Total", "sum"),
        Total_Std_Margin      = ("Std Margin", "sum"),
        Avg_Discount          = ("Discount %", "mean"),
        Lines_Below_Med       = ("Price_Band",  lambda s: s.isin(_BELOW_MED).sum()),
        Lines_Below_P25       = ("Price_Band",  lambda s: s.isin(_BELOW_P25).sum()),
        Lines_Zero            = ("Price_Band",  lambda s: (s == "Zero ASP").sum()),
        Mkt_Uplift_Med        = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P25        = ("Mkt_Uplift_P25", "sum"),
        Peer_Uplift_Med       = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P25       = ("Peer_Uplift_P25", "sum"),
        Lines_Below_Peer      = ("Peer_Band", lambda s: s.isin(_BELOW_PEER_MED).sum()),
        Lines_Below_Peer_P25  = ("Peer_Band", lambda s: s.isin(_BELOW_PEER_P25).sum()),
    ).reset_index()
    c["Wtd_Margin_Pct"] = _safe_div(c["Total_Std_Margin"], c["Total_Revenue"])
    c = c[[
        "Business Partner", "Account_Type", "GPO", "Revenue_Tier",
        "Product_Groups", "Total_Revenue", "Total_Std_Margin",
        "Wtd_Margin_Pct", "Avg_Discount", "Total_Lines",
        "Lines_Below_Med", "Lines_Below_P25", "Lines_Zero",
        "Mkt_Uplift_Med", "Mkt_Uplift_P25",
        "Lines_Below_Peer", "Lines_Below_Peer_P25",
        "Peer_Uplift_Med", "Peer_Uplift_P25",
    ]]
    return c.sort_values("Peer_Uplift_P25", ascending=False)


def comp_type_uplift(fl):
    print("Building component-type uplift ...")
    ct = fl.groupby("Pricing Component Type").agg(
        Total_Lines       = ("Invoice ASP", "size"),
        Total_Revenue     = ("Invoice Total", "sum"),
        Total_Std_Margin  = ("Std Margin", "sum"),
        Lines_Below_Med   = ("Price_Band", lambda s: s.isin(_BELOW_MED).sum()),
        Lines_Below_P25   = ("Price_Band", lambda s: s.isin(_BELOW_P25).sum()),
        Mkt_Uplift_Med    = ("Mkt_Uplift_Med", "sum"),
        Mkt_Uplift_P25    = ("Mkt_Uplift_P25", "sum"),
        Peer_Uplift_Med   = ("Peer_Uplift_Med", "sum"),
        Peer_Uplift_P25   = ("Peer_Uplift_P25", "sum"),
    ).reset_index()
    ct["Wtd_Margin_Pct"]    = _safe_div(ct["Total_Std_Margin"], ct["Total_Revenue"])
    ct["Pct_Below_Mkt_Med"] = ct["Lines_Below_Med"] / ct["Total_Lines"]
    return ct.sort_values("Peer_Uplift_P25", ascending=False)


# ═══════════════════════════════════════════════════════════════════
# 6. EXECUTIVE SUMMARY KPIs
# ═══════════════════════════════════════════════════════════════════
def build_summary(df, bands, fl, cust, ct):
    n = len(fl)
    tr = df["Invoice Total"].sum()
    peer_levels = fl["Peer_Level"].value_counts()
    return dict(
        total_lines=n, total_rev=tr,
        n_cust=df["Business Partner"].nunique(),
        n_comp=df["Pricing Component"].nunique(),
        below_med=fl["Price_Band"].isin(_BELOW_MED).sum(),
        below_p25=fl["Price_Band"].isin(_BELOW_P25).sum(),
        zero=int((fl["Price_Band"] == "Zero ASP").sum()),
        mkt_up_med=fl["Mkt_Uplift_Med"].sum(),
        mkt_up_p25=fl["Mkt_Uplift_P25"].sum(),
        peer_up_med=fl["Peer_Uplift_Med"].sum(),
        peer_up_p25=fl["Peer_Uplift_P25"].sum(),
        peer_below_med=fl["Peer_Band"].isin(_BELOW_PEER_MED).sum(),
        peer_levels=peer_levels,
        top5_cust=cust.head(5)[["Business Partner","Peer_Uplift_P25"]].values.tolist(),
        top5_ct=ct.head(5)[["Pricing Component Type","Peer_Uplift_P25"]].values.tolist(),
        top5_cov=bands[bands["Record_Lines"] >= 10].nlargest(5, "CoV")[
            ["Pricing Component","CoV","Record_Lines","Total_Revenue"]].values.tolist(),
    )


# ═══════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════
def write_output(path, bands, fl, cust, ct, S):
    print(f"Writing  {path} ...")
    wr = pd.ExcelWriter(path, engine="xlsxwriter")
    wb = wr.book

    # -- Formats --
    H  = wb.add_format({"bold":1,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    HP = wb.add_format({"bold":1,"bg_color":"#2E75B6","font_color":"#FFFFFF","border":1,
                         "text_wrap":1,"valign":"vcenter","font_name":"Arial","font_size":10})
    C  = wb.add_format({"num_format":"$#,##0",   "font_name":"Arial","font_size":10})
    C2 = wb.add_format({"num_format":"$#,##0.00","font_name":"Arial","font_size":10})
    P  = wb.add_format({"num_format":"0.0%",     "font_name":"Arial","font_size":10})
    P2 = wb.add_format({"num_format":"0.00%",    "font_name":"Arial","font_size":10})
    N  = wb.add_format({"num_format":"#,##0",    "font_name":"Arial","font_size":10})
    D  = wb.add_format({"num_format":"0.00",     "font_name":"Arial","font_size":10})
    T  = wb.add_format({"font_name":"Arial","font_size":10})
    TI = wb.add_format({"bold":1,"font_name":"Arial","font_size":14,"font_color":"#1F3864"})
    SB = wb.add_format({"bold":1,"font_name":"Arial","font_size":11,"font_color":"#1F3864","bottom":1})
    LB = wb.add_format({"font_name":"Arial","font_size":10,"indent":1})
    KV = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"$#,##0"})
    KN = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"#,##0"})
    KP = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"num_format":"0.0%"})
    NT = wb.add_format({"font_name":"Arial","font_size":9,"italic":1,"text_wrap":1,"font_color":"#555"})
    IH = wb.add_format({"bold":1,"font_name":"Arial","font_size":10,"bg_color":"#D6E4F0",
                         "border":1,"text_wrap":1,"valign":"vcenter"})
    IC = wb.add_format({"font_name":"Arial","font_size":10,"text_wrap":1,"valign":"top","border":1})
    IT = wb.add_format({"font_name":"Arial","font_size":10,"bold":1,"valign":"top","border":1})

    def _ws(df, name, cols, tab="#4472C4", fz=1, pc=None):
        """Write a DataFrame sheet; pc = set of 0-indexed peer-header columns."""
        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        pc = pc or set()
        for i, (cn, w, f) in enumerate(cols):
            ws.write(0, i, cn, HP if i in pc else H); ws.set_column(i, i, w, f)

    # ────────────── Information ──────────────────────────────────────
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",30); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Price Band / ASP Benchmarking Analysis",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2
    wi.merge_range(r,1,r,2,"Workbook Sheets",SB); r+=1
    wi.write(r,1,"Sheet",IH); wi.write(r,2,"Description",IH); r+=1
    for nm,ds in [
        ("Information","Overview of each tab and key term definitions."),
        ("Executive Summary","Headline KPIs under market and peer benchmarks."),
        ("Component Price Bands","Percentile ASP benchmarks per Pricing Component (market-wide)."),
        ("Customer Benchmark","Per-customer summary with Account Type, GPO, revenue tier, and dual uplift."),
        ("Flagged Records","Lines below global P25 with market and peer benchmarks side by side.  "
         "Peer columns use lighter blue headers.  Includes the Peer Level column showing which "
         "cascade level was used (GPO + Account Type + Tier / GPO + Account Type / GPO / Account Type / Market)."),
        ("Uplift by Comp Type","Market and peer uplift aggregated by Pricing Component Type."),
    ]:
        wi.write(r,1,nm,IC); wi.write(r,2,ds,IC); wi.set_row(r,48); r+=1
    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Market Benchmark","ASP percentiles across ALL customers for a Pricing Component."),
        ("Peer Benchmark","ASP percentiles among customers sharing the same GPO, Account Type, "
         "and/or Revenue Tier.  Uses the most specific grouping available with enough peers."),
        ("Peer Level","Indicates which cascade level was used: GPO + Account Type + Tier "
         "(most specific), GPO + Account Type, GPO only, Account Type only, or Market (fallback)."),
        ("Revenue Tier",f"Customers divided into {NUM_TIERS} tiers by total revenue (Tier 1 = largest).  "
         "Used as a peer-grouping dimension alongside GPO and Account Type."),
        ("Account Type","Facility classification (e.g., HOSPITAL, ASC, CLINIC)."),
        ("GPO / Buying Grp 1","Group Purchasing Organisation affiliation."),
        (f"Min Peer Group ({MIN_PEER_GROUP})",f"A peer group needs at least {MIN_PEER_GROUP} distinct "
         "customers to be used; smaller groups cascade to the next level."),
        ("Invoice ASP","Average Selling Price per unit (Invoice Total / Units)."),
        ("P10–P90","Percentile benchmarks of Invoice ASP within the relevant population."),
        ("CoV","Coefficient of Variation (Std Dev / Mean ASP)."),
        ("Uplift to P25","Incremental revenue if the ASP were raised to the P25 benchmark, holding volume constant."),
        ("Uplift to Median","Incremental revenue if the ASP were raised to the Median benchmark, holding volume constant."),
    ]:
        wi.write(r,1,tm,IT); wi.write(r,2,df_,IC); wi.set_row(r,38); r+=1

    # ────────────── Executive Summary ────────────────────────────────
    ws = wb.add_worksheet("Executive Summary")
    ws.hide_gridlines(2); ws.set_tab_color("#1F3864")
    ws.set_column("A:A",4); ws.set_column("B:B",50)
    ws.set_column("C:C",22); ws.set_column("D:D",22); ws.set_column("E:E",22)
    r=1; ws.merge_range(r,1,r,4,"Price Band / ASP Benchmarking Analysis",TI); r+=1
    ws.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2

    ws.merge_range(r,1,r,4,"Dataset Overview",SB); r+=1
    for l,v,f in [("Total Records",S["total_lines"],KN),("Total Invoice Revenue",S["total_rev"],KV),
                   ("Unique Customers",S["n_cust"],KN),("Unique Pricing Components",S["n_comp"],KN),
                   ("Revenue Tiers",NUM_TIERS,KN)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Market Benchmark",SB); r+=1
    for l,v,f in [("Lines Below Market Median",S["below_med"],KN),
                   ("  -- as % of Total",S["below_med"]/max(S["total_lines"],1),KP),
                   ("Lines Below Market P25",S["below_p25"],KN),
                   ("Lines at Zero ASP",S["zero"],KN),
                   ("Mkt Uplift to Median",S["mkt_up_med"],KV),
                   ("Mkt Uplift to P25",S["mkt_up_p25"],KV),
                   ("  -- as % of Revenue",S["mkt_up_p25"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Peer Benchmark",SB); r+=1
    for l,v,f in [("Lines Below Peer Median",S["peer_below_med"],KN),
                   ("Peer Uplift to Median",S["peer_up_med"],KV),
                   ("Peer Uplift to P25",S["peer_up_p25"],KV),
                   ("  -- as % of Revenue",S["peer_up_p25"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1
    # Peer-level breakdown
    r+=1; ws.merge_range(r,1,r,4,"Peer Level Distribution",SB); r+=1
    ws.write(r,1,"Peer Level",H); ws.write(r,2,"Records",H); r+=1
    for lvl in ["GPO + Account Type + Tier","GPO + Account Type","GPO","Account Type","Market"]:
        cnt = S["peer_levels"].get(lvl, 0)
        ws.write(r,1,lvl,T); ws.write(r,2,cnt,N); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Customers by Peer Uplift to P25",SB); r+=1
    ws.write(r,1,"Customer",H); ws.write(r,2,"Peer Uplift to P25",H); r+=1
    for nm,v in S["top5_cust"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Component Types by Peer Uplift to P25",SB); r+=1
    ws.write(r,1,"Component Type",H); ws.write(r,2,"Peer Uplift to P25",H); r+=1
    for nm,v in S["top5_ct"]:
        ws.write(r,1,nm,T); ws.write(r,2,v,C); r+=1

    r+=1; ws.merge_range(r,1,r,4,"Top 5 Highest Price Variation (>=10 lines)",SB); r+=1
    ws.write(r,1,"Pricing Component",H); ws.write(r,2,"CoV",H)
    ws.write(r,3,"Lines",H); ws.write(r,4,"Revenue",H); r+=1
    for comp,cov,lines,rev in S["top5_cov"]:
        ws.write(r,1,comp,T); ws.write(r,2,cov,D)
        ws.write(r,3,lines,N); ws.write(r,4,rev,C); r+=1

    r+=2; ws.merge_range(r,1,r,4,"Methodology",SB); r+=1
    ws.merge_range(r,1,r+6,4,
        "Market benchmarks: ASP percentiles across ALL customers per Pricing Component.  "
        "Peer benchmarks: ASP percentiles among customers sharing the same GPO, Account Type, "
        "and/or Revenue Tier per Pricing Component, using the most specific grouping with at "
        f"least {MIN_PEER_GROUP} customers.  Cascade: GPO + Account Type + Tier -> "
        "GPO + Account Type -> GPO only -> Account Type only -> Market (fallback).  "
        f"Revenue tiers ({NUM_TIERS}): Tier 1 = largest customers by total invoice revenue.  "
        "Uplift = incremental revenue if below-benchmark ASPs were raised to the benchmark, "
        "holding unit volume constant.", NT)

    # ────────────── Component Price Bands ────────────────────────────
    _ws(bands, "Component Price Bands", [
        ("Pricing Component",32,T),("Component Type",30,T),("Material Group",20,T),
        ("Product Group",18,T),("Lines",8,N),("Customers",10,N),
        ("Total Revenue",16,C),("Total Units",12,N),("Mean ASP",12,C2),
        ("P10 ASP",12,C2),("P25 ASP",12,C2),("Median ASP",12,C2),
        ("P75 ASP",12,C2),("P90 ASP",12,C2),("Min ASP",12,C2),
        ("Max ASP",12,C2),("Std Dev ASP",12,C2),("Avg Discount",12,P2),
        ("Median Discount",14,P2),("Avg Margin %",12,P),("Median Margin %",14,P),
        ("CoV",8,D),("ASP Range ($)",14,C2),("Range / Median",14,D),
    ])
    wr.sheets["Component Price Bands"].conditional_format(1,21,len(bands),21,{
        "type":"3_color_scale","min_color":"#63BE7B","mid_color":"#FFEB84","max_color":"#F8696B"})

    # ────────────── Customer Benchmark ───────────────────────────────
    _ws(cust, "Customer Benchmark", [
        ("Business Partner",38,T),("Account Type",16,T),("GPO",18,T),
        ("Revenue Tier",12,T),("Product Groups",14,N),
        ("Total Revenue",16,C),("Total Std Margin",16,C),
        ("Wtd Margin %",12,P),("Avg Discount",12,P2),("Total Lines",10,N),
        ("Lines Below Mkt Med",16,N),("Lines Below Mkt P25",16,N),("Lines Zero ASP",12,N),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Lines Below Peer Med",16,N),("Lines Below Peer P25",16,N),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
    ], tab="#548235", pc={15,16,17,18})
    wr.sheets["Customer Benchmark"].conditional_format(
        1,18,len(cust),18,{"type":"data_bar","bar_color":"#548235"})

    # ────────────── Flagged Records ──────────────────────────────────
    bel = fl[fl["Price_Band"].isin(_BELOW_P25)].copy()
    bel.sort_values("Peer_Uplift_P25", ascending=False, inplace=True)
    fo = bel[[
        "Business Partner",ACCT_TYPE_COL,GPO_COL,"Revenue_Tier","Pricing Component","Pricing Component Type",
        "Material Group","Product Group",
        "Selling UOM Unit Qty","Invoice Total","Invoice ASP",
        "Discount %","Margin %","Price_Band",
        "Median_ASP","ASP_vs_Mkt_Med_Pct","P25_ASP","ASP_vs_Mkt_P25_Pct",
        "Peer_Median","ASP_vs_Peer_Med_Pct","Peer_P25","ASP_vs_Peer_P25_Pct",
        "Peer_Custs","Peer_Level",
        "Mkt_Uplift_Med","Mkt_Uplift_P25","Peer_Uplift_Med","Peer_Uplift_P25",
    ]].copy()
    fo.columns = [
        "Business Partner","Account Type","GPO","Revenue Tier","Pricing Component","Component Type",
        "Material Group","Product Group",
        "Units","Invoice Total","Invoice ASP",
        "Discount %","Margin %","Price Band",
        "Mkt Median ASP","ASP vs Mkt Med %","Mkt P25 ASP","ASP vs Mkt P25 %",
        "Peer Median ASP","ASP vs Peer Med %","Peer P25 ASP","ASP vs Peer P25 %",
        "Peer Customers","Peer Level",
        "Mkt Uplift to Median","Mkt Uplift to P25","Peer Uplift to Median","Peer Uplift to P25",
    ]
    fcols = [
        ("Business Partner",38,T),("Account Type",16,T),
        ("GPO",18,T),("Revenue Tier",12,T),
        ("Pricing Component",30,T),("Component Type",28,T),
        ("Material Group",20,T),("Product Group",18,T),
        ("Units",8,N),("Invoice Total",14,C),("Invoice ASP",12,C2),
        ("Discount %",12,P2),("Margin %",10,P),("Price Band",14,T),
        ("Mkt Median ASP",14,C2),("ASP vs Mkt Med %",16,P),
        ("Mkt P25 ASP",12,C2),("ASP vs Mkt P25 %",14,P),
        ("Peer Median ASP",14,C2),("ASP vs Peer Med %",16,P),
        ("Peer P25 ASP",12,C2),("ASP vs Peer P25 %",14,P),
        ("Peer Customers",14,N),("Peer Level",18,T),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
    ]
    _ws(fo,"Flagged Records",fcols,tab="#C00000",
        pc={18,19,20,21,22,23,26,27})
    wf = wr.sheets["Flagged Records"]
    for v,clr in [("Zero ASP","#C00000"),("Below P10","#ED7D31"),("Below P25","#FFC000")]:
        wf.conditional_format(1,13,len(fo),13,{
            "type":"text","criteria":"containing","value":v,
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF",
                                     "font_name":"Arial","font_size":10})})

    # ────────────── Uplift by Comp Type ──────────────────────────────
    _ws(ct, "Uplift by Comp Type", [
        ("Component Type",30,T),("Total Lines",10,N),("Total Revenue",16,C),
        ("Total Std Margin",16,C),
        ("Lines Below Mkt Med",16,N),("Lines Below Mkt P25",14,N),
        ("Mkt Uplift to Median",18,C),("Mkt Uplift to P25",16,C),
        ("Peer Uplift to Median",18,C),("Peer Uplift to P25",16,C),
        ("Wtd Margin %",12,P),("% Lines Below Mkt Med",18,P),
    ], tab="#7030A0", pc={8,9})

    wr.close()
    print(f"Done -- {path}")


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    warnings.filterwarnings("ignore", category=UserWarning)
    path = sys.argv[1] if len(sys.argv) > 1 else _pick_file()
    if not os.path.isfile(path): sys.exit(f"Not found: {path}")

    out = str(Path(path).parent / f"{Path(path).stem}_PriceBandAnalysis.xlsx")

    df = load_data(path)
    df = assign_revenue_tiers(df)
    bands = build_market_bands(df)
    l1, l2, l3, l4 = build_peer_bands(df)
    fl = flag_records(df, bands, l1, l2, l3, l4)
    cust = customer_benchmark(fl)
    ct = comp_type_uplift(fl)
    S = build_summary(df, bands, fl, cust, ct)

    write_output(out, bands, fl, cust, ct, S)

    print(f"\n{'='*65}")
    print("  HEADLINE RESULTS")
    print(f"{'='*65}")
    print(f"  Total Revenue:              ${S['total_rev']:>14,.0f}")
    print(f"  Market Uplift to P25:       ${S['mkt_up_p25']:>14,.0f}  "
          f"({S['mkt_up_p25']/S['total_rev']:.1%})")
    print(f"  Peer Uplift to P25:         ${S['peer_up_p25']:>14,.0f}  "
          f"({S['peer_up_p25']/S['total_rev']:.1%})")
    print(f"  Lines Below Peer Median:    {S['peer_below_med']:>14,}  "
          f"({S['peer_below_med']/S['total_lines']:.1%})")
    print(f"\n  Peer Level Distribution:")
    for lvl in ["GPO + Account Type + Tier","GPO + Account Type","GPO","Account Type","Market"]:
        cnt = S["peer_levels"].get(lvl, 0)
        print(f"    {lvl:<30s}  {cnt:>8,}  ({cnt/S['total_lines']:.1%})")
    print(f"{'='*65}")


if __name__ == "__main__":
    main()
