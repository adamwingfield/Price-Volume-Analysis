"""
Customer Profitability Tiering Analysis
=======================================
Segments customers into a Revenue x Margin quadrant matrix to identify
accounts that are under-performing on margin relative to their revenue
contribution, and quantifies the margin-recovery opportunity from
bringing low-margin accounts closer to their peer benchmarks.

Quadrants
---------
  Q1  High Revenue / High Margin   -- protect & grow
  Q2  High Revenue / Low Margin    -- priority for price correction
  Q3  Low Revenue  / High Margin   -- maintain pricing discipline
  Q4  Low Revenue  / Low Margin    -- rationalise or enforce standard pricing

Outputs
-------
  Excel workbook with 7 data sheets plus a Quadrant Map scatter plot
  saved as a separate PNG file and displayed automatically.

Usage
-----
    python customer_profitability_tiering.py [path_to_file.xlsx]
    If no path is given a file-picker dialog opens.

Requirements
------------
    pip install pandas openpyxl xlsxwriter matplotlib
"""

import sys, os, subprocess, warnings, datetime
from pathlib import Path

# -- Dependency bootstrap ------------------------------------------------
_PKGS = {"pandas": "pandas", "openpyxl": "openpyxl",
         "xlsxwriter": "xlsxwriter", "matplotlib": "matplotlib"}
def _ensure():
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

# ═══════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════
# Quadrant boundary percentiles
REVENUE_PERCENTILE = 80     # top half = "High Revenue"
MARGIN_PERCENTILE  = 50     # top half = "High Margin"

# Scatter-plot axis limits (decimals: 0.5 = 50%, 1.0 = 100%).
# Set to None to auto-scale from the data.
CHART_MARGIN_MIN  = 0.0     # Y-axis floor
CHART_MARGIN_MAX  = 1.05    # Y-axis ceiling
CHART_REVENUE_MIN = None    # X-axis floor; None = auto
CHART_REVENUE_MAX = None    # X-axis ceiling; None = auto

# Source column names for optional enrichment dimensions
GPO_COL       = "GPO / Buying Grp 1"
ACCT_TYPE_COL = "Account Type"

# ═══════════════════════════════════════════════════════════════════
# EXCLUSION FILTERS
# ═══════════════════════════════════════════════════════════════════
EXCLUDE_BUSINESS_PARTNER  = []
EXCLUDE_PRICING_COMPONENT = ["ROSA BRAIN - SERVICE"]
EXCLUDE_COMPONENT_TYPE    = []
EXCLUDE_MATERIAL_GROUP    = []
EXCLUDE_PRODUCT_GROUP     = []
EXCLUDE_GPO               = []
EXCLUDE_ACCOUNT_TYPE      = []

# ═══════════════════════════════════════════════════════════════════
# QUADRANT COLOURS (shared by workbook formatting and chart)
# ═══════════════════════════════════════════════════════════════════
Q_COLORS = {
    "Q1: High Rev / High Margin": "#548235",
    "Q2: High Rev / Low Margin":  "#C00000",
    "Q3: Low Rev / High Margin":  "#4472C4",
    "Q4: Low Rev / Low Margin":   "#BF8F00",
}
Q_ORDER = list(Q_COLORS.keys())

# ═══════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════
def _safe_div(num, denom):
    return num / denom.replace(0, np.nan)

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
        sys.exit("tkinter unavailable -- pass file path as argument.")

# ═══════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════
REQUIRED_COLS = [
    "Business Partner", "Product Group",
    "Pricing Component Type", "Pricing Component",
    "Selling UOM Unit Qty", "Discount %",
    "Invoice ASP", "Invoice Total",
    "Standard COGS", "Std Margin", "Margin %",
]
OPTIONAL_COLS = [ACCT_TYPE_COL, GPO_COL, "Material Group", "Market Segment"]

def load_data(path: str) -> pd.DataFrame:
    print(f"Loading  {path} ...")
    df = pd.read_excel(path, sheet_name="Data Table")
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing: sys.exit(f"ERROR -- missing columns: {missing}")
    avail_opt = [c for c in OPTIONAL_COLS if c in df.columns]
    print(f"  {len(df):,} rows loaded  |  optional columns: {avail_opt}")

    filters = {
        "Business Partner": EXCLUDE_BUSINESS_PARTNER,
        "Pricing Component": EXCLUDE_PRICING_COMPONENT,
        "Pricing Component Type": EXCLUDE_COMPONENT_TYPE,
        "Product Group": EXCLUDE_PRODUCT_GROUP,
    }
    if GPO_COL in df.columns:       filters[GPO_COL] = EXCLUDE_GPO
    if ACCT_TYPE_COL in df.columns: filters[ACCT_TYPE_COL] = EXCLUDE_ACCOUNT_TYPE
    if "Material Group" in df.columns: filters["Material Group"] = EXCLUDE_MATERIAL_GROUP
    for col, vals in filters.items():
        if vals and col in df.columns:
            pre = len(df); df = df[~df[col].isin(vals)]; n = pre - len(df)
            if n: print(f"  Excluded {n:,} rows  ({col} in {vals})")
    print(f"  {len(df):,} rows after filters  |  "
          f"{df['Business Partner'].nunique()} customers")
    return df


# ═══════════════════════════════════════════════════════════════════
# 1. CUSTOMER PROFITABILITY TABLE
# ═══════════════════════════════════════════════════════════════════
def build_customer_profile(df: pd.DataFrame) -> pd.DataFrame:
    print("Building customer profitability profiles ...")
    agg = {
        "Total_Lines":     ("Invoice ASP", "size"),
        "Total_Revenue":   ("Invoice Total", "sum"),
        "Total_Std_Margin":("Std Margin", "sum"),
        "Total_COGS":      ("Standard COGS", "sum"),
        "Total_Units":     ("Selling UOM Unit Qty", "sum"),
        "Avg_Discount":    ("Discount %", "mean"),
        "Med_Margin_Pct":  ("Margin %", "median"),
        "Product_Groups":  ("Product Group", "nunique"),
        "Comp_Types":      ("Pricing Component Type", "nunique"),
        "Components":      ("Pricing Component", "nunique"),
    }
    if ACCT_TYPE_COL in df.columns: agg["Account_Type"] = (ACCT_TYPE_COL, "first")
    if GPO_COL in df.columns:       agg["GPO"] = (GPO_COL, "first")

    c = df.groupby("Business Partner").agg(**agg).reset_index()
    c["Wtd_Margin_Pct"] = _safe_div(c["Total_Std_Margin"], c["Total_Revenue"])
    c["Avg_ASP"]        = _safe_div(c["Total_Revenue"], c["Total_Units"])
    c["Avg_COGS_Unit"]  = _safe_div(c["Total_COGS"].abs(), c["Total_Units"])
    c["Rev_per_Line"]   = _safe_div(c["Total_Revenue"], c["Total_Lines"])

    c.sort_values("Total_Revenue", ascending=False, inplace=True)
    c["Rev_Rank"] = range(1, len(c) + 1)
    total_rev = c["Total_Revenue"].sum()
    c["Rev_Share"]     = c["Total_Revenue"] / total_rev
    c["Cum_Rev_Share"] = c["Rev_Share"].cumsum()

    rev_t = np.percentile(c["Total_Revenue"], REVENUE_PERCENTILE)
    mar_t = np.percentile(c["Wtd_Margin_Pct"].dropna(), MARGIN_PERCENTILE)
    print(f"  Revenue threshold (P{REVENUE_PERCENTILE}): ${rev_t:,.0f}")
    print(f"  Margin threshold  (P{MARGIN_PERCENTILE}): {mar_t:.1%}")

    hi_r = c["Total_Revenue"]   >= rev_t
    hi_m = c["Wtd_Margin_Pct"]  >= mar_t
    c["Quadrant"] = "Q4: Low Rev / Low Margin"
    c.loc[hi_r & hi_m,  "Quadrant"] = "Q1: High Rev / High Margin"
    c.loc[hi_r & ~hi_m, "Quadrant"] = "Q2: High Rev / Low Margin"
    c.loc[~hi_r & hi_m, "Quadrant"] = "Q3: Low Rev / High Margin"

    overall_med = c["Wtd_Margin_Pct"].median()
    c["Margin_Gap_pp"]       = c["Wtd_Margin_Pct"] - overall_med
    c["Margin_Recov_to_Med"] = 0.0
    below = c["Wtd_Margin_Pct"] < overall_med
    c.loc[below, "Margin_Recov_to_Med"] = (
        (overall_med - c.loc[below, "Wtd_Margin_Pct"]) * c.loc[below, "Total_Revenue"])

    c.attrs["rev_threshold"]     = rev_t
    c.attrs["margin_threshold"]  = mar_t
    c.attrs["overall_med_margin"] = overall_med
    return c


# ═══════════════════════════════════════════════════════════════════
# 2. QUADRANT SUMMARY
# ═══════════════════════════════════════════════════════════════════
def build_quadrant_summary(cust):
    print("Building quadrant summary ...")
    # Revenue-weighted discount: weight each customer's avg discount by
    # their revenue.  Clip extreme outlier discounts (e.g. -32,550%)
    # to the -100% to 0% range before weighting so that data-quality
    # anomalies don't contaminate the quadrant average.
    cust = cust.copy()
    cust["_Disc_Clipped"] = cust["Avg_Discount"].clip(lower=-1.0, upper=0.0)
    cust["_Disc_x_Rev"]   = cust["_Disc_Clipped"] * cust["Total_Revenue"]

    qs = cust.groupby("Quadrant").agg(
        Customers      = ("Business Partner", "size"),
        Total_Revenue  = ("Total_Revenue", "sum"),
        Total_Margin   = ("Total_Std_Margin", "sum"),
        Avg_Wtd_Margin = ("Wtd_Margin_Pct", "mean"),
        Med_Wtd_Margin = ("Wtd_Margin_Pct", "median"),
        _Disc_x_Rev    = ("_Disc_x_Rev", "sum"),
        Avg_PG_Count   = ("Product_Groups", "mean"),
        Margin_Recovery= ("Margin_Recov_to_Med", "sum"),
    ).reset_index()

    # Derived columns
    qs["Wtd_Margin_Pct"] = _safe_div(qs["Total_Margin"], qs["Total_Revenue"])
    qs["Avg_Discount"]   = _safe_div(qs["_Disc_x_Rev"], qs["Total_Revenue"])
    qs["Rev_Share"]      = qs["Total_Revenue"] / qs["Total_Revenue"].sum()

    # Explicit column order to match the Excel header list
    qs = qs[["Quadrant", "Customers", "Total_Revenue", "Total_Margin",
             "Wtd_Margin_Pct", "Avg_Wtd_Margin", "Med_Wtd_Margin",
             "Avg_Discount", "Avg_PG_Count", "Margin_Recovery", "Rev_Share"]]
    return qs.sort_values("Quadrant")


# ═══════════════════════════════════════════════════════════════════
# 3. PRODUCT GROUP MIX
# ═══════════════════════════════════════════════════════════════════
def build_product_mix(df, cust):
    print("Building product group mix ...")
    pg = df.groupby(["Business Partner","Product Group"]).agg(
        PG_Revenue=("Invoice Total","sum"), PG_Margin=("Std Margin","sum"),
        PG_Lines=("Invoice ASP","size"),
    ).reset_index()
    pg["PG_Wtd_Margin"] = _safe_div(pg["PG_Margin"], pg["PG_Revenue"])
    ctx = ["Business Partner","Quadrant","Total_Revenue","Wtd_Margin_Pct","Rev_Rank"]
    if "Account_Type" in cust.columns: ctx.append("Account_Type")
    if "GPO" in cust.columns: ctx.append("GPO")
    pg = pg.merge(cust[ctx], on="Business Partner", how="left")
    pg["PG_Rev_Share"] = _safe_div(pg["PG_Revenue"], pg["Total_Revenue"])
    return pg.sort_values(["Rev_Rank","PG_Revenue"], ascending=[True,False])


# ═══════════════════════════════════════════════════════════════════
# 4. PRIORITY ACCOUNTS (Q2 deep-dive)
# ═══════════════════════════════════════════════════════════════════
def build_priority_accounts(df, cust):
    print("Building priority account detail ...")
    q2_bp = cust[cust["Quadrant"]=="Q2: High Rev / Low Margin"]["Business Partner"]
    d = df[df["Business Partner"].isin(q2_bp)]
    det = d.groupby(["Business Partner","Pricing Component Type"]).agg(
        Product_Group=("Product Group","first"), Lines=("Invoice ASP","size"),
        Revenue=("Invoice Total","sum"), Margin=("Std Margin","sum"),
        Avg_ASP=("Invoice ASP","mean"), Avg_Discount=("Discount %","mean"),
        Med_Margin_Pct=("Margin %","median"),
    ).reset_index()
    det["Wtd_Margin_Pct"] = _safe_div(det["Margin"], det["Revenue"])
    ctx = ["Business Partner","Quadrant","Total_Revenue","Wtd_Margin_Pct",
           "Rev_Rank","Margin_Recov_to_Med"]
    if "Account_Type" in cust.columns: ctx.append("Account_Type")
    if "GPO" in cust.columns: ctx.append("GPO")
    det = det.merge(cust[ctx].rename(columns={"Wtd_Margin_Pct":"Cust_Wtd_Margin"}),
                    on="Business Partner", how="left")
    return det.sort_values(["Rev_Rank","Revenue"], ascending=[True,False])


# ═══════════════════════════════════════════════════════════════════
# 5. MARGIN DISTRIBUTION BANDS
# ═══════════════════════════════════════════════════════════════════
def build_margin_distribution(cust):
    print("Building margin distribution ...")
    bins   = [-np.inf, 0, 0.25, 0.50, 0.60, 0.70, 0.80, 0.85, 0.90, 0.95, 1.0, np.inf]
    labels = ["< 0%","0-25%","25-50%","50-60%","60-70%",
              "70-80%","80-85%","85-90%","90-95%","95-100%","> 100%"]
    c2 = cust.copy()
    c2["Margin_Band"] = pd.cut(c2["Wtd_Margin_Pct"], bins=bins, labels=labels)
    dist = c2.groupby("Margin_Band", observed=False).agg(
        Customers=("Business Partner","size"), Total_Revenue=("Total_Revenue","sum"),
        Total_Margin=("Total_Std_Margin","sum"), Avg_Margin=("Wtd_Margin_Pct","mean"),
        Med_Discount=("Avg_Discount","median"),
    ).reset_index()
    dist["Pct_Customers"] = dist["Customers"] / dist["Customers"].sum()
    dist["Pct_Revenue"]   = dist["Total_Revenue"] / dist["Total_Revenue"].sum()
    dist["Wtd_Margin_Pct"] = _safe_div(dist["Total_Margin"], dist["Total_Revenue"])
    return dist


# ═══════════════════════════════════════════════════════════════════
# 6. EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════════
def build_summary(cust, qs):
    tr = cust["Total_Revenue"].sum()
    tm = cust["Total_Std_Margin"].sum()
    q2 = qs[qs["Quadrant"]=="Q2: High Rev / Low Margin"]
    return dict(
        n_cust=len(cust), total_rev=tr, total_margin=tm,
        overall_wtd=tm/tr if tr else 0,
        rev_threshold=cust.attrs.get("rev_threshold",0),
        margin_threshold=cust.attrs.get("margin_threshold",0),
        q2_custs=int(q2["Customers"].sum()) if len(q2) else 0,
        q2_rev=float(q2["Total_Revenue"].sum()) if len(q2) else 0,
        q2_recov=float(q2["Margin_Recovery"].sum()) if len(q2) else 0,
        total_recov=cust["Margin_Recov_to_Med"].sum(),
        top10_q2=cust[cust["Quadrant"]=="Q2: High Rev / Low Margin"].head(10)[
            ["Business Partner","Total_Revenue","Wtd_Margin_Pct",
             "Margin_Recov_to_Med"]].values.tolist(),
    )


# ═══════════════════════════════════════════════════════════════════
# EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════
def write_output(path, cust, qs, pg_mix, priority, margin_dist, S):
    print(f"Writing  {path} ...")
    wr = pd.ExcelWriter(path, engine="xlsxwriter")
    wb = wr.book

    H  = wb.add_format({"bold":1,"bg_color":"#1F3864","font_color":"#FFFFFF","border":1,
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

    def _ws(df, name, cols, tab="#4472C4", fz=1):
        df.to_excel(wr, sheet_name=name, startrow=1, index=False, header=False)
        ws = wr.sheets[name]; ws.set_tab_color(tab)
        ws.freeze_panes(1, fz); ws.autofilter(0, 0, len(df), len(df.columns)-1)
        for i, (cn, w, f) in enumerate(cols):
            ws.write(0, i, cn, H); ws.set_column(i, i, w, f)

    # ── Information ──
    wi = wb.add_worksheet("Information")
    wi.hide_gridlines(2); wi.set_tab_color("#305496")
    wi.set_column("A:A",4); wi.set_column("B:B",28); wi.set_column("C:C",95)
    r=1; wi.merge_range(r,1,r,2,"Customer Profitability Tiering Analysis",TI); r+=1
    wi.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2
    wi.merge_range(r,1,r,2,"Workbook Sheets",SB); r+=1
    wi.write(r,1,"Sheet",IH); wi.write(r,2,"Description",IH); r+=1
    for nm,ds in [
        ("Information","This sheet. Overview of tabs and key terms."),
        ("Executive Summary","Headline KPIs, quadrant totals, and top 10 Q2 accounts."),
        ("Customer Profitability","One row per customer with revenue, margin, quadrant, and recovery estimate."),
        ("Quadrant Summary","Aggregate metrics per quadrant."),
        ("Product Group Mix","Customer x Product Group revenue and margin detail."),
        ("Priority Accounts","Q2 accounts broken out by Pricing Component Type."),
        ("Margin Distribution","Customers grouped into margin bands with revenue share."),
    ]:
        wi.write(r,1,nm,IC); wi.write(r,2,ds,IC); wi.set_row(r,45); r+=1
    r+=1; wi.merge_range(r,1,r,2,"Key Terms",SB); r+=1
    wi.write(r,1,"Term",IH); wi.write(r,2,"Definition",IH); r+=1
    for tm,df_ in [
        ("Quadrant",f"Customers placed into four groups based on P{REVENUE_PERCENTILE} revenue "
         f"and P{MARGIN_PERCENTILE} margin thresholds."),
        ("Q1: High Rev / High Margin","Core accounts. Protect and grow."),
        ("Q2: High Rev / Low Margin","Priority targets for price correction or mix improvement."),
        ("Q3: Low Rev / High Margin","Good pricing. Potential growth targets."),
        ("Q4: Low Rev / Low Margin","Candidates for standard pricing or rationalisation."),
        ("Weighted Margin %","Total Std Margin / Total Revenue. Revenue-weighted."),
        ("Margin Gap (pp)","Customer margin minus overall median margin."),
        ("Margin Recovery to Median","Incremental margin if customer margin raised to median, holding revenue constant."),
        ("Revenue Rank","Customers ranked 1-N by descending revenue."),
        ("Cumulative Revenue Share","Running total of revenue share from rank 1 down."),
    ]:
        wi.write(r,1,tm,IT); wi.write(r,2,df_,IC); wi.set_row(r,40); r+=1

    # ── Executive Summary ──
    ws = wb.add_worksheet("Executive Summary")
    ws.hide_gridlines(2); ws.set_tab_color("#1F3864")
    ws.set_column("A:A",4); ws.set_column("B:B",50)
    ws.set_column("C:C",22); ws.set_column("D:D",22); ws.set_column("E:E",22)
    r=1; ws.merge_range(r,1,r,4,"Customer Profitability Tiering Analysis",TI); r+=1
    ws.write(r,1,f"Generated: {datetime.datetime.now():%d %b %Y %H:%M}",NT); r+=2
    ws.merge_range(r,1,r,4,"Dataset Overview",SB); r+=1
    for l,v,f in [("Total Customers",S["n_cust"],KN),("Total Invoice Revenue",S["total_rev"],KV),
                   ("Total Standard Margin",S["total_margin"],KV),("Overall Wtd Margin %",S["overall_wtd"],KP),
                   (f"Revenue Threshold (P{REVENUE_PERCENTILE})",S["rev_threshold"],KV),
                   (f"Margin Threshold (P{MARGIN_PERCENTILE})",S["margin_threshold"],KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1
    r+=1; ws.merge_range(r,1,r,4,"Quadrant Summary",SB); r+=1
    ws.write(r,1,"Quadrant",H); ws.write(r,2,"Customers",H)
    ws.write(r,3,"Revenue",H); ws.write(r,4,"Wtd Margin %",H); r+=1
    for _,qr in qs.iterrows():
        ws.write(r,1,qr["Quadrant"],T); ws.write(r,2,int(qr["Customers"]),N)
        ws.write(r,3,qr["Total_Revenue"],C); ws.write(r,4,qr["Wtd_Margin_Pct"],P); r+=1
    r+=1; ws.merge_range(r,1,r,4,"Margin Recovery Opportunity",SB); r+=1
    for l,v,f in [("Q2 Customers",S["q2_custs"],KN),("Q2 Revenue",S["q2_rev"],KV),
                   ("Q2 Recovery to Median",S["q2_recov"],KV),
                   ("Total Recovery (all below-median)",S["total_recov"],KV),
                   ("  -- as % of Revenue",S["total_recov"]/max(S["total_rev"],1),KP)]:
        ws.write(r,1,l,LB); ws.write(r,2,v,f); r+=1
    r+=1; ws.merge_range(r,1,r,4,"Top 10 Q2 Accounts by Recovery",SB); r+=1
    ws.write(r,1,"Customer",H); ws.write(r,2,"Revenue",H)
    ws.write(r,3,"Wtd Margin %",H); ws.write(r,4,"Recovery",H); r+=1
    for nm,rev,wm,rec in S["top10_q2"]:
        ws.write(r,1,nm,T); ws.write(r,2,rev,C); ws.write(r,3,wm,P); ws.write(r,4,rec,C); r+=1
    r+=2; ws.merge_range(r,1,r,4,"Methodology",SB); r+=1
    ws.merge_range(r,1,r+4,4,
        f"Customers assigned to quadrants using P{REVENUE_PERCENTILE} revenue "
        f"(${S['rev_threshold']:,.0f}) and P{MARGIN_PERCENTILE} margin "
        f"({S['margin_threshold']:.1%}) thresholds.  Recovery estimates the "
        f"incremental margin if below-median customers' margin were raised to the "
        f"overall median, holding revenue constant.", NT)

    # ── Customer Profitability ──
    cust_out = cust.copy()
    cc = [("Business Partner",38,T)]
    if "Account_Type" in cust_out.columns: cc.append(("Account Type",16,T))
    if "GPO" in cust_out.columns: cc.append(("GPO",18,T))
    cc += [("Quadrant",30,T),("Rev Rank",8,N),("Total Revenue",16,C),
           ("Total Std Margin",16,C),("Wtd Margin %",12,P),("Med Margin %",12,P),
           ("Avg Discount",12,P2),("Total Lines",10,N),("Total Units",12,N),
           ("Avg ASP",12,C2),("Rev per Line",12,C),("Product Groups",14,N),
           ("Component Types",14,N),("Components",12,N),("Revenue Share",12,P),
           ("Cum Rev Share",12,P),("Margin Gap (pp)",14,P),("Recovery to Median",16,C)]
    col_map = {"Business Partner":"Business Partner","Account Type":"Account_Type","GPO":"GPO",
               "Quadrant":"Quadrant","Rev Rank":"Rev_Rank","Total Revenue":"Total_Revenue",
               "Total Std Margin":"Total_Std_Margin","Wtd Margin %":"Wtd_Margin_Pct",
               "Med Margin %":"Med_Margin_Pct","Avg Discount":"Avg_Discount",
               "Total Lines":"Total_Lines","Total Units":"Total_Units","Avg ASP":"Avg_ASP",
               "Rev per Line":"Rev_per_Line","Product Groups":"Product_Groups",
               "Component Types":"Comp_Types","Components":"Components",
               "Revenue Share":"Rev_Share","Cum Rev Share":"Cum_Rev_Share",
               "Margin Gap (pp)":"Margin_Gap_pp","Recovery to Median":"Margin_Recov_to_Med"}
    sel = [v for k,v in col_map.items() if v in cust_out.columns and any(x[0]==k for x in cc)]
    cust_out = cust_out[sel]
    cc = [x for x in cc if col_map.get(x[0]) in sel]
    _ws(cust_out, "Customer Profitability", cc, tab="#548235")
    ws_cp = wr.sheets["Customer Profitability"]
    q_col = next(i for i,(cn,_,_) in enumerate(cc) if cn=="Quadrant")
    for qn,clr in Q_COLORS.items():
        ws_cp.conditional_format(1,q_col,len(cust_out),q_col,{
            "type":"text","criteria":"containing","value":qn[:2],
            "format":wb.add_format({"bg_color":clr,"font_color":"#FFFFFF","font_name":"Arial","font_size":10})})
    rec_col = next(i for i,(cn,_,_) in enumerate(cc) if cn=="Recovery to Median")
    ws_cp.conditional_format(1,rec_col,len(cust_out),rec_col,{"type":"data_bar","bar_color":"#C00000"})

    # ── Quadrant Summary ──
    _ws(qs, "Quadrant Summary", [("Quadrant",30,T),("Customers",10,N),("Total Revenue",16,C),
        ("Total Margin",16,C),("Wtd Margin %",12,P),("Avg Margin",12,P),("Med Margin",12,P),
        ("Avg Discount",12,P2),("Avg Product Groups",16,D),("Margin Recovery",16,C),
        ("Revenue Share",12,P)], tab="#7030A0")

    # ── Product Group Mix ──
    pgc = [("Business Partner",38,T)]
    if "Account_Type" in pg_mix.columns: pgc.append(("Account Type",16,T))
    if "GPO" in pg_mix.columns: pgc.append(("GPO",18,T))
    pgc += [("Quadrant",30,T),("Cust Revenue",14,C),("Cust Margin %",12,P),
            ("Product Group",20,T),("PG Revenue",14,C),("PG Margin",14,C),
            ("PG Margin %",10,P),("PG Lines",8,N),("PG Rev Share",12,P)]
    pg_map = {"Business Partner":"Business Partner","Account Type":"Account_Type","GPO":"GPO",
              "Quadrant":"Quadrant","Cust Revenue":"Total_Revenue","Cust Margin %":"Wtd_Margin_Pct",
              "Product Group":"Product Group","PG Revenue":"PG_Revenue","PG Margin":"PG_Margin",
              "PG Margin %":"PG_Wtd_Margin","PG Lines":"PG_Lines","PG Rev Share":"PG_Rev_Share"}
    pg_sel = [v for k,v in pg_map.items() if v in pg_mix.columns]
    pgc = [x for x in pgc if pg_map.get(x[0]) in pg_sel]
    _ws(pg_mix[pg_sel], "Product Group Mix", pgc, tab="#4472C4")

    # ── Priority Accounts ──
    prc = [("Business Partner",38,T)]
    if "Account_Type" in priority.columns: prc.append(("Account Type",16,T))
    if "GPO" in priority.columns: prc.append(("GPO",18,T))
    prc += [("Cust Revenue",14,C),("Cust Margin %",12,P),("Recovery to Med",14,C),
            ("Component Type",30,T),("Product Group",18,T),("Lines",8,N),
            ("Revenue",14,C),("Margin",14,C),("Wtd Margin %",12,P),
            ("Avg ASP",12,C2),("Avg Discount",12,P2),("Med Margin %",12,P)]
    pr_map = {"Business Partner":"Business Partner","Account Type":"Account_Type","GPO":"GPO",
              "Cust Revenue":"Total_Revenue","Cust Margin %":"Cust_Wtd_Margin",
              "Recovery to Med":"Margin_Recov_to_Med","Component Type":"Pricing Component Type",
              "Product Group":"Product_Group","Lines":"Lines","Revenue":"Revenue","Margin":"Margin",
              "Wtd Margin %":"Wtd_Margin_Pct","Avg ASP":"Avg_ASP","Avg Discount":"Avg_Discount",
              "Med Margin %":"Med_Margin_Pct"}
    pr_sel = [v for k,v in pr_map.items() if v in priority.columns]
    prc = [x for x in prc if pr_map.get(x[0]) in pr_sel]
    _ws(priority[pr_sel], "Priority Accounts", prc, tab="#C00000")

    # ── Margin Distribution ──
    _ws(margin_dist, "Margin Distribution", [("Margin Band",12,T),("Customers",10,N),
        ("Total Revenue",16,C),("Total Margin",16,C),("Wtd Margin %",12,P),
        ("Avg Margin",12,P),("Med Discount",12,P2),("% Customers",12,P),
        ("% Revenue",12,P)], tab="#BF8F00")

    wr.close()
    print("  Workbook saved.")


# ═══════════════════════════════════════════════════════════════════
# QUADRANT MAP (matplotlib scatter)
# ═══════════════════════════════════════════════════════════════════
def build_quadrant_chart(cust, S, chart_path):
    """Render a Revenue-vs-Margin scatter, save as PNG, and display."""
    # Let matplotlib auto-detect the best backend.  On a desktop
    # (Windows/macOS/Linux with display) this will typically be TkAgg
    # and plt.show() will open an interactive window.  In headless
    # environments it falls back to Agg and plt.show() is a no-op.
    import matplotlib
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker

    print("Building quadrant map chart ...")
    fig, ax = plt.subplots(figsize=(12, 7.5))

    for q in Q_ORDER:
        sub = cust[cust["Quadrant"] == q]
        if sub.empty: continue
        ax.scatter(sub["Total_Revenue"], sub["Wtd_Margin_Pct"] * 100,
                   c=Q_COLORS[q], label=q, s=40, alpha=0.7,
                   edgecolors="white", linewidths=0.4)

    # Threshold lines
    rev_t = S["rev_threshold"]
    mar_t = S["margin_threshold"] * 100
    ax.axvline(rev_t, color="#888888", lw=1.2, ls="--",
               label=f"Revenue P{REVENUE_PERCENTILE}")
    ax.axhline(mar_t, color="#888888", lw=1.2, ls="--",
               label=f"Margin P{MARGIN_PERCENTILE}")

    # Apply configured axis limits
    xlim, ylim = ax.get_xlim(), ax.get_ylim()
    if CHART_REVENUE_MIN is not None: xlim = (CHART_REVENUE_MIN, xlim[1])
    if CHART_REVENUE_MAX is not None: xlim = (xlim[0], CHART_REVENUE_MAX)
    if CHART_MARGIN_MIN  is not None: ylim = (CHART_MARGIN_MIN * 100, ylim[1])
    if CHART_MARGIN_MAX  is not None: ylim = (ylim[0], CHART_MARGIN_MAX * 100)
    ax.set_xlim(xlim); ax.set_ylim(ylim)

    # Quadrant corner labels
    px = (xlim[1]-xlim[0]) * 0.02
    py = (ylim[1]-ylim[0]) * 0.03
    for (txt, x, y, ha, va), clr in zip([
        ("Q1: Protect & Grow",              xlim[1]-px, ylim[1]-py, "right","top"),
        ("Q2: Priority - Price Correction",  xlim[1]-px, ylim[0]+py, "right","bottom"),
        ("Q3: Maintain Discipline",          xlim[0]+px, ylim[1]-py, "left", "top"),
        ("Q4: Rationalise / Enforce",        xlim[0]+px, ylim[0]+py, "left", "bottom"),
    ], [Q_COLORS[q] for q in Q_ORDER]):
        ax.text(x, y, txt, fontsize=9, fontweight="bold", color=clr,
                ha=ha, va=va, alpha=0.5)

    ax.set_title("Customer Profitability Quadrant Map",
                 fontsize=15, fontweight="bold", color="#1F3864", pad=14)
    ax.set_xlabel("Total Revenue ($)", fontsize=11, labelpad=8)
    ax.set_ylabel("Weighted Margin (%)", fontsize=11, labelpad=8)
    ax.set_xscale("linear")
    ax.set_yscale("linear")
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"${x:,.0f}"))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda y,_: f"{y:.0f}%"))
    ax.grid(True, alpha=0.3)
    ax.set_facecolor("#FAFAFA")
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.08),
              ncol=3, fontsize=9, frameon=True, fancybox=True)
    fig.tight_layout(rect=[0, 0.04, 1, 1])

    fig.savefig(chart_path, dpi=180, bbox_inches="tight")
    print(f"  Chart saved: {Path(chart_path).name}")

    # Display the chart interactively.  On desktop environments this
    # opens a window; in headless environments it is silently skipped.
    try:
        plt.show()
    except Exception:
        pass
    plt.close(fig)


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    warnings.filterwarnings("ignore", category=UserWarning)
    path = sys.argv[1] if len(sys.argv) > 1 else _pick_file()
    if not os.path.isfile(path): sys.exit(f"Not found: {path}")

    stem    = Path(path).stem
    out_dir = Path(path).parent
    out_xl  = str(out_dir / f"{stem}_ProfitabilityTiering.xlsx")
    out_png = str(out_dir / f"{stem}_QuadrantMap.png")

    df    = load_data(path)
    cust  = build_customer_profile(df)
    qs    = build_quadrant_summary(cust)
    pgmix = build_product_mix(df, cust)
    pri   = build_priority_accounts(df, cust)
    mdist = build_margin_distribution(cust)
    S     = build_summary(cust, qs)

    write_output(out_xl, cust, qs, pgmix, pri, mdist, S)
    build_quadrant_chart(cust, S, out_png)

    print(f"\n{'='*65}")
    print("  HEADLINE RESULTS")
    print(f"{'='*65}")
    print(f"  Customers:                  {S['n_cust']:>14,}")
    print(f"  Total Revenue:              ${S['total_rev']:>14,.0f}")
    print(f"  Overall Wtd Margin:          {S['overall_wtd']:>13.1%}")
    print(f"  Q2 Customers (Hi Rev/Lo Mar):{S['q2_custs']:>13,}")
    print(f"  Q2 Revenue:                 ${S['q2_rev']:>14,.0f}")
    print(f"  Q2 Margin Recovery:         ${S['q2_recov']:>14,.0f}")
    print(f"  Total Recovery (all):       ${S['total_recov']:>14,.0f}  "
          f"({S['total_recov']/S['total_rev']:.1%} of revenue)")
    print(f"{'='*65}")


if __name__ == "__main__":
    main()
