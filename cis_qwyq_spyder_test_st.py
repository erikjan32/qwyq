# -*- coding: utf-8 -*-
"""
Created on Thu Dec  4 07:39:52 2025

@author: ejynt
"""


#"""
#CIS QWYQ (Spyder-friendly)
#--------------------------
#Single-file script you can run inside Spyder IDE:
#- Put this .py next to your enhanced template (cis_qwyq_template_enhanced.xlsx)
#- Edit the CONFIG section below (paths, score_method, output_tag)
#- Click "Run" in Spyder or call main() from the console

#You can also import this file and call the functions:
#   from cis_qwyq_spyder import load_inputs, score_controls, select_portfolio, build_roadmap, save_outputs

#Dependencies:
#   pip install pandas numpy openpyxl matplotlib
#"""

import os
import streamlit as st
import matplotlib
import plost
from datetime import datetime
import pandas as pd
import numpy as np
import math

# -----------------------------
# CONFIG (edit here in Spyder)
# -----------------------------
#INPUT_XLSX = "C:/Users/ejynt/OneDrive/QWYQ/Sources/inputs/cis_qwyq_template_enhanced_test_IG1.xlsx" # path to the enhanced template
INPUT_XLSX = "cis_qwyq_template_enhanced_test_IG1.xlsx" # path to the enhanced template
SCORE_METHOD = "z"                               # "z" or "pct"
OUTPUT_TAG = datetime.today().strftime("%y%m%d") # string used in filenames

# Visualization toggles (PNG charts)
MAKE_PLOTS = True
PLOT_DIR = "C:/Users/ejynt/OneDrive/QWYQ/Sources/outputs"  # where to write PNGs

# -----------------------------
# Core constants
# -----------------------------
CRITERIA = [
    "RiskReduction","QuickWins","OperationalComplexity",
    "MaturityTrend","CostEfficiency","RegulatoryAlignment","BusinessCriticality"
]

DEFAULT_WEIGHTS = {
    "RiskReduction": 0.25, "QuickWins": 0.10, "OperationalComplexity": 0.10,
    "MaturityTrend": 0.10, "CostEfficiency": 0.10, "RegulatoryAlignment": 0.20, "BusinessCriticality": 0.15
}

# -----------------------------
# Utility functions
# -----------------------------
def zscore(series: pd.Series) -> pd.Series:
    mu = series.mean(skipna=True)
    sd = series.std(skipna=True, ddof=0)
    if sd == 0 or np.isnan(sd):
        return pd.Series(np.zeros(len(series)), index=series.index)
    return (series - mu) / sd

def pct_rank(series: pd.Series) -> pd.Series:
    return series.rank(pct=True, method="average")

def normalize_weights(weights_df: pd.DataFrame | None, overrides: dict | None = None) -> dict:
    w = dict(DEFAULT_WEIGHTS)
    if weights_df is not None and not weights_df.empty:
        for _, r in weights_df.iterrows():
            c = str(r.get("Criterion"))
            if c in w:
                try:
                    w[c] = float(r.get("Weight", w[c]))
                except Exception:
                    pass
    if overrides:
        for k, v in overrides.items():
            if k in w:
                w[k] = float(v)
    s = sum(w.values())
    if s <= 0:
        raise ValueError("Weights must sum to a positive number")
    return {k: v/s for k, v in w.items()}

# -----------------------------
# IO functions
# -----------------------------
def load_inputs(path_xlsx: str):
    """Load the enhanced template sheets."""
    if not os.path.exists(path_xlsx):
        raise FileNotFoundError(f"Input not found: {path_xlsx}")
    controls = pd.read_excel(path_xlsx, sheet_name="Controls")
    weights  = pd.read_excel(path_xlsx, sheet_name="Weights")
    constraints = pd.read_excel(path_xlsx, sheet_name="Constraints")
    params = pd.read_excel(path_xlsx, sheet_name="Params")
    return controls, weights, constraints, params

def save_outputs(base_name: str, output_tag: str, ranked: pd.DataFrame, portfolio: pd.DataFrame,
                 dom_agg: pd.DataFrame, roadmap: pd.DataFrame, weights_used: dict,
                 plot_paths: dict | None = None, outdir: str = "."):
    """Save CSV, Excel, and optionally include plot paths in a simple summary print."""
    csv_scores = os.path.join(outdir, f"C:/Users/ejynt/OneDrive/QWYQ/Sources/outputs/{base_name}_scores_{output_tag}.csv")
    xlsx_all   = os.path.join(outdir, f"C:/Users/ejynt/OneDrive/QWYQ/Sources/outputs/{base_name}_outputs_{output_tag}.xlsx")

    ranked.to_csv(csv_scores, index=False)
    with pd.ExcelWriter(xlsx_all, engine="openpyxl") as writer:
        ranked.to_excel(writer, index=False, sheet_name="Ranking")
        portfolio.to_excel(writer, index=False, sheet_name="Portfolio")
        dom_agg.to_excel(writer, index=False, sheet_name="DomainEffort")
        roadmap.to_excel(writer, index=False, sheet_name="Roadmap")
        pd.DataFrame({"Criterion": list(weights_used.keys()), "WeightNormalized": list(weights_used.values())}).to_excel(
            writer, index=False, sheet_name="WeightsUsed"
        )
    print("Saved:")
    print(" -", csv_scores)
    print(" -", xlsx_all)
    if plot_paths:
        for k, p in plot_paths.items():
            print(f" - {k}: {p}")

# -----------------------------
# Scoring & selection functions
# -----------------------------
def parse_params(params_df: pd.DataFrame):
    total_budget = float(params_df.loc[params_df["Parameter"]=="TotalBudgetDays", "Value"].iloc[0])
    quarters_str = str(params_df.loc[params_df["Parameter"]=="Quarters", "Value"].iloc[0])
    quarters = [q.strip() for q in quarters_str.split(",") if q.strip()]
    return total_budget, quarters

def score_controls(df_controls: pd.DataFrame, weights: dict, score_method: str = "z") -> pd.DataFrame:
    df = df_controls.copy()
    for c in CRITERIA + ["EffortDays"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    # invert OperationalComplexity (lower is better)
    df["OperationalComplexity_Scored"] = -df["OperationalComplexity"]

    comp_map = {
        "RiskReduction": df["RiskReduction"],
        "QuickWins": df["QuickWins"],
        "OperationalComplexity": df["OperationalComplexity_Scored"],
        "MaturityTrend": df["MaturityTrend"],
        "CostEfficiency": df["CostEfficiency"],
        "RegulatoryAlignment": df["RegulatoryAlignment"],
        "BusinessCriticality": df["BusinessCriticality"],
    }
    scores = {}
    for k, s in comp_map.items():
        series = s.fillna(s.median())
        scores[k] = zscore(series) if score_method=="z" else series.rank(pct=True, method="average")
    scores_df = pd.DataFrame(scores)
    total = sum(scores_df[c] * weights[c] for c in CRITERIA)
    ranked = pd.concat([df[["ControlID","ControlName","Domain"] + CRITERIA + ["EffortDays"]], scores_df.add_suffix("_Score")], axis=1)
    ranked["QWYQ_TotalScore"] = total
    ranked = ranked.sort_values("QWYQ_TotalScore", ascending=False).reset_index(drop=True)
    ranked["Rank"] = np.arange(1, len(ranked)+1)
    return ranked

def select_portfolio(ranked: pd.DataFrame, constraints_df: pd.DataFrame, total_budget: float):
    domains = constraints_df["Domain"].tolist()
    min_share = {r["Domain"]: float(r["MinShare"]) for _, r in constraints_df.iterrows()}
    max_share = {r["Domain"]: float(r["MaxShare"]) for _, r in constraints_df.iterrows()}

    min_budget = {d: total_budget * min_share.get(d, 0.0) for d in domains}
    max_budget = {d: total_budget * max_share.get(d, 1.0) for d in domains}

    remaining = total_budget
    spent = {d: 0.0 for d in domains}
    selected_rows = []

    # pass 1: satisfy minimums
    for d in domains:
        domain_items = ranked[ranked["Domain"]==d].copy().sort_values("QWYQ_TotalScore", ascending=False)
        for _, row in domain_items.iterrows():
            if spent[d] >= min_budget[d] or remaining <= 0:
                break
            if row["EffortDays"] <= 0:
                continue
            if spent[d] + row["EffortDays"] <= max_budget[d] and row["EffortDays"] <= remaining:
                selected_rows.append(row)
                spent[d] += row["EffortDays"]
                remaining -= row["EffortDays"]

    # pass 2: fill remaining by global rank
    for _, row in ranked.iterrows():
        d = row["Domain"]
        if row["EffortDays"] <= 0 or remaining <= 0:
            continue
        if any((s["ControlID"]==row["ControlID"]) for s in selected_rows):
            continue
        if spent[d] + row["EffortDays"] <= max_budget[d] and row["EffortDays"] <= remaining:
            selected_rows.append(row)
            spent[d] += row["EffortDays"]
            remaining -= row["EffortDays"]

    portfolio = pd.DataFrame(selected_rows).sort_values("Rank").reset_index(drop=True)
    if not portfolio.empty:
        dom_agg = portfolio.groupby("Domain")["EffortDays"].sum().reset_index().rename(columns={"EffortDays":"EffortUsed"})
    else:
        dom_agg = pd.DataFrame(columns=["Domain","EffortUsed"])
    return portfolio, dom_agg, spent, remaining

def build_roadmap(portfolio: pd.DataFrame, quarters: list, slot_days: int = 30):
    if portfolio.empty:
        return pd.DataFrame(columns=["ControlID","ControlName","Domain","Rank","EffortDays","StartQuarter","EndQuarter"])
    qslots = quarters.copy()
    schedule = []
    q_index = 0
    for _, r in portfolio.iterrows():
        days = float(r["EffortDays"])
        slots_needed = max(1, int(math.ceil(days / slot_days)))
        start_q = qslots[q_index % len(qslots)]
        end_q = qslots[(q_index + slots_needed - 1) % len(qslots)]
        schedule.append({
            "ControlID": r["ControlID"],
            "ControlName": r["ControlName"],
            "Domain": r["Domain"],
            "Rank": r["Rank"],
            "EffortDays": days,
            "StartQuarter": start_q,
            "EndQuarter": end_q
        })
        q_index += slots_needed
    road = pd.DataFrame(schedule).sort_values(["StartQuarter","Rank"])
    return road

# -----------------------------
# Optional plotting helpers
# -----------------------------
def plot_effort_by_domain(portfolio: pd.DataFrame, outpath: str):
    if portfolio.empty:
        return None
    agg = portfolio.groupby("Domain")["EffortDays"].sum().sort_values(ascending=False)
    import matplotlib.pyplot as plt
    plt.figure()
    agg.plot(kind="bar")
    plt.title("Effort by Domain")
    plt.ylabel("EffortDays")
    plt.tight_layout()
    plt.savefig(outpath)
    plt.close()
    return outpath

def plot_heatmap_top10(ranked: pd.DataFrame, outpath: str):
    top = ranked.head(10)
    crit = CRITERIA
    data = top[crit].to_numpy(dtype=float)
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()
    im = ax.imshow(data, aspect="auto")
    ax.set_xticks(range(len(crit)))
    ax.set_xticklabels(crit, rotation=45, ha="right")
    ax.set_yticks(range(len(top)))
    ax.set_yticklabels(top["ControlID"].astype(str)+" "+top["ControlName"].astype(str))
    plt.title("Criteria Heatmap (Top 10)")
    plt.tight_layout()
    plt.savefig(outpath)
    plt.close()
    return outpath

def plot_radar_by_domain(ranked: pd.DataFrame, outpath: str):
    if ranked.empty:
        return None
    domains = sorted(ranked["Domain"].unique().tolist())
    crit = CRITERIA
    import numpy as np
    angles = np.linspace(0, 2*np.pi, len(crit), endpoint=False).tolist()
    angles += angles[:1]
    import matplotlib.pyplot as plt
    fig = plt.figure()
    ax = plt.subplot(111, polar=True)
    for d in domains:
        subset = ranked[ranked["Domain"]==d]
        vals = [subset[c].mean() for c in crit]
        vals += vals[:1]
        ax.plot(angles, vals, label=d)
        ax.fill(angles, vals, alpha=0.05)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(crit)
    plt.title("Average Criteria per Domain (Radar)")
    plt.legend(loc="upper right", bbox_to_anchor=(1.3, 1.1))
    plt.tight_layout()
    plt.savefig(outpath, bbox_inches="tight")
    plt.close()
    return outpath


# -----------------------------
# Main runner (for Spyder "Run file" or interactive use)
# -----------------------------
def main(input_xlsx: str | None = None, score_method: str | None = None, output_tag: str | None = None,
         make_plots: bool | None = None, plot_dir: str | None = None):
    path = input_xlsx or INPUT_XLSX
    method = (score_method or SCORE_METHOD).lower()
    tag = output_tag or OUTPUT_TAG
    do_plots = MAKE_PLOTS if make_plots is None else bool(make_plots)
    pdir = plot_dir or PLOT_DIR

    controls, weights_df, constraints, params = load_inputs(path)
    weights = normalize_weights(weights_df, overrides=None)
    total_budget, quarters = parse_params(params)

    ranked = score_controls(controls, weights, score_method=method)
    portfolio, dom_agg, spent, remaining = select_portfolio(ranked, constraints, total_budget)
    roadmap = build_roadmap(portfolio, quarters)


    if st.checkbox('Show ranked controls'):
             st.subheader('Ranked controls')
             st.write(ranked)
         
    if st.checkbox('Show aggregated domains'):
          st.subheader('Agg domains')
          st.write(dom_agg)   

    'scores by bar'
    st.bar_chart(ranked, x='ControlName', y=('RiskReduction_Score', 'QuickWins_Score', 'OperationalComplexity_Score', 'CostEfficiency_Score', 'RegulatoryAlignment_Score', 'BusinessCriticality_Score'),  horizontal=True) 
    'scores by scatter'
    st.scatter_chart(ranked, x='ControlName', y=('RiskReduction_Score', 'QuickWins_Score', 'OperationalComplexity_Score', 'CostEfficiency_Score', 'RegulatoryAlignment_Score', 'BusinessCriticality_Score'))
    
   
        
    st.bar_chart(dom_agg, x='Domain', y='EffortUsed')
    
    if st.checkbox('Show aroadmap'):
        st.subheader('Roadmap')
        st.write(roadmap)
        
    plost.line_chart(
        ranked,
        x='ControlName',
        y=('RiskReduction_Score', 'QuickWins_Score', 'OperationalComplexity_Score'),
        pan_zoom='minimap',  # 👈 This is magic!
        )
    
    plot_paths = {}
    if do_plots:
        base = os.path.splitext(os.path.basename(path))[0]
        eff = os.path.join(pdir, f"{base}_effort_domain_{tag}.png")
        heat = os.path.join(pdir, f"{base}_heatmap_top10_{tag}.png")
        rad = os.path.join(pdir, f"{base}_radar_domain_{tag}.png")
        if plot_effort_by_domain(portfolio, eff): plot_paths["effort_by_domain"] = eff
        if plot_heatmap_top10(ranked, heat): plot_paths["heatmap_top10"] = heat
        if plot_radar_by_domain(ranked, rad): plot_paths["radar_by_domain"] = rad

  
    
    base_name = os.path.splitext(os.path.basename(path))[0]
    save_outputs(base_name, tag, ranked, portfolio, dom_agg, roadmap, weights, plot_paths, outdir=".")
    return ranked, portfolio, dom_agg, roadmap, weights, plot_paths

# When running the file directly in Spyder
if __name__ == "__main__":
    main()
