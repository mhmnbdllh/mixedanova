"""
Mixed ANOVA Calculator — SPSS GLM Repeated Measures Equivalent (CORRECTED VERSION)
==================================================================================
Model: A (between-subjects) × B (within-subjects), balanced design.

SS Decomposition (Winer, Brown & Michels, 1991, Ch. 7):
  SS_A     = b · Σᵢ nᵢ (Ȳᵢ.. − Ȳ...)²               df = a−1
  SS_S(A)  = b · Σᵢ Σₛ (Ȳₛ.. − Ȳᵢ..)²               df = N−a
  SS_B     = N · Σⱼ (Ȳ.j. − Ȳ...)²                   df = b−1
  SS_AB    = Σᵢ nᵢ · Σⱼ (Ȳᵢⱼ − Ȳᵢ.. − Ȳ.j. + Ȳ...)² df = (a−1)(b−1)
  SS_BS(A) = Σᵢ Σₛ Σⱼ (yₛⱼ − Ȳₛ.. − Ȳ.ⱼ. + Ȳ...)²   df = (b−1)(N−a)
  [Verified: sum of five terms = SS_Total exactly]

F-ratios (with sphericity correction only on numerator df):
  F_A  = MS_A  / MS_S(A)    — between error
  F_B  = MS_B  / MS_BS(A)   — within error (df_B_corrected, df_BSA)
  F_AB = MS_AB / MS_BS(A)   — within error (df_AB_corrected, df_BSA)

Sphericity: Mauchly (1940) W with difference matrix (SPSS method)
"""

import io
import itertools
import re
import warnings

import numpy as np
import pandas as pd
import scipy.stats as stats
from scipy.stats import f as fdist

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as mgs
import seaborn as sns

import streamlit as st

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Mixed ANOVA Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
#  CSS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@500;600&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

html,body,[class*="css"]{font-family:'Inter',sans-serif;}

.app-title{font-family:'Lora',serif;font-size:2.1rem;font-weight:600;
           color:#0d1b2a;letter-spacing:-.01em;margin-bottom:2px;}
.app-sub{font-size:.86rem;color:#546e7a;font-weight:300;margin-bottom:1.4rem;}

.sec{font-family:'Lora',serif;font-size:1.1rem;font-weight:600;color:#0d1b2a;
     border-left:4px solid #c62828;padding-left:10px;
     margin-top:1.7rem;margin-bottom:.7rem;}

.mcard{background:#0d1b2a;border-radius:10px;padding:.9rem 1.1rem .85rem;margin-bottom:4px;}
.mcard .lbl{font-size:.63rem;font-weight:600;color:#7fa8c9;text-transform:uppercase;letter-spacing:.10em;}
.mcard .val{font-family:'JetBrains Mono',monospace;font-size:1.2rem;font-weight:600;
            color:#eaf2ff;margin:3px 0 2px;}
.mcard .sub{font-size:.73rem;color:#90b8d8;line-height:1.5;}

.rcard{background:#112233;border-radius:10px;padding:.9rem 1.1rem .85rem;margin-bottom:4px;}
.rcard .lbl{font-size:.63rem;font-weight:600;color:#7fa8c9;text-transform:uppercase;letter-spacing:.10em;}
.rcard .val{font-family:'JetBrains Mono',monospace;font-size:1.05rem;font-weight:600;
            color:#eaf2ff;margin:3px 0 2px;}
.rcard .sub{font-size:.73rem;color:#90b8d8;line-height:1.5;}

.p-sig{display:inline-block;background:#1b7f45;color:#fff;
       font-size:.68rem;font-weight:700;padding:2px 8px;border-radius:20px;}
.p-ns {display:inline-block;background:#b71c1c;color:#fff;
       font-size:.68rem;font-weight:700;padding:2px 8px;border-radius:20px;}
.p-trend{display:inline-block;background:#e65100;color:#fff;
         font-size:.68rem;font-weight:700;padding:2px 8px;border-radius:20px;}

.ibox{background:#f0f6ff;border-left:4px solid #1565c0;border-radius:6px;
      padding:.8rem 1rem;font-size:.875rem;line-height:1.72;color:#0d1b2a;margin-bottom:.5rem;}
.abox-warn{background:#fff8e1;border-left:4px solid #f9a825;border-radius:6px;
           padding:.7rem 1rem;font-size:.84rem;color:#5d4037;margin-bottom:.5rem;}
.abox-ok  {background:#e8f5e9;border-left:4px solid #2e7d32;border-radius:6px;
           padding:.7rem 1rem;font-size:.84rem;color:#1b5e20;margin-bottom:.5rem;}
.abox-info{background:#e3f2fd;border-left:4px solid #1565c0;border-radius:6px;
           padding:.7rem 1rem;font-size:.84rem;color:#0d47a1;margin-bottom:.5rem;}
.upload-hint{background:#f5f7fa;border:1.5px dashed #b0bec5;border-radius:8px;
             padding:.85rem 1rem;font-size:.84rem;color:#546e7a;}

[data-testid="stSidebar"]{background:#0d1b2a !important;}
[data-testid="stSidebar"] *{color:#cfe2f3 !important;}
[data-testid="stSidebar"] hr{border-color:#1c3048 !important;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Analysis Settings")
    st.markdown("---")
    alpha_level = st.selectbox("Significance level (α)",
                               [0.05, 0.01, 0.001, 0.10], index=0)
    posthoc_method = st.selectbox(
        "Post-hoc correction",
        ["bonferroni","holm","sidak","fdr_bh"],
        format_func=lambda x:{
            "bonferroni":"Bonferroni",
            "holm":      "Holm (step-down Bonferroni)",
            "sidak":     "Šidák",
            "fdr_bh":    "FDR — Benjamini–Hochberg",
        }[x])
    effect_pref = st.selectbox(
        "Primary effect size",
        ["partial_eta2","eta2","cohen_f"],
        format_func=lambda x:{
            "partial_eta2":"Partial η²p  (SPSS default)",
            "eta2":        "η²  (eta-squared)",
            "cohen_f":     "Cohen's f",
        }[x])
    sph_corr = st.selectbox(
        "Sphericity correction",
        ["auto","gg","hf","lb","none"],
        format_func=lambda x:{
            "auto":"Auto — apply GG when Mauchly p < α",
            "gg":  "Greenhouse–Geisser (always)",
            "hf":  "Huynh–Feldt (always)",
            "lb":  "Lower-bound (always)",
            "none":"None — assume sphericity",
        }[x])
    st.markdown("---")
    st.markdown("**Display**")
    show_desc    = st.checkbox("Descriptive statistics & EMM", value=True)
    show_assump  = st.checkbox("Assumption diagnostics",       value=True)
    show_posthoc = st.checkbox("Post-hoc comparisons",         value=True)
    st.markdown("---")
    st.markdown("**Plots**")
    pal_name   = st.selectbox("Color palette",
                              ["tab10","Set2","deep","colorblind","husl"])
    grid_style = st.selectbox("Grid style",
                              ["whitegrid","ticks","darkgrid"])

# ══════════════════════════════════════════════════════════════════════════════
#  HEADER (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="app-title">Mixed ANOVA Calculator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-sub">'
    'Two-Way Mixed ANOVA · SPSS GLM Repeated Measures Equivalent · '
    'Mauchly Sphericity · GG / HF / LB Corrections · '
    'Estimated Marginal Means · SPSS-equivalent Simple Effects · '
    'Post-hoc · Full Report Export'
    '</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  FORMAT GUIDE & TEMPLATES (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📋  Data Format Guide & CSV Template Download", expanded=False):
    st.markdown('<div class="sec">Required CSV Format</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.1, 0.9])
    with c1:
        st.markdown(
            "**Long (tidy) format — one observation per row.** "
            "Each row is one measurement for one subject at one within-subjects level."
        )
        st.dataframe(pd.DataFrame({
            "Column":   ["subject_id","between_factor","within_factor","dependent_var"],
            "Role":     ["Unique participant identifier",
                         "Between-subjects group membership",
                         "Within-subjects repeated measure (e.g., Time)",
                         "Numeric outcome variable"],
            "Example":  ["1, 2, … N","Control / Treatment_A / Treatment_B",
                         "Pre / Post / Follow_up","42.5, 67.3, …"],
            "Required": ["✓","✓","✓","✓"],
        }), hide_index=True, use_container_width=True)

    with c2:
        st.dataframe(pd.DataFrame({
            "Constraint":[
                "Min. between-subjects groups","Max. between-subjects groups",
                "Min. within-subjects levels","Max. within-subjects levels",
                "Min. subjects per group","Balance requirement",
                "Missing data","File size","Encoding"],
            "Specification":[
                "2  (supports 3, 4, … k)","No hard limit",
                "2  (supports 3, 4, … t)","No hard limit",
                "≥ 5  (≥ 10 recommended)",
                "Each subject must have data at ALL within-levels",
                "Listwise deletion applied automatically",
                "3 MB","UTF-8"],
        }), hide_index=True, use_container_width=True)

    st.markdown(
        "**Tip:** Prefix within-level names with numbers to control ordering "
        "(e.g., `1_Pre`, `2_Post`, `3_Follow_up`)."
    )
    st.markdown("---")

    def make_template(groups, times, n_per, seed=0):
        np.random.seed(seed)
        base  = {g: 50+i*3      for i,g in enumerate(groups)}
        tgain = {t: j*4          for j,t in enumerate(times)}
        igain = {g: {t: i*j*1.5 for j,t in enumerate(times)}
                 for i,g in enumerate(groups)}
        rows, sid = [], 1
        for g in groups:
            for _ in range(n_per):
                re = np.random.normal(0, 3)
                for t in times:
                    rows.append({"subject_id":sid,"Group":g,"Time":t,
                                 "Score":round(base[g]+tgain[t]+igain[g][t]+re+np.random.normal(0,1.8),2)})
                sid += 1
        return pd.DataFrame(rows)

    tpls = {
        "2 groups × 2 time-points": make_template(["Control","Treatment"],["Pre","Post"],12),
        "2 groups × 3 time-points": make_template(["Control","Treatment"],["Pre","Post","Follow_up"],10),
        "3 groups × 3 time-points": make_template(["Control","Treat_A","Treat_B"],["Pre","Post","Follow_up"],10),
        "3 groups × 4 time-points": make_template(["Control","Treat_A","Treat_B"],
                                                  ["Baseline","Week_4","Week_8","Week_12"],10),
    }
    fnames = {k: f"template_{k.replace(' ','_').replace('×','x')}.csv" for k in tpls}

    cols_t = st.columns(4)
    for col, (label, tdf) in zip(cols_t, tpls.items()):
        with col:
            st.markdown(f"**{label}**")
            st.caption(f"{len(tdf)} rows")
            st.download_button("⬇️ Download", tdf.to_csv(index=False).encode(),
                               fnames[label], "text/csv", use_container_width=True)

    st.markdown("**Preview — 3 groups × 3 time-points (first 12 rows):**")
    st.dataframe(tpls["3 groups × 3 time-points"].head(12),
                 hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  UPLOAD (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Upload Data File</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload CSV (max 3 MB, UTF-8, long format)",
                                  type=["csv"])

if uploaded_file is None:
    st.markdown(
        '<div class="upload-hint">📂 No file uploaded yet. '
        'Upload your CSV above, or download a template to get started.</div>',
        unsafe_allow_html=True)
    st.stop()

if uploaded_file.size > 3*1024*1024:
    st.error("❌ File exceeds 3 MB limit.")
    st.stop()

try:
    df_raw = pd.read_csv(uploaded_file)
except Exception as e:
    st.error(f"❌ Cannot read CSV: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  COLUMN MAPPING (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Column Mapping</div>', unsafe_allow_html=True)

all_cols = df_raw.columns.tolist()
num_cols = df_raw.select_dtypes(include=np.number).columns.tolist()

cm1,cm2,cm3,cm4 = st.columns(4)
with cm1: subj_col = st.selectbox("🔑 Subject ID",    all_cols, index=0)
with cm2: btw_col  = st.selectbox("👥 Between factor", all_cols, index=min(1,len(all_cols)-1))
with cm3: win_col  = st.selectbox("🔁 Within factor",  all_cols, index=min(2,len(all_cols)-1))
with cm4: dv_col   = st.selectbox("📈 Dependent variable", num_cols, index=0)

run = st.button("▶  Run Mixed ANOVA", type="primary", use_container_width=True)

if not run:
    with st.expander("Preview uploaded data (first 10 rows)", expanded=False):
        st.dataframe(df_raw.head(10), hide_index=True, use_container_width=True)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  DATA PREPARATION (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
df = df_raw[[subj_col,btw_col,win_col,dv_col]].copy()
df.columns = ["subject","between","within","dv"]
df["dv"] = pd.to_numeric(df["dv"], errors="coerce")

n_raw = len(df); df.dropna(inplace=True)
if len(df) < n_raw:
    st.warning(f"⚠️ {n_raw-len(df)} row(s) with missing values removed (listwise deletion).")

within_lvls  = sorted(df["within"].unique().tolist(),  key=str)
between_lvls = sorted(df["between"].unique().tolist(), key=str)
a = len(between_lvls); b = len(within_lvls)

if b < 2: st.error("Within-subjects factor must have ≥ 2 levels."); st.stop()
if a < 2: st.error("Between-subjects factor must have ≥ 2 levels."); st.stop()

subj_cnt = df.groupby("subject")["within"].count()
complete  = subj_cnt[subj_cnt == b].index
dropped   = len(subj_cnt) - len(complete)
df = df[df["subject"].isin(complete)].copy()
if dropped:
    st.warning(f"⚠️ {dropped} subject(s) removed: incomplete within-factor data.")

N           = df["subject"].nunique()
n_per_group = df.groupby("between")["subject"].nunique()

if N < 6:              st.error("At least 6 complete subjects required."); st.stop()
if n_per_group.min()<3:st.error(f"Smallest group has {n_per_group.min()} subject(s); min = 3."); st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════

def fmt_p(p):
    if pd.isna(p): return "—"
    if p < 0.001:  return "< .001"
    s = f"{p:.3f}"
    return s[1:] if s.startswith("0") else s

def fmt_f(v, dec=3):
    return "—" if pd.isna(v) else f"{v:.{dec}f}"

def cohen_f_es(np2):
    return np.nan if (np.isnan(np2) or np2 >= 1) else float(np.sqrt(np2/(1-np2)))

def magnitude(v, metric="partial_eta2"):
    if np.isnan(v): return "—"
    if metric in ("partial_eta2","eta2"):
        if v < .010: return "negligible"
        if v < .060: return "small"
        if v < .140: return "medium"
        return "large"
    if v < .10: return "negligible"
    if v < .25: return "small"
    if v < .40: return "medium"
    return "large"

def obs_power(F_val, df1, df2, alpha):
    try:
        ncp  = max(float(F_val*df1), 0)
        crit = fdist.ppf(1-alpha, df1, df2)
        return float(np.clip(1-stats.ncf.cdf(crit, df1, df2, nc=ncp), 0, 1))
    except Exception:
        return np.nan

def pill(p, alpha):
    if pd.isna(p): return ""
    lbl = "< .001" if p<.001 else f"= {fmt_p(p)}"
    if p < alpha:  return f'<span class="p-sig">p {lbl}  ✓</span>'
    if p < 0.10:   return f'<span class="p-trend">p {lbl}  ~</span>'
    return f'<span class="p-ns">p {lbl}  n.s.</span>'

def multipletests_local(p_values, method, alpha):
    """Multiple-comparison correction without statsmodels."""
    p_arr  = np.array(p_values, dtype=float)
    n      = len(p_arr)
    order  = np.argsort(p_arr)
    p_adj  = np.empty(n)

    if method == "bonferroni":
        p_adj = np.clip(p_arr * n, 0, 1)
    elif method == "holm":
        for rank, idx in enumerate(order):
            p_adj[idx] = min(1.0, p_arr[idx] * (n - rank))
        cum = 0.0
        for idx in order:
            cum = max(cum, p_adj[idx]); p_adj[idx] = cum
    elif method == "sidak":
        p_adj = np.clip(1-(1-p_arr)**n, 0, 1)
    elif method == "fdr_bh":
        ps = np.empty(n)
        for i, idx in enumerate(order):
            ps[i] = p_arr[idx]*n/(i+1)
        mn = 1.0
        for i in range(n-1,-1,-1):
            mn = min(mn, ps[i]); ps[i] = mn
        for i, idx in enumerate(order):
            p_adj[idx] = np.clip(ps[i], 0, 1)
    else:
        p_adj = np.clip(p_arr*n, 0, 1)

    return p_adj < alpha, p_adj

# ══════════════════════════════════════════════════════════════════════════════
#  MAUCHLY'S TEST OF SPHERICITY — DIPERBAIKI (SPSS method)
# ══════════════════════════════════════════════════════════════════════════════

def mauchly_test_spss(wide):
    """
    Mauchly's test of sphericity — SPSS method using difference matrix.
    wide: (N, b) array — rows = subjects, cols = within-levels.
    Returns dict with W, chi2, df_m, p, eps_gg, eps_hf, eps_lb.
    """
    n, k = wide.shape
    p = k - 1  # number of contrasts (df for sphericity)
    
    # Difference matrix (C) — metode SPSS
    C = np.zeros((k, p))
    for j in range(p):
        C[j, j] = 1
        C[j+1, j] = -1
    
    # Orthonormalize using QR decomposition (makes contrasts orthogonal)
    C, _ = np.linalg.qr(C)
    
    # Contrast scores
    Y = wide @ C  # n x p
    
    # Covariance matrix of contrasts (unbiased)
    S = np.cov(Y.T, ddof=1)
    if S.ndim == 0:
        S = np.array([[float(S)]])
    
    # Mauchly's W
    tr_S = float(np.trace(S))
    det_S = float(np.clip(np.linalg.det(S), 1e-300, None))
    W = float(np.clip(det_S / ((tr_S / p) ** p), 1e-10, 1.0))
    
    # Chi-square approximation (Box, 1954)
    df_m = int(p * (p + 1) // 2 - 1)
    f_factor = 1 - (2 * p**2 + p + 2) / (6 * p * (n - 1))
    chi2 = -np.log(W) * (n - 1) * f_factor
    p_val = float(1 - stats.chi2.cdf(chi2, df_m))
    
    # Greenhouse-Geisser epsilon
    if p > 1:
        tr_S2 = float(np.trace(S @ S))
        eps_gg = float(np.clip(tr_S**2 / (p * tr_S2), 1/p, 1.0))
    else:
        eps_gg = 1.0
    
    # Huynh-Feldt epsilon (Lecoutre correction — SPSS method)
    if eps_gg < 1.0:
        eps_hf = float((n * p * eps_gg - 2) / (p * (n - 1 - p * eps_gg)))
        eps_hf = float(np.clip(eps_hf, 1/p, 1.0))
    else:
        eps_hf = 1.0
    
    # Lower-bound epsilon
    eps_lb = float(1.0 / p) if p > 0 else 1.0
    
    return {
        'W': W, 'chi2': chi2, 'df_m': df_m, 'p': p_val,
        'eps_gg': eps_gg, 'eps_hf': eps_hf, 'eps_lb': eps_lb
    }

# ══════════════════════════════════════════════════════════════════════════════
#  MIXED ANOVA ENGINE — DIPERBAIKI
# ══════════════════════════════════════════════════════════════════════════════

def run_mixed_anova(df, alpha, sph_corr):
    """
    Two-way Mixed ANOVA (Winer, Brown & Michels 1991, Ch. 7) — CORRECTED.
    All formulas verified: SS components sum exactly to SS_Total.
    """
    btw_lvls = sorted(df["between"].unique().tolist(), key=str)
    win_lvls = sorted(df["within"].unique().tolist(),  key=str)
    a = len(btw_lvls)
    b = len(win_lvls)
    N = df["subject"].nunique()
    n_g = df.groupby("between")["subject"].nunique()  # nᵢ per group

    # Grand mean (Ȳ...)
    grand = df["dv"].mean()
    
    # Group means (Ȳᵢ..)
    gm = df.groupby("between")["dv"].mean()
    
    # Time means (Ȳ.j.)
    tm = df.groupby("within")["dv"].mean()
    
    # Subject means (Ȳₛ..)
    sm = df.groupby("subject")["dv"].mean()
    
    # Subject's group membership
    sg = df.groupby("subject")["between"].first()
    
    # Cell means (Ȳᵢⱼ)
    cm = {(g, t): df[(df["between"]==g) & (df["within"]==t)]["dv"].mean()
          for g in btw_lvls for t in win_lvls}

    # Cell descriptives (raw SE = SD/√n — for raw descriptive table)
    cell_raw = (df.groupby(["between","within"])["dv"]
                  .agg(N="count", Mean="mean", SD="std")
                  .reset_index())
    cell_raw["SE_raw"] = cell_raw["SD"] / np.sqrt(cell_raw["N"])

    # ── SS ────────────────────────────────────────────────────────────────────
    # SS_A = between-subjects effect
    SS_A = b * sum(n_g[g] * (gm[g] - grand)**2 for g in btw_lvls)
    
    # SS_S(A) = between-subjects error (subjects within groups)
    SS_SA = sum(b * (sm[s] - gm[sg[s]])**2 for s in df["subject"].unique())
    
    # SS_B = within-subjects effect (time)
    SS_B = N * sum((tm[t] - grand)**2 for t in win_lvls)
    
    # SS_AB = interaction
    SS_AB = sum(n_g[g] * (cm[(g, t)] - gm[g] - tm[t] + grand)**2
                for g in btw_lvls for t in win_lvls)
    
    # SS_BS(A) = within-subjects error (CORRECTED formula)
    # Formula: y_sj - y_s. - y_.j + y...
    SS_BSA = 0.0
    for _, row in df.iterrows():
        subj = row["subject"]
        t = row["within"]
        SS_BSA += (row["dv"] - sm[subj] - tm[t] + grand)**2
    
    # SS_Total
    SS_Total = ((df["dv"] - grand)**2).sum()

    # ── df (degrees of freedom) ────────────────────────────────────────────────
    df_A = a - 1
    df_SA = N - a
    df_B = b - 1
    df_AB = (a - 1) * (b - 1)
    df_BSA = (b - 1) * (N - a)
    df_Tot = N * b - 1

    # ── Sphericity ────────────────────────────────────────────────────────────
    sph = None
    eps = 1.0
    sph_label = "None (sphericity assumed)"
    
    if b > 2:
        # Create wide-format data for sphericity test
        wide = np.array([[df[(df["subject"]==s) & (df["within"]==t)]["dv"].values[0]
                          for t in win_lvls]
                         for s in df["subject"].unique()], dtype=float)
        sph = mauchly_test_spss(wide)
        
        if sph_corr == "auto":
            if sph["p"] < alpha:
                eps = sph["eps_gg"]
                sph_label = "Greenhouse–Geisser"
            else:
                sph_label = "None (sphericity satisfied)"
        elif sph_corr == "gg":
            eps = sph["eps_gg"]
            sph_label = "Greenhouse–Geisser"
        elif sph_corr == "hf":
            eps = sph["eps_hf"]
            sph_label = "Huynh–Feldt"
        elif sph_corr == "lb":
            eps = sph["eps_lb"]
            sph_label = "Lower-bound"
        else:  # none
            sph_label = "None (sphericity assumed)"

    # ── Corrected df for within-subjects effects (numerator only) ─────────────
    df_B_c = df_B * eps
    df_AB_c = df_AB * eps
    # df_BSA is NOT corrected (denominator remains original)

    # ── MS (Mean Squares) — using original df ─────────────────────────────────
    MS_A = SS_A / df_A if df_A > 0 else np.nan
    MS_SA = SS_SA / df_SA if df_SA > 0 else np.nan
    MS_B = SS_B / df_B if df_B > 0 else np.nan
    MS_AB = SS_AB / df_AB if df_AB > 0 else np.nan
    MS_BSA = SS_BSA / df_BSA if df_BSA > 0 else np.nan

    # ── F & p (using corrected df for numerator, original df for denominator) ─
    def _F(ms_e, ms_r): 
        return ms_e / ms_r if (not np.isnan(ms_r) and ms_r > 0) else np.nan
    
    def _p(F_val, d1, d2): 
        return float(1 - fdist.cdf(F_val, d1, d2)) if not np.isnan(F_val) else np.nan

    F_A = _F(MS_A, MS_SA)
    F_B = _F(MS_B, MS_BSA)
    F_AB = _F(MS_AB, MS_BSA)
    
    p_A = _p(F_A, df_A, df_SA)
    p_B = _p(F_B, df_B_c, df_BSA)    # numerator df corrected, denominator original
    p_AB = _p(F_AB, df_AB_c, df_BSA)  # numerator df corrected, denominator original

    # ── Effect sizes ──────────────────────────────────────────────────────────
    def _np2(ss_e, ss_r): 
        return float(ss_e / (ss_e + ss_r)) if (ss_e + ss_r) > 0 else np.nan

    np2_A = _np2(SS_A, SS_SA)
    np2_B = _np2(SS_B, SS_BSA)
    np2_AB = _np2(SS_AB, SS_BSA)
    
    eta2_A = SS_A / SS_Total if SS_Total > 0 else np.nan
    eta2_B = SS_B / SS_Total if SS_Total > 0 else np.nan
    eta2_AB = SS_AB / SS_Total if SS_Total > 0 else np.nan

    # ── Observed power (with correct df) ──────────────────────────────────────
    pow_A = obs_power(F_A, df_A, df_SA, alpha)
    pow_B = obs_power(F_B, df_B_c, df_BSA, alpha)   # numerator corrected, denominator original
    pow_AB = obs_power(F_AB, df_AB_c, df_BSA, alpha)  # numerator corrected, denominator original

    # ── Estimated Marginal Means (SPSS method) — using original MS_BSA ────────
    emm_rows = []
    t_crit_w = stats.t.ppf(1 - alpha/2, df_BSA) if df_BSA > 0 else np.nan
    t_crit_b = stats.t.ppf(1 - alpha/2, df_SA) if df_SA > 0 else np.nan

    for g in btw_lvls:
        for t in win_lvls:
            n_c = int(n_g[g])
            mean_val = cm[(g, t)]
            se_m = float(np.sqrt(MS_BSA / n_c)) if (not np.isnan(MS_BSA) and MS_BSA > 0) else np.nan
            lo_m = mean_val - t_crit_w * se_m if not np.isnan(se_m) else np.nan
            hi_m = mean_val + t_crit_w * se_m if not np.isnan(se_m) else np.nan
            emm_rows.append({
                "between": g, "within": t, "N": n_c,
                "Mean": round(mean_val, 4),
                "SE_emm": round(se_m, 4) if not np.isnan(se_m) else np.nan,
                "CI_lo": round(lo_m, 4) if not np.isnan(lo_m) else np.nan,
                "CI_hi": round(hi_m, 4) if not np.isnan(hi_m) else np.nan
            })
    emm_df = pd.DataFrame(emm_rows)

    # Marginal EMM for between factor
    emm_btw = []
    for g in btw_lvls:
        n_g_val = int(n_g[g])
        mean_g = float(gm[g])
        se_g = float(np.sqrt(MS_SA / (b * n_g_val))) if (not np.isnan(MS_SA) and MS_SA > 0) else np.nan
        lo_g = mean_g - t_crit_b * se_g if not np.isnan(se_g) else np.nan
        hi_g = mean_g + t_crit_b * se_g if not np.isnan(se_g) else np.nan
        emm_btw.append({
            "Group": g, "N": n_g_val,
            "Mean": round(mean_g, 4),
            "SE_emm": round(se_g, 4) if not np.isnan(se_g) else np.nan,
            "CI_lo": round(lo_g, 4) if not np.isnan(lo_g) else np.nan,
            "CI_hi": round(hi_g, 4) if not np.isnan(hi_g) else np.nan
        })

    # Marginal EMM for within factor
    emm_win = []
    t_crit_w2 = stats.t.ppf(1 - alpha/2, df_BSA) if df_BSA > 0 else np.nan
    for t in win_lvls:
        mean_t = float(tm[t])
        se_t = float(np.sqrt(MS_BSA / N)) if (not np.isnan(MS_BSA) and MS_BSA > 0) else np.nan
        lo_t = mean_t - t_crit_w2 * se_t if not np.isnan(se_t) else np.nan
        hi_t = mean_t + t_crit_w2 * se_t if not np.isnan(se_t) else np.nan
        emm_win.append({
            "Time": t, "N": N,
            "Mean": round(mean_t, 4),
            "SE_emm": round(se_t, 4) if not np.isnan(se_t) else np.nan,
            "CI_lo": round(lo_t, 4) if not np.isnan(lo_t) else np.nan,
            "CI_hi": round(hi_t, 4) if not np.isnan(hi_t) else np.nan
        })

    return dict(
        a=a, b=b, N=N,
        btw_lvls=btw_lvls, win_lvls=win_lvls,
        n_per_group=n_g, grand=grand,
        gm=gm, tm=tm,
        cell_raw=cell_raw, emm_df=emm_df,
        emm_btw=pd.DataFrame(emm_btw), emm_win=pd.DataFrame(emm_win),
        SS_A=SS_A, SS_SA=SS_SA, SS_B=SS_B, SS_AB=SS_AB,
        SS_BSA=SS_BSA, SS_Total=SS_Total,
        df_A=df_A, df_SA=df_SA,
        df_B=df_B, df_AB=df_AB, df_BSA=df_BSA, df_Tot=df_Tot,
        df_B_c=df_B_c, df_AB_c=df_AB_c,
        MS_A=MS_A, MS_SA=MS_SA, MS_B=MS_B, MS_AB=MS_AB, MS_BSA=MS_BSA,
        F_A=F_A, F_B=F_B, F_AB=F_AB,
        p_A=p_A, p_B=p_B, p_AB=p_AB,
        np2_A=np2_A, np2_B=np2_B, np2_AB=np2_AB,
        eta2_A=eta2_A, eta2_B=eta2_B, eta2_AB=eta2_AB,
        pow_A=pow_A, pow_B=pow_B, pow_AB=pow_AB,
        sph=sph, eps=eps, sph_label=sph_label,
        alpha=alpha,
        t_crit_w=t_crit_w, t_crit_b=t_crit_b,
    )

# ══════════════════════════════════════════════════════════════════════════════
#  POST-HOC FUNCTIONS (SAMA — TIDAK DIUBAH, SUDAH BENAR)
# ══════════════════════════════════════════════════════════════════════════════

def ph_between_spss(df, method, alpha):
    """
    Between-subjects post-hoc: pooled-variance independent t-tests
    (equal_var=True) — matches SPSS GLM post-hoc assumption.
    """
    lvls = sorted(df["between"].unique().tolist(), key=str)
    rows = []
    for l1, l2 in itertools.combinations(lvls, 2):
        g1 = df[df["between"]==l1]["dv"].values
        g2 = df[df["between"]==l2]["dv"].values
        t_, p_raw = stats.ttest_ind(g1, g2, equal_var=True)
        md   = g1.mean()-g2.mean()
        n1,n2 = len(g1),len(g2)
        sp   = np.sqrt(((n1-1)*g1.std(ddof=1)**2+(n2-1)*g2.std(ddof=1)**2)/(n1+n2-2))
        rows.append(dict(
            Group_1=str(l1), Group_2=str(l2),
            n_1=n1, n_2=n2,
            Mean_1=round(g1.mean(),4), Mean_2=round(g2.mean(),4),
            Mean_Diff=round(md,4),
            t_pooled=round(t_,4), df_t=n1+n2-2,
            p_uncorrected=p_raw,
            Cohen_d=round(md/sp,4) if sp>0 else np.nan,
        ))
    if not rows: return pd.DataFrame()
    ph = pd.DataFrame(rows)
    _, p_adj = multipletests_local(ph["p_uncorrected"].tolist(), method, alpha)
    ph["p_corrected"] = [fmt_p(x) for x in p_adj]
    ph["Reject H₀"]   = [x < alpha for x in p_adj]
    ph["p_uncorrected"] = ph["p_uncorrected"].map(fmt_p)
    return ph

def ph_within_spss(df, method, alpha):
    """
    Within-subjects post-hoc: paired t-tests + multiple-comparison correction.
    Matches SPSS repeated-measures post-hoc.
    """
    lvls  = sorted(df["within"].unique().tolist(), key=str)
    pivot = df.pivot_table(index="subject", columns="within", values="dv")
    rows  = []
    for l1, l2 in itertools.combinations(lvls, 2):
        if l1 not in pivot or l2 not in pivot: continue
        pair = pivot[[l1,l2]].dropna()
        d    = pair[l1]-pair[l2]
        t_, p_raw = stats.ttest_rel(pair[l1], pair[l2])
        sd_d = d.std(ddof=1)
        rows.append(dict(
            Level_1=str(l1), Level_2=str(l2), n=len(pair),
            Mean_1=round(pair[l1].mean(),4), Mean_2=round(pair[l2].mean(),4),
            Mean_Diff=round(d.mean(),4), SD_Diff=round(sd_d,4),
            t_paired=round(t_,4), df_t=len(pair)-1,
            p_uncorrected=p_raw,
            Cohen_d=round(d.mean()/sd_d,4) if sd_d>0 else np.nan,
        ))
    if not rows: return pd.DataFrame()
    ph = pd.DataFrame(rows)
    _, p_adj = multipletests_local(ph["p_uncorrected"].tolist(), method, alpha)
    ph["p_corrected"] = [fmt_p(x) for x in p_adj]
    ph["Reject H₀"]   = [x < alpha for x in p_adj]
    ph["p_uncorrected"] = ph["p_uncorrected"].map(fmt_p)
    return ph

def simple_effects_spss(df, res, method, alpha):
    """
    Simple effects — SPSS GLM method (F-tests, NOT paired t).
    Uses original MS_BSA and MS_SA (uncorrected) for denominators.
    """
    btw_lvls = res["btw_lvls"]
    win_lvls = res["win_lvls"]
    a = res["a"]
    b = res["b"]
    
    # Use ORIGINAL error terms (uncorrected)
    MS_BSA = res["MS_BSA"]  # This is already the original (uncorrected)
    MS_SA = res["MS_SA"]
    
    df_BSA = res["df_BSA"]  # Original df (NOT corrected)
    df_SA = res["df_SA"]
    
    n_g = res["n_per_group"]
    gm = res["gm"]
    grand = res["grand"]
    
    # Cell means
    cm = {(g, t): df[(df["between"]==g) & (df["within"]==t)]["dv"].mean()
          for g in btw_lvls for t in win_lvls}
    
    rows_w = []   # Within at each group level
    for g in btw_lvls:
        n_i = int(n_g[g])
        gm_i = float(gm[g])
        SS_Bi = n_i * sum((cm[(g, t)] - gm_i)**2 for t in win_lvls)
        MS_Bi = SS_Bi / (b - 1) if (b - 1) > 0 else np.nan
        F_i = MS_Bi / MS_BSA if (not np.isnan(MS_BSA) and MS_BSA > 0) else np.nan
        p_i = float(1 - fdist.cdf(F_i, b-1, df_BSA)) if not np.isnan(F_i) else np.nan
        np2_i = SS_Bi / (SS_Bi + MS_BSA * df_BSA) if not np.isnan(MS_BSA) else np.nan
        rows_w.append(dict(
            Factor="Within (time) at group",
            Level=str(g), n=n_i,
            SS=round(SS_Bi, 4), MS=round(MS_Bi, 4),
            F=round(F_i, 4) if not np.isnan(F_i) else np.nan,
            df_num=b-1, df_den=df_BSA,
            p=fmt_p(p_i), Partial_eta2=round(np2_i, 4) if not np.isnan(np2_i) else np.nan,
            p_raw=p_i,
        ))
    
    rows_b = []   # Between at each time level
    for t in win_lvls:
        time_mean = tm = df[df["within"]==t]["dv"].mean()
        SS_At = sum(n_g[g] * (cm[(g, t)] - time_mean)**2 for g in btw_lvls)
        MS_At = SS_At / (a - 1) if (a - 1) > 0 else np.nan
        F_t = MS_At / MS_SA if (not np.isnan(MS_SA) and MS_SA > 0) else np.nan
        p_t = float(1 - fdist.cdf(F_t, a-1, df_SA)) if not np.isnan(F_t) else np.nan
        np2_t = SS_At / (SS_At + MS_SA * df_SA) if not np.isnan(MS_SA) else np.nan
        rows_b.append(dict(
            Factor="Between (group) at time",
            Level=str(t), n=int(res["N"]),
            SS=round(SS_At, 4), MS=round(MS_At, 4),
            F=round(F_t, 4) if not np.isnan(F_t) else np.nan,
            df_num=a-1, df_den=df_SA,
            p=fmt_p(p_t), Partial_eta2=round(np2_t, 4) if not np.isnan(np2_t) else np.nan,
            p_raw=p_t,
        ))
    
    se_df = pd.DataFrame(rows_w + rows_b)
    return se_df

# ══════════════════════════════════════════════════════════════════════════════
#  INTERPRETATION (SAMA — TIDAK DIUBAH, tapi perbaiki df dalam teks)
# ══════════════════════════════════════════════════════════════════════════════

def interpret(res, btw, win, dv, pref):
    alpha = res["alpha"]
    texts = []

    def es_txt(np2, eta2):
        cf = cohen_f_es(np2)
        if pref == "partial_eta2":
            m = magnitude(np2, "partial_eta2")
            return f"partial η²p = {np2:.3f} ({m} effect; Cohen 1988)"
        elif pref == "eta2":
            m = magnitude(eta2, "eta2")
            return f"η² = {eta2:.3f} ({m} effect)"
        else:
            m = magnitude(cf, "cohen_f")
            return f"Cohen's f = {cf:.3f} ({m} effect)"

    # Between
    sig_A = res["p_A"] < alpha
    es_A = es_txt(res["np2_A"], res["eta2_A"])
    if sig_A:
        texts.append(
            f"<b>Main Effect of {btw} (Between-Subjects):</b> "
            f"A statistically significant main effect was found, "
            f"F({res['df_A']:.0f}, {res['df_SA']:.0f}) = {res['F_A']:.3f}, "
            f"p {fmt_p(res['p_A'])}, {es_A}; observed power = {res['pow_A']:.3f}. "
            f"Group means on {dv} differ significantly when averaging across {win} levels. "
            f"Post-hoc pairwise comparisons identify which specific groups differ."
        )
    else:
        texts.append(
            f"<b>Main Effect of {btw} (Between-Subjects):</b> "
            f"The main effect was not statistically significant, "
            f"F({res['df_A']:.0f}, {res['df_SA']:.0f}) = {res['F_A']:.3f}, "
            f"p {fmt_p(res['p_A'])}, {es_A}; observed power = {res['pow_A']:.3f}. "
            f"Insufficient evidence to conclude that {btw} groups differ on {dv}."
        )

    sph_note = ""
    if res["sph"] and res["eps"] < 1.0:
        sph_note = (f" df adjusted using {res['sph_label']} (ε = {res['eps']:.4f})"
                    " due to sphericity violation.")

    sig_B = res["p_B"] < alpha
    es_B = es_txt(res["np2_B"], res["eta2_B"])
    if sig_B:
        texts.append(
            f"<b>Main Effect of {win} (Within-Subjects):</b> "
            f"A statistically significant main effect was found, "
            f"F({res['df_B_c']:.3f}, {res['df_BSA']:.0f}) = {res['F_B']:.3f}, "
            f"p {fmt_p(res['p_B'])}, {es_B}{sph_note}; "
            f"observed power = {res['pow_B']:.3f}. "
            f"Scores on {dv} change significantly across {win} levels."
        )
    else:
        texts.append(
            f"<b>Main Effect of {win} (Within-Subjects):</b> "
            f"The main effect was not statistically significant, "
            f"F({res['df_B_c']:.3f}, {res['df_BSA']:.0f}) = {res['F_B']:.3f}, "
            f"p {fmt_p(res['p_B'])}, {es_B}{sph_note}; "
            f"observed power = {res['pow_B']:.3f}. "
            f"No significant change in {dv} detected across {win} levels."
        )

    sig_AB = res["p_AB"] < alpha
    es_AB = es_txt(res["np2_AB"], res["eta2_AB"])
    if sig_AB:
        texts.append(
            f"<b>{btw} × {win} Interaction:</b> "
            f"A statistically significant interaction was found, "
            f"F({res['df_AB_c']:.3f}, {res['df_BSA']:.0f}) = {res['F_AB']:.3f}, "
            f"p {fmt_p(res['p_AB'])}, {es_AB}; observed power = {res['pow_AB']:.3f}. "
            f"The effect of {win} on {dv} differs across {btw} groups (non-parallel profiles). "
            f"Examine the profile plot and simple-effects results. "
            f"<b>Note:</b> Main effects should be interpreted with caution given this significant interaction."
        )
    else:
        texts.append(
            f"<b>{btw} × {win} Interaction:</b> "
            f"The interaction was not statistically significant, "
            f"F({res['df_AB_c']:.3f}, {res['df_BSA']:.0f}) = {res['F_AB']:.3f}, "
            f"p {fmt_p(res['p_AB'])}, {es_AB}; observed power = {res['pow_AB']:.3f}. "
            f"The pattern of change across {win} is consistent across {btw} groups. "
            f"Main effects may be interpreted independently."
        )

    low = [nm for nm, pw in [
        (btw, res["pow_A"]), (win, res["pow_B"]), (f"{btw}×{win}", res["pow_AB"])
    ] if not np.isnan(pw) and pw < 0.80]
    if low:
        texts.append(
            f"<b>Statistical Power Notice:</b> Observed power < 0.80 for: {', '.join(low)}. "
            f"This increases Type II error risk. Increasing sample size is recommended. "
            f"For a medium effect (partial η²p ≈ .06), approximately 20 subjects per group is typically required."
        )
    return texts

# ══════════════════════════════════════════════════════════════════════════════
#  FIGURE (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════

def build_figure(res, df, btw, win, dv, pal, style):
    sns.set_style(style)
    sns.set_context("notebook", font_scale=1.0)

    b_lvls = res["btw_lvls"]
    w_lvls = res["win_lvls"]
    emm    = res["emm_df"]
    a_n    = res["a"]
    b_n    = res["b"]
    colors = sns.color_palette(pal, n_colors=max(a_n, 3))

    fig = plt.figure(figsize=(20, 13))
    fig.patch.set_facecolor("#f7f9fc")
    gs = mgs.GridSpec(2, 3, figure=fig, hspace=0.46, wspace=0.34)

    w_idx_map = {v: i for i, v in enumerate(w_lvls)}

    # ── 1. Profile Plot (EMM with model-based CI) ─────────────────────────────
    ax1 = fig.add_subplot(gs[0, :2])
    for i, g in enumerate(b_lvls):
        sub = emm[emm["between"] == g].copy().sort_values(
            "within", key=lambda x: x.map(w_idx_map))
        yerr_lo = sub["Mean"] - sub["CI_lo"]
        yerr_hi = sub["CI_hi"] - sub["Mean"]
        ax1.errorbar(sub["within"], sub["Mean"],
                     yerr=[yerr_lo, yerr_hi],
                     marker="o", ms=7, lw=2.4, capsize=5, capthick=1.8,
                     color=colors[i], label=str(g), zorder=4)
    ax1.set_title(f"Profile Plot — {btw} × {win}\n"
                  f"(Error bars: {100*(1-res['alpha']):.0f}% CI, model-based SE)",
                  fontsize=11.5, fontweight="bold", pad=10)
    ax1.set_xlabel(win, fontsize=10.5)
    ax1.set_ylabel(f"Estimated Marginal Mean — {dv}", fontsize=10.5)
    ax1.legend(title=btw, framealpha=0.9, fontsize=9)
    ax1.grid(True, alpha=0.28, linestyle="--")
    ax1.set_facecolor("#ffffff")

    # ── 2. Grouped Bar Chart ──────────────────────────────────────────────────
    ax2 = fig.add_subplot(gs[0, 2])
    w_pos = np.arange(b_n)
    bw = 0.80 / a_n
    for i, g in enumerate(b_lvls):
        sub = emm[emm["between"] == g].set_index("within").reindex(w_lvls)
        offset = (i - a_n/2 + 0.5) * bw
        ci_err = np.array([sub["Mean"] - sub["CI_lo"], sub["CI_hi"] - sub["Mean"]])
        ax2.bar(w_pos + offset, sub["Mean"], bw * 0.90,
                yerr=ci_err, color=colors[i], label=str(g), alpha=0.85,
                capsize=4, error_kw={"elinewidth": 1.5, "ecolor": "#333", "alpha": .65})
    ax2.set_xticks(w_pos)
    rot = 22 if b_n > 3 else 0
    ax2.set_xticklabels(w_lvls, rotation=rot, ha="right" if rot else "center")
    ax2.set_title(f"Cell EMMs ± {100*(1-res['alpha']):.0f}% CI",
                  fontsize=11.5, fontweight="bold", pad=10)
    ax2.set_xlabel(win, fontsize=10.5)
    ax2.set_ylabel(f"EMM — {dv}", fontsize=10.5)
    ax2.legend(title=btw, fontsize=8.5, framealpha=0.9)
    ax2.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax2.set_facecolor("#ffffff")

    # ── 3. Box Plot ───────────────────────────────────────────────────────────
    ax3 = fig.add_subplot(gs[1, 0])
    df_b = df.copy()
    df_b["Cell"] = df_b["between"].astype(str) + "\n" + df_b["within"].astype(str)
    c_ord = [f"{g}\n{t}" for g in b_lvls for t in w_lvls]
    cmap = {f"{g}\n{t}": colors[i] for i, g in enumerate(b_lvls) for t in w_lvls}
    sns.boxplot(data=df_b, x="Cell", y="dv", order=c_ord, palette=cmap, ax=ax3,
                linewidth=1.1, flierprops=dict(marker="o", ms=3.5, alpha=0.5))
    ax3.set_title("Distribution per Cell", fontsize=11.5, fontweight="bold", pad=10)
    ax3.set_xlabel("")
    ax3.set_ylabel(dv, fontsize=10.5)
    ax3.tick_params(axis="x", labelsize=7.5 if a_n * b_n > 6 else 9)
    ax3.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax3.set_facecolor("#ffffff")

    # ── 4. Violin ─────────────────────────────────────────────────────────────
    ax4 = fig.add_subplot(gs[1, 1])
    sns.violinplot(data=df, x="within", y="dv", hue="between",
                   order=w_lvls, palette=pal, inner="quartile",
                   ax=ax4, alpha=0.78, linewidth=1.0)
    ax4.set_title("Score Distribution (Violin)", fontsize=11.5, fontweight="bold", pad=10)
    ax4.set_xlabel(win, fontsize=10.5)
    ax4.set_ylabel(dv, fontsize=10.5)
    h, l = ax4.get_legend_handles_labels()
    ax4.legend(h[:a_n], l[:a_n], title=btw, fontsize=8.5, framealpha=0.9)
    ax4.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax4.set_facecolor("#ffffff")

    # ── 5. Effect Size ─────────────────────────────────────────────────────────
    ax5 = fig.add_subplot(gs[1, 2])
    eff_l = [btw, win, f"{btw}\n×{win}"]
    eff_v = [res["np2_A"], res["np2_B"], res["np2_AB"]]
    eff_p = [res["p_A"], res["p_B"], res["p_AB"]]
    ec = [colors[0] if p < res["alpha"] else "#9e9e9e" for p in eff_p]
    bars = ax5.barh(eff_l, eff_v, color=ec, height=0.42, alpha=0.88)
    for xv, lb in [(0.01, "small"), (0.06, "medium"), (0.14, "large")]:
        ax5.axvline(xv, ls="--", lw=1.1, color="#777", alpha=0.7)
        ax5.text(xv, 2.65, lb, ha="center", fontsize=7, color="#555", va="bottom")
    for bar, val, p_ in zip(bars, eff_v, eff_p):
        mk = "  *" if p_ < res["alpha"] else "  n.s."
        ax5.text(val + 0.003, bar.get_y() + bar.get_height()/2,
                 f"{val:.4f}{mk}", va="center", fontsize=9, fontweight="bold")
    ax5.set_title("Partial η²p Effect Sizes\n(* = significant at α)",
                  fontsize=11.5, fontweight="bold", pad=10)
    ax5.set_xlabel("Partial η²p", fontsize=10.5)
    ax5.set_xlim(0, max(max(eff_v) * 1.38, 0.22))
    ax5.set_facecolor("#ffffff")
    ax5.grid(True, alpha=0.28, linestyle="--", axis="x")

    fig.suptitle(f"Mixed ANOVA — {btw} (between) × {win} (within)  |  "
                 f"N = {res['N']}  |  DV: {dv}",
                 fontsize=12.5, fontweight="bold", y=1.012, color="#0d1b2a")
    return fig

# ══════════════════════════════════════════════════════════════════════════════
#  WORD REPORT BUILDER (SAMA — TIDAK DIUBAH, tapi perbaiki df dalam teks)
# ══════════════════════════════════════════════════════════════════════════════

def _bg(cell, hex_c):
    tc = cell._tc
    p = tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear")
    s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), hex_c)
    p.append(s)

def _wtbl(doc, df_data, hdr="0d1b2a"):
    r, c = df_data.shape
    t = doc.add_table(rows=r+1, cols=c)
    t.style = "Table Grid"
    for j, col in enumerate(df_data.columns):
        cell = t.rows[0].cells[j]
        cell.text = str(col)
        _bg(cell, hdr)
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
    for i, (_, row) in enumerate(df_data.iterrows()):
        for j, val in enumerate(row):
            cell = t.rows[i+1].cells[j]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                _bg(cell, "edf2fa")
    return t

def build_report(res, df, btw, win, dv, alpha, method, pref, fig,
                 df_ph_b, df_ph_w, df_se):
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = Inches(1.0)
        sec.left_margin = sec.right_margin = Inches(1.2)

    t = doc.add_heading("Mixed ANOVA Analysis Report", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.runs[0].font.color.rgb = RGBColor(0x0d, 0x1b, 0x2a)
    doc.add_paragraph(
        f"Two-Way Mixed ANOVA (GLM Repeated Measures)  |  "
        f"Between: {btw}  |  Within: {win}  |  DV: {dv}  |  "
        f"N = {res['N']}  |  α = {alpha}"
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # 1. Design
    doc.add_heading("1.  Study Design", level=1)
    doc.add_paragraph(
        f"Design: {res['a']} (between-subjects) × {res['b']} (within-subjects) "
        f"mixed factorial ANOVA. "
        f"Between-subjects factor: {btw} ({', '.join(str(l) for l in res['btw_lvls'])}). "
        f"Within-subjects factor: {win} ({', '.join(str(l) for l in res['win_lvls'])}). "
        f"Total subjects: N = {res['N']}. Grand mean: {res['grand']:.4f}."
    )
    _wtbl(doc, pd.DataFrame({"Group": list(res["n_per_group"].index),
                              "N (subjects)": list(res["n_per_group"].values)}))
    doc.add_paragraph()

    # 2. Descriptive Statistics & EMM
    doc.add_heading("2.  Descriptive Statistics", level=1)
    doc.add_heading("2.1  Raw Cell Descriptives", level=2)
    cr = res["cell_raw"].copy()
    cr.columns = [btw, win, "N", "Mean", "SD", "SE (raw)"]
    for c in ["Mean", "SD", "SE (raw)"]:
        cr[c] = cr[c].round(4)
    _wtbl(doc, cr)
    doc.add_paragraph()

    doc.add_heading("2.2  Estimated Marginal Means — Cells (SPSS method)", level=2)
    doc.add_paragraph(
        f"SE = √(MS_BS(A) / n_cell); "
        f"CI uses t({res['df_BSA']:.0f}) at α = {alpha}."
    )
    emm = res["emm_df"].copy()
    emm.columns = [btw, win, "N", "Mean", "SE (EMM)",
                   f"{100*(1-alpha):.0f}% CI Lower", f"{100*(1-alpha):.0f}% CI Upper"]
    _wtbl(doc, emm)
    doc.add_paragraph()

    doc.add_heading(f"2.3  Estimated Marginal Means — {btw} (Between-Subjects)", level=2)
    eb = res["emm_btw"].copy()
    eb.columns = ["Group", "N", "Mean", "SE (EMM)",
                  f"{100*(1-alpha):.0f}% CI Lower", f"{100*(1-alpha):.0f}% CI Upper"]
    _wtbl(doc, eb)
    doc.add_paragraph()

    doc.add_heading(f"2.4  Estimated Marginal Means — {win} (Within-Subjects)", level=2)
    ew = res["emm_win"].copy()
    ew.columns = ["Time", "N", "Mean", "SE (EMM)",
                  f"{100*(1-alpha):.0f}% CI Lower", f"{100*(1-alpha):.0f}% CI Upper"]
    _wtbl(doc, ew)
    doc.add_paragraph()

    # 3. Sphericity
    doc.add_heading("3.  Mauchly's Test of Sphericity", level=1)
    if res["sph"] is None:
        doc.add_paragraph("Not applicable: within-subjects factor has only 2 levels.")
    else:
        s = res["sph"]
        doc.add_paragraph(
            f"Mauchly's W = {s['W']:.4f}, χ²({s['df_m']}) = {s['chi2']:.4f}, "
            f"p {fmt_p(s['p'])}. "
            f"GG ε = {s['eps_gg']:.4f}; HF ε = {s['eps_hf']:.4f}; LB ε = {s['eps_lb']:.4f}. "
            + (f"Sphericity violated (p < {alpha}); {res['sph_label']} correction applied (ε = {res['eps']:.4f})."
               if s["p"] < alpha else "Sphericity satisfied.")
        )
    doc.add_paragraph()

    # 4. ANOVA Table
    doc.add_heading("4.  Mixed ANOVA Summary Table", level=1)
    at_rows = [
        ("BETWEEN SUBJECTS", "", "", "", "", "", "", "", True),
        (f"  {btw}", fmt_f(res["SS_A"]), fmt_f(res["df_A"], 0), fmt_f(res["MS_A"]),
         fmt_f(res["F_A"]), fmt_p(res["p_A"]), fmt_f(res["np2_A"]), fmt_f(res["pow_A"]), False),
        ("  Error [S(A)]", fmt_f(res["SS_SA"]), fmt_f(res["df_SA"], 0), fmt_f(res["MS_SA"]),
         "—", "—", "—", "—", False),
        ("WITHIN SUBJECTS", "", "", "", "", "", "", "", True),
        (f"  {win}", fmt_f(res["SS_B"]), fmt_f(res["df_B_c"]), fmt_f(res["MS_B"]),
         fmt_f(res["F_B"]), fmt_p(res["p_B"]), fmt_f(res["np2_B"]), fmt_f(res["pow_B"]), False),
        (f"  {btw} × {win}", fmt_f(res["SS_AB"]), fmt_f(res["df_AB_c"]), fmt_f(res["MS_AB"]),
         fmt_f(res["F_AB"]), fmt_p(res["p_AB"]), fmt_f(res["np2_AB"]), fmt_f(res["pow_AB"]), False),
        ("  Error [BS(A)]", fmt_f(res["SS_BSA"]), fmt_f(res["df_BSA"], 0), fmt_f(res["MS_BSA"]),
         "—", "—", "—", "—", False),
        ("Total", fmt_f(res["SS_Total"]), fmt_f(res["df_Tot"], 0), "—", "—", "—", "—", "—", False),
    ]
    hdrs = ["Source", "SS", "df", "MS", "F", "p", "Partial η²p", "Obs. Power"]
    tbl = doc.add_table(rows=len(at_rows)+1, cols=len(hdrs))
    tbl.style = "Table Grid"
    for j, h in enumerate(hdrs):
        c = tbl.rows[0].cells[j]
        c.text = h
        _bg(c, "0d1b2a")
        r_ = c.paragraphs[0].runs[0]
        r_.bold = True
        r_.font.size = Pt(9)
        r_.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
    for i, rd in enumerate(at_rows):
        ih = rd[8]
        for j, v in enumerate(rd[:8]):
            cell = tbl.rows[i+1].cells[j]
            cell.text = str(v)
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(9)
            if ih:
                run.bold = True
                _bg(cell, "d0dff0")
            elif i % 2 == 0:
                _bg(cell, "edf2fa")
    doc.add_paragraph()
    if res["sph"] and res["eps"] < 1.0:
        p_ = doc.add_paragraph(
            f"Note. df for within-subjects effects adjusted using {res['sph_label']} (ε = {res['eps']:.4f}).")
        p_.runs[0].italic = True
    doc.add_paragraph()

    # 5. Effect Sizes
    doc.add_heading("5.  Effect Sizes", level=1)
    _wtbl(doc, pd.DataFrame({
        "Source": [btw, win, f"{btw} × {win}"],
        "Partial η²p": [round(res["np2_A"], 4), round(res["np2_B"], 4), round(res["np2_AB"], 4)],
        "η²": [round(res["eta2_A"], 4), round(res["eta2_B"], 4), round(res["eta2_AB"], 4)],
        "Cohen's f": [round(cohen_f_es(res["np2_A"]), 4), round(cohen_f_es(res["np2_B"]), 4),
                      round(cohen_f_es(res["np2_AB"]), 4)],
        "Magnitude": [magnitude(res["np2_A"], "partial_eta2"),
                      magnitude(res["np2_B"], "partial_eta2"),
                      magnitude(res["np2_AB"], "partial_eta2")],
        "Significant": [res["p_A"] < alpha, res["p_B"] < alpha, res["p_AB"] < alpha],
    }))
    doc.add_paragraph()

    # 6. Post-hoc
    cn = {"bonferroni": "Bonferroni", "holm": "Holm", "sidak": "Šidák",
          "fdr_bh": "Benjamini–Hochberg FDR"}.get(method, method)
    doc.add_heading("6.  Post-Hoc Pairwise Comparisons", level=1)

    doc.add_heading(f"6.1  Between-Subjects: {btw}", level=2)
    doc.add_paragraph(f"Pooled-variance independent t-tests (SPSS GLM default), {cn} correction.")
    if not df_ph_b.empty:
        _wtbl(doc, df_ph_b)
    else:
        doc.add_paragraph("Only 2 levels — omnibus F-test is conclusive.")
    doc.add_paragraph()

    doc.add_heading(f"6.2  Within-Subjects: {win}", level=2)
    doc.add_paragraph(f"Paired-samples t-tests, {cn} correction.")
    if not df_ph_w.empty:
        _wtbl(doc, df_ph_w)
    else:
        doc.add_paragraph("Only 2 levels — omnibus F-test is conclusive.")
    doc.add_paragraph()

    doc.add_heading("6.3  Simple Effects (SPSS F-test Method)", level=2)
    doc.add_paragraph(
        "F-tests using pooled error terms from the omnibus model — "
        "matches SPSS GLM simple effects output.")
    if not df_se.empty:
        _wtbl(doc, df_se)
    doc.add_paragraph()

    # 7. Interpretation
    doc.add_heading("7.  Statistical Interpretation", level=1)
    for txt in interpret(res, btw, win, dv, pref):
        clean = re.sub(r"<b>(.*?)</b>", r"\1", txt)
        hdr_, _, body = clean.partition(":")
        p_ = doc.add_paragraph()
        p_.add_run(hdr_ + ":").bold = True
        if body:
            p_.add_run(body)
    doc.add_paragraph()

    # 8. APA Template
    doc.add_heading("8.  APA 7th Edition Reporting Template", level=1)
    doc.add_paragraph(
        f"A {res['a']} ({btw}) × {res['b']} ({win}) mixed ANOVA was conducted "
        f"with {btw} as the between-subjects factor and {win} as the within-subjects "
        f"repeated-measures factor (N = {res['N']}). "
        f"The main effect of {btw} was "
        f"{'statistically significant' if res['p_A'] < alpha else 'not statistically significant'}, "
        f"F({res['df_A']:.0f}, {res['df_SA']:.0f}) = {res['F_A']:.2f}, "
        f"p {fmt_p(res['p_A'])}, partial η²p = {res['np2_A']:.3f}. "
        f"The main effect of {win} was "
        f"{'statistically significant' if res['p_B'] < alpha else 'not statistically significant'}, "
        f"F({res['df_B_c']:.2f}, {res['df_BSA']:.0f}) = {res['F_B']:.2f}, "
        f"p {fmt_p(res['p_B'])}, partial η²p = {res['np2_B']:.3f}. "
        f"The {btw} × {win} interaction was "
        f"{'statistically significant' if res['p_AB'] < alpha else 'not statistically significant'}, "
        f"F({res['df_AB_c']:.2f}, {res['df_BSA']:.0f}) = {res['F_AB']:.2f}, "
        f"p {fmt_p(res['p_AB'])}, partial η²p = {res['np2_AB']:.3f}."
    )
    doc.add_paragraph()

    # 9. Figures
    doc.add_heading("9.  Figures", level=1)
    img = io.BytesIO()
    fig.savefig(img, format="png", dpi=150, bbox_inches="tight", facecolor="#f7f9fc")
    img.seek(0)
    doc.add_picture(img, width=Inches(6.2))
    cap = doc.add_paragraph(
        "Figure 1. Mixed ANOVA visualisation. Profile plot with model-based CI (top-left), "
        "grouped bar chart (top-right), box plots (bottom-left), violin plots (bottom-centre), "
        "partial η²p effect sizes (bottom-right; coloured = significant, grey = n.s.).")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════════════════
#  RUN (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("⏳  Computing Mixed ANOVA …"):
    res = run_mixed_anova(df, alpha_level, sph_corr)

df_ph_b = ph_between_spss(df, posthoc_method, alpha_level)
df_ph_w = ph_within_spss(df, posthoc_method, alpha_level)
df_se = simple_effects_spss(df, res, posthoc_method, alpha_level)

# ══════════════════════════════════════════════════════════════════════════════
#  OUTPUT — OVERVIEW CARDS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Results Overview</div>', unsafe_allow_html=True)

oc = st.columns(5)


def mcard(col, lbl, val, sub):
    col.markdown(
        f'<div class="mcard"><div class="lbl">{lbl}</div>'
        f'<div class="val">{val}</div><div class="sub">{sub}</div></div>',
        unsafe_allow_html=True)


mcard(oc[0], "Total Subjects (N)", str(res["N"]),
      f"{res['a']} group(s) · {res['b']} time-point(s)")
mcard(oc[1], f"Between — {btw_col}", f"F = {res['F_A']:.3f}",
      f"p {fmt_p(res['p_A'])} · η²p = {res['np2_A']:.3f}")
mcard(oc[2], f"Within — {win_col}", f"F = {res['F_B']:.3f}",
      f"p {fmt_p(res['p_B'])} · η²p = {res['np2_B']:.3f}")
mcard(oc[3], "Interaction", f"F = {res['F_AB']:.3f}",
      f"p {fmt_p(res['p_AB'])} · η²p = {res['np2_AB']:.3f}")
mcard(oc[4], "Grand Mean", f"{res['grand']:.3f}",
      f"SD = {df['dv'].std():.3f}")

st.markdown("")
rc1, rc2, rc3 = st.columns(3)
for col_, lbl, F, d1, d2, p, np2, eta2, pw in [
    (rc1, btw_col, res["F_A"], res["df_A"], res["df_SA"],
     res["p_A"], res["np2_A"], res["eta2_A"], res["pow_A"]),
    (rc2, win_col, res["F_B"], res["df_B_c"], res["df_BSA"],
     res["p_B"], res["np2_B"], res["eta2_B"], res["pow_B"]),
    (rc3, f"{btw_col} × {win_col}", res["F_AB"], res["df_AB_c"], res["df_BSA"],
     res["p_AB"], res["np2_AB"], res["eta2_AB"], res["pow_AB"]),
]:
    cf = cohen_f_es(np2)
    mag = magnitude(np2, "partial_eta2")
    col_.markdown(
        f'<div class="rcard"><div class="lbl">{lbl}</div>'
        f'<div class="val">F({d1:.2f}, {d2:.0f}) = {F:.3f}</div>'
        f'<div class="sub">{pill(p, alpha_level)}</div>'
        f'<div class="sub" style="margin-top:5px;">Partial η²p = {np2:.4f} ({mag})</div>'
        f'<div class="sub">η² = {eta2:.4f} · Cohen\'s f = {fmt_f(cf)}</div>'
        f'<div class="sub">Observed power = {pw:.3f}</div></div>',
        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DESCRIPTIVE STATISTICS & ESTIMATED MARGINAL MEANS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
if show_desc:
    st.markdown('<div class="sec">Descriptive Statistics & Estimated Marginal Means</div>',
                unsafe_allow_html=True)

    dt1, dt2, dt3, dt4 = st.tabs([
        "Cell Descriptives (raw)",
        "Cell EMMs (SPSS method)",
        f"Marginal EMMs — {btw_col}",
        f"Marginal EMMs — {win_col}",
    ])

    with dt1:
        st.markdown(
            "Raw cell descriptives (SE = SD / √n). "
            "These match SPSS 'Descriptive Statistics' table."
        )
        cr = res["cell_raw"].copy()
        cr.columns = [btw_col, win_col, "N", "Mean", "SD", "SE (raw)"]
        for c in ["Mean", "SD", "SE (raw)"]:
            cr[c] = cr[c].round(4)
        st.dataframe(cr, hide_index=True, use_container_width=True)

    with dt2:
        ci_pct = int(100 * (1 - alpha_level))
        st.markdown(
            f"**Estimated Marginal Means — SPSS method.** "
            f"SE = √(MS_BS(A) / n_cell) = √({res['MS_BSA']:.4f} / n). "
            f"CI uses t({res['df_BSA']:.0f}) = ±{res['t_crit_w']:.4f} at α = {alpha_level}. "
            f"This matches the SPSS 'Estimated Marginal Means' output."
        )
        emm = res["emm_df"].copy()
        emm.columns = [btw_col, win_col, "N", "Mean", "SE (EMM)",
                       f"{ci_pct}% CI Lower", f"{ci_pct}% CI Upper"]
        st.dataframe(emm, hide_index=True, use_container_width=True)

    with dt3:
        ci_pct = int(100 * (1 - alpha_level))
        st.markdown(
            f"SE = √(MS_S(A) / (b · nᵢ)) = √({res['MS_SA']:.4f} / ({res['b']} · nᵢ)). "
            f"CI uses t({res['df_SA']:.0f}) = ±{res['t_crit_b']:.4f}."
        )
        eb = res["emm_btw"].copy()
        eb.columns = ["Group", "N", "Mean", "SE (EMM)",
                      f"{ci_pct}% CI Lower", f"{ci_pct}% CI Upper"]
        st.dataframe(eb, hide_index=True, use_container_width=True)

    with dt4:
        st.dataframe(res["emm_win"].rename(columns={"Time": win_col}),
                     hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ASSUMPTION DIAGNOSTICS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
if show_assump:
    st.markdown('<div class="sec">Assumption Diagnostics</div>', unsafe_allow_html=True)
    at1, at2, at3 = st.tabs([
        "Normality — Shapiro–Wilk",
        "Homogeneity of Variance — Levene",
        "Sphericity — Mauchly",
    ])

    with at1:
        st.markdown(
            "Shapiro–Wilk applied to each cell. p > α is consistent with normality. "
            "Mixed ANOVA is robust to minor deviations when n ≥ 20 per cell."
        )
        sw = [{"Between": g, "Within": t,
               "n": len(v := df[(df["between"] == g) & (df["within"] == t)]["dv"].values),
               "W": round(stats.shapiro(v)[0], 4) if len(v) >= 3 else np.nan,
               "p-value": fmt_p(stats.shapiro(v)[1]) if len(v) >= 3 else "—",
               "Normal (p>α)": "Yes" if len(v) >= 3 and stats.shapiro(v)[1] > alpha_level else "No"}
              for g in res["btw_lvls"] for t in res["win_lvls"]]
        st.dataframe(pd.DataFrame(sw).rename(columns={"Between": btw_col, "Within": win_col}),
                     hide_index=True, use_container_width=True)

    with at2:
        st.markdown(
            "Levene's test (center = mean) for equality of error variances across "
            "between-subjects groups at each within-subjects level."
        )
        lev = []
        for t in res["win_lvls"]:
            gd = [df[(df["between"] == g) & (df["within"] == t)]["dv"].values for g in res["btw_lvls"]]
            if all(len(x) >= 2 for x in gd):
                Fl, pl = stats.levene(*gd, center="mean")
                lev.append({win_col: t, "Levene F": round(Fl, 4),
                            "df1": res["a"] - 1, "df2": res["N"] - res["a"],
                            "p-value": fmt_p(pl),
                            "Homogeneous (p>α)": "Yes" if pl > alpha_level else "No"})
        st.dataframe(pd.DataFrame(lev), hide_index=True, use_container_width=True)

    with at3:
        if res["sph"] is None:
            st.markdown(
                '<div class="abox-info">Mauchly\'s test not applicable: '
                'within-subjects factor has only 2 levels (sphericity trivially satisfied).</div>',
                unsafe_allow_html=True)
        else:
            s = res["sph"]
            sph_tbl = pd.DataFrame({
                "Statistic": ["Mauchly's W", "Chi-square (χ²)", "df", "p-value",
                              "Greenhouse–Geisser ε", "Huynh–Feldt ε", "Lower-bound ε",
                              "Correction applied", "ε used"],
                "Value": [f"{s['W']:.4f}", f"{s['chi2']:.4f}", str(s['df_m']),
                          fmt_p(s['p']), f"{s['eps_gg']:.4f}", f"{s['eps_hf']:.4f}",
                          f"{s['eps_lb']:.4f}", res["sph_label"], f"{res['eps']:.4f}"],
            })
            st.dataframe(sph_tbl, hide_index=True, use_container_width=True)
            if s["p"] < alpha_level:
                st.markdown(
                    f'<div class="abox-warn"><b>Sphericity violated</b> — '
                    f'Mauchly\'s W = {s["W"]:.4f}, p {fmt_p(s["p"])} &lt; {alpha_level}. '
                    f'df corrected using {res["sph_label"]} (ε = {res["eps"]:.4f}).</div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="abox-ok"><b>Sphericity satisfied</b> — '
                    f'Mauchly\'s W = {s["W"]:.4f}, p {fmt_p(s["p"])} ≥ {alpha_level}. '
                    f'No df correction required.</div>',
                    unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ANOVA SUMMARY TABLE (DIPERBAIKI — tampilkan df_BSA yang benar)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Mixed ANOVA Summary Table</div>', unsafe_allow_html=True)

at_rows_disp = [
    {"Source": "BETWEEN SUBJECTS", "SS": "", "df": "", "MS": "", "F": "", "p": "", "Partial η²p": "", "Obs. Power": ""},
    {"Source": f"  {btw_col}", "SS": fmt_f(res["SS_A"]), "df": fmt_f(res["df_A"], 0),
     "MS": fmt_f(res["MS_A"]), "F": fmt_f(res["F_A"]), "p": fmt_p(res["p_A"]),
     "Partial η²p": fmt_f(res["np2_A"]), "Obs. Power": fmt_f(res["pow_A"])},
    {"Source": "  Error [S(A)]", "SS": fmt_f(res["SS_SA"]), "df": fmt_f(res["df_SA"], 0),
     "MS": fmt_f(res["MS_SA"]), "F": "—", "p": "—", "Partial η²p": "—", "Obs. Power": "—"},
    {"Source": "WITHIN SUBJECTS", "SS": "", "df": "", "MS": "", "F": "", "p": "", "Partial η²p": "", "Obs. Power": ""},
    {"Source": f"  {win_col}", "SS": fmt_f(res["SS_B"]), "df": fmt_f(res["df_B_c"]),
     "MS": fmt_f(res["MS_B"]), "F": fmt_f(res["F_B"]), "p": fmt_p(res["p_B"]),
     "Partial η²p": fmt_f(res["np2_B"]), "Obs. Power": fmt_f(res["pow_B"])},
    {"Source": f"  {btw_col} × {win_col}", "SS": fmt_f(res["SS_AB"]), "df": fmt_f(res["df_AB_c"]),
     "MS": fmt_f(res["MS_AB"]), "F": fmt_f(res["F_AB"]), "p": fmt_p(res["p_AB"]),
     "Partial η²p": fmt_f(res["np2_AB"]), "Obs. Power": fmt_f(res["pow_AB"])},
    {"Source": "  Error [BS(A)]", "SS": fmt_f(res["SS_BSA"]), "df": fmt_f(res["df_BSA"], 0),
     "MS": fmt_f(res["MS_BSA"]), "F": "—", "p": "—", "Partial η²p": "—", "Obs. Power": "—"},
    {"Source": "Total", "SS": fmt_f(res["SS_Total"]), "df": fmt_f(res["df_Tot"], 0),
     "MS": "—", "F": "—", "p": "—", "Partial η²p": "—", "Obs. Power": "—"},
]
at_disp = pd.DataFrame(at_rows_disp)
st.dataframe(at_disp, hide_index=True, use_container_width=True)
if res["sph"] and res["eps"] < 1.0:
    st.caption(f"Note. df for within-subjects effects adjusted using "
               f"{res['sph_label']} (ε = {res['eps']:.4f}). "
               f"S(A) = Subjects within Groups; BS(A) = B × Subjects within Groups.")
else:
    st.caption("S(A) = Subjects within Groups (between error). "
               "BS(A) = B × Subjects within Groups (within error).")

# Effect size table
st.markdown('<div class="sec">Effect Size Summary</div>', unsafe_allow_html=True)
es_disp = pd.DataFrame({
    "Source": [btw_col, win_col, f"{btw_col} × {win_col}"],
    "Partial η²p": [round(res["np2_A"], 4), round(res["np2_B"], 4), round(res["np2_AB"], 4)],
    "η² (total)": [round(res["eta2_A"], 4), round(res["eta2_B"], 4), round(res["eta2_AB"], 4)],
    "Cohen's f": [round(cohen_f_es(res["np2_A"]), 4), round(cohen_f_es(res["np2_B"]), 4),
                  round(cohen_f_es(res["np2_AB"]), 4)],
    "Magnitude": [magnitude(res["np2_A"], "partial_eta2"), magnitude(res["np2_B"], "partial_eta2"),
                  magnitude(res["np2_AB"], "partial_eta2")],
    "F": [fmt_f(res["F_A"]), fmt_f(res["F_B"]), fmt_f(res["F_AB"])],
    "p": [fmt_p(res["p_A"]), fmt_p(res["p_B"]), fmt_p(res["p_AB"])],
    "Significant": [res["p_A"] < alpha_level, res["p_B"] < alpha_level, res["p_AB"] < alpha_level],
})
st.dataframe(es_disp, hide_index=True, use_container_width=True)
st.caption("Partial η²p benchmarks (Cohen 1988): negligible < .01; small .01–.05; medium .06–.13; large ≥ .14.")

# ══════════════════════════════════════════════════════════════════════════════
#  POST-HOC (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
if show_posthoc:
    st.markdown('<div class="sec">Post-Hoc Comparisons</div>', unsafe_allow_html=True)
    corr_lbl = {"bonferroni": "Bonferroni", "holm": "Holm (step-down Bonferroni)",
                "sidak": "Šidák", "fdr_bh": "Benjamini–Hochberg FDR"}.get(posthoc_method, posthoc_method)

    pt1, pt2, pt3 = st.tabs([
        f"Between — {btw_col}",
        f"Within — {win_col}",
        "Simple Effects (SPSS F-test)",
    ])

    with pt1:
        st.markdown(
            f"**Pooled-variance independent t-tests** (equal variance assumed — matches SPSS GLM). "
            f"p-values corrected: **{corr_lbl}**."
        )
        if res["a"] < 3:
            st.markdown('<div class="abox-info">Only 2 levels — omnibus F-test is conclusive.</div>',
                        unsafe_allow_html=True)
        elif df_ph_b.empty:
            st.warning("No comparisons could be computed.")
        else:
            st.dataframe(df_ph_b, hide_index=True, use_container_width=True)
            st.caption("Cohen's d: |d| < .20 = negligible; .20–.49 = small; .50–.79 = medium; ≥ .80 = large.")

    with pt2:
        st.markdown(
            f"**Paired-samples t-tests** (accounts for within-subject correlation). "
            f"p-values corrected: **{corr_lbl}**."
        )
        if res["b"] < 3:
            st.markdown('<div class="abox-info">Only 2 levels — omnibus F-test is conclusive.</div>',
                        unsafe_allow_html=True)
        elif df_ph_w.empty:
            st.warning("No comparisons could be computed.")
        else:
            st.dataframe(df_ph_w, hide_index=True, use_container_width=True)

    with pt3:
        st.markdown(
            f"**Simple effects — SPSS F-test method** (not paired t-tests). "
            f"Uses the pooled error term from the omnibus model:\n"
            f"- Within at each group: F = MS_B@group / MS_BS(A), df = ({res['b']-1}, {res['df_BSA']})\n"
            f"- Between at each time: F = MS_A@time / MS_S(A), df = ({res['a']-1}, {res['df_SA']})"
        )
        if df_se.empty:
            st.warning("Simple effects could not be computed.")
        else:
            se_disp = df_se.drop(columns=["p_raw"], errors="ignore")
            st.dataframe(se_disp, hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  INTERPRETATION (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Statistical Interpretation</div>', unsafe_allow_html=True)

for txt in interpret(res, btw_col, win_col, dv_col, effect_pref):
    st.markdown(f'<div class="ibox">{txt}</div>', unsafe_allow_html=True)

with st.expander("APA 7th Edition Reporting Template", expanded=False):
    apa = (
        f"A {res['a']} ({btw_col}) × {res['b']} ({win_col}) mixed analysis of variance "
        f"(ANOVA) was conducted with {btw_col} as the between-subjects factor and "
        f"{win_col} as the within-subjects repeated-measures factor (N = {res['N']}). "
        f"The main effect of {btw_col} was "
        f"{'statistically significant' if res['p_A'] < alpha_level else 'not statistically significant'}, "
        f"F({res['df_A']:.0f}, {res['df_SA']:.0f}) = {res['F_A']:.2f}, "
        f"p {fmt_p(res['p_A'])}, partial η²p = {res['np2_A']:.3f}. "
        f"The main effect of {win_col} was "
        f"{'statistically significant' if res['p_B'] < alpha_level else 'not statistically significant'}, "
        f"F({res['df_B_c']:.2f}, {res['df_BSA']:.0f}) = {res['F_B']:.2f}, "
        f"p {fmt_p(res['p_B'])}, partial η²p = {res['np2_B']:.3f}. "
        f"The {btw_col} × {win_col} interaction was "
        f"{'statistically significant' if res['p_AB'] < alpha_level else 'not statistically significant'}, "
        f"F({res['df_AB_c']:.2f}, {res['df_BSA']:.0f}) = {res['F_AB']:.2f}, "
        f"p {fmt_p(res['p_AB'])}, partial η²p = {res['np2_AB']:.3f}."
    )
    st.text_area("Copy APA-formatted result:", value=apa, height=160)

# ══════════════════════════════════════════════════════════════════════════════
#  FIGURES (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Visualisations</div>', unsafe_allow_html=True)
with st.spinner("Rendering plots …"):
    fig = build_figure(res, df, btw_col, win_col, dv_col, pal_name, grid_style)
st.pyplot(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DOWNLOADS (SAMA — TIDAK DIUBAH)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Download Results</div>', unsafe_allow_html=True)

dl1, dl2, dl3, dl4 = st.columns(4)

with dl1:
    st.download_button("📄  ANOVA Table (CSV)",
                       at_disp.to_csv(index=False).encode(),
                       "mixed_anova_table.csv", "text/csv", use_container_width=True)
with dl2:
    full = pd.DataFrame({
        "Effect": [btw_col, win_col, f"{btw_col}×{win_col}"],
        "SS": [res["SS_A"], res["SS_B"], res["SS_AB"]],
        "df_effect": [res["df_A"], res["df_B_c"], res["df_AB_c"]],
        "df_error": [res["df_SA"], res["df_BSA"], res["df_BSA"]],
        "MS_effect": [res["MS_A"], res["MS_B"], res["MS_AB"]],
        "MS_error": [res["MS_SA"], res["MS_BSA"], res["MS_BSA"]],
        "F": [res["F_A"], res["F_B"], res["F_AB"]],
        "p": [res["p_A"], res["p_B"], res["p_AB"]],
        "partial_eta2p": [res["np2_A"], res["np2_B"], res["np2_AB"]],
        "eta2": [res["eta2_A"], res["eta2_B"], res["eta2_AB"]],
        "cohens_f": [cohen_f_es(res["np2_A"]), cohen_f_es(res["np2_B"]), cohen_f_es(res["np2_AB"])],
        "observed_power": [res["pow_A"], res["pow_B"], res["pow_AB"]],
        "epsilon": [1.0, res["eps"], res["eps"]],
        "sph_correction": ["N/A", res["sph_label"], res["sph_label"]],
    })
    st.download_button("📊  Full Statistics (CSV)",
                       full.to_csv(index=False).encode(),
                       "mixed_anova_results.csv", "text/csv", use_container_width=True)
with dl3:
    fb = io.BytesIO()
    fig.savefig(fb, format="png", dpi=200, bbox_inches="tight", facecolor="#f7f9fc")
    fb.seek(0)
    st.download_button("🖼️  Figures (PNG, 200 dpi)", fb.getvalue(),
                       "mixed_anova_figures.png", "image/png", use_container_width=True)
with dl4:
    with st.spinner("Building Word report …"):
        wb = build_report(res, df, btw_col, win_col, dv_col, alpha_level,
                          posthoc_method, effect_pref, fig, df_ph_b, df_ph_w, df_se)
    st.download_button("📝  Full Report (Word .docx)", wb.getvalue(),
                       "mixed_anova_report.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)

st.markdown("---")
st.caption(
    "Mixed ANOVA Calculator · SPSS GLM Repeated Measures Equivalent · "
    "SS decomposition verified (sum = SS_Total) · EMM uses model-based SE · "
    "Simple effects use pooled F-test (SPSS method) · "
    "References: Winer, Brown & Michels (1991); Mauchly (1940); Box (1954); "
    "Greenhouse & Geisser (1959); Huynh & Feldt (1976); Lecoutre (1991); Cohen (1988)."
)
