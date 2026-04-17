"""
Mixed ANOVA Calculator
======================
SPSS GLM Repeated Measures equivalent.

Statistical model (Winer, Brown & Michels, 1991, Ch. 7):
  Two-way mixed factorial: A (between-subjects) × B (within-subjects)

SS Decomposition — verified to satisfy SS_Total = SS_A + SS_S(A) + SS_B + SS_AB + SS_BS(A):
  SS_A     = b · Σᵢ nᵢ (Ȳᵢ.. − Ȳ...)²             df = a − 1
  SS_S(A)  = b · Σᵢ Σₛ (Ȳₛ.. − Ȳᵢ..)²            df = N − a
  SS_B     = N · Σⱼ (Ȳ.j. − Ȳ...)²               df = b − 1
  SS_AB    = Σᵢ nᵢ · Σⱼ (Ȳᵢⱼ − Ȳᵢ.. − Ȳ.j. + Ȳ...)²  df = (a−1)(b−1)
  SS_BS(A) = Σᵢ Σₛ Σⱼ (yₛⱼ − Ȳₛ.. − Ȳᵢⱼ + Ȳᵢ..)²    df = (b−1)(N−a)

F-ratios:
  F_A  = MS_A  / MS_S(A)    (between error)
  F_B  = MS_B  / MS_BS(A)   (within error)
  F_AB = MS_AB / MS_BS(A)   (within error)

Effect sizes:
  Partial η²p = SS_effect / (SS_effect + SS_error_for_that_effect)  [SPSS default]
  η²          = SS_effect / SS_Total
  Cohen's f   = √(η²p / (1 − η²p))

Sphericity: Mauchly (1940) W, Box (1954) χ² approximation, GG (1959), HF (1976/Lecoutre 1991)
Post-hoc:   Bonferroni, Holm, Šidák, Benjamini–Hochberg FDR
"""

# ── stdlib ─────────────────────────────────────────────────────────────────────
import io
import itertools
import re
import warnings

# ── third-party ────────────────────────────────────────────────────────────────
import numpy as np
import pandas as pd
import scipy.stats as stats
from scipy.stats import f as fdist

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import seaborn as sns

import streamlit as st

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Mixed ANOVA Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
#  GLOBAL CSS  (no raw Markdown tables ever printed in the app)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lora:wght@500;600&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* ── App header ── */
.app-title {
    font-family: 'Lora', serif; font-size: 2.2rem; font-weight: 600;
    color: #0d1b2a; letter-spacing: -0.01em; margin-bottom: 2px;
}
.app-sub {
    font-size: 0.88rem; color: #546e7a; font-weight: 300; margin-bottom: 1.5rem;
}

/* ── Section headers ── */
.sec {
    font-family: 'Lora', serif; font-size: 1.15rem; font-weight: 600;
    color: #0d1b2a; border-left: 4px solid #d32f2f; padding-left: 10px;
    margin-top: 1.8rem; margin-bottom: 0.75rem;
}

/* ── Metric card ── */
.mcard {
    background: #0d1b2a; border-radius: 10px;
    padding: 0.9rem 1.1rem 0.85rem; margin-bottom: 4px;
}
.mcard .mc-lbl { font-size: 0.64rem; font-weight:600; color:#7fa8c9;
                 text-transform:uppercase; letter-spacing:.10em; }
.mcard .mc-val { font-family:'JetBrains Mono',monospace;
                 font-size:1.25rem; font-weight:600; color:#eaf2ff; margin:3px 0 2px; }
.mcard .mc-sub { font-size:0.74rem; color:#90b8d8; line-height:1.5; }

/* ── Result card (effect row) ── */
.rcard {
    background:#112233; border-radius:10px;
    padding:0.9rem 1.1rem 0.85rem; margin-bottom:4px;
}
.rcard .mc-lbl { font-size:0.64rem; font-weight:600; color:#7fa8c9;
                 text-transform:uppercase; letter-spacing:.10em; }
.rcard .mc-val { font-family:'JetBrains Mono',monospace;
                 font-size:1.1rem; font-weight:600; color:#eaf2ff; margin:3px 0 2px; }
.rcard .mc-sub { font-size:0.74rem; color:#90b8d8; line-height:1.55; }

/* ── Significance pills ── */
.p-sig   {display:inline-block;background:#1b7f45;color:#fff;
           font-size:0.69rem;font-weight:700;padding:2px 8px;border-radius:20px;}
.p-ns    {display:inline-block;background:#b71c1c;color:#fff;
           font-size:0.69rem;font-weight:700;padding:2px 8px;border-radius:20px;}
.p-trend {display:inline-block;background:#e65100;color:#fff;
           font-size:0.69rem;font-weight:700;padding:2px 8px;border-radius:20px;}

/* ── Interpretation panel ── */
.ibox {
    background:#f0f6ff; border-left:4px solid #1565c0;
    border-radius:6px; padding:.8rem 1rem;
    font-size:0.875rem; line-height:1.72; color:#0d1b2a;
    margin-bottom:.55rem;
}
.ibox b { color:#0d1b2a; }

/* ── Alert panels ── */
.abox-warn {
    background:#fff8e1; border-left:4px solid #f9a825;
    border-radius:6px; padding:.7rem 1rem;
    font-size:0.85rem; color:#5d4037; margin-bottom:.5rem;
}
.abox-ok {
    background:#e8f5e9; border-left:4px solid #2e7d32;
    border-radius:6px; padding:.7rem 1rem;
    font-size:0.85rem; color:#1b5e20; margin-bottom:.5rem;
}
.abox-info {
    background:#e3f2fd; border-left:4px solid #1565c0;
    border-radius:6px; padding:.7rem 1rem;
    font-size:0.85rem; color:#0d47a1; margin-bottom:.5rem;
}

/* ── Upload hint ── */
.upload-hint {
    background:#f5f7fa; border:1.5px dashed #b0bec5;
    border-radius:8px; padding:.85rem 1rem;
    font-size:0.85rem; color:#546e7a;
}

/* ── Sidebar ── */
[data-testid="stSidebar"]          { background:#0d1b2a !important; }
[data-testid="stSidebar"] *        { color:#cfe2f3 !important; }
[data-testid="stSidebar"] hr      { border-color:#1c3048 !important; }
[data-testid="stSidebar"] .stSelectbox > div > div { background:#1c3048; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR  ── all analysis settings
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Analysis Settings")
    st.markdown("---")

    alpha_level = st.selectbox(
        "Significance level (α)",
        options=[0.05, 0.01, 0.001, 0.10], index=0,
        help="Type I error rate applied to all hypothesis tests and corrections.")

    posthoc_method = st.selectbox(
        "Post-hoc correction method",
        options=["bonferroni", "holm", "sidak", "fdr_bh"],
        format_func=lambda x: {
            "bonferroni": "Bonferroni",
            "holm":       "Holm (step-down Bonferroni)",
            "sidak":      "Šidák",
            "fdr_bh":     "FDR — Benjamini–Hochberg",
        }[x],
        help="Multiple-comparison correction for post-hoc pairwise tests.")

    effect_pref = st.selectbox(
        "Primary effect size metric",
        options=["partial_eta2", "eta2", "cohen_f"],
        format_func=lambda x: {
            "partial_eta2": "Partial η²p  (SPSS default)",
            "eta2":         "η²  (eta-squared)",
            "cohen_f":      "Cohen's f",
        }[x])

    sph_corr = st.selectbox(
        "Sphericity correction",
        options=["auto", "gg", "hf", "lb", "none"],
        format_func=lambda x: {
            "auto": "Auto — apply GG when Mauchly p < α",
            "gg":   "Greenhouse–Geisser  (always)",
            "hf":   "Huynh–Feldt  (always)",
            "lb":   "Lower-bound  (always)",
            "none": "None — assume sphericity",
        }[x],
        help="Applied only to within-subjects and interaction effects.")

    st.markdown("---")
    st.markdown("**Display Options**")
    show_desc    = st.checkbox("Descriptive statistics",  value=True)
    show_assump  = st.checkbox("Assumption diagnostics",  value=True)
    show_posthoc = st.checkbox("Post-hoc comparisons",   value=True)

    st.markdown("---")
    st.markdown("**Plot Appearance**")
    pal_name   = st.selectbox("Color palette",
                              ["tab10", "Set2", "deep", "colorblind", "husl"])
    grid_style = st.selectbox("Grid style",
                              ["whitegrid", "ticks", "darkgrid"])

# ══════════════════════════════════════════════════════════════════════════════
#  APP HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="app-title">Mixed ANOVA Calculator</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-sub">'
    'Two-Way Mixed ANOVA · SPSS GLM Repeated Measures Equivalent · '
    'Mauchly Sphericity · GG / HF / LB Corrections · '
    'Post-hoc Comparisons · Effect Sizes · Full Report Export'
    '</div>',
    unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  FORMAT GUIDE & TEMPLATE
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📋  Data Format Guide & CSV Template Download", expanded=False):
    st.markdown('<div class="sec">Required CSV Format</div>', unsafe_allow_html=True)

    col_info, col_rules = st.columns([1.1, 0.9])

    with col_info:
        st.markdown(
            "Your data must be in **long (tidy) format** — one observation per row. "
            "Each row is one measurement for one subject at one level of the within-subjects factor."
        )
        st.dataframe(
            pd.DataFrame({
                "Column name": ["subject_id", "between_factor", "within_factor", "dependent_var"],
                "Role": [
                    "Unique participant identifier",
                    "Between-subjects group membership",
                    "Within-subjects repeated measure (e.g., Time)",
                    "Numeric outcome variable",
                ],
                "Example": ["1, 2, … N", "Control / Treatment_A / Treatment_B",
                            "Pre / Post / Follow_up", "42.5, 67.3, …"],
                "Required": ["✓", "✓", "✓", "✓"],
            }),
            hide_index=True, use_container_width=True,
        )

    with col_rules:
        st.dataframe(
            pd.DataFrame({
                "Design constraint": [
                    "Min. between-subjects groups",
                    "Max. between-subjects groups",
                    "Min. within-subjects levels",
                    "Max. within-subjects levels",
                    "Min. subjects per group",
                    "Balance requirement",
                    "Missing data",
                    "File size limit",
                    "File encoding",
                ],
                "Specification": [
                    "2  (supports 3, 4, … k)",
                    "No hard limit",
                    "2  (supports 3, 4, … t)",
                    "No hard limit",
                    "≥ 5  (≥ 10 recommended)",
                    "Each subject must have data at ALL within-levels",
                    "Listwise deletion applied automatically",
                    "3 MB",
                    "UTF-8",
                ],
            }),
            hide_index=True, use_container_width=True,
        )

    st.markdown("---")
    st.markdown("**Important:** The order of levels for the within-subjects factor is determined "
                "alphabetically / lexicographically. Prefix with numbers (e.g., `1_Pre`, `2_Post`) "
                "to control order if needed.")
    st.markdown("---")

    # ── Generate templates ─────────────────────────────────────────────────────
    def make_template(groups, times, n_per):
        np.random.seed(0)
        base  = {g: 50 + i*3 for i, g in enumerate(groups)}
        tgain = {t: j*4 for j, t in enumerate(times)}
        igain = {g: {t: i*j*1.5 for j, t in enumerate(times)}
                 for i, g in enumerate(groups)}
        rows, sid = [], 1
        for g in groups:
            for _ in range(n_per):
                re = np.random.normal(0, 3)
                for t in times:
                    v = base[g] + tgain[t] + igain[g][t] + re + np.random.normal(0, 1.8)
                    rows.append({"subject_id": sid, "Group": g, "Time": t,
                                 "Score": round(v, 2)})
                sid += 1
        return pd.DataFrame(rows)

    tpl2x2 = make_template(["Control","Treatment"], ["Pre","Post"], 12)
    tpl2x3 = make_template(["Control","Treatment"], ["Pre","Post","Follow_up"], 10)
    tpl3x3 = make_template(["Control","Treat_A","Treat_B"], ["Pre","Post","Follow_up"], 10)
    tpl3x4 = make_template(["Control","Treat_A","Treat_B"],
                           ["Baseline","Week_4","Week_8","Week_12"], 10)

    tc1, tc2, tc3, tc4 = st.columns(4)
    with tc1:
        st.markdown("**2 groups × 2 time-points**")
        st.caption(f"{len(tpl2x2)} rows — N = 24")
        st.download_button("⬇️ Download", tpl2x2.to_csv(index=False).encode(),
                           "template_2x2.csv", "text/csv", use_container_width=True)
    with tc2:
        st.markdown("**2 groups × 3 time-points**")
        st.caption(f"{len(tpl2x3)} rows — N = 20")
        st.download_button("⬇️ Download", tpl2x3.to_csv(index=False).encode(),
                           "template_2x3.csv", "text/csv", use_container_width=True)
    with tc3:
        st.markdown("**3 groups × 3 time-points**")
        st.caption(f"{len(tpl3x3)} rows — N = 30")
        st.download_button("⬇️ Download", tpl3x3.to_csv(index=False).encode(),
                           "template_3x3.csv", "text/csv", use_container_width=True)
    with tc4:
        st.markdown("**3 groups × 4 time-points**")
        st.caption(f"{len(tpl3x4)} rows — N = 30")
        st.download_button("⬇️ Download", tpl3x4.to_csv(index=False).encode(),
                           "template_3x4.csv", "text/csv", use_container_width=True)

    st.markdown("---")
    st.markdown("**Preview of 3-group × 3-time-point template (first 15 rows):**")
    st.dataframe(tpl3x3.head(15), hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Upload Your Data</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload CSV  (max 3 MB, UTF-8, long format)",
    type=["csv"],
    help="See the format guide above for column requirements.",
)

if uploaded_file is None:
    st.markdown(
        '<div class="upload-hint">📂 No file uploaded yet. '
        'Upload your CSV above, or download one of the templates to get started.</div>',
        unsafe_allow_html=True)
    st.stop()

if uploaded_file.size > 3 * 1024 * 1024:
    st.error("❌  File exceeds the 3 MB limit. Please reduce the dataset size.")
    st.stop()

try:
    df_raw = pd.read_csv(uploaded_file)
except Exception as exc:
    st.error(f"❌  Could not read CSV: {exc}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  COLUMN MAPPING
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Column Mapping</div>', unsafe_allow_html=True)

all_cols  = df_raw.columns.tolist()
num_cols  = df_raw.select_dtypes(include=np.number).columns.tolist()

cm1, cm2, cm3, cm4 = st.columns(4)
with cm1:
    subj_col = st.selectbox("🔑  Subject ID", all_cols, index=0,
                            help="Unique identifier per participant.")
with cm2:
    btw_col = st.selectbox("👥  Between-subjects factor", all_cols,
                           index=min(1, len(all_cols)-1),
                           help="Grouping variable (e.g., treatment condition).")
with cm3:
    win_col = st.selectbox("🔁  Within-subjects factor", all_cols,
                           index=min(2, len(all_cols)-1),
                           help="Repeated-measure variable (e.g., time-point).")
with cm4:
    dv_col = st.selectbox("📈  Dependent variable", num_cols,
                          index=0, help="Numeric outcome measure.")

run = st.button("▶  Run Mixed ANOVA", type="primary", use_container_width=True)

if not run:
    with st.expander("Preview uploaded data (first 10 rows)", expanded=False):
        st.dataframe(df_raw.head(10), hide_index=True, use_container_width=True)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  DATA CLEANING & VALIDATION
# ══════════════════════════════════════════════════════════════════════════════
df = df_raw[[subj_col, btw_col, win_col, dv_col]].copy()
df.columns = ["subject", "between", "within", "dv"]
df["dv"] = pd.to_numeric(df["dv"], errors="coerce")

n_raw = len(df)
df.dropna(inplace=True)
if len(df) < n_raw:
    st.warning(f"⚠️  {n_raw - len(df)} row(s) with missing values removed (listwise deletion).")

within_lvls  = sorted(df["within"].unique().tolist(),  key=str)
between_lvls = sorted(df["between"].unique().tolist(), key=str)
a = len(between_lvls)
b = len(within_lvls)

if b < 2:
    st.error("Within-subjects factor must have at least 2 levels.")
    st.stop()
if a < 2:
    st.error("Between-subjects factor must have at least 2 levels.")
    st.stop()

# Listwise: keep only subjects with a complete set of within-level observations
subj_counts    = df.groupby("subject")["within"].count()
complete_subjs = subj_counts[subj_counts == b].index
n_dropped_subj = len(subj_counts) - len(complete_subjs)
df = df[df["subject"].isin(complete_subjs)].copy()
if n_dropped_subj:
    st.warning(f"⚠️  {n_dropped_subj} subject(s) removed: incomplete within-factor data.")

N           = df["subject"].nunique()
n_per_group = df.groupby("between")["subject"].nunique()

if N < 6:
    st.error("At least 6 complete subjects required for a valid analysis.")
    st.stop()
if n_per_group.min() < 3:
    st.error(f"Smallest group has {n_per_group.min()} subject(s). Minimum required: 3.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def fmt_p(p: float) -> str:
    """APA-style p-value string."""
    if pd.isna(p):  return "—"
    if p < 0.001:   return "< .001"
    s = f"{p:.3f}"
    return s[1:] if s.startswith("0") else s   # strip leading zero → ".035"

def fmt_f(v, dec: int = 3) -> str:
    if pd.isna(v): return "—"
    return f"{v:.{dec}f}"

def cohen_f(np2: float) -> float:
    if np.isnan(np2) or np2 >= 1.0: return np.nan
    return float(np.sqrt(np2 / (1.0 - np2)))

def magnitude(val: float, metric: str = "partial_eta2") -> str:
    """Cohen (1988) magnitude benchmarks."""
    if np.isnan(val): return "—"
    if metric in ("partial_eta2", "eta2"):
        if val < 0.010: return "negligible"
        if val < 0.060: return "small"
        if val < 0.140: return "medium"
        return "large"
    # cohen_f
    if val < 0.10: return "negligible"
    if val < 0.25: return "small"
    if val < 0.40: return "medium"
    return "large"

def obs_power(F_val: float, df1: float, df2: float, alpha: float) -> float:
    """Non-central F distribution power estimate."""
    try:
        ncp  = max(float(F_val * df1), 0.0)
        crit = fdist.ppf(1.0 - alpha, df1, df2)
        return float(np.clip(1.0 - stats.ncf.cdf(crit, df1, df2, nc=ncp), 0.0, 1.0))
    except Exception:
        return np.nan

def pill(p: float, alpha: float) -> str:
    if pd.isna(p): return ""
    lbl = "< .001" if p < 0.001 else f"= {fmt_p(p)}"
    if p < alpha:   return f'<span class="p-sig">p {lbl}  ✓</span>'
    if p < 0.10:    return f'<span class="p-trend">p {lbl}  ~</span>'
    return f'<span class="p-ns">p {lbl}  n.s.</span>'

def multipletests_local(p_values: list, method: str, alpha: float):
    """
    Vectorised multiple-testing correction — no statsmodels dependency.
    Supports: bonferroni, holm, sidak, fdr_bh.
    """
    p_arr = np.array(p_values, dtype=float)
    n     = len(p_arr)
    order = np.argsort(p_arr)
    p_adj = np.empty(n)

    if method == "bonferroni":
        p_adj = np.clip(p_arr * n, 0, 1)

    elif method == "holm":
        # Holm (1979): step-down, reject while p_{(i)} ≤ α / (n − i + 1)
        for rank, idx in enumerate(order):
            p_adj[idx] = min(1.0, p_arr[idx] * (n - rank))
        # enforce monotonicity
        cum_max = 0.0
        for idx in order:
            cum_max = max(cum_max, p_adj[idx])
            p_adj[idx] = cum_max

    elif method == "sidak":
        p_adj = np.clip(1.0 - (1.0 - p_arr) ** n, 0, 1)

    elif method == "fdr_bh":
        # Benjamini–Hochberg (1995)
        p_adj_sorted = np.empty(n)
        for i, idx in enumerate(order):
            p_adj_sorted[i] = p_arr[idx] * n / (i + 1)
        # enforce monotone decrease from right
        running_min = 1.0
        for i in range(n - 1, -1, -1):
            running_min = min(running_min, p_adj_sorted[i])
            p_adj_sorted[i] = running_min
        # map back to original order
        for i, idx in enumerate(order):
            p_adj[idx] = np.clip(p_adj_sorted[i], 0, 1)
    else:
        p_adj = np.clip(p_arr * n, 0, 1)  # fallback = bonferroni

    reject = p_adj < alpha
    return reject, p_adj

# ══════════════════════════════════════════════════════════════════════════════
#  MAUCHLY'S TEST OF SPHERICITY
# ══════════════════════════════════════════════════════════════════════════════

def mauchly_test(wide: np.ndarray) -> dict:
    """
    Mauchly (1940) W via Box (1954) chi-square approximation — SPSS formula.

    Parameters
    ----------
    wide : (N_subjects, b) array — each row is one subject's repeated measures.

    Returns
    -------
    dict: W, chi2, df_m, p, eps_gg, eps_hf, eps_lb
    """
    n, k = wide.shape
    p = k - 1   # number of contrasts

    # Orthonormal Helmert contrast matrix  (k × p)
    C = np.zeros((k, p))
    for j in range(p):
        C[:j+1, j] =  1.0 / np.sqrt((j + 1) * (j + 2))
        C[j+1,  j] = -np.sqrt((j + 1) / (j + 2))

    Y = wide @ C                           # (n, p) — contrast scores
    S = np.cov(Y.T, ddof=1)               # (p, p) — unbiased covariance
    if S.ndim == 0:
        S = np.array([[float(S)]])

    det_S   = float(max(np.linalg.det(S), 1e-300))
    trace_S = float(np.trace(S))

    W = float(np.clip(det_S / (trace_S / p) ** p, 1e-10, 1.0))

    # Box (1954) chi-square approximation
    f_factor = 1.0 - (2.0 * p**2 + p + 2) / (6.0 * p * (n - 1))
    df_m     = int(p * (p + 1) / 2 - 1)
    chi2_val = -np.log(W) * (n - 1) * f_factor
    p_val    = float(1.0 - stats.chi2.cdf(chi2_val, df_m))

    # Greenhouse-Geisser ε (1959)
    trace_S2 = float(np.trace(S @ S))
    eps_gg   = float(np.clip(trace_S**2 / (p * trace_S2), 1.0 / p, 1.0))

    # Huynh-Feldt ε — Lecoutre (1991) correction, identical to SPSS output
    hf_num = n * p * eps_gg - 2.0
    hf_den = p * (n - 1.0 - p * eps_gg)
    eps_hf = float(np.clip(hf_num / hf_den if hf_den != 0 else 1.0, 1.0 / p, 1.0))

    eps_lb = float(1.0 / p)   # lower-bound = 1 / (b − 1)

    return dict(W=W, chi2=chi2_val, df_m=df_m, p=p_val,
                eps_gg=eps_gg, eps_hf=eps_hf, eps_lb=eps_lb)

# ══════════════════════════════════════════════════════════════════════════════
#  MIXED ANOVA ENGINE  (verified: SS components sum exactly to SS_Total)
# ══════════════════════════════════════════════════════════════════════════════

def run_mixed_anova(df: pd.DataFrame, alpha: float, sph_corr: str) -> dict:
    """
    Two-way Mixed ANOVA: A (between) × B (within).
    Source: Winer, Brown & Michels (1991), Chapter 7.
    """
    between_lvls = sorted(df["between"].unique().tolist(), key=str)
    within_lvls  = sorted(df["within"].unique().tolist(),  key=str)
    a = len(between_lvls)
    b = len(within_lvls)
    N = df["subject"].nunique()
    n_g = df.groupby("between")["subject"].nunique()   # nᵢ per group

    grand  = df["dv"].mean()
    g_mean = df.groupby("between")["dv"].mean()         # Ȳᵢ..
    t_mean = df.groupby("within")["dv"].mean()          # Ȳ.j.
    s_mean = df.groupby("subject")["dv"].mean()         # Ȳₛ..
    s_grp  = df.groupby("subject")["between"].first()  # subject → group map

    # Cell means  Ȳᵢⱼ
    cell_mean = {
        (g, t): df[(df["between"] == g) & (df["within"] == t)]["dv"].mean()
        for g in between_lvls for t in within_lvls
    }

    # Cell descriptives (for tables / plots)
    cell_stats = (
        df.groupby(["between", "within"])["dv"]
          .agg(N="count", Mean="mean", SD="std")
          .reset_index()
    )
    cell_stats["SE"]  = cell_stats["SD"] / np.sqrt(cell_stats["N"])
    cell_stats["CI_lo"] = cell_stats["Mean"] - 1.96 * cell_stats["SE"]
    cell_stats["CI_hi"] = cell_stats["Mean"] + 1.96 * cell_stats["SE"]

    # ── SS_A  (between-subjects main effect) ─────────────────────────────────
    # b · Σᵢ nᵢ (Ȳᵢ.. − Ȳ...)²
    SS_A = b * sum(n_g[g] * (g_mean[g] - grand) ** 2 for g in between_lvls)

    # ── SS_S(A)  (subjects-within-groups; between error) ─────────────────────
    # b · Σᵢ Σₛ (Ȳₛ.. − Ȳᵢ..)²
    SS_SA = 0.0
    for subj in df["subject"].unique():
        g = s_grp[subj]
        SS_SA += b * (s_mean[subj] - g_mean[g]) ** 2

    # ── SS_B  (within-subjects main effect) ──────────────────────────────────
    # N · Σⱼ (Ȳ.j. − Ȳ...)²
    SS_B = N * sum((t_mean[t] - grand) ** 2 for t in within_lvls)

    # ── SS_AB  (interaction) ──────────────────────────────────────────────────
    # Σᵢ nᵢ · Σⱼ (Ȳᵢⱼ − Ȳᵢ.. − Ȳ.j. + Ȳ...)²
    SS_AB = 0.0
    for g in between_lvls:
        for t in within_lvls:
            SS_AB += n_g[g] * (cell_mean[(g, t)] - g_mean[g]
                                - t_mean[t] + grand) ** 2

    # ── SS_BS(A)  (within error / residual) ───────────────────────────────────
    # Σᵢ Σₛ Σⱼ (yₛⱼ − Ȳₛ.. − Ȳᵢⱼ + Ȳᵢ..)²
    SS_BSA = 0.0
    for _, row in df.iterrows():
        subj = row["subject"]
        g    = row["between"]
        t    = row["within"]
        SS_BSA += (row["dv"] - s_mean[subj]
                   - cell_mean[(g, t)] + g_mean[g]) ** 2

    SS_Total = ((df["dv"] - grand) ** 2).sum()

    # ── Degrees of freedom (unadjusted) ──────────────────────────────────────
    df_A   = a - 1
    df_SA  = N - a
    df_B   = b - 1
    df_AB  = (a - 1) * (b - 1)
    df_BSA = (b - 1) * (N - a)
    df_Tot = N * b - 1

    # ── Sphericity test ───────────────────────────────────────────────────────
    wide = np.array([
        [df[(df["subject"] == s) & (df["within"] == t)]["dv"].values[0]
         for t in within_lvls]
        for s in df["subject"].unique()
    ], dtype=float)   # shape (N, b)

    sph = None
    eps = 1.0
    sph_label = "None"

    if b > 2:
        sph = mauchly_test(wide)
        if sph_corr == "auto":
            if sph["p"] < alpha:
                eps       = sph["eps_gg"]
                sph_label = "Greenhouse–Geisser"
            else:
                sph_label = "None (sphericity satisfied)"
        elif sph_corr == "gg":
            eps       = sph["eps_gg"];  sph_label = "Greenhouse–Geisser"
        elif sph_corr == "hf":
            eps       = sph["eps_hf"];  sph_label = "Huynh–Feldt"
        elif sph_corr == "lb":
            eps       = sph["eps_lb"];  sph_label = "Lower-bound"
        # sph_corr == "none" → eps stays 1.0

    # ── Corrected df (within-subjects effects only) ───────────────────────────
    df_B_c   = df_B   * eps
    df_AB_c  = df_AB  * eps
    df_BSA_c = df_BSA * eps

    # ── Mean squares ──────────────────────────────────────────────────────────
    MS_A   = SS_A   / df_A      if df_A      > 0 else np.nan
    MS_SA  = SS_SA  / df_SA     if df_SA     > 0 else np.nan
    MS_B   = SS_B   / df_B_c    if df_B_c    > 0 else np.nan
    MS_AB  = SS_AB  / df_AB_c   if df_AB_c   > 0 else np.nan
    MS_BSA = SS_BSA / df_BSA_c  if df_BSA_c  > 0 else np.nan

    # ── F-ratios ──────────────────────────────────────────────────────────────
    def safe_f(ms_eff, ms_err):
        if np.isnan(ms_eff) or np.isnan(ms_err) or ms_err <= 0:
            return np.nan
        return ms_eff / ms_err

    F_A  = safe_f(MS_A,  MS_SA)
    F_B  = safe_f(MS_B,  MS_BSA)
    F_AB = safe_f(MS_AB, MS_BSA)

    def safe_p(F, d1, d2):
        if np.isnan(F): return np.nan
        return float(1.0 - fdist.cdf(F, d1, d2))

    p_A  = safe_p(F_A,  df_A,    df_SA)
    p_B  = safe_p(F_B,  df_B_c,  df_BSA_c)
    p_AB = safe_p(F_AB, df_AB_c, df_BSA_c)

    # ── Effect sizes ──────────────────────────────────────────────────────────
    # Partial η²p = SS_effect / (SS_effect + SS_error_for_that_effect)
    def safe_es(ss_eff, ss_err):
        d = ss_eff + ss_err
        return float(ss_eff / d) if d > 0 else np.nan

    np2_A  = safe_es(SS_A,  SS_SA)
    np2_B  = safe_es(SS_B,  SS_BSA)
    np2_AB = safe_es(SS_AB, SS_BSA)

    # η² (proportion of total variance)
    eta2_A  = SS_A  / SS_Total if SS_Total > 0 else np.nan
    eta2_B  = SS_B  / SS_Total if SS_Total > 0 else np.nan
    eta2_AB = SS_AB / SS_Total if SS_Total > 0 else np.nan

    # Observed power
    pow_A  = obs_power(F_A,  df_A,    df_SA,    alpha)
    pow_B  = obs_power(F_B,  df_B_c,  df_BSA_c, alpha)
    pow_AB = obs_power(F_AB, df_AB_c, df_BSA_c, alpha)

    return dict(
        # Design
        a=a, b=b, N=N,
        between_lvls=between_lvls, within_lvls=within_lvls,
        n_per_group=n_g, grand=grand,
        cell_stats=cell_stats, g_mean=g_mean, t_mean=t_mean,
        # SS
        SS_A=SS_A, SS_SA=SS_SA, SS_B=SS_B, SS_AB=SS_AB,
        SS_BSA=SS_BSA, SS_Total=SS_Total,
        # df
        df_A=df_A, df_SA=df_SA,
        df_B=df_B, df_AB=df_AB, df_BSA=df_BSA, df_Tot=df_Tot,
        df_B_c=df_B_c, df_AB_c=df_AB_c, df_BSA_c=df_BSA_c,
        # MS
        MS_A=MS_A, MS_SA=MS_SA,
        MS_B=MS_B, MS_AB=MS_AB, MS_BSA=MS_BSA,
        # F & p
        F_A=F_A,   F_B=F_B,   F_AB=F_AB,
        p_A=p_A,   p_B=p_B,   p_AB=p_AB,
        # Effect sizes
        np2_A=np2_A,   np2_B=np2_B,   np2_AB=np2_AB,
        eta2_A=eta2_A, eta2_B=eta2_B, eta2_AB=eta2_AB,
        # Power
        pow_A=pow_A, pow_B=pow_B, pow_AB=pow_AB,
        # Sphericity
        sph=sph, eps=eps, sph_label=sph_label,
        alpha=alpha,
    )

# ══════════════════════════════════════════════════════════════════════════════
#  POST-HOC FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def ph_between(df: pd.DataFrame, method: str, alpha: float) -> pd.DataFrame:
    """Welch independent-samples t-tests for between-subjects factor levels."""
    lvls = sorted(df["between"].unique().tolist(), key=str)
    rows = []
    for l1, l2 in itertools.combinations(lvls, 2):
        g1 = df[df["between"] == l1]["dv"].values
        g2 = df[df["between"] == l2]["dv"].values
        t_, p_raw = stats.ttest_ind(g1, g2, equal_var=False)
        md  = g1.mean() - g2.mean()
        psd = np.sqrt(((len(g1)-1)*g1.std(ddof=1)**2 +
                       (len(g2)-1)*g2.std(ddof=1)**2) / (len(g1)+len(g2)-2))
        rows.append(dict(
            Group_1=str(l1), Group_2=str(l2),
            n_1=int(len(g1)), n_2=int(len(g2)),
            Mean_1=round(g1.mean(), 4), Mean_2=round(g2.mean(), 4),
            Mean_Diff=round(md, 4),
            t_Welch=round(t_, 4),
            df_approx=round(len(g1)+len(g2)-2, 1),
            p_uncorrected=p_raw,
            Cohen_d=round(md / psd, 4) if psd > 0 else np.nan,
        ))
    if not rows:
        return pd.DataFrame()
    ph = pd.DataFrame(rows)
    _, p_adj = multipletests_local(ph["p_uncorrected"].tolist(), method, alpha)
    ph["p_corrected"] = [fmt_p(x) for x in p_adj]
    ph["Reject H₀"]   = [x < alpha for x in p_adj]
    ph["p_uncorrected"] = ph["p_uncorrected"].map(fmt_p)
    return ph

def ph_within(df: pd.DataFrame, method: str, alpha: float) -> pd.DataFrame:
    """Paired-samples t-tests for within-subjects factor levels."""
    lvls  = sorted(df["within"].unique().tolist(), key=str)
    pivot = df.pivot_table(index="subject", columns="within", values="dv")
    rows  = []
    for l1, l2 in itertools.combinations(lvls, 2):
        if l1 not in pivot.columns or l2 not in pivot.columns:
            continue
        pair = pivot[[l1, l2]].dropna()
        d    = pair[l1] - pair[l2]
        t_, p_raw = stats.ttest_rel(pair[l1], pair[l2])
        sd_d = d.std(ddof=1)
        rows.append(dict(
            Level_1=str(l1), Level_2=str(l2),
            n=int(len(pair)),
            Mean_1=round(pair[l1].mean(), 4),
            Mean_2=round(pair[l2].mean(), 4),
            Mean_Diff=round(d.mean(), 4),
            SD_Diff=round(sd_d, 4),
            t_paired=round(t_, 4),
            df_t=int(len(pair)-1),
            p_uncorrected=p_raw,
            Cohen_d=round(d.mean() / sd_d, 4) if sd_d > 0 else np.nan,
        ))
    if not rows:
        return pd.DataFrame()
    ph = pd.DataFrame(rows)
    _, p_adj = multipletests_local(ph["p_uncorrected"].tolist(), method, alpha)
    ph["p_corrected"] = [fmt_p(x) for x in p_adj]
    ph["Reject H₀"]   = [x < alpha for x in p_adj]
    ph["p_uncorrected"] = ph["p_uncorrected"].map(fmt_p)
    return ph

def simple_effects(df: pd.DataFrame, method: str, alpha: float) -> pd.DataFrame:
    """
    Simple effects analysis: effect of within-factor at each level of
    between-factor. Paired t-tests with global correction.
    """
    between_lvls = sorted(df["between"].unique().tolist(), key=str)
    within_lvls  = sorted(df["within"].unique().tolist(),  key=str)
    rows = []
    for g in between_lvls:
        sub   = df[df["between"] == g]
        pivot = sub.pivot_table(index="subject", columns="within", values="dv")
        for l1, l2 in itertools.combinations(within_lvls, 2):
            if l1 not in pivot or l2 not in pivot:
                continue
            pair = pivot[[l1, l2]].dropna()
            if len(pair) < 2:
                continue
            d     = pair[l1] - pair[l2]
            t_, p_raw = stats.ttest_rel(pair[l1], pair[l2])
            sd_d  = d.std(ddof=1)
            rows.append(dict(
                Between_Group=str(g),
                Level_1=str(l1), Level_2=str(l2),
                n=int(len(pair)),
                Mean_1=round(pair[l1].mean(), 4),
                Mean_2=round(pair[l2].mean(), 4),
                Mean_Diff=round(d.mean(), 4),
                t_paired=round(t_, 4),
                df_t=int(len(pair)-1),
                p_uncorrected=p_raw,
                Cohen_d=round(d.mean() / sd_d, 4) if sd_d > 0 else np.nan,
            ))
    if not rows:
        return pd.DataFrame()
    se = pd.DataFrame(rows)
    _, p_adj = multipletests_local(se["p_uncorrected"].tolist(), method, alpha)
    se["p_corrected"] = [fmt_p(x) for x in p_adj]
    se["Reject H₀"]   = [x < alpha for x in p_adj]
    se["p_uncorrected"] = se["p_uncorrected"].map(fmt_p)
    return se

# ══════════════════════════════════════════════════════════════════════════════
#  INTERPRETATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def interpret(res: dict, btw: str, win: str, dv: str, pref: str) -> list:
    alpha = res["alpha"]
    texts = []

    def es_text(np2, eta2):
        cf  = cohen_f(np2)
        if pref == "partial_eta2":
            mag = magnitude(np2, "partial_eta2")
            return f"partial \u03b7\u00b2p\u202f=\u202f{np2:.3f} ({mag} effect, Cohen\u202f1988)"
        elif pref == "eta2":
            mag = magnitude(eta2, "eta2")
            return f"\u03b7\u00b2\u202f=\u202f{eta2:.3f} ({mag} effect)"
        else:
            mag = magnitude(cf, "cohen_f")
            return f"Cohen\u2019s\u202ff\u202f=\u202f{cf:.3f} ({mag} effect)"

    # ── Between ───────────────────────────────────────────────────────────────
    sig_A = res["p_A"] < alpha
    es_A  = es_text(res["np2_A"], res["eta2_A"])
    if sig_A:
        texts.append(
            f"<b>Main Effect of {btw} (Between-Subjects):</b> "
            f"A statistically significant main effect of {btw} was found, "
            f"F({res['df_A']:.0f},\u202f{res['df_SA']:.0f})\u202f=\u202f{res['F_A']:.3f}, "
            f"p\u202f{fmt_p(res['p_A'])}, {es_A}; "
            f"observed power\u202f=\u202f{res['pow_A']:.3f}. "
            f"Group means on {dv} differ significantly when averaging across {win} levels. "
            f"Post-hoc pairwise comparisons are recommended to identify which groups differ."
        )
    else:
        texts.append(
            f"<b>Main Effect of {btw} (Between-Subjects):</b> "
            f"The main effect of {btw} was not statistically significant, "
            f"F({res['df_A']:.0f},\u202f{res['df_SA']:.0f})\u202f=\u202f{res['F_A']:.3f}, "
            f"p\u202f{fmt_p(res['p_A'])}, {es_A}; "
            f"observed power\u202f=\u202f{res['pow_A']:.3f}. "
            f"There is insufficient evidence to conclude that {btw} groups differ "
            f"on {dv} when averaging across {win} levels."
        )

    # ── Within ────────────────────────────────────────────────────────────────
    sph_note = ""
    if res["sph"] is not None and res["eps"] < 1.0:
        sph_note = (
            f" Degrees of freedom were adjusted using the "
            f"{res['sph_label']} procedure "
            f"(\u03b5\u202f=\u202f{res['eps']:.4f}), because Mauchly\u2019s test "
            f"indicated a violation of sphericity."
        )
    sig_B = res["p_B"] < alpha
    es_B  = es_text(res["np2_B"], res["eta2_B"])
    if sig_B:
        texts.append(
            f"<b>Main Effect of {win} (Within-Subjects):</b> "
            f"A statistically significant main effect of {win} was found, "
            f"F({res['df_B_c']:.3f},\u202f{res['df_BSA_c']:.3f})\u202f=\u202f{res['F_B']:.3f}, "
            f"p\u202f{fmt_p(res['p_B'])}, {es_B}{sph_note}; "
            f"observed power\u202f=\u202f{res['pow_B']:.3f}. "
            f"Scores on {dv} change significantly across {win} levels when "
            f"averaging across {btw} groups. "
            f"Post-hoc pairwise comparisons are recommended."
        )
    else:
        texts.append(
            f"<b>Main Effect of {win} (Within-Subjects):</b> "
            f"The main effect of {win} was not statistically significant, "
            f"F({res['df_B_c']:.3f},\u202f{res['df_BSA_c']:.3f})\u202f=\u202f{res['F_B']:.3f}, "
            f"p\u202f{fmt_p(res['p_B'])}, {es_B}{sph_note}; "
            f"observed power\u202f=\u202f{res['pow_B']:.3f}. "
            f"No significant change in {dv} was detected across {win} levels."
        )

    # ── Interaction ────────────────────────────────────────────────────────────
    sig_AB = res["p_AB"] < alpha
    es_AB  = es_text(res["np2_AB"], res["eta2_AB"])
    if sig_AB:
        texts.append(
            f"<b>{btw}\u202f\u00d7\u202f{win} Interaction:</b> "
            f"A statistically significant interaction was found, "
            f"F({res['df_AB_c']:.3f},\u202f{res['df_BSA_c']:.3f})\u202f=\u202f{res['F_AB']:.3f}, "
            f"p\u202f{fmt_p(res['p_AB'])}, {es_AB}; "
            f"observed power\u202f=\u202f{res['pow_AB']:.3f}. "
            f"The effect of {win} on {dv} differs across levels of {btw} "
            f"(non-parallel profiles). Examine the profile plot and simple-effects "
            f"analyses to characterize the interaction. "
            f"<b>Caution:</b> main effects should be interpreted in the context of "
            f"this significant interaction."
        )
    else:
        texts.append(
            f"<b>{btw}\u202f\u00d7\u202f{win} Interaction:</b> "
            f"The interaction was not statistically significant, "
            f"F({res['df_AB_c']:.3f},\u202f{res['df_BSA_c']:.3f})\u202f=\u202f{res['F_AB']:.3f}, "
            f"p\u202f{fmt_p(res['p_AB'])}, {es_AB}; "
            f"observed power\u202f=\u202f{res['pow_AB']:.3f}. "
            f"The pattern of change across {win} levels is consistent (parallel) "
            f"across {btw} groups. Main effects may be interpreted independently."
        )

    # ── Power warning ─────────────────────────────────────────────────────────
    low = [nm for nm, pw in [
        (btw, res["pow_A"]),
        (win, res["pow_B"]),
        (f"{btw}\u202f\u00d7\u202f{win}", res["pow_AB"]),
    ] if not np.isnan(pw) and pw < 0.80]
    if low:
        texts.append(
            f"<b>Statistical Power Notice:</b> "
            f"Observed power\u202f&lt;\u202f0.80 for: {', '.join(low)}. "
            f"This elevates Type\u202fII error risk (false negative). "
            f"Increasing sample size is recommended. For a medium effect "
            f"(partial\u202f\u03b7\u00b2p\u202f\u2248\u202f.06), approximately "
            f"20 subjects per group is typically required."
        )
    return texts

# ══════════════════════════════════════════════════════════════════════════════
#  FIGURE BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def build_figure(res, df, btw, win, dv, pal, style):
    sns.set_style(style)
    sns.set_context("notebook", font_scale=1.0)

    b_lvls = res["between_lvls"]
    w_lvls = res["within_lvls"]
    cs     = res["cell_stats"]
    a      = res["a"]; b_n = res["b"]
    colors = sns.color_palette(pal, n_colors=max(a, 3))

    fig = plt.figure(figsize=(20, 13))
    fig.patch.set_facecolor("#f7f9fc")
    gs_main = gridspec.GridSpec(2, 3, figure=fig, hspace=0.46, wspace=0.34)

    # ── 1: Profile Plot ───────────────────────────────────────────────────────
    ax1 = fig.add_subplot(gs_main[0, :2])
    w_order_idx = {v: i for i, v in enumerate(w_lvls)}
    for i, g in enumerate(b_lvls):
        sub = cs[cs["between"] == g].copy()
        sub = sub.sort_values("within", key=lambda x: x.map(w_order_idx))
        ax1.errorbar(
            sub["within"], sub["Mean"], yerr=1.96 * sub["SE"],
            marker="o", ms=7, lw=2.4, capsize=5, capthick=1.8,
            color=colors[i], label=str(g), zorder=4,
        )
    ax1.set_title(f"Profile Plot — {btw} × {win}\n(Error bars: 95% CI of the mean)",
                  fontsize=11.5, fontweight="bold", pad=10)
    ax1.set_xlabel(win, fontsize=10.5)
    ax1.set_ylabel(f"Mean {dv}", fontsize=10.5)
    ax1.legend(title=btw, framealpha=0.9, fontsize=9)
    ax1.grid(True, alpha=0.28, linestyle="--")
    ax1.set_facecolor("#ffffff")

    # ── 2: Grouped Bar Chart ──────────────────────────────────────────────────
    ax2 = fig.add_subplot(gs_main[0, 2])
    w_idx = np.arange(b_n); bw = 0.80 / a
    for i, g in enumerate(b_lvls):
        sub = cs[cs["between"] == g].set_index("within").reindex(w_lvls)
        offset = (i - a / 2.0 + 0.5) * bw
        ax2.bar(w_idx + offset, sub["Mean"], bw * 0.90,
                yerr=1.96 * sub["SE"].values,
                color=colors[i], label=str(g), alpha=0.85,
                capsize=4, error_kw={"elinewidth": 1.5, "ecolor": "#333", "alpha": 0.65})
    ax2.set_xticks(w_idx)
    rot = 22 if b_n > 3 else 0
    ax2.set_xticklabels(w_lvls, rotation=rot, ha="right" if rot else "center")
    ax2.set_title("Cell Means ± 95% CI", fontsize=11.5, fontweight="bold", pad=10)
    ax2.set_xlabel(win, fontsize=10.5)
    ax2.set_ylabel(f"Mean {dv}", fontsize=10.5)
    ax2.legend(title=btw, fontsize=8.5, framealpha=0.9)
    ax2.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax2.set_facecolor("#ffffff")

    # ── 3: Box Plot ───────────────────────────────────────────────────────────
    ax3 = fig.add_subplot(gs_main[1, 0])
    df_b = df.copy()
    df_b["Cell"] = df_b["between"].astype(str) + "\n" + df_b["within"].astype(str)
    c_order = [f"{g}\n{t}" for g in b_lvls for t in w_lvls]
    cmap    = {f"{g}\n{t}": colors[i]
               for i, g in enumerate(b_lvls) for t in w_lvls}
    sns.boxplot(data=df_b, x="Cell", y="dv", order=c_order,
                palette=cmap, ax=ax3, linewidth=1.1,
                flierprops=dict(marker="o", ms=3.5, alpha=0.5))
    ax3.set_title("Distribution per Cell", fontsize=11.5, fontweight="bold", pad=10)
    ax3.set_xlabel(""); ax3.set_ylabel(dv, fontsize=10.5)
    ax3.tick_params(axis="x", labelsize=7.5 if a * b_n > 6 else 9)
    ax3.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax3.set_facecolor("#ffffff")

    # ── 4: Violin Plot ────────────────────────────────────────────────────────
    ax4 = fig.add_subplot(gs_main[1, 1])
    sns.violinplot(data=df, x="within", y="dv", hue="between",
                   order=w_lvls, palette=pal,
                   inner="quartile", ax=ax4, alpha=0.78, linewidth=1.0)
    ax4.set_title("Score Distribution (Violin)", fontsize=11.5, fontweight="bold", pad=10)
    ax4.set_xlabel(win, fontsize=10.5)
    ax4.set_ylabel(dv, fontsize=10.5)
    h, lbl = ax4.get_legend_handles_labels()
    ax4.legend(h[:a], lbl[:a], title=btw, fontsize=8.5, framealpha=0.9)
    ax4.grid(True, alpha=0.28, linestyle="--", axis="y")
    ax4.set_facecolor("#ffffff")

    # ── 5: Effect-Size Summary ────────────────────────────────────────────────
    ax5 = fig.add_subplot(gs_main[1, 2])
    eff_lbls = [btw, win, f"{btw}\n×{win}"]
    eff_vals = [res["np2_A"], res["np2_B"], res["np2_AB"]]
    eff_ps   = [res["p_A"],   res["p_B"],   res["p_AB"]]
    ec       = [colors[0] if p < res["alpha"] else "#9e9e9e" for p in eff_ps]
    bars = ax5.barh(eff_lbls, eff_vals, color=ec, height=0.42, alpha=0.88)
    for xv, lb in [(0.01, "small"), (0.06, "medium"), (0.14, "large")]:
        ax5.axvline(xv, ls="--", lw=1.1, color="#777", alpha=0.7)
        ax5.text(xv, 2.65, lb, ha="center", fontsize=7, color="#555", va="bottom")
    for bar, val, p_ in zip(bars, eff_vals, eff_ps):
        mk = "  *" if p_ < res["alpha"] else "  n.s."
        ax5.text(val + 0.003, bar.get_y() + bar.get_height() / 2,
                 f"{val:.4f}{mk}", va="center", fontsize=9, fontweight="bold")
    ax5.set_title("Partial η²p Effect Sizes\n(* = significant at α)",
                  fontsize=11.5, fontweight="bold", pad=10)
    ax5.set_xlabel("Partial η²p", fontsize=10.5)
    ax5.set_xlim(0, max(max(eff_vals) * 1.38, 0.22))
    ax5.set_facecolor("#ffffff")
    ax5.grid(True, alpha=0.28, linestyle="--", axis="x")

    fig.suptitle(
        f"Mixed ANOVA — {btw} (between) × {win} (within)  |  "
        f"N = {res['N']}  |  DV: {dv}",
        fontsize=12.5, fontweight="bold", y=1.012, color="#0d1b2a")
    return fig

# ══════════════════════════════════════════════════════════════════════════════
#  WORD REPORT BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def _cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def _word_table(doc, df_data: pd.DataFrame, hdr_bg="0d1b2a"):
    r, c = df_data.shape
    tbl  = doc.add_table(rows=r + 1, cols=c)
    tbl.style = "Table Grid"
    for j, col_name in enumerate(df_data.columns):
        cell = tbl.rows[0].cells[j]
        cell.text = str(col_name)
        _cell_bg(cell, hdr_bg)
        run = cell.paragraphs[0].runs[0]
        run.bold = True; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
    for i, (_, row) in enumerate(df_data.iterrows()):
        for j, val in enumerate(row):
            cell = tbl.rows[i + 1].cells[j]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if i % 2 == 0:
                _cell_bg(cell, "edf2fa")
    return tbl

def build_report(res, df, btw, win, dv, alpha, method, pref, fig,
                 df_ph_b, df_ph_w, df_se) -> io.BytesIO:
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = sec.bottom_margin = Inches(1.0)
        sec.left_margin = sec.right_margin = Inches(1.2)

    # Title
    t = doc.add_heading("Mixed ANOVA Analysis Report", 0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.runs[0].font.color.rgb = RGBColor(0x0d, 0x1b, 0x2a)
    doc.add_paragraph(
        f"Two-Way Mixed ANOVA (GLM Repeated Measures)  |  "
        f"Between: {btw}  |  Within: {win}  |  DV: {dv}  |  "
        f"N = {res['N']}  |  \u03b1 = {alpha}"
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # 1. Design
    doc.add_heading("1.  Study Design", level=1)
    doc.add_paragraph(
        f"Design: {res['a']}\u202f(between-subjects)\u202f\u00d7\u202f"
        f"{res['b']}\u202f(within-subjects) mixed factorial ANOVA. "
        f"Between-subjects factor: {btw} "
        f"({', '.join(str(l) for l in res['between_lvls'])}). "
        f"Within-subjects factor: {win} "
        f"({', '.join(str(l) for l in res['within_lvls'])}). "
        f"Total subjects: N\u202f=\u202f{res['N']}. "
        f"Grand mean: {res['grand']:.4f}."
    )
    ng = pd.DataFrame({
        "Group": list(res["n_per_group"].index),
        "N (subjects)": list(res["n_per_group"].values),
    })
    _word_table(doc, ng)
    doc.add_paragraph()

    # 2. Descriptive Statistics
    doc.add_heading("2.  Descriptive Statistics", level=1)
    desc = res["cell_stats"].copy()
    desc.columns = [btw, win, "N", "Mean", "SD", "SE", "95% CI Lower", "95% CI Upper"]
    for c in ["Mean", "SD", "SE", "95% CI Lower", "95% CI Upper"]:
        desc[c] = desc[c].round(4)
    _word_table(doc, desc)
    doc.add_paragraph()

    # 3. Assumption Tests
    doc.add_heading("3.  Assumption Tests", level=1)
    doc.add_heading("3.1  Mauchly's Test of Sphericity", level=2)
    if res["sph"] is None:
        doc.add_paragraph(
            "Mauchly's test is not applicable when the within-subjects factor "
            "has only 2 levels (sphericity is trivially satisfied with 1 df).")
    else:
        s = res["sph"]
        doc.add_paragraph(
            f"Mauchly\u2019s W\u202f=\u202f{s['W']:.4f}, "
            f"\u03c7\u00b2({s['df_m']})\u202f=\u202f{s['chi2']:.4f}, "
            f"p\u202f{fmt_p(s['p'])}. "
            f"Greenhouse\u2013Geisser \u03b5\u202f=\u202f{s['eps_gg']:.4f}; "
            f"Huynh\u2013Feldt \u03b5\u202f=\u202f{s['eps_hf']:.4f}; "
            f"Lower-bound \u03b5\u202f=\u202f{s['eps_lb']:.4f}. "
            + (f"Sphericity violated (p\u202f<\u202f{alpha}); "
               f"{res['sph_label']} correction applied "
               f"(\u03b5\u202f=\u202f{res['eps']:.4f})."
               if s["p"] < alpha
               else "Sphericity assumption satisfied (no correction applied).")
        )
    doc.add_paragraph()

    # 4. ANOVA Table
    doc.add_heading("4.  Mixed ANOVA Summary Table", level=1)
    at_rows = [
        ("BETWEEN SUBJECTS",            "",                   "",                    "",                    "",                    "",                    "",                    "",                True),
        (f"  {btw}",                    fmt_f(res["SS_A"]),   fmt_f(res["df_A"],0),  fmt_f(res["MS_A"]),    fmt_f(res["F_A"]),     fmt_p(res["p_A"]),     fmt_f(res["np2_A"]),   fmt_f(res["pow_A"]),False),
        ("  Error [S(A)]",              fmt_f(res["SS_SA"]),  fmt_f(res["df_SA"],0), fmt_f(res["MS_SA"]),   "—","—","—","—",        False),
        ("WITHIN SUBJECTS",             "",                   "",                    "",                    "",                    "",                    "",                    "",                True),
        (f"  {win}",                    fmt_f(res["SS_B"]),   fmt_f(res["df_B_c"]),  fmt_f(res["MS_B"]),    fmt_f(res["F_B"]),     fmt_p(res["p_B"]),     fmt_f(res["np2_B"]),   fmt_f(res["pow_B"]),False),
        (f"  {btw} \u00d7 {win}",      fmt_f(res["SS_AB"]),  fmt_f(res["df_AB_c"]), fmt_f(res["MS_AB"]),   fmt_f(res["F_AB"]),    fmt_p(res["p_AB"]),    fmt_f(res["np2_AB"]),  fmt_f(res["pow_AB"]),False),
        ("  Error [BS(A)]",             fmt_f(res["SS_BSA"]), fmt_f(res["df_BSA_c"]),fmt_f(res["MS_BSA"]),  "—","—","—","—",        False),
        ("Total",                       fmt_f(res["SS_Total"]),fmt_f(res["df_Tot"],0),"—","—","—","—","—",  False),
    ]
    hdrs = ["Source","SS","df","MS","F","p","Partial η²p","Obs. Power"]
    tbl  = doc.add_table(rows=len(at_rows) + 1, cols=len(hdrs))
    tbl.style = "Table Grid"
    for j, h in enumerate(hdrs):
        c = tbl.rows[0].cells[j]; c.text = h
        _cell_bg(c, "0d1b2a")
        r_ = c.paragraphs[0].runs[0]
        r_.bold = True; r_.font.size = Pt(9)
        r_.font.color.rgb = RGBColor(0xff, 0xff, 0xff)
    for i, row_d in enumerate(at_rows):
        is_hdr = row_d[8]
        for j, val in enumerate(row_d[:8]):
            cell = tbl.rows[i+1].cells[j]; cell.text = str(val)
            run_ = cell.paragraphs[0].runs[0]; run_.font.size = Pt(9)
            if is_hdr:
                run_.bold = True; _cell_bg(cell, "d0dff0")
            elif i % 2 == 0:
                _cell_bg(cell, "edf2fa")
    doc.add_paragraph()
    if res["sph"] and res["eps"] < 1.0:
        note = doc.add_paragraph(
            f"Note. df for within-subjects effects adjusted using "
            f"{res['sph_label']} (\u03b5 = {res['eps']:.4f}).")
        note.runs[0].italic = True
    doc.add_paragraph()

    # 5. Effect Sizes
    doc.add_heading("5.  Effect Sizes", level=1)
    es_df = pd.DataFrame({
        "Source": [btw, win, f"{btw} \u00d7 {win}"],
        "Partial η²p": [round(res["np2_A"],4),  round(res["np2_B"],4),  round(res["np2_AB"],4)],
        "η²":          [round(res["eta2_A"],4), round(res["eta2_B"],4), round(res["eta2_AB"],4)],
        "Cohen's f":   [round(cohen_f(res["np2_A"]),4),
                        round(cohen_f(res["np2_B"]),4),
                        round(cohen_f(res["np2_AB"]),4)],
        "Magnitude":   [magnitude(res["np2_A"],"partial_eta2"),
                        magnitude(res["np2_B"],"partial_eta2"),
                        magnitude(res["np2_AB"],"partial_eta2")],
        "Significant": [res["p_A"]<alpha, res["p_B"]<alpha, res["p_AB"]<alpha],
    })
    _word_table(doc, es_df); doc.add_paragraph()

    # 6. Post-hoc
    cn = {"bonferroni":"Bonferroni","holm":"Holm","sidak":"Šidák",
          "fdr_bh":"Benjamini–Hochberg FDR"}.get(method, method)
    doc.add_heading("6.  Post-Hoc Pairwise Comparisons", level=1)
    doc.add_heading(f"6.1  Between-Subjects: {btw}", level=2)
    doc.add_paragraph(f"Welch's independent-samples t-tests, {cn} correction.")
    if not df_ph_b.empty: _word_table(doc, df_ph_b)
    else: doc.add_paragraph("Only 2 levels — omnibus F-test is conclusive.")
    doc.add_paragraph()
    doc.add_heading(f"6.2  Within-Subjects: {win}", level=2)
    doc.add_paragraph(f"Paired-samples t-tests, {cn} correction.")
    if not df_ph_w.empty: _word_table(doc, df_ph_w)
    else: doc.add_paragraph("Only 2 levels — omnibus F-test is conclusive.")
    doc.add_paragraph()
    doc.add_heading("6.3  Simple Effects (Interaction Decomposition)", level=2)
    doc.add_paragraph(
        f"Effect of {win} at each level of {btw}; "
        f"paired t-tests with {cn} correction applied globally.")
    if not df_se.empty: _word_table(doc, df_se)
    doc.add_paragraph()

    # 7. Interpretation
    doc.add_heading("7.  Statistical Interpretation", level=1)
    for txt in interpret(res, btw, win, dv, pref):
        clean = re.sub(r"<b>(.*?)</b>", r"\1", txt)
        hdr_txt, _, body = clean.partition(":")
        p_ = doc.add_paragraph()
        p_.add_run(hdr_txt + ":").bold = True
        if body: p_.add_run(body)
    doc.add_paragraph()

    # 8. APA 7 Template
    doc.add_heading("8.  APA 7th Edition Reporting Template", level=1)
    doc.add_paragraph(
        f"A {res['a']} ({btw}) \u00d7 {res['b']} ({win}) mixed analysis of variance "
        f"(ANOVA) was conducted with {btw} as the between-subjects factor and "
        f"{win} as the within-subjects repeated-measures factor "
        f"(N\u202f=\u202f{res['N']}). "
        f"The main effect of {btw} was "
        f"{'statistically significant' if res['p_A']<alpha else 'not statistically significant'}, "
        f"F({res['df_A']:.0f},\u202f{res['df_SA']:.0f})\u202f=\u202f{res['F_A']:.2f}, "
        f"p\u202f{fmt_p(res['p_A'])}, "
        f"\u03b7\u00b2p\u202f=\u202f{res['np2_A']:.3f}. "
        f"The main effect of {win} was "
        f"{'statistically significant' if res['p_B']<alpha else 'not statistically significant'}, "
        f"F({res['df_B_c']:.2f},\u202f{res['df_BSA_c']:.2f})\u202f=\u202f{res['F_B']:.2f}, "
        f"p\u202f{fmt_p(res['p_B'])}, "
        f"\u03b7\u00b2p\u202f=\u202f{res['np2_B']:.3f}. "
        f"The {btw}\u00d7{win} interaction was "
        f"{'statistically significant' if res['p_AB']<alpha else 'not statistically significant'}, "
        f"F({res['df_AB_c']:.2f},\u202f{res['df_BSA_c']:.2f})\u202f=\u202f{res['F_AB']:.2f}, "
        f"p\u202f{fmt_p(res['p_AB'])}, "
        f"\u03b7\u00b2p\u202f=\u202f{res['np2_AB']:.3f}."
    )
    doc.add_paragraph()

    # 9. Figures
    doc.add_heading("9.  Figures", level=1)
    img = io.BytesIO()
    fig.savefig(img, format="png", dpi=150, bbox_inches="tight", facecolor="#f7f9fc")
    img.seek(0)
    doc.add_picture(img, width=Inches(6.2))
    cap = doc.add_paragraph(
        f"Figure 1. Mixed ANOVA visualisation panel. "
        f"Profile plot (top-left), grouped bar chart (top-right), "
        f"box plots per cell (bottom-left), violin plots (bottom-centre), "
        f"partial \u03b7\u00b2p effect sizes (bottom-right; filled = significant, grey = n.s.).")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════════════════
#  RUN ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("⏳  Computing Mixed ANOVA …"):
    res = run_mixed_anova(df, alpha_level, sph_corr)

# Compute post-hoc regardless of toggle (needed for download)
df_ph_b = ph_between(df, posthoc_method, alpha_level)
df_ph_w = ph_within(df, posthoc_method, alpha_level)
df_se   = simple_effects(df, posthoc_method, alpha_level)

# ══════════════════════════════════════════════════════════════════════════════
#  OUTPUT — OVERVIEW CARDS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Results Overview</div>', unsafe_allow_html=True)

oc = st.columns(5)
def mcard(col, lbl, val, sub):
    col.markdown(
        f'<div class="mcard"><div class="mc-lbl">{lbl}</div>'
        f'<div class="mc-val">{val}</div>'
        f'<div class="mc-sub">{sub}</div></div>',
        unsafe_allow_html=True)

mcard(oc[0], "Total Subjects (N)", str(res["N"]),
      f"{res['a']} group(s) &nbsp;&bull;&nbsp; {res['b']} time-point(s)")
mcard(oc[1], f"Between — {btw_col}", f"F = {res['F_A']:.3f}",
      f"p {fmt_p(res['p_A'])} &nbsp;&bull;&nbsp; η²p = {res['np2_A']:.3f}")
mcard(oc[2], f"Within — {win_col}", f"F = {res['F_B']:.3f}",
      f"p {fmt_p(res['p_B'])} &nbsp;&bull;&nbsp; η²p = {res['np2_B']:.3f}")
mcard(oc[3], "Interaction", f"F = {res['F_AB']:.3f}",
      f"p {fmt_p(res['p_AB'])} &nbsp;&bull;&nbsp; η²p = {res['np2_AB']:.3f}")
mcard(oc[4], "Grand Mean", f"{res['grand']:.3f}",
      f"SD = {df['dv'].std():.3f}")

# Detailed result cards (one per effect)
st.markdown("")
rc1, rc2, rc3 = st.columns(3)
for col_, lbl, F, d1, d2, p, np2, eta2, pw in [
    (rc1, btw_col,
     res["F_A"],  res["df_A"],    res["df_SA"],    res["p_A"],
     res["np2_A"], res["eta2_A"], res["pow_A"]),
    (rc2, win_col,
     res["F_B"],  res["df_B_c"],  res["df_BSA_c"], res["p_B"],
     res["np2_B"], res["eta2_B"], res["pow_B"]),
    (rc3, f"{btw_col} × {win_col}",
     res["F_AB"], res["df_AB_c"], res["df_BSA_c"],  res["p_AB"],
     res["np2_AB"],res["eta2_AB"],res["pow_AB"]),
]:
    cf  = cohen_f(np2)
    mag = magnitude(np2, "partial_eta2")
    col_.markdown(
        f'<div class="rcard">'
        f'<div class="mc-lbl">{lbl}</div>'
        f'<div class="mc-val">F({d1:.2f}, {d2:.2f}) = {F:.3f}</div>'
        f'<div class="mc-sub">{pill(p, alpha_level)}</div>'
        f'<div class="mc-sub" style="margin-top:5px;">'
        f'Partial η²p = {np2:.4f} ({mag})</div>'
        f'<div class="mc-sub">η² = {eta2:.4f} &nbsp;|&nbsp; '
        f"Cohen's f = {fmt_f(cf)}</div>"
        f'<div class="mc-sub">Observed power = {pw:.3f}</div>'
        f'</div>',
        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DESCRIPTIVE STATISTICS  (rendered as dataframes, never raw markdown tables)
# ══════════════════════════════════════════════════════════════════════════════
if show_desc:
    st.markdown('<div class="sec">Descriptive Statistics</div>', unsafe_allow_html=True)
    d_tab1, d_tab2, d_tab3 = st.tabs([
        "Cell Statistics",
        f"Marginal Means — {btw_col}",
        f"Marginal Means — {win_col}",
    ])

    with d_tab1:
        disp = res["cell_stats"].copy()
        disp.columns = [btw_col, win_col, "N", "Mean", "SD", "SE",
                        "95% CI Lower", "95% CI Upper"]
        for c in ["Mean","SD","SE","95% CI Lower","95% CI Upper"]:
            disp[c] = disp[c].round(4)
        st.dataframe(disp, hide_index=True, use_container_width=True)
        st.caption("95% CI = Mean ± 1.96 × SE (cell-level standard error).")

    with d_tab2:
        mm_b = df.groupby("between")["dv"].agg(
            N="count", Mean="mean", SD="std",
            SE=lambda x: x.std(ddof=1) / np.sqrt(len(x)),
            Median="median", Min="min", Max="max",
        ).reset_index()
        mm_b.columns = [btw_col, "N", "Mean", "SD", "SE", "Median", "Min", "Max"]
        for c in ["Mean","SD","SE","Median","Min","Max"]:
            mm_b[c] = mm_b[c].round(4)
        st.dataframe(mm_b, hide_index=True, use_container_width=True)

    with d_tab3:
        mm_w = df.groupby("within")["dv"].agg(
            N="count", Mean="mean", SD="std",
            SE=lambda x: x.std(ddof=1) / np.sqrt(len(x)),
            Median="median", Min="min", Max="max",
        ).reset_index()
        mm_w.columns = [win_col, "N", "Mean", "SD", "SE", "Median", "Min", "Max"]
        for c in ["Mean","SD","SE","Median","Min","Max"]:
            mm_w[c] = mm_w[c].round(4)
        st.dataframe(mm_w, hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ASSUMPTION DIAGNOSTICS  (all results as st.dataframe, never raw markdown)
# ══════════════════════════════════════════════════════════════════════════════
if show_assump:
    st.markdown('<div class="sec">Assumption Diagnostics</div>', unsafe_allow_html=True)
    a_tab1, a_tab2, a_tab3 = st.tabs([
        "Normality — Shapiro–Wilk",
        "Homogeneity of Variance — Levene",
        "Sphericity — Mauchly",
    ])

    with a_tab1:
        st.markdown(
            "Shapiro–Wilk normality test applied to each cell. "
            "A non-significant result (p > α) is consistent with normality. "
            "Mixed ANOVA is robust to minor deviations when n ≥ 20 per cell "
            "(central limit theorem)."
        )
        sw_rows = []
        for g in res["between_lvls"]:
            for t in res["within_lvls"]:
                vals = df[(df["between"] == g) & (df["within"] == t)]["dv"].values
                if len(vals) < 3: continue
                W_sw, p_sw = stats.shapiro(vals)
                sw_rows.append({
                    btw_col: g, win_col: t,
                    "n": len(vals),
                    "Shapiro-Wilk W": round(W_sw, 4),
                    "p-value": fmt_p(p_sw),
                    "Normal (p > α)": "Yes" if p_sw > alpha_level else "No",
                })
        st.dataframe(pd.DataFrame(sw_rows), hide_index=True, use_container_width=True)

    with a_tab2:
        st.markdown(
            "Levene's test (center = mean) for equality of error variances across "
            "between-subjects groups at each within-subjects level. "
            "A non-significant result supports the assumption of homogeneity."
        )
        lev_rows = []
        for t in res["within_lvls"]:
            gdata = [df[(df["between"] == g) & (df["within"] == t)]["dv"].values
                     for g in res["between_lvls"]]
            if all(len(g_) >= 2 for g_ in gdata):
                F_lev, p_lev = stats.levene(*gdata, center="mean")
                lev_rows.append({
                    win_col: t,
                    "Levene F": round(F_lev, 4),
                    "df1": res["a"] - 1,
                    "df2": res["N"] - res["a"],
                    "p-value": fmt_p(p_lev),
                    "Homogeneous (p > α)": "Yes" if p_lev > alpha_level else "No",
                })
        st.dataframe(pd.DataFrame(lev_rows), hide_index=True, use_container_width=True)

    with a_tab3:
        if res["sph"] is None:
            st.markdown(
                '<div class="abox-info">Mauchly\'s test is not applicable when '
                'the within-subjects factor has only 2 levels — sphericity is '
                'trivially satisfied (only 1 df).</div>',
                unsafe_allow_html=True)
        else:
            s = res["sph"]
            sph_df = pd.DataFrame({
                "Statistic": [
                    "Mauchly's W",
                    "Chi-square (χ²) approximation",
                    "Degrees of freedom",
                    "p-value",
                    "Greenhouse–Geisser ε",
                    "Huynh–Feldt ε",
                    "Lower-bound ε",
                    "Correction applied",
                    "ε used",
                ],
                "Value": [
                    f"{s['W']:.4f}",
                    f"{s['chi2']:.4f}",
                    str(s["df_m"]),
                    fmt_p(s["p"]),
                    f"{s['eps_gg']:.4f}",
                    f"{s['eps_hf']:.4f}",
                    f"{s['eps_lb']:.4f}",
                    res["sph_label"],
                    f"{res['eps']:.4f}",
                ],
            })
            st.dataframe(sph_df, hide_index=True, use_container_width=True)

            if s["p"] < alpha_level:
                st.markdown(
                    f'<div class="abox-warn"><b>Sphericity assumption violated</b> — '
                    f'Mauchly\'s W\u202f=\u202f{s["W"]:.4f}, '
                    f'p\u202f{fmt_p(s["p"])}\u202f&lt;\u202f{alpha_level}. '
                    f'Degrees of freedom for within-subjects effects have been corrected '
                    f'using the {res["sph_label"]} procedure '
                    f'(&epsilon;\u202f=\u202f{res["eps"]:.4f}).</div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="abox-ok"><b>Sphericity assumption satisfied</b> — '
                    f'Mauchly\'s W\u202f=\u202f{s["W"]:.4f}, '
                    f'p\u202f{fmt_p(s["p"])}\u202f&ge;\u202f{alpha_level}. '
                    f'No correction to degrees of freedom is required.</div>',
                    unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ANOVA SUMMARY TABLE  (st.dataframe — no raw markdown)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Mixed ANOVA Summary Table</div>', unsafe_allow_html=True)

at_rows_disp = [
    {"Source": "BETWEEN SUBJECTS",         "SS":"","df":"","MS":"","F":"","p":"","Partial η²p":"","Obs. Power":""},
    {"Source": f"  {btw_col}",             "SS":fmt_f(res["SS_A"]),  "df":fmt_f(res["df_A"],0),
     "MS":fmt_f(res["MS_A"]),   "F":fmt_f(res["F_A"]),   "p":fmt_p(res["p_A"]),
     "Partial η²p":fmt_f(res["np2_A"]),   "Obs. Power":fmt_f(res["pow_A"])},
    {"Source": "  Error [S(A)]",           "SS":fmt_f(res["SS_SA"]), "df":fmt_f(res["df_SA"],0),
     "MS":fmt_f(res["MS_SA"]),  "F":"—","p":"—","Partial η²p":"—","Obs. Power":"—"},
    {"Source": "WITHIN SUBJECTS",          "SS":"","df":"","MS":"","F":"","p":"","Partial η²p":"","Obs. Power":""},
    {"Source": f"  {win_col}",             "SS":fmt_f(res["SS_B"]),  "df":fmt_f(res["df_B_c"]),
     "MS":fmt_f(res["MS_B"]),   "F":fmt_f(res["F_B"]),   "p":fmt_p(res["p_B"]),
     "Partial η²p":fmt_f(res["np2_B"]),   "Obs. Power":fmt_f(res["pow_B"])},
    {"Source": f"  {btw_col} × {win_col}","SS":fmt_f(res["SS_AB"]), "df":fmt_f(res["df_AB_c"]),
     "MS":fmt_f(res["MS_AB"]),  "F":fmt_f(res["F_AB"]),  "p":fmt_p(res["p_AB"]),
     "Partial η²p":fmt_f(res["np2_AB"]),  "Obs. Power":fmt_f(res["pow_AB"])},
    {"Source": "  Error [BS(A)]",          "SS":fmt_f(res["SS_BSA"]),"df":fmt_f(res["df_BSA_c"]),
     "MS":fmt_f(res["MS_BSA"]), "F":"—","p":"—","Partial η²p":"—","Obs. Power":"—"},
    {"Source": "Total",                    "SS":fmt_f(res["SS_Total"]),"df":fmt_f(res["df_Tot"],0),
     "MS":"—","F":"—","p":"—","Partial η²p":"—","Obs. Power":"—"},
]
at_disp = pd.DataFrame(at_rows_disp)
st.dataframe(at_disp, hide_index=True, use_container_width=True)

if res["sph"] and res["eps"] < 1.0:
    st.caption(
        f"Note. df for within-subjects effects adjusted using "
        f"{res['sph_label']} correction (ε = {res['eps']:.4f}). "
        f"S(A) = Subjects within Groups (between error); "
        f"BS(A) = B × Subjects within Groups (within error)."
    )
else:
    st.caption(
        "Error terms: S(A) = Subjects within Groups (between-subjects); "
        "BS(A) = B × Subjects within Groups (within-subjects)."
    )

# ── Effect Size Table ─────────────────────────────────────────────────────────
st.markdown('<div class="sec">Effect Size Summary</div>', unsafe_allow_html=True)

es_disp = pd.DataFrame({
    "Source":        [btw_col, win_col, f"{btw_col} × {win_col}"],
    "Partial η²p":   [round(res["np2_A"],4),  round(res["np2_B"],4),  round(res["np2_AB"],4)],
    "η² (total)":    [round(res["eta2_A"],4), round(res["eta2_B"],4), round(res["eta2_AB"],4)],
    "Cohen's f":     [round(cohen_f(res["np2_A"]),4),
                      round(cohen_f(res["np2_B"]),4),
                      round(cohen_f(res["np2_AB"]),4)],
    "Magnitude":     [magnitude(res["np2_A"],"partial_eta2"),
                      magnitude(res["np2_B"],"partial_eta2"),
                      magnitude(res["np2_AB"],"partial_eta2")],
    "F":             [fmt_f(res["F_A"]), fmt_f(res["F_B"]), fmt_f(res["F_AB"])],
    "p":             [fmt_p(res["p_A"]), fmt_p(res["p_B"]), fmt_p(res["p_AB"])],
    "Significant":   [res["p_A"] < alpha_level,
                      res["p_B"] < alpha_level,
                      res["p_AB"] < alpha_level],
})
st.dataframe(es_disp, hide_index=True, use_container_width=True)
st.caption(
    "Cohen (1988) benchmarks for partial η²p: negligible < .01; "
    "small .01–.05; medium .06–.13; large ≥ .14. "
    "Partial η²p is the SPSS default effect size for mixed ANOVA."
)

# ══════════════════════════════════════════════════════════════════════════════
#  POST-HOC COMPARISONS
# ══════════════════════════════════════════════════════════════════════════════
if show_posthoc:
    st.markdown('<div class="sec">Post-Hoc Pairwise Comparisons</div>', unsafe_allow_html=True)

    corr_lbl = {"bonferroni":"Bonferroni","holm":"Holm (step-down Bonferroni)",
                "sidak":"Šidák","fdr_bh":"Benjamini–Hochberg FDR"}.get(
                    posthoc_method, posthoc_method)

    ph_t1, ph_t2, ph_t3 = st.tabs([
        f"Between — {btw_col}",
        f"Within — {win_col}",
        "Simple Effects (Interaction)",
    ])

    with ph_t1:
        st.markdown(
            f"**Welch's independent-samples t-tests** (corrects for potential "
            f"heterogeneity of variance). "
            f"p-values adjusted using **{corr_lbl}**."
        )
        if res["a"] < 3:
            st.markdown(
                '<div class="abox-info">Only 2 levels — the omnibus F-test '
                'directly tests this comparison. No further correction needed.</div>',
                unsafe_allow_html=True)
        elif df_ph_b.empty:
            st.warning("No comparisons could be computed.")
        else:
            st.dataframe(df_ph_b, hide_index=True, use_container_width=True)
            st.caption(
                "Cohen's d benchmarks: |d| < .20 = negligible; "
                ".20–.49 = small; .50–.79 = medium; ≥ .80 = large.")

    with ph_t2:
        st.markdown(
            f"**Paired-samples t-tests** (accounts for within-subject correlation). "
            f"p-values adjusted using **{corr_lbl}**."
        )
        if res["b"] < 3:
            st.markdown(
                '<div class="abox-info">Only 2 levels — the omnibus F-test '
                'directly tests this comparison.</div>',
                unsafe_allow_html=True)
        elif df_ph_w.empty:
            st.warning("No comparisons could be computed.")
        else:
            st.dataframe(df_ph_w, hide_index=True, use_container_width=True)

    with ph_t3:
        st.markdown(
            f"Effect of **{win_col}** at each level of **{btw_col}** — "
            f"decomposes the {btw_col} × {win_col} interaction. "
            f"Paired t-tests; **{corr_lbl}** correction applied globally "
            f"across all comparisons."
        )
        if df_se.empty:
            st.warning("Simple effects could not be computed.")
        else:
            st.dataframe(df_se, hide_index=True, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  INTERPRETATION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Statistical Interpretation</div>', unsafe_allow_html=True)

for txt in interpret(res, btw_col, win_col, dv_col, effect_pref):
    st.markdown(f'<div class="ibox">{txt}</div>', unsafe_allow_html=True)

with st.expander("APA 7th Edition Reporting Template", expanded=False):
    apa = (
        f"A {res['a']} ({btw_col}) \u00d7 {res['b']} ({win_col}) mixed analysis of "
        f"variance (ANOVA) was conducted with {btw_col} as the between-subjects factor "
        f"and {win_col} as the within-subjects repeated-measures factor "
        f"(N = {res['N']}). "
        f"The main effect of {btw_col} was "
        f"{'statistically significant' if res['p_A']<alpha_level else 'not statistically significant'}, "
        f"F({res['df_A']:.0f}, {res['df_SA']:.0f}) = {res['F_A']:.2f}, "
        f"p {fmt_p(res['p_A'])}, partial \u03b7\u00b2p = {res['np2_A']:.3f}. "
        f"The main effect of {win_col} was "
        f"{'statistically significant' if res['p_B']<alpha_level else 'not statistically significant'}, "
        f"F({res['df_B_c']:.2f}, {res['df_BSA_c']:.2f}) = {res['F_B']:.2f}, "
        f"p {fmt_p(res['p_B'])}, partial \u03b7\u00b2p = {res['np2_B']:.3f}. "
        f"The {btw_col} \u00d7 {win_col} interaction was "
        f"{'statistically significant' if res['p_AB']<alpha_level else 'not statistically significant'}, "
        f"F({res['df_AB_c']:.2f}, {res['df_BSA_c']:.2f}) = {res['F_AB']:.2f}, "
        f"p {fmt_p(res['p_AB'])}, partial \u03b7\u00b2p = {res['np2_AB']:.3f}."
    )
    st.text_area("Copy the APA-formatted result:", value=apa, height=160)

# ══════════════════════════════════════════════════════════════════════════════
#  FIGURES
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Visualisations</div>', unsafe_allow_html=True)

with st.spinner("Rendering plots …"):
    fig = build_figure(res, df, btw_col, win_col, dv_col, pal_name, grid_style)

st.pyplot(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DOWNLOADS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec">Download Results</div>', unsafe_allow_html=True)

dl1, dl2, dl3, dl4 = st.columns(4)

with dl1:
    st.download_button(
        "📄  ANOVA Table (CSV)",
        data=at_disp.to_csv(index=False).encode("utf-8"),
        file_name="mixed_anova_table.csv", mime="text/csv",
        use_container_width=True)

with dl2:
    full_csv = pd.DataFrame({
        "Effect":           [btw_col, win_col, f"{btw_col}×{win_col}"],
        "SS":               [res["SS_A"],   res["SS_B"],   res["SS_AB"]],
        "df_effect":        [res["df_A"],   res["df_B_c"], res["df_AB_c"]],
        "df_error":         [res["df_SA"],  res["df_BSA_c"],res["df_BSA_c"]],
        "MS_effect":        [res["MS_A"],   res["MS_B"],   res["MS_AB"]],
        "MS_error":         [res["MS_SA"],  res["MS_BSA"], res["MS_BSA"]],
        "F":                [res["F_A"],    res["F_B"],    res["F_AB"]],
        "p":                [res["p_A"],    res["p_B"],    res["p_AB"]],
        "partial_eta2p":    [res["np2_A"],  res["np2_B"],  res["np2_AB"]],
        "eta2":             [res["eta2_A"], res["eta2_B"], res["eta2_AB"]],
        "cohens_f":         [cohen_f(res["np2_A"]),
                             cohen_f(res["np2_B"]),
                             cohen_f(res["np2_AB"])],
        "observed_power":   [res["pow_A"],  res["pow_B"],  res["pow_AB"]],
        "epsilon":          [1.0, res["eps"], res["eps"]],
        "sph_correction":   ["N/A", res["sph_label"], res["sph_label"]],
    })
    st.download_button(
        "📊  Full Statistics (CSV)",
        data=full_csv.to_csv(index=False).encode("utf-8"),
        file_name="mixed_anova_full_results.csv", mime="text/csv",
        use_container_width=True)

with dl3:
    fig_buf = io.BytesIO()
    fig.savefig(fig_buf, format="png", dpi=200,
                bbox_inches="tight", facecolor="#f7f9fc")
    fig_buf.seek(0)
    st.download_button(
        "🖼️  Figures (PNG, 200 dpi)",
        data=fig_buf.getvalue(),
        file_name="mixed_anova_figures.png", mime="image/png",
        use_container_width=True)

with dl4:
    with st.spinner("Building Word report …"):
        word_buf = build_report(
            res, df, btw_col, win_col, dv_col,
            alpha_level, posthoc_method, effect_pref, fig,
            df_ph_b, df_ph_w, df_se,
        )
    st.download_button(
        "📝  Full Report (Word .docx)",
        data=word_buf.getvalue(),
        file_name="mixed_anova_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "Mixed ANOVA Calculator  ·  SPSS GLM Repeated Measures Equivalent  ·  "
    "SS decomposition verified: SS_A + SS_S(A) + SS_B + SS_AB + SS_BS(A) = SS_Total  ·  "
    "References: Winer, Brown & Michels (1991); Mauchly (1940); Box (1954); "
    "Greenhouse & Geisser (1959); Huynh & Feldt (1976); Lecoutre (1991); Cohen (1988); "
    "Holm (1979); Benjamini & Hochberg (1995)."
)
