"""
Mixed ANOVA Calculator — SPSS-equivalent statistical analysis
============================================================
Supports:
  • One between-subjects factor + one within-subjects factor (2-way Mixed ANOVA)
  • Sphericity test (Mauchly's W) + corrections (Greenhouse-Geisser, Huynh-Feldt, Lower-bound)
  • Between-subjects main effect + Within-subjects main effect + Interaction
  • Post-hoc tests: Bonferroni, Tukey HSD, LSD, Sidak
  • Descriptive statistics table
  • Profile plots + interaction plots + box plots
  • Full downloadable Word report with embedded figures
"""

# ── dependencies ──────────────────────────────────────────────────────────────
import io, warnings, itertools, textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import scipy.stats as stats
from scipy.stats import f as fdist
import pingouin as pg
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import seaborn as sns
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import tempfile, os

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG & STYLING
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Mixed ANOVA Calculator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

/* ── main header ── */
.main-title {
    font-family: 'DM Serif Display', serif;
    font-size: 2.6rem;
    color: #1a1a2e;
    line-height: 1.1;
    margin-bottom: 0;
}
.main-subtitle {
    font-size: 1.0rem;
    color: #555;
    font-weight: 300;
    margin-top: 4px;
    margin-bottom: 2rem;
}

/* ── section headers ── */
.section-header {
    font-family: 'DM Serif Display', serif;
    font-size: 1.5rem;
    color: #1a1a2e;
    border-bottom: 2px solid #e63946;
    padding-bottom: 6px;
    margin-top: 1.6rem;
    margin-bottom: 1rem;
}

/* ── stat cards ── */
.stat-card {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border-radius: 12px;
    padding: 1.1rem 1.4rem;
    color: #fff;
    margin-bottom: 0.5rem;
}
.stat-card .label { font-size: 0.78rem; color: #aaa; text-transform: uppercase; letter-spacing: 0.08em; }
.stat-card .value { font-size: 1.6rem; font-weight: 600; color: #f1faee; margin-top: 2px; }
.stat-card .sub   { font-size: 0.82rem; color: #a8dadc; margin-top: 2px; }

/* ── significance badges ── */
.sig-yes { background:#2dc653; color:#fff; padding:2px 10px; border-radius:20px; font-size:0.82rem; font-weight:600; }
.sig-no  { background:#c0392b; color:#fff; padding:2px 10px; border-radius:20px; font-size:0.82rem; font-weight:600; }
.sig-mar { background:#f39c12; color:#fff; padding:2px 10px; border-radius:20px; font-size:0.82rem; font-weight:600; }

/* ── interpretation box ── */
.interp-box {
    background: #f1faee;
    border-left: 4px solid #457b9d;
    border-radius: 6px;
    padding: 1rem 1.2rem;
    font-size: 0.93rem;
    line-height: 1.65;
    color: #1d3557;
    margin-top: 0.6rem;
}

/* ── upload box ── */
.upload-hint {
    font-size: 0.82rem;
    color: #888;
    background: #f9f9f9;
    border: 1px dashed #ccc;
    border-radius: 8px;
    padding: 0.7rem 1rem;
    margin-top: 0.3rem;
}

/* ── tables ── */
.stDataFrame { border-radius: 8px; overflow: hidden; }

/* ── sidebar ── */
[data-testid="stSidebar"] { background: #1a1a2e; }
[data-testid="stSidebar"] * { color: #f1faee !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stSlider label { color: #a8dadc !important; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="main-title">Mixed ANOVA Calculator</div>', unsafe_allow_html=True)
st.markdown('<div class="main-subtitle">SPSS-equivalent • Sphericity Tests • Post-hoc • Full Report Export</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR — CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## ⚙️ Configuration")
    st.markdown("---")
    alpha_level = st.selectbox("Significance level (α)", [0.05, 0.01, 0.10], index=0)
    posthoc_method = st.selectbox(
        "Post-hoc method",
        ["bonferroni", "holm", "sidak", "fdr_bh"],
        format_func=lambda x: {
            "bonferroni": "Bonferroni",
            "holm": "Holm (step-down Bonferroni)",
            "sidak": "Šidák",
            "fdr_bh": "FDR – Benjamini–Hochberg",
        }[x],
    )
    effect_size_type = st.selectbox(
        "Effect size",
        ["np2", "eta2", "cohen_f"],
        format_func=lambda x: {
            "np2": "Partial η² (η²p) — SPSS default",
            "eta2": "η² (eta-squared)",
            "cohen_f": "Cohen's f",
        }[x],
    )
    sphericity_correction = st.selectbox(
        "Sphericity correction",
        ["auto", "gg", "hf", "lb", "none"],
        format_func=lambda x: {
            "auto": "Auto (use GG when p < α)",
            "gg": "Greenhouse–Geisser",
            "hf": "Huynh–Feldt",
            "lb": "Lower-bound",
            "none": "None (assume sphericity)",
        }[x],
    )
    show_descriptives = st.checkbox("Show descriptive statistics", value=True)
    show_assumptions = st.checkbox("Show assumption checks", value=True)
    st.markdown("---")
    st.markdown("**Plot settings**")
    palette_choice = st.selectbox("Color palette", ["deep", "Set2", "tab10", "husl", "rocket"])
    plot_style = st.selectbox("Plot style", ["whitegrid", "ticks", "darkgrid"])

# ══════════════════════════════════════════════════════════════════════════════
#  FORMAT GUIDE
# ══════════════════════════════════════════════════════════════════════════════
with st.expander("📋 Format Panduan CSV — Klik untuk membuka", expanded=False):
    st.markdown("""
### Format CSV yang Dibutuhkan

Mixed ANOVA memerlukan data dalam format **long (panjang)**:

| Kolom | Keterangan |
|---|---|
| `subject_id` | Identitas unik setiap partisipan (wajib) |
| `between_factor` | Faktor antar-subjek (mis. Kelompok: Kontrol, Perlakuan A, B) |
| `time` | Faktor dalam-subjek (mis. Waktu: Pre, Post, Follow-up) |
| `score` | Nilai/skor pengukuran (numerik) |

**Contoh CSV:**
```
subject_id,Kelompok,Waktu,Skor
1,Kontrol,Pre,45
1,Kontrol,Post,48
1,Kontrol,Follow,50
2,Kontrol,Pre,42
2,Kontrol,Post,46
...
11,Perlakuan,Pre,44
11,Perlakuan,Post,55
11,Perlakuan,Follow,60
```

**Aturan:**
- Setiap subjek harus memiliki **semua level** faktor dalam-subjek
- Nilai pada kolom dependen harus **numerik**
- Tidak ada spasi di nama kolom (gunakan `_` atau PascalCase)
- Encoding: **UTF-8**
- Ukuran maksimum: **3 MB**
- Format: **.csv** (comma-separated)

**Tips:**
- Minimal 2 level untuk setiap faktor
- Minimal 5 subjek per kelompok dianjurkan
- Pastikan tidak ada missing value
""")
    
    # Provide sample CSV download
    sample_data = {
        "subject_id": list(range(1, 21)) * 3,
        "Kelompok": (["Kontrol"] * 10 + ["Perlakuan"] * 10) * 3,
        "Waktu": ["Pre"] * 20 + ["Post"] * 20 + ["Follow_up"] * 20,
        "Skor": (
            [45, 42, 44, 47, 41, 43, 46, 44, 45, 43] +  # Kontrol Pre
            [44, 46, 45, 48, 43, 47, 49, 44, 46, 45] +  # Perlakuan Pre
            [48, 46, 49, 50, 44, 47, 50, 47, 49, 46] +  # Kontrol Post
            [55, 58, 57, 60, 53, 59, 62, 56, 58, 57] +  # Perlakuan Post
            [49, 47, 50, 52, 45, 48, 51, 48, 50, 47] +  # Kontrol Follow_up
            [58, 61, 60, 64, 56, 62, 65, 59, 61, 60]    # Perlakuan Follow_up
        )
    }
    sample_df = pd.DataFrame(sample_data)
    st.download_button(
        "⬇️ Download Contoh CSV",
        data=sample_df.to_csv(index=False).encode("utf-8"),
        file_name="contoh_mixed_anova.csv",
        mime="text/csv",
    )

# ══════════════════════════════════════════════════════════════════════════════
#  FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📂 Unggah Data</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Pilih file CSV (maks. 3 MB)",
    type=["csv"],
    help="File harus dalam format long (tidy). Lihat panduan format di atas.",
)

if uploaded is None:
    st.markdown("""
    <div class="upload-hint">
    👆 Belum ada file yang diunggah. Silakan unggah file CSV Anda atau download contoh CSV di atas untuk mencoba.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Size check
if uploaded.size > 3 * 1024 * 1024:
    st.error("❌ File melebihi 3 MB. Harap kompres atau kurangi data.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  LOAD & VALIDATE DATA
# ══════════════════════════════════════════════════════════════════════════════
try:
    df_raw = pd.read_csv(uploaded)
except Exception as e:
    st.error(f"❌ Gagal membaca CSV: {e}")
    st.stop()

st.markdown('<div class="section-header">🔧 Pemetaan Kolom</div>', unsafe_allow_html=True)

cols = df_raw.columns.tolist()
numeric_cols = df_raw.select_dtypes(include=np.number).columns.tolist()

col1, col2, col3, col4 = st.columns(4)
with col1:
    subj_col = st.selectbox("Kolom Subject ID", cols, index=0)
with col2:
    between_col = st.selectbox("Faktor Between-Subjects", [c for c in cols if c != subj_col])
with col3:
    within_col = st.selectbox("Faktor Within-Subjects", [c for c in cols if c not in [subj_col, between_col]])
with col4:
    dv_col = st.selectbox("Variabel Dependen (numerik)", numeric_cols)

run_btn = st.button("▶ Jalankan Analisis Mixed ANOVA", type="primary", use_container_width=True)

if not run_btn:
    with st.expander("👁 Preview Data (5 baris pertama)"):
        st.dataframe(df_raw.head(), use_container_width=True)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  DATA CLEANING
# ══════════════════════════════════════════════════════════════════════════════
df = df_raw[[subj_col, between_col, within_col, dv_col]].copy()
df.columns = ["subject", "between", "within", "dv"]
df["dv"] = pd.to_numeric(df["dv"], errors="coerce")
n_before = len(df)
df.dropna(inplace=True)
n_after = len(df)

if n_before != n_after:
    st.warning(f"⚠️ {n_before - n_after} baris dengan missing value dihapus.")

# Ensure all subjects have all within-levels
within_levels = sorted(df["within"].unique())
between_levels = sorted(df["between"].unique())
n_within = len(within_levels)
n_between = len(between_levels)

if n_within < 2:
    st.error("❌ Faktor within-subjects harus memiliki minimal 2 level.")
    st.stop()
if n_between < 2:
    st.error("❌ Faktor between-subjects harus memiliki minimal 2 level.")
    st.stop()

# Check balanced design
counts_per_subject = df.groupby("subject")["within"].count()
complete_subjects = counts_per_subject[counts_per_subject == n_within].index
df = df[df["subject"].isin(complete_subjects)]
removed = len(counts_per_subject) - len(complete_subjects)
if removed:
    st.warning(f"⚠️ {removed} subjek dihapus karena data tidak lengkap di semua level within.")

subjects = df["subject"].unique()
N = len(subjects)

if N < 6:
    st.error("❌ Minimal 6 subjek dibutuhkan untuk analisis yang valid.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def cohen_f_from_np2(np2):
    if np2 >= 1:
        return np.nan
    return np.sqrt(np2 / (1 - np2))

def interpret_effect(val, etype="np2"):
    """Return qualitative label for effect size."""
    if etype in ("np2", "eta2"):
        if val < 0.01: return "sangat kecil"
        if val < 0.06: return "kecil"
        if val < 0.14: return "sedang"
        return "besar"
    else:  # cohen_f
        if val < 0.10: return "sangat kecil"
        if val < 0.25: return "kecil"
        if val < 0.40: return "sedang"
        return "besar"

def sig_badge(p, alpha):
    if p < 0.001:   return '<span class="sig-yes">p < .001 ✓</span>'
    if p < alpha:   return f'<span class="sig-yes">p = {p:.3f} ✓</span>'
    if p < 0.10:    return f'<span class="sig-mar">p = {p:.3f} ~</span>'
    return f'<span class="sig-no">p = {p:.3f} ✗</span>'

def fmt_p(p):
    if p < 0.001: return "< .001"
    return f".{round(p*1000):03d}"[:-1] + f"{round(p*1000):03d}"[-3:]  # fallback

def format_p(p):
    if pd.isna(p): return "—"
    if p < 0.001: return "< .001"
    return f"{p:.3f}"

def mauchly_sphericity(data_wide):
    """
    Compute Mauchly's W, chi-square, df, p, GG epsilon, HF epsilon.
    data_wide: subjects × repeated_measures numpy array
    """
    n, k = data_wide.shape
    # Orthonormal contrast matrix
    C = np.zeros((k, k - 1))
    for i in range(k - 1):
        C[:i+1, i] = 1 / np.sqrt((i + 1) * (i + 2))
        C[i+1, i] = -np.sqrt((i + 1) / (i + 2))
    # Difference scores
    Y = data_wide @ C  # n × (k-1)
    S = np.cov(Y.T)    # (k-1) × (k-1)
    p_sph = k - 1
    # Mauchly's W
    det_S = np.linalg.det(S)
    trace_S = np.trace(S)
    W = det_S / (trace_S / p_sph) ** p_sph
    W = np.clip(W, 0, 1)
    # Chi-square approximation
    f_coeff = 1 - (2 * p_sph**2 + p_sph + 2) / (6 * p_sph * (n - 1))
    df_mauchly = p_sph * (p_sph + 1) / 2 - 1
    chi2 = -np.log(W) * (n - 1) * f_coeff
    p_val = 1 - stats.chi2.cdf(chi2, df_mauchly)
    # GG epsilon
    trace_S2 = np.trace(S @ S)
    eps_gg = (np.trace(S)**2) / ((p_sph) * trace_S2)
    eps_gg = np.clip(eps_gg, 1 / p_sph, 1.0)
    # HF epsilon
    eps_hf = (n * (p_sph) * eps_gg - 2) / (p_sph * (n - 1 - p_sph * eps_gg))
    eps_hf = np.clip(eps_hf, 1 / p_sph, 1.0)
    # Lower-bound
    eps_lb = 1.0 / p_sph
    return dict(W=W, chi2=chi2, df=df_mauchly, p=p_val,
                eps_gg=eps_gg, eps_hf=eps_hf, eps_lb=eps_lb)


def run_mixed_anova(df, alpha, sph_corr):
    """
    Full Mixed ANOVA computation matching SPSS output.
    Returns dict with all tables.
    """
    subjects   = df["subject"].unique()
    N          = len(subjects)
    a          = len(df["between"].unique())   # between levels
    b          = len(df["within"].unique())    # within levels
    between_lv = sorted(df["between"].unique())
    within_lv  = sorted(df["within"].unique())

    # Grand mean
    grand_mean = df["dv"].mean()

    # ── pivot to wide (subjects × within) ──────────────────────────────────
    df_pivot = df.pivot_table(index=["subject", "between"],
                              columns="within", values="dv").reset_index()
    df_pivot.columns.name = None
    within_cols_w = within_lv

    # ── cell means ──────────────────────────────────────────────────────────
    cell_means = df.groupby(["between", "within"])["dv"].agg(["mean", "std", "count"]).reset_index()
    cell_means.columns = ["between", "within", "mean", "std", "n"]
    cell_means["se"] = cell_means["std"] / np.sqrt(cell_means["n"])

    # Subject means
    subj_means = df.groupby("subject")["dv"].mean()
    between_means = df.groupby("between")["dv"].mean()
    within_means  = df.groupby("within")["dv"].mean()
    n_per_between = df.groupby("between")["subject"].nunique()

    # ── SS Decomposition ────────────────────────────────────────────────────
    # SS_total
    SS_total = ((df["dv"] - grand_mean) ** 2).sum()

    # SS_between (A)
    SS_A = b * sum(
        n_per_between[g] * (between_means[g] - grand_mean) ** 2
        for g in between_lv
    )

    # SS_subjects within groups (error between)
    SS_S_A = 0
    for g in between_lv:
        subj_g = df[df["between"] == g]["subject"].unique()
        subj_means_g = df[df["between"] == g].groupby("subject")["dv"].mean()
        SS_S_A += b * ((subj_means_g - between_means[g]) ** 2).sum()

    # SS_within (B)
    SS_B = a * N / a * sum(
        (within_means[t] - grand_mean) ** 2
        for t in within_lv
    )
    # Correct formula
    SS_B = 0
    for t in within_lv:
        SS_B += (df[df["within"] == t]["dv"].mean() - grand_mean) ** 2
    SS_B = SS_B * N

    # SS_AB (interaction)
    SS_AB = 0
    for g in between_lv:
        for t in within_lv:
            cell_m = df[(df["between"] == g) & (df["within"] == t)]["dv"].mean()
            SS_AB += (cell_m - between_means[g] - within_means[t] + grand_mean) ** 2
    SS_AB = SS_AB * (N / len(between_lv))

    # SS_error_within (B × S/A)
    SS_BxS_A = 0
    for subj in subjects:
        row = df[df["subject"] == subj]
        g = row["between"].iloc[0]
        subj_mean = subj_means[subj]
        for t in within_lv:
            val = row[row["within"] == t]["dv"].values
            if len(val) == 0:
                continue
            v = val[0]
            cell_m = df[(df["between"] == g) & (df["within"] == t)]["dv"].mean()
            # Rumus residual yang benar:
            residual = v - cell_m - subj_mean + grand_mean
            SS_BxS_A += residual ** 2

    # ── Degrees of freedom ──────────────────────────────────────────────────
    df_A    = a - 1
    df_S_A  = N - a           # subjects within groups
    df_B    = b - 1
    df_AB   = (a - 1) * (b - 1)
    df_BxSA = (b - 1) * (N - a)

    # ── Mauchly / Sphericity ─────────────────────────────────────────────────
    # Build wide matrix per between group and pool
    wide_matrices = {}
    for g in between_lv:
        subj_g = df_pivot[df_pivot["between"] == g]
        mat = subj_g[within_cols_w].values.astype(float)
        wide_matrices[g] = mat

    all_wide = np.vstack(list(wide_matrices.values()))
    # Center within groups before pooling
    sph_result = mauchly_sphericity(all_wide) if b > 2 else None

    # Determine epsilon
    if b == 2 or sph_result is None:
        eps_use = 1.0
        sph_assumed = True
    else:
        p_sph = sph_result["p"]
        if sph_corr == "auto":
            if p_sph < alpha:
                eps_use = sph_result["eps_gg"]
                sph_assumed = False
            else:
                eps_use = 1.0
                sph_assumed = True
        elif sph_corr == "gg":
            eps_use = sph_result["eps_gg"]
            sph_assumed = False
        elif sph_corr == "hf":
            eps_use = sph_result["eps_hf"]
            sph_assumed = False
        elif sph_corr == "lb":
            eps_use = sph_result["eps_lb"]
            sph_assumed = False
        else:
            eps_use = 1.0
            sph_assumed = True

    # ── Adjusted df for within effects ──────────────────────────────────────
    df_B_adj    = df_B    * eps_use
    df_AB_adj   = df_AB   * eps_use
    df_BxSA_adj = df_BxSA * eps_use

    # ── MS ──────────────────────────────────────────────────────────────────
    MS_A     = SS_A    / df_A
    MS_S_A   = SS_S_A  / df_S_A
    MS_B     = SS_B    / df_B_adj if df_B_adj > 0 else 0
    MS_AB    = SS_AB   / df_AB_adj if df_AB_adj > 0 else 0
    MS_BxSA  = SS_BxS_A / df_BxSA_adj if df_BxSA_adj > 0 else 0

    # ── F ratios ─────────────────────────────────────────────────────────────
    F_A   = MS_A   / MS_S_A
    F_B   = MS_B   / MS_BxSA if MS_BxSA > 0 else np.nan
    F_AB  = MS_AB  / MS_BxSA if MS_BxSA > 0 else np.nan

    p_A   = 1 - fdist.cdf(F_A,  df_A,    df_S_A)
    p_B   = 1 - fdist.cdf(F_B,  df_B_adj, df_BxSA_adj)
    p_AB  = 1 - fdist.cdf(F_AB, df_AB_adj, df_BxSA_adj)

    # ── Effect sizes (partial eta-squared) ──────────────────────────────────
    np2_A  = SS_A   / (SS_A  + SS_S_A)
    np2_B  = SS_B   / (SS_B  + SS_BxS_A)
    np2_AB = SS_AB  / (SS_AB + SS_BxS_A)

    # Observed power (approximate via non-central F)
    def obs_power(F, df1, df2, alpha_lvl):
        try:
            ncp = F * df1
            crit = fdist.ppf(1 - alpha_lvl, df1, df2)
            power = 1 - stats.ncf.cdf(crit, df1, df2, nc=ncp)
            return float(np.clip(power, 0, 1))
        except Exception:
            return np.nan

    pow_A  = obs_power(F_A,  df_A,    df_S_A,    alpha)
    pow_B  = obs_power(F_B,  df_B_adj, df_BxSA_adj, alpha)
    pow_AB = obs_power(F_AB, df_AB_adj, df_BxSA_adj, alpha)

    # ── Build ANOVA table ────────────────────────────────────────────────────
    anova_table = pd.DataFrame([
        {
            "Source": f"Between Subjects",
            "": "",
            "SS": np.nan, "df": np.nan, "MS": np.nan,
            "F": np.nan, "p": np.nan, "η²p": np.nan, "Power": np.nan,
            "_header": True,
        },
        {
            "Source": f"  {between_col}",
            "": "",
            "SS": SS_A,   "df": df_A,   "MS": MS_A,
            "F": F_A,     "p": p_A,     "η²p": np2_A,  "Power": pow_A,
            "_header": False,
        },
        {
            "Source": "  Error (Between)",
            "": "",
            "SS": SS_S_A, "df": df_S_A, "MS": MS_S_A,
            "F": np.nan,  "p": np.nan,  "η²p": np.nan, "Power": np.nan,
            "_header": False,
        },
        {
            "Source": "Within Subjects",
            "": "",
            "SS": np.nan, "df": np.nan, "MS": np.nan,
            "F": np.nan, "p": np.nan, "η²p": np.nan, "Power": np.nan,
            "_header": True,
        },
        {
            "Source": f"  {within_col}",
            "": "",
            "SS": SS_B,   "df": df_B_adj,  "MS": MS_B,
            "F": F_B,     "p": p_B,        "η²p": np2_B,  "Power": pow_B,
            "_header": False,
        },
        {
            "Source": f"  {between_col} × {within_col}",
            "": "",
            "SS": SS_AB,  "df": df_AB_adj, "MS": MS_AB,
            "F": F_AB,    "p": p_AB,       "η²p": np2_AB, "Power": pow_AB,
            "_header": False,
        },
        {
            "Source": "  Error (Within)",
            "": "",
            "SS": SS_BxS_A, "df": df_BxSA_adj, "MS": MS_BxSA,
            "F": np.nan,    "p": np.nan,        "η²p": np.nan, "Power": np.nan,
            "_header": False,
        },
        {
            "Source": "Total",
            "": "",
            "SS": SS_total, "df": N*b - 1, "MS": np.nan,
            "F": np.nan, "p": np.nan, "η²p": np.nan, "Power": np.nan,
            "_header": False,
        },
    ])

    return dict(
        anova_table=anova_table,
        cell_means=cell_means,
        between_means=between_means,
        within_means=within_means,
        grand_mean=grand_mean,
        N=N, a=a, b=b,
        between_lv=between_lv,
        within_lv=within_lv,
        SS_A=SS_A, SS_B=SS_B, SS_AB=SS_AB, SS_S_A=SS_S_A, SS_BxS_A=SS_BxS_A,
        df_A=df_A, df_B=df_B, df_AB=df_AB, df_S_A=df_S_A, df_BxSA=df_BxSA,
        df_B_adj=df_B_adj, df_AB_adj=df_AB_adj, df_BxSA_adj=df_BxSA_adj,
        MS_A=MS_A, MS_B=MS_B, MS_AB=MS_AB, MS_S_A=MS_S_A, MS_BxSA=MS_BxSA,
        F_A=F_A, F_B=F_B, F_AB=F_AB,
        p_A=p_A, p_B=p_B, p_AB=p_AB,
        np2_A=np2_A, np2_B=np2_B, np2_AB=np2_AB,
        pow_A=pow_A, pow_B=pow_B, pow_AB=pow_AB,
        sph_result=sph_result, eps_use=eps_use, sph_assumed=sph_assumed,
        n_per_between=n_per_between,
        alpha=alpha,
    )


def run_posthoc(df, factor_col, dv_col, method, alpha):
    """Pairwise comparisons with correction."""
    levels = sorted(df[factor_col].unique())
    pairs = list(itertools.combinations(levels, 2))
    rows = []
    for (l1, l2) in pairs:
        g1 = df[df[factor_col] == l1][dv_col].values
        g2 = df[df[factor_col] == l2][dv_col].values
        t, p_raw = stats.ttest_ind(g1, g2)
        mean_diff = g1.mean() - g2.mean()
        pooled_sd = np.sqrt((g1.std()**2 + g2.std()**2) / 2)
        cohen_d = mean_diff / pooled_sd if pooled_sd > 0 else np.nan
        rows.append({
            "Group 1": l1, "Group 2": l2,
            "Mean Diff": mean_diff, "t": t,
            "p (raw)": p_raw, "Cohen's d": cohen_d,
        })
    ph_df = pd.DataFrame(rows)
    if len(ph_df) == 0:
        return ph_df
    # Apply correction
    from statsmodels.stats.multitest import multipletests
    reject, p_corr, _, _ = multipletests(ph_df["p (raw)"], method=method, alpha=alpha)
    ph_df["p (corrected)"] = p_corr
    ph_df["Sig."] = reject
    return ph_df


def run_posthoc_within(df, within_col, dv_col, subj_col, method, alpha):
    """Paired post-hoc for within-subjects factor."""
    levels = sorted(df[within_col].unique())
    pairs = list(itertools.combinations(levels, 2))
    rows = []
    for (l1, l2) in pairs:
        paired = df.pivot_table(index=subj_col, columns=within_col, values=dv_col)
        if l1 not in paired.columns or l2 not in paired.columns:
            continue
        both = paired[[l1, l2]].dropna()
        t, p_raw = stats.ttest_rel(both[l1], both[l2])
        mean_diff = (both[l1] - both[l2]).mean()
        sd_diff   = (both[l1] - both[l2]).std()
        cohen_d   = mean_diff / sd_diff if sd_diff > 0 else np.nan
        rows.append({
            "Level 1": l1, "Level 2": l2,
            "Mean Diff": mean_diff, "t": t,
            "p (raw)": p_raw, "Cohen's d": cohen_d,
        })
    ph_df = pd.DataFrame(rows)
    if len(ph_df) == 0:
        return ph_df
    from statsmodels.stats.multitest import multipletests
    reject, p_corr, _, _ = multipletests(ph_df["p (raw)"], method=method, alpha=alpha)
    ph_df["p (corrected)"] = p_corr
    ph_df["Sig."] = reject
    return ph_df


# ══════════════════════════════════════════════════════════════════════════════
#  RUN ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
with st.spinner("⏳ Menghitung Mixed ANOVA ..."):
    res = run_mixed_anova(df, alpha_level, sphericity_correction)

# ══════════════════════════════════════════════════════════════════════════════
#  OVERVIEW CARDS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📊 Ringkasan Hasil</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
def card(col, label, value, sub=""):
    col.markdown(f"""
    <div class="stat-card">
      <div class="label">{label}</div>
      <div class="value">{value}</div>
      <div class="sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

card(c1, "N Subjek", res["N"], f"{res['a']} kelompok")
card(c2, between_col, f"F = {res['F_A']:.3f}", sig_badge(res['p_A'], alpha_level).replace('<span', '<span style="font-size:0.7rem"').replace("</span>", "</span>"))
card(c3, within_col, f"F = {res['F_B']:.3f}", sig_badge(res['p_B'], alpha_level).replace('<span', '<span style="font-size:0.7rem"').replace("</span>", "</span>"))
card(c4, "Interaksi", f"F = {res['F_AB']:.3f}", sig_badge(res['p_AB'], alpha_level).replace('<span', '<span style="font-size:0.7rem"').replace("</span>", "</span>"))
card(c5, "Grand Mean", f"{res['grand_mean']:.2f}", f"SD = {df['dv'].std():.2f}")

# ══════════════════════════════════════════════════════════════════════════════
#  DESCRIPTIVE STATISTICS
# ══════════════════════════════════════════════════════════════════════════════
if show_descriptives:
    st.markdown('<div class="section-header">📋 Statistik Deskriptif</div>', unsafe_allow_html=True)

    desc = df.groupby(["between", "within"])["dv"].agg(
        N="count", Mean="mean", SD="std", SE=lambda x: x.std()/np.sqrt(len(x)),
        Min="min", Max="max", Median="median"
    ).reset_index()
    desc.columns = [between_col, within_col, "N", "Mean", "SD", "SE", "Min", "Max", "Median"]

    # CI 95%
    desc["CI_lower"] = desc["Mean"] - 1.96 * desc["SE"]
    desc["CI_upper"] = desc["Mean"] + 1.96 * desc["SE"]
    desc["95% CI"] = desc.apply(lambda r: f"[{r['CI_lower']:.2f}, {r['CI_upper']:.2f}]", axis=1)

    display_desc = desc[[between_col, within_col, "N", "Mean", "SD", "SE", "95% CI", "Min", "Max", "Median"]].copy()
    for c in ["Mean", "SD", "SE", "Min", "Max", "Median"]:
        display_desc[c] = display_desc[c].round(3)

    st.dataframe(display_desc, use_container_width=True, hide_index=True)

    # Marginal means
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"**Marginal Means — {between_col}**")
        mm_b = df.groupby("between")["dv"].agg(N="count", Mean="mean", SD="std").reset_index()
        mm_b["SE"] = mm_b["SD"] / np.sqrt(mm_b["N"])
        mm_b.columns = [between_col, "N", "Mean", "SD", "SE"]
        st.dataframe(mm_b.round(3), use_container_width=True, hide_index=True)
    with col_b:
        st.markdown(f"**Marginal Means — {within_col}**")
        mm_w = df.groupby("within")["dv"].agg(N="count", Mean="mean", SD="std").reset_index()
        mm_w["SE"] = mm_w["SD"] / np.sqrt(mm_w["N"])
        mm_w.columns = [within_col, "N", "Mean", "SD", "SE"]
        st.dataframe(mm_w.round(3), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
#  ASSUMPTION CHECKS
# ══════════════════════════════════════════════════════════════════════════════
if show_assumptions:
    st.markdown('<div class="section-header">🔍 Uji Asumsi</div>', unsafe_allow_html=True)

    tab_norm, tab_homo, tab_sph = st.tabs(["Normalitas", "Homogenitas Varians", "Sphericity"])

    with tab_norm:
        st.markdown("**Shapiro-Wilk Test per Sel** (normal jika p > α)")
        norm_rows = []
        for g in res["between_lv"]:
            for t in res["within_lv"]:
                vals = df[(df["between"] == g) & (df["within"] == t)]["dv"].values
                if len(vals) >= 3:
                    W, p = stats.shapiro(vals)
                    norm_rows.append({
                        between_col: g, within_col: t,
                        "W": round(W, 4), "p": round(p, 4),
                        "Normal?": "✓ Ya" if p > alpha_level else "✗ Tidak",
                    })
        st.dataframe(pd.DataFrame(norm_rows), use_container_width=True, hide_index=True)

    with tab_homo:
        st.markdown("**Levene's Test** per level within-subjects")
        lev_rows = []
        for t in res["within_lv"]:
            groups = [df[(df["between"] == g) & (df["within"] == t)]["dv"].values
                      for g in res["between_lv"]]
            if all(len(g) >= 2 for g in groups):
                F_lev, p_lev = stats.levene(*groups)
                lev_rows.append({
                    within_col: t, "F": round(F_lev, 4),
                    "df1": res["a"] - 1,
                    "df2": res["N"] - res["a"],
                    "p": round(p_lev, 4),
                    "Homogen?": "✓ Ya" if p_lev > alpha_level else "✗ Tidak",
                })
        st.dataframe(pd.DataFrame(lev_rows), use_container_width=True, hide_index=True)

    with tab_sph:
        if res["sph_result"] is None:
            st.info("ℹ️ Uji Mauchly tidak diperlukan karena faktor within hanya memiliki 2 level.")
        else:
            s = res["sph_result"]
            st.markdown(f"""
            **Mauchly's Test of Sphericity**

            | Statistik | Nilai |
            |---|---|
            | Mauchly's W | {s['W']:.4f} |
            | Chi-square (approx.) | {s['chi2']:.4f} |
            | df | {int(s['df'])} |
            | p-value | {format_p(s['p'])} |
            | **Epsilon (GG)** | **{s['eps_gg']:.4f}** |
            | Epsilon (HF) | {s['eps_hf']:.4f} |
            | Epsilon (LB) | {s['eps_lb']:.4f} |
            """)
            if s['p'] < alpha_level:
                st.warning(f"⚠️ Asumsi sphericity **dilanggar** (p = {format_p(s['p'])} < {alpha_level}). Koreksi diterapkan dengan ε = {res['eps_use']:.4f}.")
            else:
                st.success(f"✅ Asumsi sphericity **terpenuhi** (p = {format_p(s['p'])} ≥ {alpha_level}).")

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN ANOVA TABLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📊 Tabel Mixed ANOVA</div>', unsafe_allow_html=True)

at = res["anova_table"]
display_at = at[~at["_header"]].drop(columns=["_header", ""]).copy()

def fmt_val(v, decimals=3):
    if pd.isna(v): return ""
    return f"{v:.{decimals}f}"

display_rows = []
for _, row in at.iterrows():
    if row["_header"]:
        display_rows.append({
            "Sumber": f"**{row['Source']}**",
            "SS": "", "df": "", "MS": "",
            "F": "", "p": "", "η²p": "", "Power": "",
        })
    else:
        display_rows.append({
            "Sumber": row["Source"],
            "SS": fmt_val(row["SS"]),
            "df": fmt_val(row["df"], 3),
            "MS": fmt_val(row["MS"]),
            "F": fmt_val(row["F"]),
            "p": format_p(row["p"]) if not pd.isna(row["p"]) else "",
            "η²p": fmt_val(row["η²p"]),
            "Power": fmt_val(row["Power"]),
        })

df_display = pd.DataFrame(display_rows)
st.dataframe(df_display, use_container_width=True, hide_index=True)

if not res["sph_assumed"] and res["sph_result"]:
    corr_name = {"gg": "Greenhouse-Geisser", "hf": "Huynh-Feldt", "lb": "Lower-bound"}.get(
        sphericity_correction, "Greenhouse-Geisser"
    )
    st.caption(f"ᵃ df for within-subjects effects adjusted using {corr_name} correction (ε = {res['eps_use']:.4f})")

# ══════════════════════════════════════════════════════════════════════════════
#  SIGNIFICANCE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">🎯 Uji Signifikansi</div>', unsafe_allow_html=True)

sc1, sc2, sc3 = st.columns(3)

def effect_display(np2, etype):
    if etype == "np2":
        return f"η²p = {np2:.4f} ({interpret_effect(np2, 'np2')})"
    elif etype == "eta2":
        return f"η² = {np2:.4f} ({interpret_effect(np2, 'eta2')})"
    else:
        f_val = cohen_f_from_np2(np2)
        return f"f = {f_val:.4f} ({interpret_effect(f_val, 'cohen_f')})"

for col_s, label, F, df1, df2, p, np2, power in [
    (sc1, between_col, res["F_A"], res["df_A"], res["df_S_A"], res["p_A"], res["np2_A"], res["pow_A"]),
    (sc2, within_col, res["F_B"], res["df_B_adj"], res["df_BxSA_adj"], res["p_B"], res["np2_B"], res["pow_B"]),
    (sc3, f"{between_col} × {within_col}", res["F_AB"], res["df_AB_adj"], res["df_BxSA_adj"], res["p_AB"], res["np2_AB"], res["pow_AB"]),
]:
    sig = p < alpha_level
    col_s.markdown(f"""
    <div class="stat-card">
      <div class="label">{label}</div>
      <div class="value">F({df1:.2f}, {df2:.2f}) = {F:.3f}</div>
      <div class="sub">p {format_p(p)} | {effect_display(np2, effect_size_type)}</div>
      <div class="sub">Observed Power = {power:.3f}</div>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  DESCRIPTIVE INTERPRETATION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">💬 Interpretasi Deskriptif</div>', unsafe_allow_html=True)

def gen_interpretation(res, between_col, within_col, alpha_level, effect_size_type):
    lines = []

    # Between
    sig_A = res["p_A"] < alpha_level
    eff_A = interpret_effect(res["np2_A"], "np2")
    f_A_str = f"F({res['df_A']:.0f}, {res['df_S_A']:.0f}) = {res['F_A']:.3f}, p {format_p(res['p_A'])}, η²p = {res['np2_A']:.3f}"
    if sig_A:
        lines.append(f"**Faktor Between-Subjects ({between_col}):** Analisis Mixed ANOVA menunjukkan efek utama yang **signifikan secara statistik** untuk faktor {between_col} [{f_A_str}]. Hal ini mengindikasikan bahwa terdapat perbedaan yang bermakna antara kelompok-kelompok pada variabel dependen. Besar efek tergolong **{eff_A}** (η²p = {res['np2_A']:.3f}).")
    else:
        lines.append(f"**Faktor Between-Subjects ({between_col}):** Efek utama untuk faktor {between_col} **tidak signifikan** secara statistik [{f_A_str}]. Tidak ditemukan perbedaan bermakna antar kelompok secara keseluruhan. Besar efek tergolong **{eff_A}**.")

    # Within
    sig_B = res["p_B"] < alpha_level
    eff_B = interpret_effect(res["np2_B"], "np2")
    f_B_str = f"F({res['df_B_adj']:.2f}, {res['df_BxSA_adj']:.2f}) = {res['F_B']:.3f}, p {format_p(res['p_B'])}, η²p = {res['np2_B']:.3f}"
    if sig_B:
        lines.append(f"**Faktor Within-Subjects ({within_col}):** Terdapat efek utama yang **signifikan** untuk faktor {within_col} [{f_B_str}]. Ini berarti skor berubah secara bermakna sepanjang level-level {within_col}. Besar efek tergolong **{eff_B}**.")
    else:
        lines.append(f"**Faktor Within-Subjects ({within_col}):** Efek utama untuk faktor {within_col} **tidak signifikan** [{f_B_str}]. Tidak ada perubahan bermakna sepanjang level {within_col}. Besar efek tergolong **{eff_B}**.")

    # Interaction
    sig_AB = res["p_AB"] < alpha_level
    eff_AB = interpret_effect(res["np2_AB"], "np2")
    f_AB_str = f"F({res['df_AB_adj']:.2f}, {res['df_BxSA_adj']:.2f}) = {res['F_AB']:.3f}, p {format_p(res['p_AB'])}, η²p = {res['np2_AB']:.3f}"
    if sig_AB:
        lines.append(f"**Efek Interaksi ({between_col} × {within_col}):** Ditemukan efek interaksi yang **signifikan** [{f_AB_str}]. Ini berarti pengaruh {within_col} terhadap variabel dependen **berbeda-beda tergantung pada kelompok** {between_col}. Besar efek interaksi tergolong **{eff_AB}**. Disarankan melakukan analisis post-hoc untuk mengidentifikasi pola spesifik perbedaan.")
    else:
        lines.append(f"**Efek Interaksi ({between_col} × {within_col}):** Efek interaksi **tidak signifikan** [{f_AB_str}]. Pengaruh {within_col} bersifat **konsisten** di semua kelompok {between_col}. Besar efek tergolong **{eff_AB}**.")

    # Power note
    low_power = []
    if res["pow_A"] < 0.80: low_power.append(between_col)
    if res["pow_B"] < 0.80: low_power.append(within_col)
    if res["pow_AB"] < 0.80: low_power.append(f"{between_col} × {within_col}")
    if low_power:
        lines.append(f"**Catatan Power:** Observed power < 0.80 terdeteksi pada efek: {', '.join(low_power)}. Pertimbangkan penambahan ukuran sampel untuk meningkatkan power statistik (disarankan N per kelompok ≥ 30 untuk efek sedang).")

    return lines

interp_lines = gen_interpretation(res, between_col, within_col, alpha_level, effect_size_type)
for line in interp_lines:
    st.markdown(f'<div class="interp-box">{line}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  POST-HOC TESTS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">🔬 Analisis Post-Hoc</div>', unsafe_allow_html=True)

tab_ph_b, tab_ph_w, tab_ph_i = st.tabs([
    f"Between: {between_col}",
    f"Within: {within_col}",
    "Interaction (Simple Effects)",
])

with tab_ph_b:
    if res["a"] < 3:
        st.info(f"ℹ️ Hanya 2 level pada {between_col}. Hasil uji-F sudah cukup menentukan perbedaan.")
    else:
        ph_b = run_posthoc(df, "between", "dv", posthoc_method, alpha_level)
        ph_b.columns = [between_col + " 1", between_col + " 2", "Mean Diff", "t", "p (raw)", "Cohen's d", "p (corrected)", "Sig."]
        ph_b = ph_b.round(4)
        st.dataframe(ph_b, use_container_width=True, hide_index=True)
        st.caption(f"Koreksi: {posthoc_method.capitalize()}, α = {alpha_level}")

with tab_ph_w:
    if res["b"] < 3:
        st.info(f"ℹ️ Hanya 2 level pada {within_col}. Lihat nilai F dan p untuk perbedaan.")
    else:
        ph_w = run_posthoc_within(df, "within", "dv", "subject", posthoc_method, alpha_level)
        ph_w.columns = [within_col + " 1", within_col + " 2", "Mean Diff", "t", "p (raw)", "Cohen's d", "p (corrected)", "Sig."]
        ph_w = ph_w.round(4)
        st.dataframe(ph_w, use_container_width=True, hide_index=True)
        st.caption(f"Paired t-test dengan koreksi {posthoc_method.capitalize()}")

with tab_ph_i:
    st.markdown("**Simple Effects: Pengaruh Within per Level Between**")
    se_rows = []
    for g in res["between_lv"]:
        df_g = df[df["between"] == g]
        if res["b"] < 3:
            t_vals = res["within_lv"]
            g1 = df_g[df_g["within"] == t_vals[0]]["dv"].values
            g2 = df_g[df_g["within"] == t_vals[-1]]["dv"].values
            t_s, p_s = stats.ttest_rel(g1, g2)
            se_rows.append({
                between_col: g,
                within_col: f"{t_vals[0]} vs {t_vals[-1]}",
                "t": round(t_s, 4), "p": round(p_s, 4),
                "Mean Diff": round(g1.mean() - g2.mean(), 4),
                "Sig.": p_s < alpha_level,
            })
        else:
            for (l1, l2) in itertools.combinations(res["within_lv"], 2):
                g1 = df_g[df_g["within"] == l1]["dv"].values
                g2 = df_g[df_g["within"] == l2]["dv"].values
                if len(g1) > 1 and len(g2) > 1:
                    t_s, p_s = stats.ttest_rel(g1, g2)
                    se_rows.append({
                        between_col: g,
                        within_col: f"{l1} vs {l2}",
                        "t": round(t_s, 4), "p": round(p_s, 4),
                        "Mean Diff": round(g1.mean() - g2.mean(), 4),
                        "Sig.": p_s < alpha_level,
                    })
    if se_rows:
        st.dataframe(pd.DataFrame(se_rows), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
#  VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">📈 Visualisasi</div>', unsafe_allow_html=True)

sns.set_style(plot_style)
palette = sns.color_palette(palette_choice, n_colors=max(res["a"], res["b"]))

fig = plt.figure(figsize=(18, 14))
fig.patch.set_facecolor("#fafafa")
gs = GridSpec(2, 3, figure=fig, hspace=0.40, wspace=0.35)

# ── Plot 1: Profile Plot (Interaction) ─────────────────────────────────────
ax1 = fig.add_subplot(gs[0, :2])
cell = res["cell_means"]
for i, g in enumerate(res["between_lv"]):
    sub = cell[cell["between"] == g].sort_values("within",
          key=lambda x: x.map({v: j for j, v in enumerate(res["within_lv"])}))
    ax1.errorbar(sub["within"], sub["mean"], yerr=1.96 * sub["se"],
                 marker="o", ms=8, lw=2.2, capsize=5,
                 color=palette[i], label=g, zorder=3)
ax1.set_title(f"Profile Plot: {between_col} × {within_col}", fontsize=13, fontweight="bold")
ax1.set_xlabel(within_col, fontsize=11)
ax1.set_ylabel(f"Mean {dv_col}", fontsize=11)
ax1.legend(title=between_col, framealpha=0.8)
ax1.grid(True, alpha=0.3, linestyle="--")

# ── Plot 2: Bar Chart ───────────────────────────────────────────────────────
ax2 = fig.add_subplot(gs[0, 2])
within_order = res["within_lv"]
x = np.arange(len(within_order))
bar_w = 0.8 / res["a"]
for i, g in enumerate(res["between_lv"]):
    sub = cell[cell["between"] == g].set_index("within").reindex(within_order)
    offset = (i - res["a"]/2 + 0.5) * bar_w
    ax2.bar(x + offset, sub["mean"], bar_w * 0.9, yerr=1.96 * sub["se"],
            color=palette[i], label=g, alpha=0.85, capsize=4, error_kw={"elinewidth": 1.5})
ax2.set_xticks(x)
ax2.set_xticklabels(within_order, rotation=20 if len(within_order) > 3 else 0)
ax2.set_title("Mean ± 95% CI", fontsize=13, fontweight="bold")
ax2.set_ylabel(f"Mean {dv_col}", fontsize=11)
ax2.legend(title=between_col, fontsize=9)
ax2.grid(True, alpha=0.3, linestyle="--", axis="y")

# ── Plot 3: Box Plot ─────────────────────────────────────────────────────────
ax3 = fig.add_subplot(gs[1, 0])
pivot_box = df.copy()
pivot_box["Group"] = pivot_box["between"].astype(str) + "\n" + pivot_box["within"].astype(str)
order_box = [f"{g}\n{t}" for g in res["between_lv"] for t in res["within_lv"]]
color_map = {f"{g}\n{t}": palette[i] for i, g in enumerate(res["between_lv"]) for t in res["within_lv"]}
sns.boxplot(data=pivot_box, x="Group", y="dv", order=order_box,
            palette=color_map, ax=ax3, flierprops={"marker": "o", "ms": 4})
ax3.set_title("Box Plot per Sel", fontsize=13, fontweight="bold")
ax3.set_xlabel("")
ax3.set_ylabel(f"{dv_col}", fontsize=11)
ax3.tick_params(axis="x", labelsize=8)

# ── Plot 4: Violin Plot ──────────────────────────────────────────────────────
ax4 = fig.add_subplot(gs[1, 1])
sns.violinplot(data=df, x="within", y="dv", hue="between",
               order=res["within_lv"], palette=palette_choice,
               inner="quartile", ax=ax4, alpha=0.8)
ax4.set_title("Violin Plot Distribution", fontsize=13, fontweight="bold")
ax4.set_xlabel(within_col, fontsize=11)
ax4.set_ylabel(f"{dv_col}", fontsize=11)
ax4.legend(title=between_col, fontsize=9)

# ── Plot 5: Effect Size Visual ───────────────────────────────────────────────
ax5 = fig.add_subplot(gs[1, 2])
effects = {
    between_col: res["np2_A"],
    within_col: res["np2_B"],
    f"Interaksi": res["np2_AB"],
}
colors_eff = ["#e63946" if v >= alpha_level else "#2dc653"
              for k, v in zip(effects.keys(),
                              [res["p_A"], res["p_B"], res["p_AB"]])]
bars = ax5.barh(list(effects.keys()), list(effects.values()), color=colors_eff, height=0.5, alpha=0.85)
ax5.axvline(0.01, ls="--", color="#aaa", lw=1.2, label="Kecil (0.01)")
ax5.axvline(0.06, ls="--", color="#555", lw=1.2, label="Sedang (0.06)")
ax5.axvline(0.14, ls="--", color="#111", lw=1.2, label="Besar (0.14)")
for bar, val in zip(bars, effects.values()):
    ax5.text(val + 0.002, bar.get_y() + bar.get_height()/2,
             f"{val:.4f}", va="center", fontsize=10, fontweight="bold")
ax5.set_title("Partial η² Effect Sizes", fontsize=13, fontweight="bold")
ax5.set_xlabel("η²p", fontsize=11)
ax5.legend(fontsize=8, loc="lower right")
ax5.set_xlim(0, max(max(effects.values()) * 1.3, 0.20))

plt.suptitle(f"Mixed ANOVA Results — {between_col} × {within_col}",
             fontsize=15, fontweight="bold", y=1.01)

st.pyplot(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
#  WORD REPORT GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def add_table_style(table):
    table.style = "Table Grid"
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            cell.paragraphs[0].runs[0].font.size = Pt(9) if i > 0 else Pt(9)

def build_word_report(res, df, between_col, within_col, dv_col, alpha_level,
                      posthoc_method, effect_size_type, fig):
    doc = Document()

    # Margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    # ── Title ──────────────────────────────────────────────────────────────
    title_para = doc.add_heading("Mixed ANOVA Analysis Report", 0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.runs[0]
    run.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

    doc.add_paragraph(
        f"Between-subjects factor: {between_col}  |  "
        f"Within-subjects factor: {within_col}  |  "
        f"Dependent variable: {dv_col}  |  "
        f"α = {alpha_level}  |  N = {res['N']}"
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ── Sample information ─────────────────────────────────────────────────
    doc.add_heading("1. Sample Information", level=1)
    info_table = doc.add_table(rows=4, cols=2)
    info_data = [
        ("Total Subjects (N)", str(res["N"])),
        (f"Between-subjects groups ({between_col})", ", ".join(str(l) for l in res["between_lv"])),
        (f"Within-subjects levels ({within_col})", ", ".join(str(l) for l in res["within_lv"])),
        ("Grand Mean", f"{res['grand_mean']:.3f}"),
    ]
    for i, (k, v) in enumerate(info_data):
        info_table.rows[i].cells[0].text = k
        info_table.rows[i].cells[1].text = v
        info_table.rows[i].cells[0].paragraphs[0].runs[0].bold = True
    info_table.style = "Table Grid"

    doc.add_paragraph()

    # ── Descriptive Stats ──────────────────────────────────────────────────
    doc.add_heading("2. Descriptive Statistics", level=1)
    desc2 = df.groupby(["between", "within"])["dv"].agg(
        N="count", Mean="mean", SD="std",
        SE=lambda x: x.std()/np.sqrt(len(x)), Min="min", Max="max"
    ).reset_index()
    desc2.columns = [between_col, within_col, "N", "Mean", "SD", "SE", "Min", "Max"]

    d_table = doc.add_table(rows=len(desc2) + 1, cols=len(desc2.columns))
    for j, col in enumerate(desc2.columns):
        cell = d_table.rows[0].cells[j]
        cell.text = col
        cell.paragraphs[0].runs[0].bold = True
        set_cell_bg(cell, "1a1a2e")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)
    for i, (_, row) in enumerate(desc2.iterrows()):
        for j, val in enumerate(row):
            v = f"{val:.3f}" if isinstance(val, float) else str(val)
            d_table.rows[i+1].cells[j].text = v
    d_table.style = "Table Grid"

    doc.add_paragraph()

    # ── ANOVA Table ────────────────────────────────────────────────────────
    doc.add_heading("3. Mixed ANOVA Summary Table", level=1)
    at2 = res["anova_table"]
    cols_anova = ["Sumber", "SS", "df", "MS", "F", "p", "η²p", "Power"]
    a_table = doc.add_table(rows=len(at2) + 1, cols=len(cols_anova))

    hdr_cells = a_table.rows[0].cells
    for j, c in enumerate(cols_anova):
        hdr_cells[j].text = c
        hdr_cells[j].paragraphs[0].runs[0].bold = True
        set_cell_bg(hdr_cells[j], "1a1a2e")
        hdr_cells[j].paragraphs[0].runs[0].font.color.rgb = RGBColor(0xff, 0xff, 0xff)

    for i, (_, row) in enumerate(at2.iterrows()):
        r = a_table.rows[i + 1]
        if row["_header"]:
            for cell in r.cells:
                set_cell_bg(cell, "e8e8e8")
            r.cells[0].text = row["Source"]
            r.cells[0].paragraphs[0].runs[0].bold = True
        else:
            r.cells[0].text = row["Source"]
            r.cells[1].text = fmt_val(row["SS"])
            r.cells[2].text = fmt_val(row["df"])
            r.cells[3].text = fmt_val(row["MS"])
            r.cells[4].text = fmt_val(row["F"])
            r.cells[5].text = format_p(row["p"]) if not pd.isna(row["p"]) else ""
            r.cells[6].text = fmt_val(row["η²p"])
            r.cells[7].text = fmt_val(row["Power"])
    a_table.style = "Table Grid"

    doc.add_paragraph()

    # ── Sphericity ─────────────────────────────────────────────────────────
    doc.add_heading("4. Sphericity Test (Mauchly's W)", level=1)
    if res["sph_result"] is None:
        doc.add_paragraph("Mauchly's test not applicable (within-subjects factor has only 2 levels).")
    else:
        s = res["sph_result"]
        doc.add_paragraph(
            f"Mauchly's W = {s['W']:.4f}, χ²({int(s['df'])}) = {s['chi2']:.4f}, "
            f"p {format_p(s['p'])}, ε(GG) = {s['eps_gg']:.4f}, ε(HF) = {s['eps_hf']:.4f}."
        )
        if s['p'] < alpha_level:
            doc.add_paragraph(
                f"Sphericity assumption violated (p < {alpha_level}). "
                f"Degrees of freedom corrected using ε = {res['eps_use']:.4f}."
            )
        else:
            doc.add_paragraph("Sphericity assumption satisfied.")

    doc.add_paragraph()

    # ── Interpretation ─────────────────────────────────────────────────────
    doc.add_heading("5. Interpretation", level=1)
    interp = gen_interpretation(res, between_col, within_col, alpha_level, effect_size_type)
    for line in interp:
        # Remove markdown bold markers for Word
        clean = line.replace("**", "")
        doc.add_paragraph(clean, style="List Bullet")

    doc.add_paragraph()

    # ── Figure ─────────────────────────────────────────────────────────────
    doc.add_heading("6. Figures", level=1)
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format="png", dpi=150, bbox_inches="tight",
                facecolor="#fafafa")
    img_buf.seek(0)
    doc.add_picture(img_buf, width=Inches(6.0))
    doc.add_paragraph("Figure 1. Mixed ANOVA visualization panel.").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ── APA Citation guide ──────────────────────────────────────────────────
    doc.add_heading("7. APA-style Reporting Template", level=1)
    apa = (
        f"A {res['a']}-group (between: {between_col}) × "
        f"{res['b']}-measurement (within: {within_col}) mixed ANOVA was conducted "
        f"(N = {res['N']}). "
        f"The main effect of {between_col} was {'significant' if res['p_A'] < alpha_level else 'not significant'}, "
        f"F({res['df_A']:.0f}, {res['df_S_A']:.0f}) = {res['F_A']:.2f}, "
        f"p {format_p(res['p_A'])}, η²p = {res['np2_A']:.3f}. "
        f"The main effect of {within_col} was {'significant' if res['p_B'] < alpha_level else 'not significant'}, "
        f"F({res['df_B_adj']:.2f}, {res['df_BxSA_adj']:.2f}) = {res['F_B']:.2f}, "
        f"p {format_p(res['p_B'])}, η²p = {res['np2_B']:.3f}. "
        f"The {between_col} × {within_col} interaction was "
        f"{'significant' if res['p_AB'] < alpha_level else 'not significant'}, "
        f"F({res['df_AB_adj']:.2f}, {res['df_BxSA_adj']:.2f}) = {res['F_AB']:.2f}, "
        f"p {format_p(res['p_AB'])}, η²p = {res['np2_AB']:.3f}."
    )
    doc.add_paragraph(apa)

    # ── Save ────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
#  DOWNLOAD SECTION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-header">⬇️ Unduh Hasil</div>', unsafe_allow_html=True)

dcol1, dcol2, dcol3, dcol4 = st.columns(4)

# 1. ANOVA table CSV
with dcol1:
    at_export = res["anova_table"].drop(columns=["_header", ""])
    csv_anova = at_export.to_csv(index=False).encode("utf-8")
    st.download_button("📄 ANOVA Table (CSV)", csv_anova,
                       "mixed_anova_table.csv", "text/csv", use_container_width=True)

# 2. Full results CSV
with dcol2:
    desc_export = df.groupby(["between", "within"])["dv"].agg(
        N="count", Mean="mean", SD="std",
        SE=lambda x: x.std()/np.sqrt(len(x)), Min="min", Max="max"
    ).reset_index()
    full_result = {
        "between_factor": [between_col] * 3,
        "within_factor":  [within_col]  * 3,
        "effect":         [between_col, within_col, f"{between_col}×{within_col}"],
        "F":              [res["F_A"], res["F_B"], res["F_AB"]],
        "df1":            [res["df_A"], res["df_B_adj"], res["df_AB_adj"]],
        "df2":            [res["df_S_A"], res["df_BxSA_adj"], res["df_BxSA_adj"]],
        "p":              [res["p_A"], res["p_B"], res["p_AB"]],
        "np2":            [res["np2_A"], res["np2_B"], res["np2_AB"]],
        "power":          [res["pow_A"], res["pow_B"], res["pow_AB"]],
    }
    csv_full = pd.DataFrame(full_result).to_csv(index=False).encode("utf-8")
    st.download_button("📊 Full Results (CSV)", csv_full,
                       "mixed_anova_results.csv", "text/csv", use_container_width=True)

# 3. Figure PNG
with dcol3:
    img_buf2 = io.BytesIO()
    fig.savefig(img_buf2, format="png", dpi=200, bbox_inches="tight", facecolor="#fafafa")
    img_buf2.seek(0)
    st.download_button("🖼️ Figures (PNG)", img_buf2.getvalue(),
                       "mixed_anova_figures.png", "image/png", use_container_width=True)

# 4. Word Report
with dcol4:
    with st.spinner("Menyiapkan laporan Word ..."):
        word_buf = build_word_report(
            res, df, between_col, within_col, dv_col,
            alpha_level, posthoc_method, effect_size_type, fig
        )
    st.download_button("📝 Full Report (Word)", word_buf.getvalue(),
                       "mixed_anova_report.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)

st.markdown("---")
st.caption("Mixed ANOVA Calculator • Equivalent to SPSS GLM Repeated Measures • Built with Streamlit + pingouin + scipy")
