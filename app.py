import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import warnings, io
from itertools import combinations
from scipy import stats
from scipy.stats import levene as scipy_levene

warnings.filterwarnings("ignore")

st.set_page_config(page_title="Mixed-Design ANOVA", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; font-size: 14px; }
.stApp { background: #f4f6fb; }

/* ── hero ── */
.hero { background: linear-gradient(135deg, #1a3a5c 0%, #24527a 100%);
  border-radius: 12px; padding: 2rem 2.5rem; margin-bottom: 1.5rem; position: relative; overflow: hidden; }
.hero::after { content: ''; position: absolute; right: -80px; top: -80px;
  width: 260px; height: 260px; border-radius: 50%;
  background: radial-gradient(circle, rgba(255,255,255,.07) 0%, transparent 65%); }
.hero::before { content: ''; position: absolute; bottom: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, #3b9eff, #36d399, #a78bfa, #fb923c); }
.hero-title { font-size: 1.65rem; font-weight: 700; color: #fff;
  margin: 0 0 .3rem; letter-spacing: -.2px; }
.hero-sub { color: #9ec4e8; font-size: .9rem; margin: 0 0 .75rem; }
.badge { display: inline-block; padding: 3px 11px; border-radius: 20px; font-size: .7rem;
  font-weight: 600; font-family: 'JetBrains Mono', monospace; margin: 2px 3px 0 0; }
.bd-b { background: rgba(59,158,255,.18); color: #7ec8ff; border: 1px solid rgba(59,158,255,.35); }
.bd-g { background: rgba(54,211,153,.18); color: #5de8b0; border: 1px solid rgba(54,211,153,.35); }
.bd-p { background: rgba(167,139,250,.18); color: #c4aaff; border: 1px solid rgba(167,139,250,.35); }
.bd-o { background: rgba(251,146,60,.18);  color: #fdb97a; border: 1px solid rgba(251,146,60,.35); }

/* ── cards ── */
.card { background: #fff; border: 1px solid #dde3ed; border-radius: 10px;
  padding: 1.5rem; margin-bottom: 1.25rem; box-shadow: 0 1px 4px rgba(0,0,0,.05); }
.card-title { font-size: .68rem; font-weight: 700; color: #6b7280; letter-spacing: 1.6px;
  text-transform: uppercase; margin-bottom: 1rem; padding-bottom: .5rem;
  border-bottom: 1px solid #e9ecf3; font-family: 'JetBrains Mono', monospace; }

/* ── metric row ── */
.mrow { display: grid; grid-template-columns: repeat(4,1fr); gap: 1rem; margin: 1.25rem 0; }
.mc { background: #fff; border: 1px solid #dde3ed; border-radius: 10px;
  padding: 1.1rem; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,.04); }
.mc-val { font-family: 'JetBrains Mono', monospace; font-size: 2rem;
  font-weight: 700; color: #1a3a5c; line-height: 1; }
.mc-lbl { font-size: .72rem; color: #6b7280; text-transform: uppercase;
  letter-spacing: .8px; margin-top: .35rem; }

/* ── alert boxes ── */
.ab { border-radius: 8px; padding: .9rem 1.15rem; margin: .6rem 0;
  font-size: .875rem; line-height: 1.65; border-left-width: 4px; border-left-style: solid; }
.ab-info  { background: #eff6ff; border-color: #2563eb; border: 1px solid #bfdbfe;
  border-left: 4px solid #2563eb; color: #1e3a5f; }
.ab-ok    { background: #f0fdf4; border: 1px solid #bbf7d0;
  border-left: 4px solid #16a34a; color: #14532d; }
.ab-warn  { background: #fffbeb; border: 1px solid #fde68a;
  border-left: 4px solid #f59e0b; color: #78350f; }
.ab-err   { background: #fef2f2; border: 1px solid #fecaca;
  border-left: 4px solid #dc2626; color: #7f1d1d; }

/* ── ANOVA table ── */
.spss-wrap { overflow-x: auto; border-radius: 8px;
  border: 1px solid #dde3ed; margin: .75rem 0; }
.spss-table { width: 100%; border-collapse: collapse; font-size: .84rem;
  font-family: 'JetBrains Mono', monospace; }
.spss-table caption { text-align: left; font-weight: 600; color: #1a3a5c;
  padding: .6rem .8rem; font-size: .85rem; background: #f8faff;
  border-bottom: 1px solid #dde3ed; }
.spss-table thead th { background: #1a3a5c; color: #fff; padding: 9px 14px;
  text-align: left; white-space: nowrap; font-weight: 600; font-size: .8rem; }
.spss-table tbody td { padding: 8px 14px; border-bottom: 1px solid #edf0f7; color: #1f2937; }
.spss-table tbody tr:last-child td { border-bottom: none; }
.spss-table tbody tr:nth-child(even) td { background: #f8faff; }
.spss-table tbody tr:hover td { background: #eef4ff; }
.spss-table .err-row td { color: #9ca3af; font-style: italic; background: #fafafa; }
.spss-table .err-row:hover td { background: #f3f4f6; }
.cell-sig  { color: #059669; font-weight: 700; }
.cell-sig2 { color: #2563eb; font-weight: 700; }
.cell-sig3 { color: #7c3aed; font-weight: 700; }
.cell-marg { color: #d97706; font-weight: 600; }
.cell-ns   { color: #6b7280; }

/* ── interpretation ── */
.interp-card { background: #fff; border: 1px solid #dde3ed; border-radius: 10px;
  padding: 1.25rem 1.5rem; margin: .75rem 0;
  box-shadow: 0 1px 3px rgba(0,0,0,.04); }
.interp-heading { font-size: .95rem; font-weight: 700; color: #1a3a5c;
  margin: 0 0 .55rem; border-bottom: 2px solid #e9ecf3; padding-bottom: .4rem; }
.interp-body { font-size: .875rem; line-height: 1.75; color: #374151; }
.interp-sig  { color: #059669; font-weight: 600; }
.interp-ns   { color: #6b7280; font-weight: 600; }
.interp-warn { color: #d97706; font-weight: 600; }

/* ── posthoc ── */
.ph-header { font-size: .9rem; font-weight: 700; color: #1a3a5c;
  margin: 1.1rem 0 .4rem; border-left: 3px solid #3b9eff; padding-left: .6rem; }
.ph-card { background: #f8faff; border: 1px solid #dde3ed; border-radius: 8px;
  padding: 1rem 1.2rem; margin: .5rem 0; font-size: .86rem; line-height: 1.7; color: #374151; }
.ph-label { font-weight: 600; color: #1a3a5c; margin-bottom: .25rem; }
.ph-sig-yes { color: #059669; font-weight: 700; }
.ph-sig-no  { color: #6b7280; }

/* ── APA box ── */
.apa-box { background: #f0f7ff; border: 1px solid #bfdbfe; border-radius: 8px;
  padding: 1rem 1.2rem; font-size: .875rem; line-height: 1.7; color: #1e3a5f;
  font-style: italic; }

/* ── effect size legend ── */
.es-legend { display: flex; gap: .75rem; flex-wrap: wrap; margin: .5rem 0; }
.es-chip { padding: 3px 10px; border-radius: 20px; font-size: .75rem;
  font-weight: 600; font-family: 'JetBrains Mono', monospace; }
.es-neg  { background: #f3f4f6; color: #6b7280; border: 1px solid #d1d5db; }
.es-sm   { background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }
.es-med  { background: #fffbeb; color: #d97706; border: 1px solid #fde68a; }
.es-lg   { background: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0; }

/* ── sidebar ── */
[data-testid="stSidebar"] { background: #1a3a5c !important; }
[data-testid="stSidebar"] * { color: #d1e4f5 !important; }
[data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: #2d5278 !important; }
[data-testid="stSidebar"] .stDownloadButton > button {
  background: #2d5278 !important; border-color: #3b7ab8 !important;
  color: #d1e4f5 !important; width: 100%; font-size: .8rem; }
[data-testid="stSidebar"] .stDownloadButton > button:hover {
  background: #3b7ab8 !important; }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { font-size: .82rem !important; }

/* ── tabs ── */
.stTabs [data-baseweb="tab-list"] { background: transparent; gap: 3px; border-bottom: 2px solid #dde3ed; }
.stTabs [data-baseweb="tab"] { background: transparent; border: none; color: #6b7280;
  font-size: .85rem; font-weight: 500; border-radius: 0; padding: .6rem .9rem;
  border-bottom: 2px solid transparent; margin-bottom: -2px; }
.stTabs [aria-selected="true"] { background: transparent !important; color: #1a3a5c !important;
  border-bottom: 2px solid #1a3a5c !important; font-weight: 700 !important; }

/* ── button ── */
.stButton > button { background: #1a3a5c; color: #fff; border: none; border-radius: 8px;
  font-size: .875rem; font-weight: 600; padding: .6rem 1.4rem; transition: background .18s; }
.stButton > button:hover { background: #24527a; color: #fff; }

/* ── section label ── */
.step-label { display: flex; align-items: center; gap: .5rem;
  font-size: .78rem; font-weight: 700; color: #1a3a5c;
  text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: .9rem; }
.step-num { background: #1a3a5c; color: #fff; border-radius: 50%;
  width: 20px; height: 20px; display: flex; align-items: center;
  justify-content: center; font-size: .7rem; font-weight: 700; flex-shrink: 0; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ─────────────────────────────────────────────────────────────────────────────
def fmt(x, d=3):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    if isinstance(x, (int, np.integer)):
        return str(x)
    return f"{x:.{d}f}"

def fmt_p(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "—"
    if p < .001: return "< .001"
    return f"{p:.3f}"

def sig_star(p):
    if isinstance(p, float) and np.isnan(p): return ""
    if p < .001: return "***"
    if p < .01:  return "**"
    if p < .05:  return "*"
    if p < .10:  return "†"
    return "ns"

def sig_cls(p):
    if isinstance(p, float) and np.isnan(p): return "cell-ns"
    if p < .001: return "cell-sig3"
    if p < .01:  return "cell-sig2"
    if p < .05:  return "cell-sig"
    if p < .10:  return "cell-marg"
    return "cell-ns"

def eta_label(e):
    if isinstance(e, float) and np.isnan(e): return "—"
    if e >= .14: return "Large"
    if e >= .06: return "Medium"
    if e >= .01: return "Small"
    return "Negligible"

def design_str(n_groups, n_times):
    gw = {2:"Two",3:"Three",4:"Four",5:"Five",6:"Six",7:"Seven",8:"Eight"}.get(n_groups, str(n_groups))
    tw = {2:"Two",3:"Three",4:"Four",5:"Five",6:"Six",7:"Seven",8:"Eight"}.get(n_times, str(n_times))
    return f"{gw} Groups x {tw} Time Points"


# ─────────────────────────────────────────────────────────────────────────────
# SAMPLE DATA  (4 variants to cover all common designs)
# ─────────────────────────────────────────────────────────────────────────────
def make_sample(n_groups, n_times):
    np.random.seed(2024)
    group_cfg = {
        2: [("Control", [0, 2]), ("Treatment", [0, 8])],
        3: [("Control", [0, 1, 2]), ("Dose_A", [0, 5, 9]), ("Dose_B", [0, 9, 18])],
    }
    time_names_cfg = {
        2: ["Pre", "Post"],
        3: ["Pre", "Mid", "Post"],
        4: ["Baseline", "Week4", "Week8", "Week12"],
    }
    groups = group_cfg.get(n_groups,
        [(f"Group{i+1}", list(range(0, 4*n_times, 4))) for i in range(n_groups)])
    times  = time_names_cfg.get(n_times,
        [f"T{i+1}" for i in range(n_times)])
    n_per  = 15
    rows   = []
    for g_name, g_effects in groups:
        for i in range(n_per):
            row = {"ID": f"{g_name[0]}{i+1:02d}", "Group": g_name}
            base = np.random.normal(50, 10)
            for j, t in enumerate(times):
                effect = g_effects[j] if j < len(g_effects) else g_effects[-1]
                row[t] = round(base + np.random.normal(effect, 8), 2)
            rows.append(row)
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# VALIDATION
# ─────────────────────────────────────────────────────────────────────────────
def validate_data(df, subj_col, btwn_col, time_cols):
    issues = []

    missing_cols = [c for c in [subj_col, btwn_col] + time_cols if c not in df.columns]
    if missing_cols:
        issues.append(("error",
            f"The following columns were not found in the uploaded file: "
            f"{', '.join(missing_cols)}. "
            f"Available columns are: {', '.join(df.columns.tolist())}."))
        return issues

    for c in time_cols:
        if not pd.api.types.is_numeric_dtype(df[c]):
            sample_vals = df[c].dropna().head(3).tolist()
            issues.append(("error",
                f"Column '{c}' must contain numeric values, but non-numeric data were found "
                f"(sample values: {sample_vals}). "
                f"Please ensure all time-point columns contain numbers only."))

    n_groups = df[btwn_col].nunique()
    if n_groups < 2:
        issues.append(("error",
            f"The between-subjects factor '{btwn_col}' contains only {n_groups} unique value. "
            f"A Mixed-Design ANOVA requires at least two independent groups."))
    elif n_groups > 10:
        issues.append(("warning",
            f"The between-subjects factor '{btwn_col}' contains {n_groups} groups, which is "
            f"unusually high. Please verify that this column contains categorical group labels "
            f"and not a continuous variable."))

    group_sizes = df.groupby(btwn_col)[time_cols[0]].count()
    for g, n in group_sizes.items():
        if n < 3:
            issues.append(("error",
                f"Group '{g}' contains only {n} participant(s). "
                f"Each group must have at least 3 participants for a valid analysis."))
        elif n < 10:
            issues.append(("warning",
                f"Group '{g}' contains only {n} participants. "
                f"Small sample sizes reduce statistical power and may compromise assumption tests."))

    req = [subj_col, btwn_col] + time_cols
    n_miss = df[req].isnull().any(axis=1).sum()
    if n_miss > 0:
        pct = 100 * n_miss / len(df)
        lv  = "warning" if pct > 20 else "info"
        issues.append((lv,
            f"{n_miss} of {len(df)} rows ({pct:.1f}%) contain missing values and will be "
            f"excluded using listwise deletion. "
            f"The analysis will proceed with {len(df) - n_miss} complete cases."))

    dupes = df[subj_col].duplicated()
    if dupes.any():
        dup_ids = df.loc[dupes, subj_col].unique().tolist()[:5]
        issues.append(("warning",
            f"Duplicate Subject IDs were detected: {dup_ids}. "
            f"In wide-format data, each participant should appear in exactly one row."))

    sizes = group_sizes.values
    if len(sizes) > 1 and sizes.max() > sizes.min():
        issues.append(("info",
            f"The design is unbalanced: group sizes range from {sizes.min()} to {sizes.max()}. "
            f"The analysis accommodates unbalanced designs, but balanced designs yield greater power."))

    if not issues:
        issues.append(("info",
            f"Data validation passed. "
            f"Design detected: {design_str(n_groups, len(time_cols))} "
            f"with {len(df)} participants."))
    return issues


# ─────────────────────────────────────────────────────────────────────────────
# DESCRIPTIVES
# ─────────────────────────────────────────────────────────────────────────────
def compute_descriptives(df_wide, between_col, time_cols):
    rows = []
    for g in df_wide[between_col].unique():
        sub = df_wide[df_wide[between_col] == g]
        for t in time_cols:
            v  = sub[t].dropna()
            n  = len(v)
            m  = v.mean()
            sd = v.std(ddof=1)
            se = sd / np.sqrt(n) if n > 1 else np.nan
            rows.append({"Group": g, "Time": t, "N": n,
                         "Mean": round(m, 3), "SD": round(sd, 3),
                         "SE": round(se, 3),
                         "Median": round(v.median(), 3),
                         "Min": round(v.min(), 3), "Max": round(v.max(), 3),
                         "95% CI Lower": round(m - 1.96*se, 3),
                         "95% CI Upper": round(m + 1.96*se, 3)})
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# NORMALITY
# ─────────────────────────────────────────────────────────────────────────────
def test_normality(df_wide, between_col, time_cols):
    rows = []
    for g in df_wide[between_col].unique():
        sub = df_wide[df_wide[between_col] == g]
        for t in time_cols:
            v = sub[t].dropna().values
            n = len(v)
            if n < 3:
                rows.append({"Group": g, "Time": t, "N": n, "Test": "—",
                              "Statistic": np.nan, "p-value": np.nan,
                              "Normal?": "—", "Test Selection": "n < 3, untestable"})
                continue
            if n <= 50:
                stat, p = stats.shapiro(v)
                tname   = "Shapiro-Wilk"
                note    = "n ≤ 50: SW has optimal power for small samples"
            else:
                stat, p = stats.kstest(v, "norm", args=(v.mean(), v.std()))
                tname   = "Kolmogorov-Smirnov"
                note    = "n > 50: KS used (SW loses sensitivity at large n)"
            rows.append({"Group": g, "Time": t, "N": n, "Test": tname,
                         "Statistic": round(stat, 4), "p-value": round(p, 4),
                         "Normal?": "Yes" if p > .05 else "No",
                         "Test Selection": note})
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# LEVENE
# ─────────────────────────────────────────────────────────────────────────────
def test_levene(df_wide, between_col, time_cols):
    rows = []
    for t in time_cols:
        gdata = [df_wide[df_wide[between_col] == g][t].dropna().values
                 for g in df_wide[between_col].unique()]
        if all(len(x) > 1 for x in gdata):
            stat, p = scipy_levene(*gdata, center="mean")
            rows.append({"Time Point": t, "Levene F": round(stat, 3),
                         "df1": len(gdata) - 1,
                         "df2": sum(len(x) for x in gdata) - len(gdata),
                         "p-value": round(p, 4),
                         "Equal Variances?": "Yes" if p > .05 else "No"})
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# MAUCHLY SPHERICITY
# ─────────────────────────────────────────────────────────────────────────────
def mauchly_test(df_wide, between_col, time_cols):
    b = len(time_cols)
    if b < 3:
        return None
    Y = df_wide[time_cols].values.astype(float)
    N = Y.shape[0]
    C = np.zeros((b, b - 1))
    for j in range(b - 1):
        C[:j+1, j] = 1.0 / (j + 1)
        C[j+1,  j] = -1.0
    T = Y @ C
    S = np.cov(T, rowvar=False)
    if S.ndim == 0:
        S = np.array([[S]])
    try:
        k     = b - 1
        det_S = np.linalg.det(S)
        tr_S  = np.trace(S)
        W     = max(1e-15, min(det_S / (tr_S / k) ** k, 1.0))
        df_c  = k * (k + 1) // 2 - 1
        u     = N - 1 - (2*k**2 + k + 2) / (6 * k)
        chi2  = -u * np.log(W)
        p_W   = float(1 - stats.chi2.cdf(chi2, df_c)) if df_c > 0 else np.nan
    except Exception:
        return None
    tr_S2  = np.trace(S @ S)
    eps_GG = min(max(tr_S**2 / (k * tr_S2), 1 / k), 1.0)
    eps_HF = min((N * (k * eps_GG) - 2) / (k * (N - 1 - k * eps_GG)), 1.0)
    return W, p_W, eps_GG, eps_HF, chi2, df_c


# ─────────────────────────────────────────────────────────────────────────────
# MIXED ANOVA
# ─────────────────────────────────────────────────────────────────────────────
def mixed_anova(df_wide, subject_col, between_col, time_cols):
    groups = df_wide[between_col].unique()
    a = len(groups)
    b = len(time_cols)

    gd = {g: df_wide[df_wide[between_col] == g][time_cols].values.astype(float) for g in groups}
    ng = {g: gd[g].shape[0] for g in groups}
    N  = sum(ng.values())
    all_scores = np.vstack([gd[g] for g in groups])
    GM  = np.mean(all_scores)
    gm  = {g: np.mean(gd[g]) for g in groups}
    tm  = np.mean(all_scores, axis=0)
    cm  = {g: np.mean(gd[g], axis=0) for g in groups}
    sm  = {g: np.mean(gd[g], axis=1) for g in groups}

    SS_A   = b * sum(ng[g] * (gm[g] - GM)**2 for g in groups)
    df_A   = a - 1
    SS_sWG = sum(b * np.sum((sm[g] - gm[g])**2) for g in groups)
    df_sWG = sum(ng[g] - 1 for g in groups)
    SS_B   = N * np.sum((tm - GM)**2)
    df_B   = b - 1
    SS_AB  = sum(ng[g] * np.sum((cm[g] - gm[g] - tm + GM)**2) for g in groups)
    df_AB  = (a - 1) * (b - 1)
    SS_eW  = sum((gd[g][i, j] - sm[g][i] - cm[g][j] + gm[g])**2
                 for g in groups for i in range(ng[g]) for j in range(b))
    df_eW  = sum(ng[g] - 1 for g in groups) * (b - 1)

    def ms(ss, df): return ss / df if df > 0 else np.nan
    MS_A   = ms(SS_A, df_A);    MS_sWG = ms(SS_sWG, df_sWG)
    MS_B   = ms(SS_B, df_B);    MS_AB  = ms(SS_AB, df_AB)
    MS_eW  = ms(SS_eW, df_eW)

    F_A  = MS_A  / MS_sWG if MS_sWG else np.nan
    F_B  = MS_B  / MS_eW  if MS_eW  else np.nan
    F_AB = MS_AB / MS_eW  if MS_eW  else np.nan

    def pf(F, d1, d2):
        return float(1 - stats.f.cdf(F, d1, d2)) if not np.isnan(F) else np.nan
    p_A  = pf(F_A,  df_A,  df_sWG)
    p_B  = pf(F_B,  df_B,  df_eW)
    p_AB = pf(F_AB, df_AB, df_eW)

    def eta2(sse, ssr):
        return sse / (sse + ssr) if (sse + ssr) > 0 else np.nan
    e_A  = eta2(SS_A,  SS_sWG)
    e_B  = eta2(SS_B,  SS_eW)
    e_AB = eta2(SS_AB, SS_eW)

    sph    = mauchly_test(df_wide, between_col, time_cols)
    eps_GG = 1.0
    if sph:
        _, _, eps_GG, _, _, _ = sph

    def gg_p(F, d1, d2, eps):
        return float(1 - stats.f.cdf(F, d1 * eps, d2 * eps)) if not np.isnan(F) else np.nan

    p_B_gg  = gg_p(F_B,  df_B,  df_eW, eps_GG)
    p_AB_gg = gg_p(F_AB, df_AB, df_eW, eps_GG)

    results = pd.DataFrame([
        {"Source": between_col, "SS": SS_A, "df": df_A, "df_err": df_sWG,
         "MS": MS_A, "F": F_A, "p": p_A, "p_GG": p_A,
         "eta_p2": e_A, "eps": np.nan, "rowtype": "between"},
        {"Source": "Error (Between)", "SS": SS_sWG, "df": df_sWG, "df_err": np.nan,
         "MS": MS_sWG, "F": np.nan, "p": np.nan, "p_GG": np.nan,
         "eta_p2": np.nan, "eps": np.nan, "rowtype": "err_between"},
        {"Source": "Time", "SS": SS_B, "df": df_B, "df_err": df_eW,
         "MS": MS_B, "F": F_B, "p": p_B, "p_GG": p_B_gg,
         "eta_p2": e_B, "eps": eps_GG, "rowtype": "within"},
        {"Source": f"{between_col} × Time", "SS": SS_AB, "df": df_AB, "df_err": df_eW,
         "MS": MS_AB, "F": F_AB, "p": p_AB, "p_GG": p_AB_gg,
         "eta_p2": e_AB, "eps": eps_GG, "rowtype": "interaction"},
        {"Source": "Error (Within)", "SS": SS_eW, "df": df_eW, "df_err": np.nan,
         "MS": MS_eW, "F": np.nan, "p": np.nan, "p_GG": np.nan,
         "eta_p2": np.nan, "eps": np.nan, "rowtype": "err_within"},
    ])
    return results, sph


# ─────────────────────────────────────────────────────────────────────────────
# P ADJUSTMENT
# ─────────────────────────────────────────────────────────────────────────────
def _adjust_p(ps, method):
    ps = np.array(ps, dtype=float)
    n  = len(ps)
    if n == 0:
        return ps
    if method == "bonferroni":
        return np.minimum(ps * n, 1.0)
    if method == "holm":
        order = np.argsort(ps)
        adj   = ps.copy()
        for rank, idx in enumerate(order):
            adj[idx] = min(ps[idx] * (n - rank), 1.0)
        for i in range(1, n):
            adj[order[i]] = max(adj[order[i]], adj[order[i-1]])
        return np.minimum(adj, 1.0)
    if method == "fdr_bh":
        order = np.argsort(ps)[::-1]
        adj   = ps.copy()
        mn    = 1.0
        for i, idx in enumerate(order):
            adj[idx] = min(ps[idx] * n / (n - i), mn)
            mn = adj[idx]
        return np.minimum(adj, 1.0)
    return ps  # tukey handled separately


def _tukey_p(q_stat, k, df_err):
    try:
        from scipy.stats import studentized_range
        return float(1 - studentized_range.cdf(abs(q_stat) * np.sqrt(2), k, df_err))
    except Exception:
        return min(2 * (1 - stats.t.cdf(abs(q_stat), df_err)) * k * (k-1) / 2, 1.0)


# ─────────────────────────────────────────────────────────────────────────────
# POST-HOC
# ─────────────────────────────────────────────────────────────────────────────
def run_posthoc(df_wide, between_col, time_cols, method, alpha):
    groups   = sorted(df_wide[between_col].unique())
    n_grps   = len(groups)
    rows     = []

    # 1. Between-groups at each time point
    for t in time_cols:
        pairs   = list(combinations(groups, 2))
        raw_ts, raw_ps, meta = [], [], []
        for g1, g2 in pairs:
            v1 = df_wide[df_wide[between_col] == g1][t].dropna().values
            v2 = df_wide[df_wide[between_col] == g2][t].dropna().values
            ts, pr = stats.ttest_ind(v1, v2, equal_var=True)
            raw_ts.append(ts); raw_ps.append(pr)
            meta.append((g1, g2, v1, v2, len(v1)+len(v2)-2))

        if method == "tukey":
            all_v  = [df_wide[df_wide[between_col] == g][t].dropna().values for g in groups]
            n_pool = sum(len(v) for v in all_v)
            df_e   = n_pool - n_grps
            ms_e   = (sum((len(v)-1)*np.var(v, ddof=1) for v in all_v) / df_e) if df_e > 0 else np.nan
            adj_ps = []
            for g1, g2, v1, v2, _ in meta:
                if ms_e and ms_e > 0:
                    q = abs(np.mean(v1) - np.mean(v2)) / np.sqrt(ms_e * 0.5 * (1/len(v1) + 1/len(v2)))
                    adj_ps.append(_tukey_p(q, n_grps, df_e))
                else:
                    adj_ps.append(1.0)
        else:
            adj_ps = list(_adjust_p(raw_ps, method))

        for i, (g1, g2, v1, v2, dft) in enumerate(meta):
            md   = np.mean(v1) - np.mean(v2)
            pvar = ((len(v1)-1)*np.var(v1,ddof=1) + (len(v2)-1)*np.var(v2,ddof=1)) / (len(v1)+len(v2)-2)
            d    = md / np.sqrt(pvar) if pvar > 0 else np.nan
            se   = np.sqrt(np.var(v1,ddof=1)/len(v1) + np.var(v2,ddof=1)/len(v2))
            rows.append({
                "ctype": "Between Groups", "timepoint": t, "group": "—",
                "label_a": g1, "label_b": g2,
                "mean_a": round(np.mean(v1), 3), "mean_b": round(np.mean(v2), 3),
                "meandiff": round(md, 3), "se": round(se, 3),
                "t": round(raw_ts[i], 3), "df": int(dft),
                "p_raw": round(raw_ps[i], 4), "p_adj": round(adj_ps[i], 4),
                "cohen_d": round(d, 3) if not np.isnan(d) else np.nan,
                "sig": sig_star(adj_ps[i]), "significant": adj_ps[i] < alpha})

    # 2. Within-subject time comparisons per group
    time_pairs = list(combinations(time_cols, 2))
    for g in groups:
        sub = df_wide[df_wide[between_col] == g]
        raw_ts2, raw_ps2, meta2 = [], [], []
        for t1, t2 in time_pairs:
            v1 = sub[t1].dropna().values; v2 = sub[t2].dropna().values
            nm = min(len(v1), len(v2))
            ts, pr = stats.ttest_rel(v1[:nm], v2[:nm])
            raw_ts2.append(ts); raw_ps2.append(pr)
            meta2.append((t1, t2, v1[:nm], v2[:nm], nm-1))

        n_t = len(time_cols)
        if method == "tukey":
            adj_ps2 = []
            for i, (t1, t2, v1, v2, dft) in enumerate(meta2):
                adj_ps2.append(_tukey_p(abs(raw_ts2[i]) * np.sqrt(2), n_t, dft))
        else:
            adj_ps2 = list(_adjust_p(raw_ps2, method))

        for i, (t1, t2, v1, v2, dft) in enumerate(meta2):
            diff = v1 - v2
            se   = np.std(diff, ddof=1) / np.sqrt(len(diff)) if len(diff) > 1 else np.nan
            d    = np.mean(diff) / np.std(diff, ddof=1) if np.std(diff, ddof=1) > 0 else np.nan
            rows.append({
                "ctype": "Within Time", "timepoint": "—", "group": g,
                "label_a": t1, "label_b": t2,
                "mean_a": round(np.mean(v1), 3), "mean_b": round(np.mean(v2), 3),
                "meandiff": round(np.mean(diff), 3),
                "se": round(se, 3) if not np.isnan(se) else np.nan,
                "t": round(raw_ts2[i], 3), "df": int(dft),
                "p_raw": round(raw_ps2[i], 4), "p_adj": round(adj_ps2[i], 4),
                "cohen_d": round(d, 3) if not np.isnan(d) else np.nan,
                "sig": sig_star(adj_ps2[i]), "significant": adj_ps2[i] < alpha})

    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ─────────────────────────────────────────────────────────────────────────────
# INTERPRETATION BLOCKS
# ─────────────────────────────────────────────────────────────────────────────
def build_interpretation(aov_df, between_col, n_total, alpha, group_names, time_cols):
    rd = {r["Source"]: r for _, r in aov_df.iterrows()}
    ab_src = f"{between_col} × Time"

    def info(src):
        r = rd.get(src, {})
        return (r.get("F", np.nan), r.get("p_GG", r.get("p", np.nan)),
                r.get("eta_p2", np.nan), r.get("df", np.nan), r.get("df_err", np.nan))

    F_A,  p_A,  e_A,  df_A,  dfe_A  = info(between_col)
    F_B,  p_B,  e_B,  df_B,  dfe_B  = info("Time")
    F_AB, p_AB, e_AB, df_AB, dfe_AB  = info(ab_src)

    n_g = len(group_names)
    n_t = len(time_cols)
    design = design_str(n_g, n_t)
    g_list = ", ".join(f"'{g}'" for g in group_names)
    t_list = ", ".join(f"'{t}'" for t in time_cols)

    blocks = []

    # Overview
    blocks.append({
        "heading": "Study Design Overview",
        "type": "info",
        "body": (
            f"A Mixed-Design (Split-Plot) ANOVA was conducted with a <strong>{design}</strong> design. "
            f"The between-subjects factor comprised group membership ({g_list}), "
            f"and the within-subjects factor represented the measurement occasion ({t_list}). "
            f"A total of <strong>N = {n_total}</strong> participants contributed complete data. "
            f"The significance threshold was set at &alpha; = {alpha}."
        )
    })

    # Between-subjects main effect
    if not np.isnan(p_A):
        sig_w = "statistically significant" if p_A < alpha else "not statistically significant"
        mag   = eta_label(e_A)
        if p_A < alpha:
            conc = (f"This indicates that, averaged across all {n_t} time points, "
                    f"the groups differed significantly in their overall mean score on the dependent variable.")
        else:
            conc = (f"This indicates that, averaged across all {n_t} time points, "
                    f"the groups did not differ significantly in their overall mean performance.")
        blocks.append({
            "heading": "Main Effect of Group (Between-Subjects Factor)",
            "type": "between",
            "sig": p_A < alpha,
            "body": (
                f"The main effect of the between-subjects factor (<em>Group</em>) was "
                f"<strong>{sig_w}</strong>, "
                f"<em>F</em>({fmt(df_A, 0)}, {fmt(dfe_A, 0)}) = {fmt(F_A, 2)}, "
                f"<em>p</em> = {fmt_p(p_A)}, "
                f"partial &eta;<sup>2</sup> = {fmt(e_A, 3)} ({mag} effect; Cohen, 1988). "
                f"{conc}"
            )
        })

    # Within-subjects main effect
    if not np.isnan(p_B):
        sig_w = "statistically significant" if p_B < alpha else "not statistically significant"
        mag   = eta_label(e_B)
        if p_B < alpha:
            conc = (f"This indicates that, averaged across all {n_g} groups, "
                    f"scores changed significantly across the {n_t} measurement occasions.")
        else:
            conc = (f"This indicates that, averaged across all {n_g} groups, "
                    f"scores did not change significantly over time.")
        blocks.append({
            "heading": "Main Effect of Time (Within-Subjects Factor)",
            "type": "within",
            "sig": p_B < alpha,
            "body": (
                f"The main effect of <em>Time</em> was <strong>{sig_w}</strong>, "
                f"<em>F</em>({fmt(df_B, 0)}, {fmt(dfe_B, 0)}) = {fmt(F_B, 2)}, "
                f"<em>p</em> = {fmt_p(p_B)}, "
                f"partial &eta;<sup>2</sup> = {fmt(e_B, 3)} ({mag} effect). "
                f"{conc}"
            )
        })

    # Interaction
    if not np.isnan(p_AB):
        mag = eta_label(e_AB)
        if p_AB < alpha:
            heading = "Group × Time Interaction — Statistically Significant"
            itype   = "sig"
            conc = (
                f"This significant interaction indicates that the trajectory of change over time "
                f"<strong>differed between the {n_g} groups</strong>. "
                f"The groups did not follow the same pattern of change across the {n_t} measurement occasions; "
                f"some groups changed more rapidly or in different directions than others. "
                f"Because the interaction is significant, the main effects of Group and Time "
                f"should be interpreted with caution — their interpretation depends on the level "
                f"of the other factor. "
                f"Post-hoc pairwise comparisons (reported below) identify the specific contrasts driving this pattern."
            )
        elif p_AB < .10:
            heading = "Group × Time Interaction — Marginal Trend (p < .10)"
            itype   = "marginal"
            conc = (
                f"A marginal trend toward a Group × Time interaction was observed. "
                f"This did not reach the conventional significance threshold (&alpha; = {alpha}), "
                f"but researchers may wish to examine the pattern of means in the profile plot "
                f"and consider the practical significance of the effect size "
                f"(partial &eta;<sup>2</sup> = {fmt(e_AB, 3)})."
            )
        else:
            heading = "Group × Time Interaction — Not Statistically Significant"
            itype   = "ns"
            conc = (
                f"The non-significant interaction indicates that the {n_g} groups "
                f"followed a <strong>similar pattern of change</strong> across the {n_t} time points. "
                f"Consequently, the main effects of Group and Time can be interpreted independently."
            )
        blocks.append({
            "heading": heading,
            "type": itype,
            "sig": p_AB < alpha,
            "body": (
                f"The Group × Time interaction was "
                f"<strong>{'statistically significant' if p_AB < alpha else 'not statistically significant'}</strong>, "
                f"<em>F</em>({fmt(df_AB, 0)}, {fmt(dfe_AB, 0)}) = {fmt(F_AB, 2)}, "
                f"<em>p</em> = {fmt_p(p_AB)}, "
                f"partial &eta;<sup>2</sup> = {fmt(e_AB, 3)} ({mag} effect size). "
                f"{conc}"
            )
        })

    return blocks, ab_src, p_AB


def build_apa(aov_df, between_col, n_total, n_groups, time_cols, alpha):
    ab_src = f"{between_col} × Time"
    rd     = {r["Source"]: r for _, r in aov_df.iterrows()}
    r      = rd.get(ab_src, {})
    F      = r.get("F", np.nan)
    p      = r.get("p_GG", r.get("p", np.nan))
    e      = r.get("eta_p2", np.nan)
    df1    = r.get("df", np.nan)
    df2    = r.get("df_err", np.nan)
    sig    = "significant" if (not np.isnan(p) and p < alpha) else "not significant"
    design = design_str(n_groups, len(time_cols))
    return (f"A {design} Mixed-Design ANOVA was conducted (N = {n_total}). "
            f"The Group × Time interaction was {sig}: "
            f"F({fmt(df1,0)}, {fmt(df2,0)}) = {fmt(F,2)}, "
            f"p = {fmt_p(p)}, partial η² = {fmt(e,3)}.")


# ─────────────────────────────────────────────────────────────────────────────
# ANOVA TABLE HTML
# ─────────────────────────────────────────────────────────────────────────────
def render_anova_table(aov_df, caption, use_gg):
    rows_html = ""
    for _, r in aov_df.iterrows():
        src  = r["Source"]
        ss   = r.get("SS",  np.nan)
        df_v = r.get("df",  np.nan)
        ms   = r.get("MS",  np.nan)
        F    = r.get("F",   np.nan)
        p    = r["p_GG"] if use_gg else r["p"]
        eta  = r.get("eta_p2", np.nan)
        eps  = r.get("eps", np.nan)
        rt   = r.get("rowtype", "")

        is_err = "err" in rt
        rc     = ' class="err-row"' if is_err else ""
        F_str  = fmt(F, 3) if not (isinstance(F, float) and np.isnan(F)) else "—"
        p_str  = fmt_p(p)
        star   = sig_star(p) if not (isinstance(p, float) and np.isnan(p)) else ""
        cls    = sig_cls(p)  if not (isinstance(p, float) and np.isnan(p)) else "cell-ns"
        e_str  = fmt(eta, 3) if not (isinstance(eta, float) and np.isnan(eta)) else "—"
        eps_str= fmt(eps, 3) if not (isinstance(eps, float) and np.isnan(eps)) else "—"

        rows_html += (
            f'<tr{rc}><td><strong>{src}</strong></td>'
            f'<td>{fmt(ss,3)}</td><td>{fmt(df_v,0)}</td><td>{fmt(ms,3)}</td>'
            f'<td>{F_str}</td>'
            f'<td class="{cls}">{p_str} {star}</td>'
            f'<td>{e_str}</td><td>{eps_str}</td></tr>'
        )

    return f"""
    <div class="spss-wrap">
    <table class="spss-table">
      <caption>{caption}</caption>
      <thead><tr>
        <th>Source</th><th>SS</th><th>df</th><th>MS</th>
        <th>F</th><th>p</th><th>Partial &eta;<sup>2</sup></th><th>&epsilon; (GG)</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    </div>
    <p style="font-size:.75rem;color:#6b7280;margin:.25rem 0 .75rem;">
    *** p &lt; .001 &ensp; ** p &lt; .01 &ensp; * p &lt; .05 &ensp; † p &lt; .10 &ensp; ns p &ge; .10 &ensp;
    <em>Italicised rows = error terms (not tested)</em>
    </p>"""


# ─────────────────────────────────────────────────────────────────────────────
# VISUALIZATION  – always light background
# ─────────────────────────────────────────────────────────────────────────────
PAL = ["#2563eb","#16a34a","#9333ea","#ea580c","#0891b2","#be123c","#854d0e"]

def _ax_style(ax, title="", xlabel="Time Point", ylabel="Mean Score"):
    ax.set_facecolor("#fdfdff")
    for sp in ["top","right"]:  ax.spines[sp].set_visible(False)
    for sp in ["bottom","left"]: ax.spines[sp].set_color("#d1d5db")
    ax.tick_params(colors="#374151", labelsize=9)
    ax.grid(axis="y", color="#e5e7eb", linewidth=.8, linestyle="--")
    if xlabel: ax.set_xlabel(xlabel, color="#4b5563", fontsize=9.5, labelpad=5)
    if ylabel: ax.set_ylabel(ylabel, color="#4b5563", fontsize=9.5, labelpad=5)
    if title:  ax.set_title(title, color="#1a3a5c", fontsize=10.5, fontweight="bold", pad=8)

def profile_plot(desc, between_col, dv_label):
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor("#ffffff")
    groups  = desc["Group"].unique()
    times_u = list(dict.fromkeys(desc["Time"].tolist()))

    for ax in axes:
        _ax_style(ax, ylabel=dv_label)
        ax.set_xticks(range(len(times_u)))
        ax.set_xticklabels(times_u, fontsize=9)

    # Line plot
    ax = axes[0]
    _ax_style(ax, title="Profile Plot (Mean ± 1 SE)", ylabel=dv_label)
    for i, g in enumerate(groups):
        sub = desc[desc["Group"] == g].set_index("Time").loc[times_u].reset_index()
        ax.errorbar(range(len(sub)), sub["Mean"], yerr=sub["SE"],
                    marker="o", markersize=8, linewidth=2.5, color=PAL[i % len(PAL)],
                    label=str(g), capsize=5, capthick=1.5, elinewidth=1.5,
                    markeredgecolor="#fff", markeredgewidth=1.5)
    ax.set_xticks(range(len(times_u))); ax.set_xticklabels(times_u, fontsize=9)
    ax.legend(title=between_col, fontsize=8.5, title_fontsize=8.5,
              facecolor="white", edgecolor="#d1d5db", framealpha=1)

    # Bar plot
    ax = axes[1]
    _ax_style(ax, title="Bar Chart (Mean ± 95% CI)", ylabel=dv_label)
    ngrp = len(groups); w = 0.72 / ngrp; xs = np.arange(len(times_u))
    for i, g in enumerate(groups):
        sub = desc[desc["Group"] == g].set_index("Time").loc[times_u].reset_index()
        off = (i - ngrp/2 + .5) * w
        ax.bar(xs + off, sub["Mean"], width=w*0.88, color=PAL[i % len(PAL)],
               alpha=0.82, label=str(g), edgecolor="white", linewidth=.5)
        ax.errorbar(xs + off, sub["Mean"], yerr=1.96*sub["SE"],
                    fmt="none", color="#374151", capsize=3, capthick=1, elinewidth=1, alpha=.65)
    ax.set_xticks(xs); ax.set_xticklabels(times_u, fontsize=9)
    ax.legend(title=between_col, fontsize=8.5, title_fontsize=8.5,
              facecolor="white", edgecolor="#d1d5db", framealpha=1)

    plt.tight_layout(pad=2.5)
    return fig

def dist_plots(df_wide, between_col, time_cols, dv_label):
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    n   = len(time_cols)
    fig, axes = plt.subplots(1, n, figsize=(5.5*n, 4.5))
    fig.patch.set_facecolor("#ffffff")
    if n == 1: axes = [axes]
    groups = df_wide[between_col].unique()
    for ax, t in zip(axes, time_cols):
        _ax_style(ax, title=f"Distribution: {t}", xlabel=dv_label, ylabel="Frequency")
        for i, g in enumerate(groups):
            v = df_wide[df_wide[between_col] == g][t].dropna()
            ax.hist(v, bins=12, alpha=.45, color=PAL[i % len(PAL)],
                    label=str(g), edgecolor="white", linewidth=.4)
            if len(v) > 5:
                kde = stats.gaussian_kde(v)
                xr  = np.linspace(v.min(), v.max(), 200)
                sc  = len(v) * (v.max() - v.min()) / 12
                ax.plot(xr, kde(xr)*sc, color=PAL[i % len(PAL)], linewidth=2, alpha=.9)
    axes[0].legend(title=between_col, fontsize=8, facecolor="white", edgecolor="#d1d5db")
    plt.tight_layout()
    return fig

def qq_plots(df_wide, between_col, time_cols):
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    groups = df_wide[between_col].unique()
    nr, nc = len(groups), len(time_cols)
    fig, axes = plt.subplots(nr, nc, figsize=(4.5*nc, 3.5*nr), squeeze=False)
    fig.patch.set_facecolor("#ffffff")
    for r, g in enumerate(groups):
        for c, t in enumerate(time_cols):
            ax = axes[r][c]
            v  = df_wide[df_wide[between_col] == g][t].dropna().values
            ax.set_facecolor("#fdfdff")
            for sp in ["top","right"]: ax.spines[sp].set_visible(False)
            for sp in ["bottom","left"]: ax.spines[sp].set_color("#d1d5db")
            ax.tick_params(colors="#374151", labelsize=7)
            ax.grid(color="#e5e7eb", linewidth=.6, linestyle="--")
            if len(v) >= 3:
                (osm, osr), (_s, _int, _r2) = stats.probplot(v, dist="norm")
                ax.scatter(osm, osr, color=PAL[r % len(PAL)], s=25,
                           alpha=.75, edgecolors="white", linewidths=.5, zorder=3)
                q1, q3 = np.percentile(v, [25, 75])
                th1, th3 = stats.norm.ppf([.25, .75])
                slope = (q3-q1)/(th3-th1); inter = q1 - slope*th1
                xl = np.array([osm[0], osm[-1]])
                ax.plot(xl, slope*xl+inter, color="#dc2626", lw=1.2,
                        alpha=.8, linestyle="--", zorder=2, label="Reference")
            ax.set_title(f"{g}  —  {t}", color="#1a3a5c", fontsize=8.5, fontweight="bold", pad=4)
            ax.set_xlabel("Theoretical Quantiles", color="#6b7280", fontsize=7.5)
            ax.set_ylabel("Sample Quantiles", color="#6b7280", fontsize=7.5)
    fig.suptitle("Q–Q Plots for Normality Assessment",
                 color="#1a3a5c", fontsize=11, fontweight="bold", y=1.01)
    plt.tight_layout()
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────────────────────────────────────
def build_pdf(desc, norm_df, lev_df, aov_df, sph, interp_blocks, ph_df,
              between_col, dv_label, alpha, n_groups, time_cols, method_name, apa_text):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors as rl
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                         Table, TableStyle, HRFlowable)
        from reportlab.lib.enums import TA_JUSTIFY
        from reportlab.lib.units import cm
    except ImportError:
        return None

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             rightMargin=2*cm, leftMargin=2*cm,
                             topMargin=2.2*cm, bottomMargin=2*cm)

    NAVY  = rl.HexColor("#1a3a5c")
    BLUE  = rl.HexColor("#2563eb")
    LGREY = rl.HexColor("#f8faff")
    GREY  = rl.HexColor("#6b7280")

    sH1 = ParagraphStyle("H1", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY, spaceAfter=4)
    sH2 = ParagraphStyle("H2", fontName="Helvetica-Bold", fontSize=11, textColor=NAVY,
                          spaceAfter=3, spaceBefore=14)
    sH3 = ParagraphStyle("H3", fontName="Helvetica-Bold", fontSize=9.5, textColor=BLUE,
                          spaceAfter=2, spaceBefore=8)
    sBD = ParagraphStyle("BD", fontName="Helvetica", fontSize=9, leading=14,
                          spaceAfter=4, alignment=TA_JUSTIFY)
    sNT = ParagraphStyle("NT", fontName="Helvetica-Oblique", fontSize=7.5,
                          textColor=GREY, leading=11, spaceAfter=2)

    def safe(val):
        s = (str(round(val, 3)) if isinstance(val, float) and not np.isnan(val) else str(val))
        return (s.replace("—", "--").replace("×", "x").replace("η", "eta")
                 .replace("²", "2").replace("≥", ">=").replace("≤", "<=")
                 .replace("α", "alpha").replace("ε", "epsilon").replace("†", "+")
                 .replace("\u00b1", "+/-").replace("\u2013", "-").replace("\u2014", "--"))

    def mktable(df):
        data = [[safe(c) for c in df.columns]]
        for _, row in df.iterrows():
            data.append([safe(v) for v in row])
        t = Table(data, repeatRows=1, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(-1,0), NAVY), ("TEXTCOLOR",(0,0),(-1,0), rl.white),
            ("FONTNAME",   (0,0),(-1,0), "Helvetica-Bold"), ("FONTSIZE",(0,0),(-1,-1), 7.5),
            ("ALIGN",      (0,0),(-1,-1), "CENTER"), ("VALIGN",(0,0),(-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl.white, rl.HexColor("#f0f4fa")]),
            ("GRID",(0,0),(-1,-1), .25, rl.HexColor("#d1d5db")),
            ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),5),
        ]))
        return t

    story = []
    design = design_str(n_groups, len(time_cols))

    story.append(Paragraph("Mixed-Design ANOVA Report", sH1))
    story.append(Paragraph(
        f"Design: {design}  |  Between-Subjects: {between_col}  |  "
        f"Dependent Variable: {dv_label}  |  Alpha: {alpha}", sBD))
    story.append(HRFlowable(width="100%", thickness=2, color=NAVY, spaceAfter=10))
    story.append(Spacer(1, 6))

    story.append(Paragraph("1. Descriptive Statistics", sH2))
    story.append(mktable(desc)); story.append(Spacer(1, 8))

    story.append(Paragraph("2. Normality Tests", sH2))
    story.append(Paragraph(
        "Shapiro-Wilk applied when n <= 50 per cell; Kolmogorov-Smirnov when n > 50. "
        "H0: data are normally distributed.", sBD))
    story.append(mktable(norm_df)); story.append(Spacer(1, 8))

    if not lev_df.empty:
        story.append(Paragraph("3. Levene's Test of Equality of Variances", sH2))
        story.append(Paragraph("H0: group variances are equal at each time point.", sBD))
        story.append(mktable(lev_df)); story.append(Spacer(1, 8))

    if sph:
        W, p_W, eps_GG, eps_HF, chi2, df_c = sph
        story.append(Paragraph("4. Mauchly's Test of Sphericity", sH2))
        story.append(Paragraph(
            "H0: sphericity is satisfied. If violated (p < .05), "
            "Greenhouse-Geisser (GG) or Huynh-Feldt (HF) correction is applied.", sBD))
        sdf = pd.DataFrame([{"Mauchly W": round(W,4), "Chi-sq": round(chi2,3),
                              "df": df_c, "p-value": fmt_p(p_W),
                              "GG Epsilon": round(eps_GG,4), "HF Epsilon": round(eps_HF,4)}])
        story.append(mktable(sdf)); story.append(Spacer(1, 8))

    story.append(Paragraph("5. Mixed-Design ANOVA Results", sH2))
    story.append(Paragraph(
        "Type III Sum of Squares (SPSS-compatible). "
        "GG-corrected p-values reported for within-subjects effects. "
        "Partial eta-squared is the effect size measure.", sBD))
    aov_s = aov_df[["Source","SS","df","MS","F","p_GG","eta_p2","eps"]].copy()
    aov_s.columns = ["Source","SS","df","MS","F","p (GG-corrected)","Partial eta-sq","GG Epsilon"]
    story.append(mktable(aov_s.round(4))); story.append(Spacer(1, 8))

    if not ph_df.empty:
        story.append(Paragraph(f"6. Post-Hoc Pairwise Comparisons ({method_name})", sH2))
        story.append(Paragraph("Conducted because the Group x Time interaction was significant.", sBD))
        ph_s = ph_df[["ctype","timepoint","group","label_a","label_b",
                        "mean_a","mean_b","meandiff","se","t","df",
                        "p_raw","p_adj","cohen_d","sig"]].copy()
        ph_s.columns = ["Type","Time Pt","Group","A","B",
                          "Mean A","Mean B","Diff","SE","t","df","p raw","p adj","d","Sig"]
        story.append(mktable(ph_s)); story.append(Spacer(1, 8))

    story.append(Paragraph("7. Statistical Interpretation", sH2))
    for blk in interp_blocks:
        heading_c = safe(blk["heading"])
        body_c    = (blk["body"]
                     .replace("<strong>","").replace("</strong>","")
                     .replace("<em>","").replace("</em>","")
                     .replace("<br>","  ").replace("&alpha;","alpha")
                     .replace("&eta;","eta").replace("<sup>2</sup>","2")
                     .replace("&epsilon;","epsilon").replace("&times;","x")
                     .replace("×","x"))
        body_c = safe(body_c)
        story.append(Paragraph(heading_c, sH3))
        story.append(Paragraph(body_c, sBD))
        story.append(Spacer(1, 4))

    story.append(Spacer(1, 8))
    story.append(Paragraph("APA-Style Reporting Sentence", sH3))
    story.append(Paragraph(safe(apa_text), sBD))

    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=.5, color=GREY))
    story.append(Paragraph(
        "Generated by Mixed-Design ANOVA Analyzer  |  scipy + numpy engine (SPSS-compatible)", sNT))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="padding: .5rem 0 1rem;">
          <div style="font-size: 1.05rem; font-weight: 700; letter-spacing: -.2px;">
            Mixed-Design ANOVA
          </div>
          <div style="font-size: .78rem; opacity: .65; margin-top: 3px;">
            Split-Plot · SPSS-Equivalent
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        with st.expander("How to Use This App", expanded=False):
            st.markdown("""
**Step 1 — Prepare your data**

Use wide format: one row per participant.

| ID | Group | Pre | Post |
|----|-------|-----|------|
| P1 | A | 45 | 60 |
| P2 | B | 40 | 58 |

**Step 2 — Upload your CSV**

The app detects any design automatically:
Two Groups x Two Time Points, Three Groups x Three Time Points, Two Groups x Three Time Points, etc.

**Step 3 — Map your columns**

Select which column is the Subject ID, which is the Group, and which columns represent time points.

**Step 4 — Run the analysis**

All assumption checks, ANOVA results, post-hoc tests, and interpretations are generated automatically.
            """)

        with st.expander("Statistical Methods", expanded=False):
            st.markdown("""
**ANOVA Engine**
Split-Plot ANOVA with Type III Sum of Squares (identical to SPSS).

**Normality**
Shapiro-Wilk for n ≤ 50 per cell (optimal power).
Kolmogorov-Smirnov for n > 50 (SW loses sensitivity).

**Homogeneity**
Levene's test (center = mean) at each time point.

**Sphericity** (>2 time points only)
Mauchly's W with Greenhouse-Geisser and Huynh-Feldt epsilon corrections.

**Post-Hoc Tests**
Tukey HSD — recommended for equal group sizes.
Bonferroni — most conservative.
Holm — more powerful than Bonferroni.
FDR (Benjamini-Hochberg) — controls false discovery rate.

**Effect Sizes**
Partial η² benchmarks (Cohen, 1988):
Negligible < .01 · Small ≥ .01 · Medium ≥ .06 · Large ≥ .14
            """)

        st.markdown("---")
        st.markdown(
            '<p style="font-size:.77rem;font-weight:600;margin-bottom:.5rem;">Sample Datasets</p>',
            unsafe_allow_html=True)
        st.markdown(
            '<p style="font-size:.73rem;opacity:.7;margin-bottom:.6rem;">'
            'Download any sample to explore the app. '
            'All designs (2 groups × 2 time points, 3 groups × 3 time points, etc.) are supported.'
            '</p>',
            unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button("2 Groups\n2 Time Pts",
                data=make_sample(2, 2).to_csv(index=False),
                file_name="sample_2g_2t.csv", mime="text/csv",
                use_container_width=True, key="s22")
        with c2:
            st.download_button("2 Groups\n3 Time Pts",
                data=make_sample(2, 3).to_csv(index=False),
                file_name="sample_2g_3t.csv", mime="text/csv",
                use_container_width=True, key="s23")
        c3, c4 = st.columns(2)
        with c3:
            st.download_button("3 Groups\n2 Time Pts",
                data=make_sample(3, 2).to_csv(index=False),
                file_name="sample_3g_2t.csv", mime="text/csv",
                use_container_width=True, key="s32")
        with c4:
            st.download_button("3 Groups\n3 Time Pts",
                data=make_sample(3, 3).to_csv(index=False),
                file_name="sample_3g_3t.csv", mime="text/csv",
                use_container_width=True, key="s33")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    render_sidebar()

    # ── Hero ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="hero">
      <div class="hero-title">Mixed-Design ANOVA Analyzer</div>
      <div class="hero-sub">
        Split-Plot ANOVA &ensp;·&ensp; SPSS-Equivalent &ensp;·&ensp;
        Any Design (2+ Groups × 2+ Time Points)
      </div>
      <span class="badge bd-b">Type III SS</span>
      <span class="badge bd-g">Mauchly + GG/HF</span>
      <span class="badge bd-p">Tukey HSD</span>
      <span class="badge bd-o">PDF Export</span>
    </div>""", unsafe_allow_html=True)

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="step-label"><span class="step-num">1</span> Upload Data</div>',
                unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload a wide-format CSV file",
        type=["csv"],
        label_visibility="collapsed",
        help="One row per participant. Columns: Subject ID, Group, and numeric time-point measurements.")

    if uploaded is None:
        st.markdown("""
        <div class="ab ab-info">
          <strong>No file uploaded yet.</strong><br><br>
          This app performs a Mixed-Design (Split-Plot) ANOVA for <em>any</em> combination of groups
          and time points — Two Groups × Two Time Points, Two Groups × Three Time Points,
          Three Groups × Two Time Points, Three Groups × Three Time Points, and so on.
          The design is detected automatically from your data.<br><br>
          <strong>Required CSV format:</strong><br>
          One row per participant. Columns: a Subject ID column, a Group column (categorical),
          and two or more numeric time-point columns.<br><br>
          Download a sample dataset from the sidebar to get started.
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    try:
        df = pd.read_csv(uploaded)
    except Exception as e:
        st.markdown(
            f'<div class="ab ab-err"><strong>File Read Error.</strong> '
            f'The uploaded file could not be parsed as CSV.<br>'
            f'Technical detail: {e}<br><br>'
            f'Please ensure the file is saved as a comma-separated values (.csv) file '
            f'and is not password-protected or corrupted.</div>',
            unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    if df.empty or len(df.columns) < 3:
        st.markdown(
            '<div class="ab ab-err"><strong>Invalid File.</strong> '
            'The uploaded file appears to be empty or has fewer than 3 columns. '
            'A valid wide-format CSV requires at minimum: a Subject ID column, '
            'a Group column, and at least two numeric time-point columns.</div>',
            unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    with st.expander(f"Data preview  —  {df.shape[0]} rows × {df.shape[1]} columns", expanded=True):
        st.dataframe(df.head(10), use_container_width=True)
        if df.shape[0] > 10:
            st.caption(f"Showing first 10 of {df.shape[0]} rows.")

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Column Mapping ────────────────────────────────────────────────────────
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="step-label"><span class="step-num">2</span> Map Columns & Set Options</div>',
                unsafe_allow_html=True)

    all_cols = list(df.columns)
    num_cols = list(df.select_dtypes(include=[np.number]).columns)

    col1, col2, col3 = st.columns(3)
    with col1:
        subj_col = st.selectbox("Subject ID Column", all_cols, index=0,
                                 help="Column containing unique participant identifiers.")
    with col2:
        btwn_col = st.selectbox("Between-Subjects Factor (Group)",
                                 [c for c in all_cols if c != subj_col], index=0,
                                 help="Column containing group labels (e.g., Control, Treatment).")
    with col3:
        avail    = [c for c in num_cols if c not in [subj_col, btwn_col]]
        time_cols = st.multiselect(
            "Time-Point Columns (Within-Subjects)",
            avail,
            default=avail[:min(len(avail), 4)],
            help="Select all columns representing repeated measurements. "
                 "Any number of time points is supported.")

    col4, col5, col6, col7 = st.columns(4)
    with col4:
        dv_label = st.text_input("Dependent Variable Name", "Score",
                                  help="Label for the outcome measure (used in plots and reports).")
    with col5:
        ph_method = st.selectbox(
            "Post-Hoc Method",
            ["tukey", "bonferroni", "holm", "fdr_bh"],
            format_func={"tukey": "Tukey HSD (recommended)",
                          "bonferroni": "Bonferroni",
                          "holm": "Holm (step-down)",
                          "fdr_bh": "FDR – Benjamini-Hochberg"}.get,
            index=0,
            help="Multiple comparison correction applied to post-hoc pairwise tests.")
    with col6:
        alpha = st.selectbox(
            "Significance Level (α)",
            options=[0.001, 0.01, 0.05, 0.10],
            index=2,
            format_func=lambda x: f"α = {x}",
            help="The probability threshold for declaring a result statistically significant.")
    with col7:
        use_gg = st.checkbox("Apply GG Correction", value=True,
                              help="Apply Greenhouse-Geisser correction to within-subjects p-values "
                                   "when sphericity is violated.")

    run_btn = st.button("Run Analysis", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if not run_btn:
        return

    # ── Minimal pre-check ─────────────────────────────────────────────────────
    if len(time_cols) < 2:
        st.markdown(
            '<div class="ab ab-err"><strong>Too Few Time Points Selected.</strong> '
            'Please select at least two time-point columns. '
            'A within-subjects factor requires a minimum of two levels (time points).</div>',
            unsafe_allow_html=True)
        return

    # ── Validate ──────────────────────────────────────────────────────────────
    issues  = validate_data(df, subj_col, btwn_col, time_cols)
    has_err = any(lv == "error" for lv, _ in issues)

    for lv, msg in issues:
        box_cls = {"error": "ab-err", "warning": "ab-warn", "info": "ab-info"}.get(lv, "ab-info")
        label   = {"error": "Data Error", "warning": "Warning", "info": "Note"}.get(lv, "Note")
        st.markdown(
            f'<div class="ab {box_cls}"><strong>{label}:</strong> {msg}</div>',
            unsafe_allow_html=True)

    if has_err:
        st.markdown(
            '<div class="ab ab-err"><strong>Analysis halted.</strong> '
            'Please correct the errors described above before running the analysis.</div>',
            unsafe_allow_html=True)
        return

    req_cols = [subj_col, btwn_col] + time_cols
    df_clean = df[req_cols].dropna()
    if len(df_clean) < 6:
        st.markdown(
            f'<div class="ab ab-err"><strong>Insufficient Data.</strong> '
            f'Only {len(df_clean)} complete cases remain after removing missing values. '
            f'At least 6 complete cases are required.</div>',
            unsafe_allow_html=True)
        return

    n_total     = df_clean[subj_col].nunique()
    n_groups    = df_clean[btwn_col].nunique()
    group_names = sorted(df_clean[btwn_col].unique().tolist())
    design      = design_str(n_groups, len(time_cols))
    method_nm   = {"tukey": "Tukey HSD", "bonferroni": "Bonferroni",
                    "holm": "Holm", "fdr_bh": "FDR (Benjamini-Hochberg)"}[ph_method]

    # ── Run ───────────────────────────────────────────────────────────────────
    with st.spinner("Running analysis..."):
        try:
            desc    = compute_descriptives(df_clean, btwn_col, time_cols)
            norm_df = test_normality(df_clean, btwn_col, time_cols)
            lev_df  = test_levene(df_clean, btwn_col, time_cols)
            aov_df, sph = mixed_anova(df_clean, subj_col, btwn_col, time_cols)

            ab_src  = f"{btwn_col} × Time"
            int_row = aov_df[aov_df["Source"] == ab_src]
            int_p   = float(int_row["p_GG"].values[0]) if not int_row.empty else 1.0
            run_ph  = int_p < alpha

            ph_df = (run_posthoc(df_clean, btwn_col, time_cols, ph_method, alpha)
                     if run_ph else pd.DataFrame())

            interp_blocks, _, _ = build_interpretation(
                aov_df, btwn_col, n_total, alpha, group_names, time_cols)
            apa_text = build_apa(aov_df, btwn_col, n_total, n_groups, time_cols, alpha)
        except Exception as e:
            st.markdown(
                f'<div class="ab ab-err"><strong>Analysis Error.</strong> '
                f'The computation failed with the following message:<br>'
                f'<code>{e}</code><br><br>'
                f'Possible causes: a column with constant values, a singular covariance matrix, '
                f'or extremely unbalanced group sizes.</div>',
                unsafe_allow_html=True)
            st.exception(e)
            return

    # ── Summary bar ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div class="ab ab-ok">Analysis complete — Design: <strong>{design}</strong> &ensp;|&ensp; '
        f'N = {n_total} participants &ensp;|&ensp; '
        f'{n_groups} groups &ensp;|&ensp; {len(time_cols)} time points</div>',
        unsafe_allow_html=True)

    st.markdown(f"""
    <div class="mrow">
      <div class="mc"><div class="mc-val">{n_total}</div><div class="mc-lbl">Participants</div></div>
      <div class="mc"><div class="mc-val">{n_groups}</div><div class="mc-lbl">Groups</div></div>
      <div class="mc"><div class="mc-val">{len(time_cols)}</div><div class="mc-lbl">Time Points</div></div>
      <div class="mc"><div class="mc-val">{n_total*len(time_cols)}</div><div class="mc-lbl">Observations</div></div>
    </div>""", unsafe_allow_html=True)

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tabs = st.tabs([
        "Descriptive Statistics",
        "Assumption Tests",
        "ANOVA Results",
        "Post-Hoc Tests",
        "Plots",
        "Interpretation",
        "Export"
    ])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 0 — Descriptive Statistics
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[0]:
        st.markdown('<div class="card-title">Cell Means and Descriptive Statistics</div>',
                    unsafe_allow_html=True)
        st.markdown(
            f'<div class="ab ab-info">Design: <strong>{design}</strong> &ensp;|&ensp; '
            f'Groups: {", ".join(group_names)} &ensp;|&ensp; '
            f'Time points: {", ".join(time_cols)}</div>',
            unsafe_allow_html=True)
        st.dataframe(desc, use_container_width=True, hide_index=True)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.2rem 0;'>",
                    unsafe_allow_html=True)
        st.markdown("**Marginal Means**", unsafe_allow_html=False)
        c1, c2 = st.columns(2)
        with c1:
            st.caption("By Group (averaged across all time points)")
            mg = desc.groupby("Group")[["N","Mean","SD"]].mean().round(3)
            mg["N"] = mg["N"].astype(int)
            st.dataframe(mg, use_container_width=True)
        with c2:
            st.caption("By Time Point (averaged across all groups)")
            mt = desc.groupby("Time")[["N","Mean","SD"]].mean().round(3)
            mt["N"] = mt["N"].astype(int)
            st.dataframe(mt, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — Assumptions
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[1]:

        # — Normality —
        st.markdown('<div class="card-title">A. Normality Tests</div>', unsafe_allow_html=True)
        sw_n = (norm_df["N"] <= 50).sum(); ks_n = (norm_df["N"] > 50).sum()
        st.markdown(
            f'<div class="ab ab-info">'
            f'<strong>Automatic test selection:</strong> '
            f'Shapiro-Wilk applied to {sw_n} cell(s) with n &le; 50 '
            f'(highest statistical power for small samples); '
            f'Kolmogorov-Smirnov applied to {ks_n} cell(s) with n &gt; 50 '
            f'(Shapiro-Wilk loses sensitivity at large n). '
            f'H<sub>0</sub>: the data are normally distributed. '
            f'H<sub>0</sub> is rejected if p &lt; {alpha}.</div>',
            unsafe_allow_html=True)

        norm_show = norm_df.copy()
        norm_show["p-value"] = norm_show["p-value"].apply(
            lambda x: fmt_p(x) if pd.notna(x) else "—")
        st.dataframe(norm_show, use_container_width=True, hide_index=True)

        n_fail = (norm_df["Normal?"] == "No").sum()
        if n_fail == 0 and (norm_df["Normal?"] == "Yes").any():
            st.markdown(
                '<div class="ab ab-ok">All tested cells satisfy the normality assumption (p &gt; ɑ). '
                'The ANOVA results are expected to be valid.</div>',
                unsafe_allow_html=True)
        elif n_fail > 0:
            st.markdown(
                f'<div class="ab ab-warn">{n_fail} cell(s) show evidence of non-normality. '
                f'ANOVA is generally robust to moderate violations when group sizes are equal '
                f'and n &ge; 15 per cell. For severe violations, consider a non-parametric '
                f'alternative (e.g., Friedman test for within-subjects, Kruskal-Wallis for '
                f'between-subjects).</div>',
                unsafe_allow_html=True)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.2rem 0;'>",
                    unsafe_allow_html=True)

        # — Levene —
        st.markdown('<div class="card-title">B. Levene\'s Test of Homogeneity of Variance</div>',
                    unsafe_allow_html=True)
        st.markdown(
            f'<div class="ab ab-info">'
            f'Tests whether the variance of the dependent variable is equal across groups at each '
            f'time point (between-subjects homogeneity). '
            f'H<sub>0</sub>: group variances are equal. '
            f'H<sub>0</sub> is rejected if p &lt; {alpha}. '
            f'Violation inflates the between-subjects F-ratio.</div>',
            unsafe_allow_html=True)

        if not lev_df.empty:
            lev_show = lev_df.copy()
            lev_show["p-value"] = lev_show["p-value"].apply(fmt_p)
            st.dataframe(lev_show, use_container_width=True, hide_index=True)
            if (lev_df["Equal Variances?"] == "Yes").all():
                st.markdown(
                    '<div class="ab ab-ok">Homogeneity of variance is satisfied at all time points.</div>',
                    unsafe_allow_html=True)
            else:
                n_fail_lev = (lev_df["Equal Variances?"] == "No").sum()
                st.markdown(
                    f'<div class="ab ab-warn">Unequal variances detected at {n_fail_lev} time point(s). '
                    f'Interpret the between-subjects F-ratio with caution.</div>',
                    unsafe_allow_html=True)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.2rem 0;'>",
                    unsafe_allow_html=True)

        # — Sphericity —
        if len(time_cols) > 2:
            st.markdown('<div class="card-title">C. Mauchly\'s Test of Sphericity</div>',
                        unsafe_allow_html=True)
            st.markdown(
                '<div class="ab ab-info">'
                'Sphericity assumes that the variances of the differences between all pairs of '
                'time points are equal. This assumption applies only when there are more than '
                'two within-subjects levels (time points). '
                'H<sub>0</sub>: sphericity is satisfied. Reject H<sub>0</sub> if p &lt; .05. '
                'If violated: use Greenhouse-Geisser correction when &epsilon; &lt; .75, '
                'or Huynh-Feldt when &epsilon; &ge; .75.</div>',
                unsafe_allow_html=True)

            if sph:
                W, p_W, eps_GG, eps_HF, chi2, df_c = sph
                mc1, mc2, mc3, mc4, mc5 = st.columns(5)
                mc1.metric("Mauchly's W", fmt(W, 4))
                mc2.metric("χ²", fmt(chi2, 3))
                mc3.metric("df", str(df_c))
                mc4.metric("p-value", fmt_p(p_W))
                mc5.metric("GG ε", fmt(eps_GG, 4))
                st.caption(f"Huynh-Feldt ε = {fmt(eps_HF, 4)}")

                if not np.isnan(p_W):
                    if p_W < .05:
                        rec = "Greenhouse-Geisser" if eps_GG < .75 else "Huynh-Feldt"
                        st.markdown(
                            f'<div class="ab ab-warn">'
                            f'Sphericity is <strong>violated</strong> '
                            f'(Mauchly W = {fmt(W,4)}, p = {fmt_p(p_W)}). '
                            f'<strong>{rec} correction is recommended</strong> '
                            f'(GG &epsilon; = {fmt(eps_GG,3)}). '
                            f'GG-corrected p-values are automatically applied to within-subjects '
                            f'and interaction effects when the "Apply GG Correction" option is enabled.'
                            f'</div>',
                            unsafe_allow_html=True)
                    else:
                        st.markdown(
                            f'<div class="ab ab-ok">'
                            f'Sphericity is <strong>satisfied</strong> '
                            f'(Mauchly W = {fmt(W,4)}, p = {fmt_p(p_W)}). '
                            f'No correction is required.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown(
                    '<div class="ab ab-warn">'
                    'Mauchly\'s test could not be computed. '
                    'Check that all participants have complete data at all time points.</div>',
                    unsafe_allow_html=True)
        else:
            st.markdown(
                '<div class="ab ab-info">'
                'Mauchly\'s test of sphericity is not applicable for designs with exactly '
                'two time points, because sphericity is automatically satisfied when there is '
                'only one difference score (a single pair of levels).</div>',
                unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — ANOVA Results
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[2]:
        corr_label = "Greenhouse-Geisser corrected" if use_gg else "uncorrected"
        st.markdown(
            f'<div class="ab ab-info">'
            f'<strong>Algorithm:</strong> Split-Plot ANOVA with manual SS decomposition &ensp;|&ensp; '
            f'<strong>Sum of Squares:</strong> Type III (SPSS-compatible) &ensp;|&ensp; '
            f'<strong>p-values reported:</strong> {corr_label}</div>',
            unsafe_allow_html=True)

        st.markdown(render_anova_table(
            aov_df,
            caption=f"Tests of Within-Subjects and Between-Subjects Effects — {design}",
            use_gg=use_gg),
            unsafe_allow_html=True)

        st.markdown("""
        <div class="es-legend">
          <span class="es-chip es-neg">Negligible  &lt; .01</span>
          <span class="es-chip es-sm">Small  &ge; .01</span>
          <span class="es-chip es-med">Medium  &ge; .06</span>
          <span class="es-chip es-lg">Large  &ge; .14</span>
        </div>""", unsafe_allow_html=True)

        with st.expander("View raw ANOVA data"):
            st.dataframe(aov_df, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 3 — Post-Hoc Tests
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[3]:
        if not run_ph:
            st.markdown(
                f'<div class="ab ab-info">'
                f'<strong>Post-hoc tests were not conducted.</strong><br><br>'
                f'The Group × Time interaction did not reach the significance threshold '
                f'(p = {fmt_p(int_p)}, α = {alpha}). '
                f'Pairwise post-hoc comparisons are conventionally conducted only when the '
                f'omnibus interaction effect is statistically significant, in order to '
                f'control the family-wise error rate.<br><br>'
                f'If a main effect is significant and of interest, request simple-effects '
                f'analyses or one-way comparisons at each level of the other factor.</div>',
                unsafe_allow_html=True)

        elif ph_df.empty:
            st.markdown(
                '<div class="ab ab-warn">Post-hoc comparisons could not be computed for this dataset.</div>',
                unsafe_allow_html=True)
        else:
            st.markdown(
                f'<div class="ab ab-ok">'
                f'Post-hoc pairwise comparisons were conducted following a statistically significant '
                f'Group × Time interaction (p = {fmt_p(int_p)}). '
                f'<strong>{method_nm}</strong> correction was applied to all comparisons '
                f'within each set to control the family-wise error rate.</div>',
                unsafe_allow_html=True)

            bt = ph_df[ph_df["ctype"] == "Between Groups"].copy()
            wt = ph_df[ph_df["ctype"] == "Within Time"].copy()

            if not bt.empty:
                st.markdown('<div class="ph-header">Between-Group Comparisons at Each Time Point</div>',
                            unsafe_allow_html=True)
                COLS = ["timepoint","label_a","label_b","mean_a","mean_b",
                        "meandiff","se","t","df","p_raw","p_adj","cohen_d","sig"]
                LABELS = ["Time Point","Group A","Group B","Mean A","Mean B",
                          "Mean Diff","SE","t","df","p (raw)","p (adj)","Cohen d","Sig"]
                bt_s = bt[[c for c in COLS if c in bt.columns]].copy()
                bt_s.columns = LABELS[:len(bt_s.columns)]
                bt_s["p (raw)"] = bt_s["p (raw)"].apply(fmt_p)
                bt_s["p (adj)"] = bt_s["p (adj)"].apply(fmt_p)
                st.dataframe(bt_s, use_container_width=True, hide_index=True)

                # Narrative
                for _, r in bt.iterrows():
                    md  = r["meandiff"]
                    d   = r["cohen_d"]
                    dir_w = "higher than" if md > 0 else ("lower than" if md < 0 else "equal to")
                    sig_c = "ph-sig-yes" if r["significant"] else "ph-sig-no"
                    sig_t = f"statistically significant ({r['sig']})" if r["significant"] else "not significant (ns)"
                    d_t   = (f"; Cohen's d = {fmt(abs(d),2)} [{eta_label(abs(d))} effect]"
                             if not (isinstance(d, float) and np.isnan(d)) else "")
                    st.markdown(f"""
                    <div class="ph-card">
                      <div class="ph-label">{r["timepoint"]} — {r["label_a"]} vs. {r["label_b"]}</div>
                      At <em>{r["timepoint"]}</em>, the <em>{r["label_a"]}</em> group
                      (M = {r["mean_a"]}) scored {abs(round(md, 2))} units {dir_w}
                      the <em>{r["label_b"]}</em> group (M = {r["mean_b"]}).
                      The difference was
                      <span class="{sig_c}">{sig_t}</span>
                      (p = {fmt_p(r["p_adj"])}{d_t}).
                    </div>""", unsafe_allow_html=True)

            if not wt.empty:
                st.markdown('<div class="ph-header">Within-Subject Time Comparisons per Group</div>',
                            unsafe_allow_html=True)
                COLS2   = ["group","label_a","label_b","mean_a","mean_b",
                           "meandiff","se","t","df","p_raw","p_adj","cohen_d","sig"]
                LABELS2 = ["Group","Time A","Time B","Mean A","Mean B",
                           "Mean Diff","SE","t","df","p (raw)","p (adj)","Cohen d","Sig"]
                wt_s = wt[[c for c in COLS2 if c in wt.columns]].copy()
                wt_s.columns = LABELS2[:len(wt_s.columns)]
                wt_s["p (raw)"] = wt_s["p (raw)"].apply(fmt_p)
                wt_s["p (adj)"] = wt_s["p (adj)"].apply(fmt_p)
                st.dataframe(wt_s, use_container_width=True, hide_index=True)

                for _, r in wt.iterrows():
                    md  = r["meandiff"]
                    d   = r["cohen_d"]
                    dir_w = "decreased" if md > 0 else ("increased" if md < 0 else "did not change")
                    sig_c = "ph-sig-yes" if r["significant"] else "ph-sig-no"
                    sig_t = f"statistically significant ({r['sig']})" if r["significant"] else "not significant (ns)"
                    d_t   = (f"; Cohen's d = {fmt(abs(d),2)} [{eta_label(abs(d))} effect]"
                             if not (isinstance(d, float) and np.isnan(d)) else "")
                    st.markdown(f"""
                    <div class="ph-card">
                      <div class="ph-label">Group: {r["group"]} — {r["label_a"]} → {r["label_b"]}</div>
                      Within the <em>{r["group"]}</em> group, scores {dir_w} from
                      <em>{r["label_a"]}</em> (M = {r["mean_a"]}) to
                      <em>{r["label_b"]}</em> (M = {r["mean_b"]}).
                      The change was
                      <span class="{sig_c}">{sig_t}</span>
                      (p = {fmt_p(r["p_adj"])}{d_t}).
                    </div>""", unsafe_allow_html=True)

            st.markdown(
                '<p style="font-size:.74rem;color:#6b7280;margin-top:.25rem;">'
                '*** p &lt; .001 &ensp; ** p &lt; .01 &ensp; * p &lt; .05 &ensp; '
                '† p &lt; .10 &ensp; ns p &ge; .10</p>',
                unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 4 — Plots
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[4]:
        st.markdown(
            '<div class="ab ab-info">'
            '<strong>Profile Plot:</strong> Non-parallel lines indicate a potential '
            'Group × Time interaction. Left panel error bars = ±1 SE; '
            'right panel error bars = 95% CI.</div>',
            unsafe_allow_html=True)
        fig1 = profile_plot(desc, btwn_col, dv_label)
        st.pyplot(fig1, use_container_width=True)
        plt.close(fig1)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.5rem 0;'>",
                    unsafe_allow_html=True)
        st.markdown(
            '<div class="ab ab-info">'
            '<strong>Distributions with KDE:</strong> Overlapping histograms and kernel density '
            'curves show the spread and shape of scores for each group at each time point. '
            'Large asymmetry may indicate non-normality.</div>',
            unsafe_allow_html=True)
        fig2 = dist_plots(df_clean, btwn_col, time_cols, dv_label)
        st.pyplot(fig2, use_container_width=True)
        plt.close(fig2)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.5rem 0;'>",
                    unsafe_allow_html=True)
        st.markdown(
            '<div class="ab ab-info">'
            '<strong>Q–Q Plots:</strong> Points near the diagonal dashed reference line '
            'indicate normally distributed data. Systematic departures (S-curves, heavy tails) '
            'suggest non-normality.</div>',
            unsafe_allow_html=True)
        fig3 = qq_plots(df_clean, btwn_col, time_cols)
        st.pyplot(fig3, use_container_width=True)
        plt.close(fig3)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 5 — Interpretation
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[5]:
        for blk in interp_blocks:
            itype = blk.get("type", "info")
            border_color = {
                "sig": "#16a34a", "ns": "#6b7280", "marginal": "#d97706",
                "between": "#2563eb", "within": "#9333ea", "info": "#2563eb"
            }.get(itype, "#2563eb")
            heading_color = {
                "sig": "#14532d", "ns": "#374151", "marginal": "#78350f",
                "between": "#1e3a5f", "within": "#4a1d96", "info": "#1e3a5f"
            }.get(itype, "#1e3a5f")
            st.markdown(f"""
            <div class="interp-card" style="border-left: 4px solid {border_color};">
              <div class="interp-heading" style="color:{heading_color};">{blk["heading"]}</div>
              <div class="interp-body">{blk["body"]}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.25rem 0;'>",
                    unsafe_allow_html=True)

        # Effect size table
        st.markdown("**Effect Size Summary**")
        eff_rows = []
        for _, r in aov_df.iterrows():
            if "err" in r.get("rowtype",""):
                continue
            src = r["Source"]; F = r.get("F", np.nan)
            p   = r.get("p_GG", r.get("p", np.nan)); e = r.get("eta_p2", np.nan)
            eff_rows.append({
                "Source": src,
                "F": fmt(F, 3),
                "p (GG-corrected)": fmt_p(p),
                "Partial η²": fmt(e, 4),
                "Interpretation": eta_label(e)
            })
        st.dataframe(pd.DataFrame(eff_rows), use_container_width=True, hide_index=True)

        st.markdown("""
        <div class="es-legend" style="margin-top:.5rem;">
          <span class="es-chip es-neg">Negligible &lt; .01</span>
          <span class="es-chip es-sm">Small &ge; .01</span>
          <span class="es-chip es-med">Medium &ge; .06</span>
          <span class="es-chip es-lg">Large &ge; .14</span>
        </div>""", unsafe_allow_html=True)

        st.markdown("<hr style='border:none;border-top:1px solid #edf0f7;margin:1.25rem 0;'>",
                    unsafe_allow_html=True)
        st.markdown("**APA-Style Reporting Sentence**")
        st.markdown(f'<div class="apa-box">{apa_text}</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 6 — Export
    # ══════════════════════════════════════════════════════════════════════════
    with tabs[6]:
        st.markdown(
            '<div class="ab ab-info">'
            'Download the complete analysis results in your preferred format. '
            'The PDF report contains formatted tables and full academic interpretation. '
            'The CSV file contains all raw numerical outputs for further processing.</div>',
            unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)

        with c1:
            st.markdown("**CSV Report**")
            buf_csv = io.StringIO()
            buf_csv.write(f"Mixed-Design ANOVA Report\nDesign: {design}\n"
                          f"Between-Subjects Factor: {btwn_col}\n"
                          f"Dependent Variable: {dv_label}\nAlpha: {alpha}\n\n")
            buf_csv.write("DESCRIPTIVE STATISTICS\n"); desc.to_csv(buf_csv, index=False)
            buf_csv.write("\nNORMALITY TESTS\n");     norm_df.to_csv(buf_csv, index=False)
            buf_csv.write("\nLEVENE TEST\n");          lev_df.to_csv(buf_csv, index=False)
            if sph:
                W2,p2,eg2,ef2,ch2,dc2 = sph
                pd.DataFrame([{"W":round(W2,4),"chi2":round(ch2,3),"df":dc2,
                                "p":fmt_p(p2),"GG_eps":round(eg2,4),"HF_eps":round(ef2,4)}]
                ).to_csv(buf_csv, index=False)
            buf_csv.write("\nANOVA RESULTS\n");       aov_df.to_csv(buf_csv, index=False)
            if not ph_df.empty:
                buf_csv.write("\nPOST-HOC TESTS\n");  ph_df.to_csv(buf_csv, index=False)
            buf_csv.write("\nINTERPRETATION\n")
            for blk in interp_blocks:
                buf_csv.write(f"\n[{blk['heading']}]\n")
                body_plain = (blk["body"]
                    .replace("<strong>","").replace("</strong>","")
                    .replace("<em>","").replace("</em>","")
                    .replace("&alpha;","alpha").replace("&eta;","eta")
                    .replace("<sup>2</sup>","2").replace("&epsilon;","epsilon")
                    .replace("×","x").replace("<br>","  ").replace("&ensp;"," "))
                buf_csv.write(body_plain + "\n")
            buf_csv.write(f"\nAPA SENTENCE\n{apa_text}\n")
            st.download_button("Download CSV Report", buf_csv.getvalue().encode(),
                                "mixed_anova_report.csv", "text/csv",
                                use_container_width=True)

        with c2:
            st.markdown("**PDF Report**")
            pdf_bytes = build_pdf(
                desc, norm_df, lev_df, aov_df, sph,
                interp_blocks, ph_df,
                btwn_col, dv_label, alpha, n_groups, time_cols,
                method_nm, apa_text)
            if pdf_bytes:
                st.download_button("Download PDF Report", pdf_bytes,
                                    "mixed_anova_report.pdf", "application/pdf",
                                    use_container_width=True)
            else:
                st.caption("Install reportlab for PDF export: pip install reportlab")

        with c3:
            st.markdown("**Long-Format Data**")
            df_long = df_clean.melt(
                id_vars=[subj_col, btwn_col], value_vars=time_cols,
                var_name="Time", value_name=dv_label)
            st.download_button("Download Long CSV",
                                df_long.to_csv(index=False).encode(),
                                "long_format_data.csv", "text/csv",
                                use_container_width=True)


if __name__ == "__main__":
    main()
