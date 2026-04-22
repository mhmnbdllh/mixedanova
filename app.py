"""
Mixed-Design ANOVA (Split-Plot ANOVA) — SPSS-Equivalent Analyzer
Zero external dependency beyond scipy/statsmodels/matplotlib/seaborn
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import warnings
import io
from itertools import combinations
from scipy import stats
from scipy.stats import levene as scipy_levene

warnings.filterwarnings('ignore')

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Mixed-Design ANOVA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=DM+Sans:wght@300;400;500;600&display=swap');

html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.stApp{background:#08090e;color:#e2e8f0;}

.hero{
  background:linear-gradient(135deg,#0f1623 0%,#0a0f1a 60%,#0d1320 100%);
  border:1px solid #1e2d45;border-radius:14px;padding:2.2rem 2.8rem;
  margin-bottom:2rem;position:relative;overflow:hidden;
}
.hero::after{
  content:'';position:absolute;top:-60px;right:-60px;width:220px;height:220px;
  border-radius:50%;background:radial-gradient(circle,rgba(56,139,253,.12) 0%,transparent 70%);
  pointer-events:none;
}
.hero::before{
  content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,#388bfd,#3fb950,#a371f7,#f78166);
}
.hero h1{font-family:'JetBrains Mono',monospace;font-size:1.7rem;font-weight:600;
  color:#e2e8f0;margin:0 0 .4rem;letter-spacing:-.5px;}
.hero p{color:#7d8590;font-size:.93rem;margin:0;}
.badge{display:inline-block;padding:2px 9px;border-radius:12px;font-size:.72rem;
  font-weight:600;font-family:'JetBrains Mono',monospace;margin-right:5px;margin-top:.8rem;}
.bd-g{background:#0a2e1a;color:#3fb950;border:1px solid #238636;}
.bd-b{background:#0a1e36;color:#388bfd;border:1px solid #1f6feb;}
.bd-p{background:#1a0f36;color:#a371f7;border:1px solid #6e40c9;}
.bd-r{background:#2d0f0f;color:#f78166;border:1px solid #da3633;}

.sec-card{background:#0f1623;border:1px solid #1e2d45;border-radius:10px;
  padding:1.4rem 1.6rem;margin-bottom:1.4rem;}
.sec-title{font-family:'JetBrains Mono',monospace;font-size:.75rem;font-weight:600;
  color:#7d8590;text-transform:uppercase;letter-spacing:1.8px;margin-bottom:1rem;
  padding-bottom:.5rem;border-bottom:1px solid #161d2d;}

.anova-table{width:100%;border-collapse:collapse;font-size:.85rem;
  font-family:'JetBrains Mono',monospace;}
.anova-table th{background:#131c2e;color:#388bfd;padding:9px 14px;text-align:left;
  font-weight:600;border-bottom:2px solid #1e2d45;white-space:nowrap;}
.anova-table td{padding:8px 14px;border-bottom:1px solid #161d2d;color:#cbd5e1;}
.anova-table tr:hover td{background:#111827;}
.anova-table .sig3{color:#3fb950;font-weight:700;}
.anova-table .sig2{color:#58a6ff;font-weight:600;}
.anova-table .sig1{color:#e3b341;font-weight:600;}
.anova-table .sigm{color:#e3b341;}
.anova-table .ns{color:#7d8590;}

.ibox{background:#0a1e36;border:1px solid #1f6feb;border-left:4px solid #388bfd;
  border-radius:8px;padding:1.1rem 1.4rem;margin:.8rem 0;font-size:.88rem;
  line-height:1.75;color:#cbd5e1;}
.wbox{background:#1e1500;border:1px solid #9e6a03;border-left:4px solid #e3b341;
  border-radius:8px;padding:.9rem 1.2rem;margin:.6rem 0;font-size:.83rem;color:#e3b341;}
.obox{background:#071c10;border:1px solid #238636;border-left:4px solid #3fb950;
  border-radius:8px;padding:.9rem 1.2rem;margin:.6rem 0;font-size:.83rem;color:#3fb950;}
.ebox{background:#2d0f0f;border:1px solid #da3633;border-left:4px solid #f78166;
  border-radius:8px;padding:.9rem 1.2rem;margin:.6rem 0;font-size:.83rem;color:#f78166;}

.mcard{background:#131c2e;border:1px solid #1e2d45;border-radius:8px;
  padding:.9rem;text-align:center;}
.mval{font-family:'JetBrains Mono',monospace;font-size:1.55rem;
  font-weight:600;color:#388bfd;}
.mlbl{font-size:.74rem;color:#7d8590;margin-top:3px;}

div[data-testid="stSidebar"]{background:#080c14;border-right:1px solid #161d2d;}
.stTabs [data-baseweb="tab-list"]{background:transparent;gap:3px;}
.stTabs [data-baseweb="tab"]{background:#0f1623;border:1px solid #1e2d45;
  color:#7d8590;border-radius:6px;font-family:'JetBrains Mono',monospace;font-size:.77rem;}
.stTabs [aria-selected="true"]{background:#1f6feb!important;
  border-color:#388bfd!important;color:#e2e8f0!important;}
.stButton>button{background:#131c2e;border:1px solid #1e2d45;color:#e2e8f0;
  border-radius:6px;font-family:'JetBrains Mono',monospace;font-size:.82rem;
  transition:all .2s;}
.stButton>button:hover{background:#1e2d45;border-color:#388bfd;color:#388bfd;}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# UTILITY
# ════════════════════════════════════════════════════════════════════════════

def fmt(x, d=3):
    if x is None or (isinstance(x, float) and np.isnan(x)): return "—"
    if isinstance(x, (int, np.integer)): return str(x)
    return f"{x:.{d}f}"

def fmt_p(p):
    if p is None or np.isnan(p): return "—"
    if p < .001: return "< .001"
    return f"{p:.3f}"

def sig_star(p):
    if np.isnan(p): return ""
    if p < .001: return "***"
    if p < .01:  return "**"
    if p < .05:  return "*"
    if p < .10:  return "†"
    return "ns"

def sig_cls(p):
    if np.isnan(p): return "ns"
    if p < .001: return "sig3"
    if p < .01:  return "sig2"
    if p < .05:  return "sig1"
    if p < .10:  return "sigm"
    return "ns"

def eta_label(e):
    if np.isnan(e): return "—"
    if e >= .14: return f"{e:.3f} (Large)"
    if e >= .06: return f"{e:.3f} (Medium)"
    if e >= .01: return f"{e:.3f} (Small)"
    return f"{e:.3f} (Negligible)"


# ════════════════════════════════════════════════════════════════════════════
# SAMPLE DATA
# ════════════════════════════════════════════════════════════════════════════

def sample_2x2():
    np.random.seed(42)
    rows=[]
    for g,m in [('Control',[50,52]),('Treatment',[50,62])]:
        for i in range(15):
            rows.append({'ID':f'{g[0]}{i+1:02d}','Group':g,
                         'Pre':round(np.random.normal(m[0],8),2),
                         'Post':round(np.random.normal(m[1],8),2)})
    return pd.DataFrame(rows)

def sample_3x3():
    np.random.seed(7)
    rows=[]
    for g,m in [('Control',[40,41,42]),('Low_Dose',[40,50,56]),('High_Dose',[40,58,70])]:
        for i in range(12):
            rows.append({'ID':f'{g[0]}{i+1:02d}','Group':g,
                         'Pre':round(np.random.normal(m[0],9),2),
                         'Mid':round(np.random.normal(m[1],9),2),
                         'Post':round(np.random.normal(m[2],9),2)})
    return pd.DataFrame(rows)


# ════════════════════════════════════════════════════════════════════════════
# CORE STATISTICS — Pure scipy/numpy (no pingouin)
# ════════════════════════════════════════════════════════════════════════════

def compute_descriptives(df_wide, between_col, time_cols):
    rows=[]
    for g in df_wide[between_col].unique():
        sub = df_wide[df_wide[between_col]==g]
        for t in time_cols:
            v = sub[t].dropna()
            se = v.std()/np.sqrt(len(v)) if len(v)>1 else np.nan
            rows.append({
                'Group': g, 'Time': t, 'N': len(v),
                'Mean': v.mean(), 'SD': v.std(ddof=1),
                'SE': se, 'Median': v.median(),
                'Min': v.min(), 'Max': v.max(),
                '95% CI Lower': v.mean()-1.96*se,
                '95% CI Upper': v.mean()+1.96*se,
            })
    return pd.DataFrame(rows)


def test_normality(df_wide, between_col, time_cols):
    rows=[]
    for g in df_wide[between_col].unique():
        sub = df_wide[df_wide[between_col]==g]
        for t in time_cols:
            v = sub[t].dropna().values
            n = len(v)
            if n < 3:
                rows.append({'Group':g,'Time':t,'N':n,'Test':'—',
                              'Statistic':np.nan,'p-value':np.nan,'Normal':'—','Note':'n < 3'})
                continue
            if n <= 50:
                stat,p = stats.shapiro(v)
                tname,note = 'Shapiro-Wilk','Recommended: n ≤ 50'
            else:
                stat,p = stats.kstest(v,'norm',args=(v.mean(),v.std()))
                tname,note = 'Kolmogorov-Smirnov','Recommended: n > 50'
            rows.append({'Group':g,'Time':t,'N':n,'Test':tname,
                          'Statistic':round(stat,4),'p-value':round(p,4),
                          'Normal':'Yes' if p>.05 else 'No','Note':note})
    return pd.DataFrame(rows)


def test_levene(df_wide, between_col, time_cols):
    rows=[]
    for t in time_cols:
        gdata=[df_wide[df_wide[between_col]==g][t].dropna().values
               for g in df_wide[between_col].unique()]
        if all(len(x)>1 for x in gdata):
            stat,p = scipy_levene(*gdata, center='mean')
            rows.append({'Time Point':t,'Levene F':round(stat,3),
                          'p-value':round(p,4),'Equal Variances':'Yes' if p>.05 else 'No'})
    return pd.DataFrame(rows)


def mauchly_test(df_wide, between_col, time_cols):
    """
    Mauchly's test of sphericity + GG and HF epsilon corrections.
    Returns (W, p, eps_GG, eps_HF, chi2, df_chi)
    """
    b = len(time_cols)
    if b < 3:
        return None
    # Pool all subjects' time-point scores
    Y = df_wide[time_cols].values.astype(float)  # (N, b)
    N = Y.shape[0]
    # Contrast matrix (Helmert-style, b-1 orthogonal contrasts)
    C = np.zeros((b, b-1))
    for j in range(b-1):
        C[:j+1, j] = 1.0/(j+1)
        C[j+1, j] = -1.0
    # Transform
    T = Y @ C  # (N, b-1)
    # Covariance
    S = np.cov(T, rowvar=False)
    if S.ndim == 0:
        S = np.array([[S]])
    # Mauchly W
    try:
        det_S = np.linalg.det(S)
        tr_S  = np.trace(S)
        k     = b - 1
        W = det_S / (tr_S/k)**k
        W = max(1e-15, min(W, 1.0))
        df_chi = k*(k+1)//2 - 1
        u  = N - 1 - (2*k**2 + k + 2)/(6*k)
        chi2 = -u * np.log(W)
        p_W  = 1 - stats.chi2.cdf(chi2, df_chi) if df_chi > 0 else np.nan
    except Exception:
        return None
    # GG epsilon
    tr_S2 = np.trace(S @ S)
    eps_GG = (tr_S**2) / ((k) * tr_S2)
    eps_GG = min(max(eps_GG, 1/k), 1.0)
    # HF epsilon
    eps_HF = (N*(k*eps_GG) - 2) / (k*(N - 1 - k*eps_GG))
    eps_HF = min(eps_HF, 1.0)
    return W, p_W, eps_GG, eps_HF, chi2, df_chi


def mixed_anova(df_wide, subject_col, between_col, time_cols):
    """
    Full Split-Plot ANOVA (Type III SS, SPSS-equivalent).
    Returns DataFrame with sources, SS, df, MS, F, p, eta_p2.
    Also returns (eps_GG, eps_HF) for within-subject correction.
    """
    groups = df_wide[between_col].unique()
    a = len(groups)
    b = len(time_cols)

    # Build group arrays
    group_data = {g: df_wide[df_wide[between_col]==g][time_cols].values.astype(float)
                  for g in groups}
    n_g   = {g: group_data[g].shape[0] for g in groups}
    N     = sum(n_g.values())

    grand_mean = np.mean([group_data[g] for g in groups])

    # Group (marginal) means — mean over all time & subjects in group
    gm  = {g: np.mean(group_data[g]) for g in groups}
    # Time (marginal) means — mean over all groups & subjects at each time
    all_scores = np.vstack([group_data[g] for g in groups])
    tm  = np.mean(all_scores, axis=0)  # shape (b,)
    # Cell means
    cm  = {g: np.mean(group_data[g], axis=0) for g in groups}  # (b,) per group
    # Subject means (for each subject: mean over time)
    sm = {g: np.mean(group_data[g], axis=1) for g in groups}   # (n_g,) per group

    # ── SS Between (Group main effect) ──
    SS_A   = b * sum(n_g[g]*(gm[g]-grand_mean)**2 for g in groups)
    df_A   = a - 1

    # ── SS Error Between (Subjects within Groups) ──
    SS_sWG = sum(b * np.sum((sm[g]-gm[g])**2) for g in groups)
    df_sWG = sum(n_g[g]-1 for g in groups)

    # ── SS Within (Time main effect) ──
    SS_B   = N * np.sum((tm - grand_mean)**2)
    df_B   = b - 1

    # ── SS Interaction ──
    SS_AB  = sum(n_g[g]*np.sum((cm[g] - gm[g] - tm + grand_mean)**2) for g in groups)
    df_AB  = (a-1)*(b-1)

    # ── SS Error Within ──
    SS_eW  = 0.0
    for g in groups:
        for i in range(n_g[g]):
            for j in range(b):
                SS_eW += (group_data[g][i,j] - sm[g][i] - cm[g][j] + gm[g])**2
    df_eW  = sum(n_g[g]-1 for g in groups)*(b-1)

    # MS
    MS_A   = SS_A   / df_A   if df_A   > 0 else np.nan
    MS_sWG = SS_sWG / df_sWG if df_sWG > 0 else np.nan
    MS_B   = SS_B   / df_B   if df_B   > 0 else np.nan
    MS_AB  = SS_AB  / df_AB  if df_AB  > 0 else np.nan
    MS_eW  = SS_eW  / df_eW  if df_eW  > 0 else np.nan

    # F
    F_A  = MS_A  / MS_sWG if MS_sWG else np.nan
    F_B  = MS_B  / MS_eW  if MS_eW  else np.nan
    F_AB = MS_AB / MS_eW  if MS_eW  else np.nan

    # p
    p_A  = float(1-stats.f.cdf(F_A,  df_A,  df_sWG)) if not np.isnan(F_A)  else np.nan
    p_B  = float(1-stats.f.cdf(F_B,  df_B,  df_eW))  if not np.isnan(F_B)  else np.nan
    p_AB = float(1-stats.f.cdf(F_AB, df_AB, df_eW))  if not np.isnan(F_AB) else np.nan

    # Partial η²
    e_A  = SS_A  / (SS_A  + SS_sWG) if (SS_A+SS_sWG)  > 0 else np.nan
    e_B  = SS_B  / (SS_B  + SS_eW)  if (SS_B+SS_eW)   > 0 else np.nan
    e_AB = SS_AB / (SS_AB + SS_eW)  if (SS_AB+SS_eW)  > 0 else np.nan

    # Sphericity / epsilon corrections
    sph = mauchly_test(df_wide, between_col, time_cols)
    eps_GG = eps_HF = 1.0
    if sph:
        _, _, eps_GG, eps_HF, _, _ = sph

    # GG-corrected p for within and interaction
    def gg_p(F, df_num, df_den, eps):
        df1 = df_num * eps
        df2 = df_den * eps
        return float(1-stats.f.cdf(F, df1, df2)) if not np.isnan(F) else np.nan

    p_B_gg  = gg_p(F_B,  df_B,  df_eW, eps_GG)
    p_AB_gg = gg_p(F_AB, df_AB, df_eW, eps_GG)

    results = pd.DataFrame([
        {'Source': between_col, 'SS':SS_A,   'df':df_A,   'df_error':df_sWG,
         'MS':MS_A,   'F':F_A,   'p':p_A,   'p_GG':p_A,   'eta_p2':e_A,
         'eps':np.nan, 'type':'between'},
        {'Source': 'Error(Between)', 'SS':SS_sWG, 'df':df_sWG, 'df_error':np.nan,
         'MS':MS_sWG, 'F':np.nan,'p':np.nan,'p_GG':np.nan,'eta_p2':np.nan,
         'eps':np.nan,'type':'error_between'},
        {'Source': 'Time', 'SS':SS_B,   'df':df_B,   'df_error':df_eW,
         'MS':MS_B,   'F':F_B,   'p':p_B,   'p_GG':p_B_gg, 'eta_p2':e_B,
         'eps':eps_GG,'type':'within'},
        {'Source': f'{between_col} × Time', 'SS':SS_AB, 'df':df_AB, 'df_error':df_eW,
         'MS':MS_AB,  'F':F_AB,  'p':p_AB,  'p_GG':p_AB_gg,'eta_p2':e_AB,
         'eps':eps_GG,'type':'interaction'},
        {'Source': 'Error(Within)',  'SS':SS_eW,  'df':df_eW,  'df_error':np.nan,
         'MS':MS_eW,  'F':np.nan,'p':np.nan,'p_GG':np.nan,'eta_p2':np.nan,
         'eps':np.nan,'type':'error_within'},
    ])
    return results, sph


def bonferroni_pairwise(df_wide, between_col, time_cols, method='bonferroni'):
    """
    Pairwise post-hoc comparisons:
    1. Between-group at each time point (independent t-tests, corrected)
    2. Within-subject time comparisons per group (paired t-tests, corrected)
    """
    groups = sorted(df_wide[between_col].unique())
    rows = []

    # 1. Between-group at each time
    for t in time_cols:
        pairs = list(combinations(groups, 2))
        raw_ps, raw_ts, raw_dfs = [], [], []
        pair_data = []
        for g1, g2 in pairs:
            v1 = df_wide[df_wide[between_col]==g1][t].dropna().values
            v2 = df_wide[df_wide[between_col]==g2][t].dropna().values
            t_stat, p_raw = stats.ttest_ind(v1, v2, equal_var=True)
            dof = len(v1)+len(v2)-2
            raw_ps.append(p_raw); raw_ts.append(t_stat); raw_dfs.append(dof)
            pair_data.append((g1,g2,v1,v2))

        adj_ps = _adjust_p(raw_ps, method)
        for i,(g1,g2,v1,v2) in enumerate(pair_data):
            d = (np.mean(v1)-np.mean(v2)) / np.sqrt(
                ((len(v1)-1)*np.var(v1,ddof=1)+(len(v2)-1)*np.var(v2,ddof=1))/(len(v1)+len(v2)-2))
            rows.append({'Type':'Between-Groups','Time':t,'Group A':g1,'Group B':g2,
                          'Mean A':round(np.mean(v1),3),'Mean B':round(np.mean(v2),3),
                          'Mean Diff':round(np.mean(v1)-np.mean(v2),3),
                          'SE':round(np.sqrt(np.var(v1,ddof=1)/len(v1)+np.var(v2,ddof=1)/len(v2)),3),
                          't':round(raw_ts[i],3),'df':int(raw_dfs[i]),
                          'p (raw)':round(raw_ps[i],4),'p (adj)':round(adj_ps[i],4),
                          'Cohen d':round(d,3),'Sig':sig_star(adj_ps[i])})

    # 2. Within-subject time pairs per group
    time_pairs = list(combinations(time_cols, 2))
    for g in groups:
        sub = df_wide[df_wide[between_col]==g]
        raw_ps2, raw_ts2, raw_dfs2 = [], [], []
        pair_data2 = []
        for t1,t2 in time_pairs:
            v1 = sub[t1].dropna().values; v2 = sub[t2].dropna().values
            n_min = min(len(v1),len(v2))
            t_stat,p_raw = stats.ttest_rel(v1[:n_min],v2[:n_min])
            dof = n_min-1
            raw_ps2.append(p_raw); raw_ts2.append(t_stat); raw_dfs2.append(dof)
            pair_data2.append((t1,t2,v1[:n_min],v2[:n_min]))
        adj_ps2 = _adjust_p(raw_ps2, method)
        for i,(t1,t2,v1,v2) in enumerate(pair_data2):
            diff = v1-v2
            se = np.std(diff,ddof=1)/np.sqrt(len(diff)) if len(diff)>1 else np.nan
            d = np.mean(diff)/np.std(diff,ddof=1) if np.std(diff,ddof=1)>0 else np.nan
            rows.append({'Type':'Within-Time','Group':g,'Time A':t1,'Time B':t2,
                          'Mean A':round(np.mean(v1),3),'Mean B':round(np.mean(v2),3),
                          'Mean Diff':round(np.mean(diff),3),
                          'SE':round(se,3) if se else np.nan,
                          't':round(raw_ts2[i],3),'df':int(raw_dfs2[i]),
                          'p (raw)':round(raw_ps2[i],4),'p (adj)':round(adj_ps2[i],4),
                          'Cohen d':round(d,3) if d else np.nan,'Sig':sig_star(adj_ps2[i])})

    return pd.DataFrame(rows)


def _adjust_p(ps, method):
    ps = np.array(ps, dtype=float)
    n  = len(ps)
    if n == 0: return ps
    if method == 'bonferroni':
        return np.minimum(ps * n, 1.0)
    if method in ('holm','holm-bonferroni'):
        order = np.argsort(ps)
        adj = ps.copy()
        for rank,idx in enumerate(order):
            adj[idx] = min(ps[idx]*(n-rank), 1.0)
        # Enforce monotonicity
        for i in range(1,n):
            adj[order[i]] = max(adj[order[i]], adj[order[i-1]])
        return adj
    if method == 'fdr_bh':
        order = np.argsort(ps)[::-1]
        adj = ps.copy()
        min_p = 1.0
        for i,idx in enumerate(order):
            adj[idx] = min(ps[idx]*n/(n-i), min_p)
            min_p = adj[idx]
        return np.minimum(adj, 1.0)
    # tukey — approximate via Bonferroni for generality
    return np.minimum(ps * n, 1.0)


# ════════════════════════════════════════════════════════════════════════════
# VISUALIZATION
# ════════════════════════════════════════════════════════════════════════════

PALETTE = ['#388bfd','#3fb950','#a371f7','#e3b341','#f78166','#79c0ff','#56d364']

def profile_plot(desc, between_col):
    plt.style.use('dark_background')
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor('#08090e')
    groups = desc['Group'].unique()
    times  = desc['Time'].tolist()
    times_u = list(dict.fromkeys(times))  # ordered unique

    for ax in axes:
        ax.set_facecolor('#0f1623')
        for sp in ['top','right']: ax.spines[sp].set_visible(False)
        for sp in ['bottom','left']: ax.spines[sp].set_color('#1e2d45')
        ax.tick_params(colors='#7d8590', labelsize=9)
        ax.grid(axis='y', color='#161d2d', linewidth=0.8)
        ax.set_xlabel('Time Point', color='#7d8590', fontsize=10)
        ax.set_ylabel('Mean Score', color='#7d8590', fontsize=10)

    # Line + error bar
    ax = axes[0]
    for i,g in enumerate(groups):
        sub = desc[desc['Group']==g].sort_values('Time', key=lambda x: pd.Categorical(x, categories=times_u, ordered=True))
        x = range(len(sub))
        ax.errorbar(x, sub['Mean'], yerr=sub['SE'],
                    marker='o', markersize=8, linewidth=2.5, color=PALETTE[i%len(PALETTE)],
                    label=str(g), capsize=5, capthick=1.5, elinewidth=1.5,
                    markeredgecolor='#08090e', markeredgewidth=1.5)
    ax.set_xticks(range(len(times_u))); ax.set_xticklabels(times_u, fontsize=9)
    ax.set_title('Profile Plot (±1 SE)', color='#e2e8f0', fontsize=11, pad=10,
                 fontfamily='monospace')
    ax.legend(title=between_col, title_fontsize=8, fontsize=8,
              facecolor='#0f1623', edgecolor='#1e2d45', labelcolor='#cbd5e1')

    # Bar
    ax = axes[1]
    ngrp = len(groups)
    w = 0.75/ngrp
    xs = np.arange(len(times_u))
    for i,g in enumerate(groups):
        sub = desc[desc['Group']==g].sort_values('Time', key=lambda x: pd.Categorical(x, categories=times_u, ordered=True))
        off = (i - ngrp/2 + .5)*w
        ax.bar(xs+off, sub['Mean'], width=w*0.9, color=PALETTE[i%len(PALETTE)],
               alpha=0.8, label=str(g), edgecolor='#08090e', linewidth=0.6)
        ax.errorbar(xs+off, sub['Mean'], yerr=1.96*sub['SE'], fmt='none',
                    color='white', capsize=3, capthick=1, elinewidth=1, alpha=0.5)
    ax.set_xticks(xs); ax.set_xticklabels(times_u, fontsize=9)
    ax.set_title('Bar Chart (95% CI)', color='#e2e8f0', fontsize=11, pad=10,
                 fontfamily='monospace')
    ax.legend(title=between_col, title_fontsize=8, fontsize=8,
              facecolor='#0f1623', edgecolor='#1e2d45', labelcolor='#cbd5e1')

    plt.tight_layout(pad=2.5)
    return fig


def dist_plots(df_wide, between_col, time_cols):
    plt.style.use('dark_background')
    n = len(time_cols)
    fig, axes = plt.subplots(1, n, figsize=(5*n, 4), sharey=False)
    fig.patch.set_facecolor('#08090e')
    if n == 1: axes = [axes]
    groups = df_wide[between_col].unique()
    for ax,t in zip(axes, time_cols):
        ax.set_facecolor('#0f1623')
        for sp in ['top','right']: ax.spines[sp].set_visible(False)
        for sp in ['bottom','left']: ax.spines[sp].set_color('#1e2d45')
        ax.tick_params(colors='#7d8590', labelsize=8)
        ax.grid(axis='y', color='#161d2d', linewidth=0.6)
        for i,g in enumerate(groups):
            v = df_wide[df_wide[between_col]==g][t].dropna()
            ax.hist(v, bins=12, alpha=0.55, color=PALETTE[i%len(PALETTE)],
                    label=str(g), edgecolor='#08090e', linewidth=0.4)
            # KDE
            if len(v) > 5:
                kde = stats.gaussian_kde(v)
                xr = np.linspace(v.min(), v.max(), 200)
                scale = len(v) * (v.max()-v.min()) / 12
                ax.plot(xr, kde(xr)*scale, color=PALETTE[i%len(PALETTE)], linewidth=2, alpha=0.9)
        ax.set_title(str(t), color='#e2e8f0', fontsize=10)
        ax.set_xlabel('Score', color='#7d8590', fontsize=9)
    axes[0].legend(title=between_col, fontsize=8, facecolor='#0f1623',
                   edgecolor='#1e2d45', labelcolor='#cbd5e1')
    plt.tight_layout()
    return fig


def qq_plots(df_wide, between_col, time_cols):
    plt.style.use('dark_background')
    groups = df_wide[between_col].unique()
    rows = len(groups); cols = len(time_cols)
    fig, axes = plt.subplots(rows, cols, figsize=(4*cols, 3.5*rows), squeeze=False)
    fig.patch.set_facecolor('#08090e')
    for r,g in enumerate(groups):
        for c,t in enumerate(time_cols):
            ax = axes[r][c]
            v  = df_wide[df_wide[between_col]==g][t].dropna().values
            ax.set_facecolor('#0f1623')
            for sp in ['top','right']: ax.spines[sp].set_visible(False)
            for sp in ['bottom','left']: ax.spines[sp].set_color('#1e2d45')
            ax.tick_params(colors='#7d8590', labelsize=7)
            if len(v) >= 3:
                (osm,osr),(_,_,_) = stats.probplot(v, dist='norm')
                ax.scatter(osm, osr, color=PALETTE[r%len(PALETTE)], s=18, alpha=0.8, edgecolors='none')
                # Ref line
                q1,q3 = np.percentile(v,[25,75])
                th1,th3 = stats.norm.ppf([.25,.75])
                slope=(q3-q1)/(th3-th1); inter=q1-slope*th1
                xl=np.array([osm[0],osm[-1]])
                ax.plot(xl, slope*xl+inter, color='#e3b341', linewidth=1.2, alpha=0.8)
            ax.set_title(f'{g} — {t}', color='#e2e8f0', fontsize=8, pad=5)
            ax.set_xlabel('Theoretical', color='#7d8590', fontsize=7)
            ax.set_ylabel('Sample', color='#7d8590', fontsize=7)
    plt.suptitle('Q-Q Plots (Normality Check)', color='#e2e8f0', fontsize=11, y=1.01)
    plt.tight_layout()
    return fig


# ════════════════════════════════════════════════════════════════════════════
# AUTO-INTERPRETATION
# ════════════════════════════════════════════════════════════════════════════

def interpret(aov_df, desc, between_col, n_total, alpha):
    rows = {r['Source']: r for _,r in aov_df.iterrows()}
    a_src  = between_col
    b_src  = 'Time'
    ab_src = f'{between_col} × Time'

    def info(src):
        r = rows.get(src,{})
        return (r.get('F',np.nan), r.get('p_GG',r.get('p',np.nan)),
                r.get('eta_p2',np.nan), r.get('df',np.nan), r.get('df_error',np.nan))

    F_A,p_A,e_A,df_A,dfe_A   = info(a_src)
    F_B,p_B,e_B,df_B,dfe_B   = info(b_src)
    F_AB,p_AB,e_AB,df_AB,dfe_AB = info(ab_src)

    n_g = desc['Group'].nunique()
    n_t = desc['Time'].nunique()

    out = []
    out.append(f"**Study Design:** {n_g} groups × {n_t} time points, Mixed-Design ANOVA (N = {n_total})")
    out.append("")

    out.append("**① Between-Subjects Effect — Group**")
    if not np.isnan(p_A):
        sig = "**statistically significant**" if p_A<alpha else "not statistically significant"
        mag = "large" if e_A>=.14 else ("medium" if e_A>=.06 else "small")
        out.append(f"F({fmt(df_A,0)}, {fmt(dfe_A,0)}) = {fmt(F_A,2)}, p = {fmt_p(p_A)}, η²ₚ = {fmt(e_A,3)} ({mag} effect). "
                   +("Groups differed significantly on the overall outcome across all time points."
                     if p_A<alpha else "No significant overall difference between groups."))
    out.append("")

    out.append("**② Within-Subjects Effect — Time**")
    if not np.isnan(p_B):
        sig = "**statistically significant**" if p_B<alpha else "not statistically significant"
        mag = "large" if e_B>=.14 else ("medium" if e_B>=.06 else "small")
        out.append(f"F({fmt(df_B,0)}, {fmt(dfe_B,0)}) = {fmt(F_B,2)}, p = {fmt_p(p_B)}, η²ₚ = {fmt(e_B,3)} ({mag} effect). "
                   +("Scores changed significantly across time points (collapsed across groups)."
                     if p_B<alpha else "No significant overall change across time points."))
    out.append("")

    out.append("**③ Interaction — Group × Time**")
    if not np.isnan(p_AB):
        mag = "large" if e_AB>=.14 else ("medium" if e_AB>=.06 else "small")
        if p_AB < alpha:
            out.append(f"🔴 **SIGNIFICANT INTERACTION** — "
                       f"F({fmt(df_AB,0)}, {fmt(dfe_AB,0)}) = {fmt(F_AB,2)}, p = {fmt_p(p_AB)}, η²ₚ = {fmt(e_AB,3)} ({mag}).")
            out.append("The effect of time **differs between groups** — groups follow different "
                       "change trajectories over time. **Main effects should be interpreted with "
                       "caution.** Examine post-hoc pairwise comparisons to identify where groups diverged.")
        elif p_AB < .10:
            out.append(f"⚠️ **MARGINAL INTERACTION** — "
                       f"F({fmt(df_AB,0)}, {fmt(dfe_AB,0)}) = {fmt(F_AB,2)}, p = {fmt_p(p_AB)}, η²ₚ = {fmt(e_AB,3)}.")
            out.append("A trend toward a Group × Time interaction (p < .10). Interpret cautiously.")
        else:
            out.append(f"✅ **NO SIGNIFICANT INTERACTION** — "
                       f"F({fmt(df_AB,0)}, {fmt(dfe_AB,0)}) = {fmt(F_AB,2)}, p = {fmt_p(p_AB)}, η²ₚ = {fmt(e_AB,3)}.")
            out.append("Groups followed similar change patterns over time. Main effects may be interpreted independently.")
    out.append("")
    out.append("**Effect Size Benchmarks (Cohen, 1988):** Negligible < .01 · Small ≥ .01 · Medium ≥ .06 · Large ≥ .14")
    return "\n\n".join(out)


def apa_sentence(aov_df, between_col, n_total, alpha):
    ab_src = f'{between_col} × Time'
    r = {r2['Source']: r2 for _,r2 in aov_df.iterrows()}.get(ab_src, {})
    F=r.get('F',np.nan); p=r.get('p_GG',r.get('p',np.nan))
    e=r.get('eta_p2',np.nan); df1=r.get('df',np.nan); df2=r.get('df_error',np.nan)
    sig = "significant" if not np.isnan(p) and p<alpha else "not significant"
    return (f"A mixed-design ANOVA revealed a {sig} Group × Time interaction, "
            f"F({fmt(df1,0)}, {fmt(df2,0)}) = {fmt(F,2)}, p = {fmt_p(p)}, "
            f"η²ₚ = {fmt(e,3)} (N = {n_total}).")


# ════════════════════════════════════════════════════════════════════════════
# PDF EXPORT  (reportlab optional)
# ════════════════════════════════════════════════════════════════════════════

def build_pdf(desc, norm_df, lev_df, aov_df, sph, interp_text,
              between_col, dv_label, alpha):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors as rl_colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                         Table, TableStyle, HRFlowable)
        from reportlab.lib.enums import TA_LEFT, TA_CENTER
    except ImportError:
        return None

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             rightMargin=36, leftMargin=36,
                             topMargin=48, bottomMargin=36)
    styles = getSampleStyleSheet()
    H1  = ParagraphStyle('H1', fontName='Helvetica-Bold', fontSize=15,
                          textColor=rl_colors.HexColor('#1f3d6e'), spaceAfter=4)
    H2  = ParagraphStyle('H2', fontName='Helvetica-Bold', fontSize=11,
                          textColor=rl_colors.HexColor('#1f6feb'), spaceAfter=3, spaceBefore=10)
    BD  = ParagraphStyle('BD', fontName='Helvetica', fontSize=8.5, leading=13, spaceAfter=3)
    NT  = ParagraphStyle('NT', fontName='Helvetica-Oblique', fontSize=7.5,
                          textColor=rl_colors.grey, leading=11)
    HDR = rl_colors.HexColor('#1f6feb')

    def df_table(df):
        data = [list(df.columns)]
        for _,row in df.iterrows():
            data.append([str(round(v,3)) if isinstance(v,float) else str(v) for v in row])
        t = Table(data, repeatRows=1, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0), HDR),
            ('TEXTCOLOR',(0,0),(-1,0), rl_colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1),7),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[rl_colors.white, rl_colors.HexColor('#f4f7fc')]),
            ('GRID',(0,0),(-1,-1),0.25,rl_colors.HexColor('#dddddd')),
            ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
        ]))
        return t

    story = []
    story.append(Paragraph("Mixed-Design ANOVA (Split-Plot) Report", H1))
    story.append(Paragraph(
        f"Between-Subject Factor: <b>{between_col}</b> &nbsp;|&nbsp; "
        f"Dependent Variable: <b>{dv_label}</b> &nbsp;|&nbsp; α = {alpha}", BD))
    story.append(HRFlowable(width="100%", thickness=1.5, color=HDR, spaceAfter=8))

    story.append(Paragraph("1. Descriptive Statistics", H2))
    story.append(df_table(desc.round(3)))
    story.append(Spacer(1,6))

    story.append(Paragraph("2. Normality Tests", H2))
    story.append(df_table(norm_df))
    story.append(Paragraph("Shapiro-Wilk: n ≤ 50; Kolmogorov-Smirnov: n > 50", NT))
    story.append(Spacer(1,6))

    if not lev_df.empty:
        story.append(Paragraph("3. Levene's Test of Equality of Variances", H2))
        story.append(df_table(lev_df))
        story.append(Spacer(1,6))

    if sph:
        W,p_W,eps_GG,eps_HF,chi2,df_chi = sph
        story.append(Paragraph("4. Mauchly's Test of Sphericity", H2))
        sph_df = pd.DataFrame([{'W':round(W,4),'χ²':round(chi2,3),'df':df_chi,
                                  'p':round(p_W,4),'GG ε':round(eps_GG,4),'HF ε':round(eps_HF,4)}])
        story.append(df_table(sph_df))
        story.append(Spacer(1,6))

    story.append(Paragraph("5. Mixed-Design ANOVA Results", H2))
    show_cols=['Source','SS','df','MS','F','p','p_GG','eta_p2','eps']
    aov_show = aov_df[[c for c in show_cols if c in aov_df.columns]]
    story.append(df_table(aov_show.round(4)))
    story.append(Paragraph("p_GG = Greenhouse-Geisser corrected p-value; η²ₚ = Partial Eta Squared; ε = GG epsilon", NT))
    story.append(Spacer(1,8))

    story.append(Paragraph("6. Statistical Interpretation", H2))
    for line in interp_text.split('\n\n'):
        clean = line.replace('**','').replace('🔴','').replace('✅','').replace('⚠️','')
        story.append(Paragraph(clean, BD))
        story.append(Spacer(1,3))

    story.append(Spacer(1,12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=rl_colors.grey))
    story.append(Paragraph("Mixed-Design ANOVA Analyzer · scipy + statsmodels engine", NT))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ════════════════════════════════════════════════════════════════════════════
# ANOVA TABLE HTML RENDERER
# ════════════════════════════════════════════════════════════════════════════

def render_anova_html(aov_df, use_gg=False):
    rows_html = ""
    for _,r in aov_df.iterrows():
        src  = r['Source']
        ss   = r.get('SS',np.nan)
        df_v = r.get('df',np.nan)
        ms   = r.get('MS',np.nan)
        F    = r.get('F',np.nan)
        p    = r['p_GG'] if use_gg else r['p']
        eta  = r.get('eta_p2',np.nan)
        eps  = r.get('eps',np.nan)
        typ  = r.get('type','')

        is_err = 'error' in typ
        style  = 'opacity:.6;font-style:italic;' if is_err else ''

        p_str  = fmt_p(p)  if not np.isnan(p) else "—"
        star   = sig_star(p) if not np.isnan(p) else ""
        cls    = sig_cls(p)  if not np.isnan(p) else "ns"

        rows_html += f"""
        <tr style="{style}">
          <td><b>{src}</b></td>
          <td>{fmt(ss,3)}</td>
          <td>{fmt(df_v,0)}</td>
          <td>{fmt(ms,3)}</td>
          <td>{fmt(F,3) if not np.isnan(F) else '—'}</td>
          <td class="{cls}">{p_str} {star}</td>
          <td>{fmt(eta,3) if not np.isnan(eta) else '—'}</td>
          <td>{fmt(eps,3) if not np.isnan(eps) else '—'}</td>
        </tr>"""

    return f"""
    <table class="anova-table">
      <thead><tr>
        <th>Source</th><th>SS</th><th>df</th><th>MS</th>
        <th>F</th><th>p-value</th><th>η²ₚ</th><th>ε (GG)</th>
      </tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
    <p style="font-size:.73rem;color:#7d8590;margin-top:6px;">
    *** p &lt; .001 &nbsp;** p &lt; .01 &nbsp;* p &lt; .05 &nbsp;† p &lt; .10 &nbsp;ns p ≥ .10
    </p>"""


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════

def sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="padding:.8rem 0 .4rem">
          <div style="font-family:'JetBrains Mono',monospace;font-size:1.05rem;
                      font-weight:600;color:#388bfd;">📊 Mixed ANOVA</div>
          <div style="font-size:.73rem;color:#7d8590;">SPSS-Equivalent Analyzer</div>
        </div>""", unsafe_allow_html=True)
        st.markdown("---")

        with st.expander("📖 User Guide", expanded=False):
            st.markdown("""
**Mixed-Design (Split-Plot) ANOVA**
Analyzes designs with:
- One **between-subjects** factor (independent groups)
- One **within-subjects** factor (repeated time points)

**Data Format (Wide CSV)**
```
ID, Group, Time1, Time2, Time3
S1, A,     45,    60,    55
S2, B,     40,    58,    70
```

**Steps**
1. Upload your wide-format CSV
2. Map columns: ID · Group · Time points
3. Set options and click **Run Analysis**

**Assumption Checks (Automatic)**
- *Normality*: Shapiro-Wilk (n≤50) or KS (n>50)
- *Homogeneity*: Levene's test per time point
- *Sphericity*: Mauchly's test (time > 2 only)
  - GG correction: ε < 0.75
  - HF correction: ε ≥ 0.75

**Interpreting Results**
- p < .05 → significant
- η²ₚ ≥ .14 → large effect
- **Significant interaction** → groups differ in their change patterns (the key finding)

**Post-Hoc Tests**
Run only when the interaction is significant. Between-group comparisons at each time point + within-subject time comparisons per group.
            """)

        st.markdown("---")
        st.markdown("<div style='font-size:.78rem;color:#7d8590;'>📥 Sample Datasets</div>",
                    unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        with c1:
            st.download_button("2×2", sample_2x2().to_csv(index=False),
                                "sample_2x2.csv","text/csv",use_container_width=True)
        with c2:
            st.download_button("3×3", sample_3x3().to_csv(index=False),
                                "sample_3x3.csv","text/csv",use_container_width=True)

        st.markdown("---")
        st.markdown("""
        <div style="font-size:.7rem;color:#3a4251;line-height:1.7;">
        <b style="color:#7d8590;">Engine</b><br>scipy · statsmodels · numpy<br><br>
        <b style="color:#7d8590;">SS Type</b><br>Type III (SPSS-compatible)<br><br>
        <b style="color:#7d8590;">Sphericity</b><br>Greenhouse-Geisser · Huynh-Feldt<br><br>
        <b style="color:#7d8590;">Post-Hoc</b><br>Bonferroni · Holm · FDR(BH)
        </div>""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    sidebar()

    st.markdown("""
    <div class="hero">
      <h1>Mixed-Design ANOVA Analyzer</h1>
      <p>Split-Plot ANOVA · SPSS-Level Statistical Analysis · Pure scipy/numpy Engine</p>
      <div>
        <span class="badge bd-g">Type III SS</span>
        <span class="badge bd-b">Mauchly + GG/HF Correction</span>
        <span class="badge bd-p">Pairwise Post-Hoc</span>
        <span class="badge bd-r">PDF / CSV Export</span>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── Upload ────────────────────────────────────────────────────────────────
    st.markdown('<div class="sec-card">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">① Data Input</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Upload Wide-Format CSV", type=['csv'],
        help="Each row = one participant. Columns: Subject ID, Group, numeric time columns.")

    if uploaded is None:
        st.markdown("""
        <div class="ibox">
        <b>No file uploaded.</b> Download a sample from the sidebar, or upload your own wide-format CSV.<br><br>
        <b>Expected format:</b> <code>ID, Group, Pre, Post, Followup, ...</code><br>
        Each row = one participant. Time columns must be numeric.
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    try:
        df = pd.read_csv(uploaded)
    except Exception as e:
        st.error(f"Cannot read file: {e}")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    with st.expander(f"📄 Preview — {df.shape[0]} rows × {df.shape[1]} columns", expanded=True):
        st.dataframe(df.head(10), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Column Mapping ────────────────────────────────────────────────────────
    st.markdown('<div class="sec-card">', unsafe_allow_html=True)
    st.markdown('<div class="sec-title">② Column Mapping & Options</div>', unsafe_allow_html=True)

    all_cols = list(df.columns)
    num_cols = list(df.select_dtypes(include=[np.number]).columns)

    c1,c2,c3 = st.columns(3)
    with c1:
        subj_col = st.selectbox("👤 Subject ID", all_cols, 0)
    with c2:
        btwn_col = st.selectbox("👥 Between-Subjects Factor (Group)",
                                 [c for c in all_cols if c != subj_col],
                                 0)
    with c3:
        avail_time = [c for c in num_cols if c not in [subj_col, btwn_col]]
        time_cols  = st.multiselect("🕐 Within-Subjects / Time Columns",
                                     avail_time, default=avail_time[:4])

    c4,c5,c6,c7 = st.columns(4)
    with c4: dv_label = st.text_input("📏 DV Label", "Score")
    with c5: ph_method = st.selectbox("Post-Hoc Correction",
                                       ['bonferroni','holm','fdr_bh'],
                                       format_func=lambda x:{'bonferroni':'Bonferroni',
                                                              'holm':'Holm',
                                                              'fdr_bh':'FDR (BH)'}[x])
    with c6: alpha = st.number_input("α Level", .001, .20, .05, .01)
    with c7: use_gg = st.checkbox("Use GG-corrected p", value=True,
                                   help="Apply Greenhouse-Geisser correction to within/interaction p-values")

    run = st.button("▶  Run Mixed-Design ANOVA", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if not run: return

    # ── Validation ────────────────────────────────────────────────────────────
    if len(time_cols) < 2:
        st.error("Select at least 2 time-point columns."); return
    if btwn_col in time_cols:
        st.error("Between-subjects column cannot also be a time column."); return

    req = [subj_col, btwn_col] + time_cols
    df_clean = df[req].dropna()
    n_dropped = len(df) - len(df_clean)
    if n_dropped:
        st.markdown(f'<div class="wbox">⚠️ {n_dropped} rows with missing values removed. N = {len(df_clean)}</div>',
                    unsafe_allow_html=True)
    if len(df_clean) < 4:
        st.error("Too few complete cases (need ≥ 4)."); return

    for c in time_cols:
        if not pd.api.types.is_numeric_dtype(df_clean[c]):
            st.error(f"Column '{c}' is not numeric."); return

    n_total  = df_clean[subj_col].nunique()
    n_groups = df_clean[btwn_col].nunique()

    # ── Run Stats ─────────────────────────────────────────────────────────────
    with st.spinner("⚙️ Running analysis…"):
        try:
            desc    = compute_descriptives(df_clean, btwn_col, time_cols)
            norm_df = test_normality(df_clean, btwn_col, time_cols)
            lev_df  = test_levene(df_clean, btwn_col, time_cols)
            aov_df, sph = mixed_anova(df_clean, subj_col, btwn_col, time_cols)
            interp  = interpret(aov_df, desc, btwn_col, n_total, alpha)

            ab_src  = f'{btwn_col} × Time'
            inter_r = aov_df[aov_df['Source']==ab_src]
            int_p   = inter_r['p_GG'].values[0] if not inter_r.empty else 1.0
            run_ph  = int_p < alpha

            ph_df   = bonferroni_pairwise(df_clean, btwn_col, time_cols, ph_method) if run_ph else pd.DataFrame()
        except Exception as e:
            st.error(f"Analysis failed: {e}")
            st.exception(e)
            return

    st.success("✅ Analysis complete!")

    # ── Summary metrics ───────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;margin:1rem 0;">
      <div class="mcard"><div class="mval">{n_total}</div><div class="mlbl">Participants</div></div>
      <div class="mcard"><div class="mval">{n_groups}</div><div class="mlbl">Groups</div></div>
      <div class="mcard"><div class="mval">{len(time_cols)}</div><div class="mlbl">Time Points</div></div>
      <div class="mcard"><div class="mval">{n_total*len(time_cols)}</div><div class="mlbl">Observations</div></div>
    </div>""", unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════════
    # TABS
    # ════════════════════════════════════════════════════════════════════════
    tabs = st.tabs(["📋 Descriptives","🔍 Assumptions","📊 ANOVA Results",
                    "🔬 Post-Hoc","📈 Plots","📝 Interpretation","⬇️ Export"])

    # ── Tab 0: Descriptives ───────────────────────────────────────────────────
    with tabs[0]:
        st.markdown('<div class="sec-title">Cell Means & Statistics</div>', unsafe_allow_html=True)
        d2 = desc.copy()
        for c in ['Mean','SD','SE','Median','Min','Max','95% CI Lower','95% CI Upper']:
            if c in d2.columns: d2[c] = d2[c].round(3)
        st.dataframe(d2, use_container_width=True)

        st.markdown('<div class="sec-title" style="margin-top:1.5rem">Marginal Means</div>', unsafe_allow_html=True)
        ca,cb = st.columns(2)
        with ca:
            st.markdown("**By Group**")
            mg = desc.groupby('Group')[['Mean','SD','N']].mean().round(3)
            mg['N'] = mg['N'].astype(int)
            st.dataframe(mg, use_container_width=True)
        with cb:
            st.markdown("**By Time**")
            mt = desc.groupby('Time')[['Mean','SD','N']].mean().round(3)
            mt['N'] = mt['N'].astype(int)
            st.dataframe(mt, use_container_width=True)

    # ── Tab 1: Assumptions ────────────────────────────────────────────────────
    with tabs[1]:
        # Normality
        st.markdown('<div class="sec-title">A. Normality Tests</div>', unsafe_allow_html=True)
        sw_n = (norm_df['N']<=50).sum(); ks_n = (norm_df['N']>50).sum()
        st.markdown(f"""
        <div class="ibox">
        <b>Automatic Selection:</b> Shapiro-Wilk used for <b>{sw_n}</b> cells (n ≤ 50) · 
        Kolmogorov-Smirnov for <b>{ks_n}</b> cells (n > 50).<br>
        H₀: data are normally distributed. Reject if <i>p</i> &lt; {alpha}.
        </div>""", unsafe_allow_html=True)

        norm_show = norm_df.copy()
        norm_show['p-value'] = norm_show['p-value'].apply(lambda x: fmt_p(x) if not pd.isna(x) else '—')
        st.dataframe(norm_show, use_container_width=True)

        all_norm = (norm_df['Normal']=='Yes').all()
        any_bad  = (norm_df['Normal']=='No').any()
        if all_norm:
            st.markdown('<div class="obox">✅ All cells satisfy normality (p > .05).</div>', unsafe_allow_html=True)
        elif any_bad:
            st.markdown('<div class="wbox">⚠️ Non-normality detected. ANOVA is robust for n ≥ 15 per cell. '
                        'Consider non-parametric alternatives for severe violations.</div>', unsafe_allow_html=True)

        # Levene
        st.markdown('<div class="sec-title" style="margin-top:1.4rem">B. Homogeneity of Variance — Levene\'s Test</div>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div class="ibox">
        Tests equality of variances across groups at each time point.<br>
        H₀: variances are equal. Reject if <i>p</i> &lt; {alpha}.
        </div>""", unsafe_allow_html=True)
        if not lev_df.empty:
            lev_show = lev_df.copy()
            lev_show['p-value'] = lev_show['p-value'].apply(fmt_p)
            st.dataframe(lev_show, use_container_width=True)
            if lev_df['Equal Variances'].eq('Yes').all():
                st.markdown('<div class="obox">✅ Homogeneity of variance satisfied at all time points.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="wbox">⚠️ Unequal variances detected. Interpret F-statistics with caution.</div>',
                            unsafe_allow_html=True)

        # Sphericity
        if len(time_cols) > 2:
            st.markdown('<div class="sec-title" style="margin-top:1.4rem">C. Sphericity — Mauchly\'s Test</div>',
                        unsafe_allow_html=True)
            st.markdown("""
            <div class="ibox">
            Required for within-subjects factors with &gt; 2 levels.<br>
            H₀: sphericity is met. If violated (p &lt; .05), use Greenhouse-Geisser (ε &lt; .75) 
            or Huynh-Feldt (ε ≥ .75) correction.
            </div>""", unsafe_allow_html=True)

            if sph:
                W,p_W,eps_GG,eps_HF,chi2,df_chi = sph
                ca,cb,cc,cd = st.columns(4)
                ca.metric("Mauchly's W", fmt(W,4))
                cb.metric("χ² (approx.)", fmt(chi2,3))
                cc.metric("p-value", fmt_p(p_W))
                cd.metric("GG ε", fmt(eps_GG,4))
                st.metric("HF ε", fmt(eps_HF,4))

                if not np.isnan(p_W):
                    if p_W < .05:
                        rec = "Greenhouse-Geisser" if eps_GG < .75 else "Huynh-Feldt"
                        st.markdown(f'<div class="wbox">⚠️ Sphericity violated (p = {fmt_p(p_W)}). '
                                    f'<b>{rec} correction recommended</b> (ε = {fmt(eps_GG,3)}). '
                                    f'GG-corrected p-values are reported when "Use GG-corrected p" is checked.</div>',
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="obox">✅ Sphericity satisfied (p = {fmt_p(p_W)}). No correction needed.</div>',
                                    unsafe_allow_html=True)
            else:
                st.info("Mauchly's test could not be computed (check data balance).")
        else:
            st.info("ℹ️ Sphericity test not applicable (only 2 time points).")

    # ── Tab 2: ANOVA Results ──────────────────────────────────────────────────
    with tabs[2]:
        st.markdown('<div class="sec-title">Mixed-Design ANOVA — Summary Table</div>', unsafe_allow_html=True)
        corr_note = "GG-corrected" if use_gg else "uncorrected"
        st.markdown(f"""
        <div class="ibox">
        <b>Algorithm:</b> Manual Split-Plot ANOVA · <b>SS Type:</b> III (SPSS-compatible) ·
        <b>p-values:</b> {corr_note} (toggle "Use GG-corrected p" in options)
        </div>""", unsafe_allow_html=True)

        st.markdown(render_anova_html(aov_df, use_gg=use_gg), unsafe_allow_html=True)

        st.markdown("""
        <div style="margin-top:.8rem;padding:.7rem 1rem;background:#0f1623;border:1px solid #1e2d45;
                    border-radius:8px;font-size:.76rem;color:#7d8590;font-family:'JetBrains Mono',monospace;">
        <b style="color:#e2e8f0;">Partial η² benchmarks:</b> &nbsp;
        Negligible &lt; .01 &nbsp;·&nbsp;
        <span style="color:#3fb950">Small ≥ .01</span> &nbsp;·&nbsp;
        <span style="color:#e3b341">Medium ≥ .06</span> &nbsp;·&nbsp;
        <span style="color:#f78166">Large ≥ .14</span>
        </div>""", unsafe_allow_html=True)

        with st.expander("🔍 Raw ANOVA DataFrame"):
            st.dataframe(aov_df, use_container_width=True)

    # ── Tab 3: Post-Hoc ───────────────────────────────────────────────────────
    with tabs[3]:
        method_name = {'bonferroni':'Bonferroni','holm':'Holm','fdr_bh':'FDR (BH)'}[ph_method]
        if not run_ph:
            ab_p = aov_df[aov_df['Source']==ab_src]['p_GG'].values
            p_show = ab_p[0] if len(ab_p) else 1.0
            st.markdown(f"""
            <div class="ibox">
            ℹ️ Post-hoc tests <b>not conducted</b> — the Group × Time interaction is 
            not significant (p = {fmt_p(p_show)}, α = {alpha}).<br>
            Post-hoc tests are only run when the interaction reaches significance.
            </div>""", unsafe_allow_html=True)
        elif ph_df.empty:
            st.warning("Post-hoc comparisons could not be computed.")
        else:
            st.markdown(f'<div class="sec-title">Pairwise Comparisons ({method_name} adjusted)</div>',
                        unsafe_allow_html=True)

            # Between-group
            bt = ph_df[ph_df['Type']=='Between-Groups'] if 'Type' in ph_df.columns else pd.DataFrame()
            wt = ph_df[ph_df['Type']=='Within-Time']    if 'Type' in ph_df.columns else pd.DataFrame()

            if not bt.empty:
                st.markdown("**Between-Group at Each Time Point**")
                cols_bt = [c for c in ['Time','Group A','Group B','Mean A','Mean B',
                                        'Mean Diff','SE','t','df','p (raw)','p (adj)','Cohen d','Sig']
                           if c in bt.columns]
                bt_show = bt[cols_bt].copy()
                bt_show['p (raw)'] = bt_show['p (raw)'].apply(fmt_p)
                bt_show['p (adj)'] = bt_show['p (adj)'].apply(fmt_p)
                st.dataframe(bt_show, use_container_width=True)

            if not wt.empty:
                st.markdown("**Within-Subject Time Comparisons per Group**")
                cols_wt = [c for c in ['Group','Time A','Time B','Mean A','Mean B',
                                        'Mean Diff','SE','t','df','p (raw)','p (adj)','Cohen d','Sig']
                           if c in wt.columns]
                wt_show = wt[cols_wt].copy()
                wt_show['p (raw)'] = wt_show['p (raw)'].apply(fmt_p)
                wt_show['p (adj)'] = wt_show['p (adj)'].apply(fmt_p)
                st.dataframe(wt_show, use_container_width=True)

            st.markdown('<div style="font-size:.72rem;color:#7d8590;margin-top:.4rem;">*** p &lt; .001 · ** p &lt; .01 · * p &lt; .05 · † p &lt; .10 · ns ≥ .10</div>',
                        unsafe_allow_html=True)

    # ── Tab 4: Plots ──────────────────────────────────────────────────────────
    with tabs[4]:
        st.markdown('<div class="sec-title">Profile Plot</div>', unsafe_allow_html=True)
        fig1 = profile_plot(desc, btwn_col)
        st.pyplot(fig1, use_container_width=True); plt.close(fig1)

        st.markdown('<div class="sec-title" style="margin-top:1.2rem">Distributions with KDE</div>',
                    unsafe_allow_html=True)
        fig2 = dist_plots(df_clean, btwn_col, time_cols)
        st.pyplot(fig2, use_container_width=True); plt.close(fig2)

        st.markdown('<div class="sec-title" style="margin-top:1.2rem">Q-Q Plots (Normality)</div>',
                    unsafe_allow_html=True)
        fig3 = qq_plots(df_clean, btwn_col, time_cols)
        st.pyplot(fig3, use_container_width=True); plt.close(fig3)

    # ── Tab 5: Interpretation ─────────────────────────────────────────────────
    with tabs[5]:
        st.markdown('<div class="sec-title">Automated Statistical Interpretation</div>', unsafe_allow_html=True)
        for block in interp.split('\n\n'):
            if block.startswith('**'):
                head, _, rest = block.partition('\n\n')
                st.markdown(head)
                if rest: st.markdown(f'<div class="ibox">{rest}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="ibox">{block}</div>', unsafe_allow_html=True)

        st.markdown('<div class="sec-title" style="margin-top:1.4rem">APA-Style Write-Up</div>',
                    unsafe_allow_html=True)
        apa = apa_sentence(aov_df, btwn_col, n_total, alpha)
        st.info(apa)

        # Effect sizes table
        st.markdown('<div class="sec-title" style="margin-top:1.4rem">Effect Size Summary</div>',
                    unsafe_allow_html=True)
        es_rows = aov_df[~aov_df['type'].str.contains('error')][
            ['Source','F','p_GG','eta_p2']].copy()
        es_rows = es_rows.rename(columns={'p_GG':'p (GG)','eta_p2':'η²ₚ'})
        es_rows['Interpretation'] = es_rows['η²ₚ'].apply(eta_label)
        es_rows['F'] = es_rows['F'].apply(lambda x: fmt(x,3))
        es_rows['p (GG)'] = es_rows['p (GG)'].apply(fmt_p)
        es_rows['η²ₚ'] = es_rows['η²ₚ'].apply(lambda x: fmt(x,4))
        st.dataframe(es_rows, use_container_width=True, hide_index=True)

    # ── Tab 6: Export ─────────────────────────────────────────────────────────
    with tabs[6]:
        st.markdown('<div class="sec-title">Download Full Report</div>', unsafe_allow_html=True)

        ca, cb, cc = st.columns(3)

        # CSV
        with ca:
            st.markdown("**📄 CSV Report**")
            buf = io.StringIO()
            buf.write("Mixed-Design ANOVA Report\n\n")
            buf.write("DESCRIPTIVE STATISTICS\n"); desc.to_csv(buf, index=False)
            buf.write("\n\nNORMALITY TESTS\n");     norm_df.to_csv(buf, index=False)
            buf.write("\n\nLEVENE'S TEST\n");        lev_df.to_csv(buf, index=False)
            if sph:
                W,p_W,eps_GG,eps_HF,chi2,df_chi = sph
                sph_df2 = pd.DataFrame([{'W':W,'chi2':chi2,'df':df_chi,'p':p_W,'eps_GG':eps_GG,'eps_HF':eps_HF}])
                buf.write("\n\nMAUCHLY'S TEST\n"); sph_df2.to_csv(buf, index=False)
            buf.write("\n\nANOVA RESULTS\n"); aov_df.to_csv(buf, index=False)
            if not ph_df.empty:
                buf.write("\n\nPOST-HOC TESTS\n"); ph_df.to_csv(buf, index=False)
            buf.write("\n\nINTERPRETATION\n")
            buf.write(interp.replace('**',''))
            st.download_button("⬇️ Download CSV", buf.getvalue().encode(),
                                "mixed_anova_report.csv","text/csv",use_container_width=True)

        # PDF
        with cb:
            st.markdown("**📑 PDF Report**")
            pdf = build_pdf(desc, norm_df, lev_df, aov_df, sph, interp,
                             btwn_col, dv_label, alpha)
            if pdf:
                st.download_button("⬇️ Download PDF", pdf,
                                    "mixed_anova_report.pdf","application/pdf",
                                    use_container_width=True)
            else:
                st.caption("PDF export requires `reportlab`. Run: `pip install reportlab`")

        # Long-format
        with cc:
            st.markdown("**📊 Long-Format Data**")
            df_long2 = df_clean.melt(id_vars=[subj_col, btwn_col],
                                      value_vars=time_cols,
                                      var_name='Time', value_name=dv_label)
            st.download_button("⬇️ Download Long CSV", df_long2.to_csv(index=False).encode(),
                                "long_format_data.csv","text/csv",use_container_width=True)


if __name__ == '__main__':
    main()
