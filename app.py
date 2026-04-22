"""
Mixed-Design ANOVA (Split-Plot ANOVA) Web Application
SPSS-equivalent statistical analysis using pingouin
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from scipy import stats
from scipy.stats import levene
import pingouin as pg
from itertools import combinations
import warnings
import io
import csv
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

warnings.filterwarnings('ignore')

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Mixed-Design ANOVA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

.stApp {
    background: #0d1117;
    color: #e6edf3;
}

.main-header {
    background: linear-gradient(135deg, #161b22 0%, #0d1117 100%);
    border: 1px solid #30363d;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}

.main-header::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: linear-gradient(90deg, #2ea043, #58a6ff, #bc8cff);
}

.main-header h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.8rem;
    font-weight: 600;
    color: #e6edf3;
    margin: 0 0 0.5rem 0;
    letter-spacing: -0.5px;
}

.main-header p {
    color: #8b949e;
    font-size: 0.95rem;
    margin: 0;
}

.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
    font-family: 'IBM Plex Mono', monospace;
    margin-right: 6px;
}

.badge-green { background: #1a3a2a; color: #3fb950; border: 1px solid #2ea043; }
.badge-blue  { background: #0d2137; color: #58a6ff; border: 1px solid #1f6feb; }
.badge-purple{ background: #1e1433; color: #bc8cff; border: 1px solid #8957e5; }

.section-card {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 10px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
}

.section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    font-weight: 600;
    color: #8b949e;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid #21262d;
}

.stat-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.875rem;
    font-family: 'IBM Plex Mono', monospace;
}

.stat-table th {
    background: #21262d;
    color: #58a6ff;
    padding: 8px 12px;
    text-align: left;
    font-weight: 600;
    border-bottom: 2px solid #30363d;
}

.stat-table td {
    padding: 7px 12px;
    border-bottom: 1px solid #21262d;
    color: #c9d1d9;
}

.stat-table tr:hover td { background: #1c2128; }

.sig-yes { color: #3fb950; font-weight: 600; }
.sig-no  { color: #f78166; }
.sig-mar { color: #e3b341; }

.interp-box {
    background: #0d2137;
    border: 1px solid #1f6feb;
    border-left: 4px solid #58a6ff;
    border-radius: 8px;
    padding: 1.25rem;
    margin: 1rem 0;
    font-size: 0.9rem;
    line-height: 1.7;
    color: #c9d1d9;
}

.warn-box {
    background: #2d2000;
    border: 1px solid #9e6a03;
    border-left: 4px solid #e3b341;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.75rem 0;
    font-size: 0.875rem;
    color: #e3b341;
}

.ok-box {
    background: #0a1f12;
    border: 1px solid #2ea043;
    border-left: 4px solid #3fb950;
    border-radius: 8px;
    padding: 1rem;
    margin: 0.75rem 0;
    font-size: 0.875rem;
    color: #3fb950;
}

.metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 1rem;
    margin: 1rem 0;
}

.metric-card {
    background: #21262d;
    border: 1px solid #30363d;
    border-radius: 8px;
    padding: 1rem;
    text-align: center;
}

.metric-val {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #58a6ff;
}

.metric-label {
    font-size: 0.78rem;
    color: #8b949e;
    margin-top: 4px;
}

/* Streamlit overrides */
.stButton > button {
    background: #21262d;
    border: 1px solid #30363d;
    color: #e6edf3;
    border-radius: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.85rem;
    transition: all 0.2s;
}
.stButton > button:hover {
    background: #30363d;
    border-color: #58a6ff;
    color: #58a6ff;
}

.stSelectbox > div > div, .stMultiSelect > div > div {
    background: #161b22 !important;
    border-color: #30363d !important;
    color: #e6edf3 !important;
}

div[data-testid="stSidebar"] {
    background: #0d1117;
    border-right: 1px solid #21262d;
}

.stTabs [data-baseweb="tab-list"] {
    background: transparent;
    gap: 4px;
}
.stTabs [data-baseweb="tab"] {
    background: #161b22;
    border: 1px solid #30363d;
    color: #8b949e;
    border-radius: 6px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
}
.stTabs [aria-selected="true"] {
    background: #1f6feb !important;
    border-color: #58a6ff !important;
    color: #e6edf3 !important;
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

def sig_star(p):
    if p < .001: return "***"
    if p < .01:  return "**"
    if p < .05:  return "*"
    if p < .10:  return "†"
    return "ns"

def sig_color(p):
    if p < .05:  return "sig-yes"
    if p < .10:  return "sig-mar"
    return "sig-no"

def fmt(x, dec=3):
    if pd.isna(x): return "—"
    if isinstance(x, (int, np.integer)): return str(x)
    return f"{x:.{dec}f}"

def fmt_p(p):
    if pd.isna(p): return "—"
    if p < .001: return "< .001"
    return f".{int(round(p*1000)):03d}"[:-1] if p < 1 else f"{p:.3f}"

def fmt_p_val(p):
    """Format p-value like SPSS"""
    if pd.isna(p): return "—"
    if p < .001: return "< .001"
    return f"{p:.3f}"


# ════════════════════════════════════════════════════════════════════════════
# SAMPLE DATA GENERATORS
# ════════════════════════════════════════════════════════════════════════════

def generate_sample_2x2():
    np.random.seed(42)
    n_per = 15
    data = []
    groups = ['Control', 'Treatment']
    means = {'Control': [50, 52], 'Treatment': [50, 62]}
    for g in groups:
        for i in range(n_per):
            pre  = np.random.normal(means[g][0], 8)
            post = np.random.normal(means[g][1], 8)
            data.append({'ID': f'{g[0]}{i+1:02d}', 'Group': g,
                         'Pre': round(pre,2), 'Post': round(post,2)})
    return pd.DataFrame(data)

def generate_sample_3x3():
    np.random.seed(99)
    n_per = 12
    data = []
    groups  = ['Control','Low_Dose','High_Dose']
    means   = {'Control':[40,42,41],'Low_Dose':[40,50,55],'High_Dose':[40,58,70]}
    for g in groups:
        for i in range(n_per):
            vals = [np.random.normal(means[g][t], 9) for t in range(3)]
            data.append({'ID': f'{g[0]}{i+1:02d}', 'Group': g,
                         'Pre': round(vals[0],2),
                         'Mid': round(vals[1],2),
                         'Post': round(vals[2],2)})
    return pd.DataFrame(data)


# ════════════════════════════════════════════════════════════════════════════
# DESCRIPTIVE STATISTICS
# ════════════════════════════════════════════════════════════════════════════

def compute_descriptives(df_long, between_col, within_col, dv_col):
    desc = df_long.groupby([between_col, within_col])[dv_col].agg(
        N='count', Mean='mean', SD='std', SE=lambda x: x.std()/np.sqrt(len(x)),
        Min='min', Max='max',
        Median='median'
    ).reset_index()
    desc['95% CI Lower'] = desc['Mean'] - 1.96 * desc['SE']
    desc['95% CI Upper'] = desc['Mean'] + 1.96 * desc['SE']
    return desc


# ════════════════════════════════════════════════════════════════════════════
# NORMALITY TESTS
# ════════════════════════════════════════════════════════════════════════════

def test_normality(df_long, between_col, within_col, dv_col):
    results = []
    groups  = df_long[between_col].unique()
    times   = df_long[within_col].unique()
    for g in groups:
        for t in times:
            sub = df_long[(df_long[between_col]==g)&(df_long[within_col]==t)][dv_col].dropna()
            n   = len(sub)
            # Choose test: SW for n<=50, KS for n>50 (SPSS convention)
            if n < 3:
                results.append({'Group':g,'Time':t,'N':n,'Test':'N/A',
                                 'Statistic':np.nan,'p-value':np.nan,'Normal':None,'Recommendation':''})
                continue
            if n <= 50:
                stat, p = stats.shapiro(sub)
                test_nm = 'Shapiro-Wilk'
                rec = 'SW preferred (n ≤ 50)'
            else:
                stat, p = stats.kstest(sub, 'norm', args=(sub.mean(), sub.std()))
                test_nm = 'Kolmogorov-Smirnov'
                rec = 'KS preferred (n > 50)'
            results.append({'Group':g,'Time':t,'N':n,'Test':test_nm,
                             'Statistic':stat,'p-value':p,
                             'Normal': p > .05,'Recommendation':rec})
    return pd.DataFrame(results)


# ════════════════════════════════════════════════════════════════════════════
# LEVENE'S TEST
# ════════════════════════════════════════════════════════════════════════════

def test_levene(df_long, between_col, within_col, dv_col):
    results = []
    for t in df_long[within_col].unique():
        groups_data = [
            df_long[(df_long[between_col]==g)&(df_long[within_col]==t)][dv_col].dropna().values
            for g in df_long[between_col].unique()
        ]
        if all(len(g) > 1 for g in groups_data):
            stat, p = levene(*groups_data, center='mean')
            results.append({'Time Point': t, 'Levene F': stat, 'p-value': p,
                             'Equal Variances': p > .05})
    return pd.DataFrame(results)


# ════════════════════════════════════════════════════════════════════════════
# MIXED ANOVA
# ════════════════════════════════════════════════════════════════════════════

def run_mixed_anova(df_long, subject_col, between_col, within_col, dv_col):
    aov = pg.mixed_anova(dv=dv_col, within=within_col,
                          between=between_col, subject=subject_col,
                          data=df_long, correction=True)
    return aov

def run_mauchly(df_long, subject_col, within_col, dv_col):
    try:
        sph = pg.sphericity(data=df_long, dv=dv_col,
                            within=within_col, subject=subject_col)
        return sph
    except Exception:
        return None


# ════════════════════════════════════════════════════════════════════════════
# POST-HOC TESTS
# ════════════════════════════════════════════════════════════════════════════

def run_posthoc(df_long, between_col, within_col, dv_col, subject_col, method='bonf'):
    results = []
    times  = sorted(df_long[within_col].unique())
    groups = sorted(df_long[between_col].unique())

    # Between-group comparisons at each time point
    for t in times:
        sub = df_long[df_long[within_col]==t]
        if len(groups) > 1:
            ph = pg.pairwise_tests(data=sub, dv=dv_col,
                                    between=between_col,
                                    padjust=method)
            for _, row in ph.iterrows():
                results.append({
                    'Comparison Type': 'Between-Groups',
                    'Time Point': t,
                    'Group A': row['A'], 'Group B': row['B'],
                    'Mean Diff': row['mean(A)'] - row['mean(B)'],
                    'SE': row.get('se', np.nan),
                    't': row['T'], 'df': row['dof'],
                    'p (adj)': row['p-corr'] if 'p-corr' in row else row.get('p-unc', np.nan),
                    'p (unadj)': row.get('p-unc', np.nan),
                    'Cohen d': row.get('cohen-d', np.nan),
                    'Sig': sig_star(row['p-corr'] if 'p-corr' in row else row.get('p-unc', np.nan))
                })

    # Within-subject (time) comparisons per group
    for g in groups:
        sub = df_long[df_long[between_col]==g]
        if len(times) > 1:
            ph = pg.pairwise_tests(data=sub, dv=dv_col,
                                    within=within_col, subject=subject_col,
                                    padjust=method)
            for _, row in ph.iterrows():
                results.append({
                    'Comparison Type': 'Within-Time',
                    'Group': g,
                    'Time A': row['A'], 'Time B': row['B'],
                    'Mean Diff': row['mean(A)'] - row['mean(B)'],
                    'SE': row.get('se', np.nan),
                    't': row['T'], 'df': row['dof'],
                    'p (adj)': row['p-corr'] if 'p-corr' in row else row.get('p-unc', np.nan),
                    'p (unadj)': row.get('p-unc', np.nan),
                    'Cohen d': row.get('cohen-d', np.nan),
                    'Sig': sig_star(row['p-corr'] if 'p-corr' in row else row.get('p-unc', np.nan))
                })

    return pd.DataFrame(results) if results else pd.DataFrame()


# ════════════════════════════════════════════════════════════════════════════
# INTERACTION PLOT
# ════════════════════════════════════════════════════════════════════════════

def make_interaction_plot(desc, between_col, within_col, title='Profile Plot'):
    plt.style.use('dark_background')
    fig, axes = plt.subplots(1, 2, figsize=(13, 5))
    fig.patch.set_facecolor('#0d1117')

    palette = ['#58a6ff','#3fb950','#bc8cff','#e3b341','#f78166','#79c0ff']
    groups  = desc[between_col].unique()
    times   = desc[within_col].unique()

    for ax in axes:
        ax.set_facecolor('#161b22')
        ax.spines['bottom'].set_color('#30363d')
        ax.spines['left'].set_color('#30363d')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.tick_params(colors='#8b949e', labelsize=9)
        ax.xaxis.label.set_color('#8b949e')
        ax.yaxis.label.set_color('#8b949e')
        ax.grid(axis='y', color='#21262d', alpha=0.8, linewidth=0.7)

    # Left: Lines with error bars (±1 SE)
    ax1 = axes[0]
    for i, g in enumerate(groups):
        sub = desc[desc[between_col]==g]
        x   = range(len(sub))
        ax1.errorbar(x, sub['Mean'], yerr=sub['SE'],
                     marker='o', markersize=7, linewidth=2.5,
                     color=palette[i % len(palette)],
                     label=str(g), capsize=5, capthick=1.5,
                     elinewidth=1.5, markeredgewidth=2,
                     markeredgecolor='#0d1117')
    ax1.set_xticks(range(len(times)))
    ax1.set_xticklabels(times, fontsize=9)
    ax1.set_xlabel('Time Point', fontsize=10)
    ax1.set_ylabel('Mean Score', fontsize=10)
    ax1.set_title('Profile Plot (±1 SE)', color='#e6edf3', fontsize=11, pad=10)
    ax1.legend(title=between_col, title_fontsize=8, fontsize=8,
               facecolor='#21262d', edgecolor='#30363d', labelcolor='#c9d1d9')

    # Right: Bar chart with CI
    ax2   = axes[1]
    width = 0.8 / len(groups)
    x_pos = np.arange(len(times))
    for i, g in enumerate(groups):
        sub = desc[desc[between_col]==g].sort_values(within_col)
        offs = (i - len(groups)/2 + .5) * width
        bars = ax2.bar(x_pos + offs, sub['Mean'], width=width*0.9,
                       color=palette[i % len(palette)], alpha=0.85,
                       label=str(g), edgecolor='#0d1117', linewidth=0.5)
        ax2.errorbar(x_pos + offs, sub['Mean'],
                     yerr=1.96*sub['SE'], fmt='none',
                     color='white', capsize=3, capthick=1, elinewidth=1, alpha=0.6)
    ax2.set_xticks(x_pos)
    ax2.set_xticklabels(times, fontsize=9)
    ax2.set_xlabel('Time Point', fontsize=10)
    ax2.set_ylabel('Mean Score', fontsize=10)
    ax2.set_title('Bar Chart (95% CI)', color='#e6edf3', fontsize=11, pad=10)
    ax2.legend(title=between_col, title_fontsize=8, fontsize=8,
               facecolor='#21262d', edgecolor='#30363d', labelcolor='#c9d1d9')

    plt.suptitle(title, color='#e6edf3', fontsize=13, fontweight='600', y=1.02)
    plt.tight_layout()
    return fig


# ════════════════════════════════════════════════════════════════════════════
# INTERPRETATION ENGINE
# ════════════════════════════════════════════════════════════════════════════

def generate_interpretation(aov, desc, between_col, within_col, n_total):
    lines = []
    aov_dict = {}
    for _, row in aov.iterrows():
        src = str(row.get('Source', row.get('source', ''))).strip()
        aov_dict[src] = row

    # Identify source names flexibly
    src_names = list(aov_dict.keys())
    between_src = next((s for s in src_names if between_col.lower() in s.lower()), src_names[0] if src_names else None)
    within_src  = next((s for s in src_names if 'time' in s.lower() and 'inter' not in s.lower()), next((s for s in src_names if within_col.lower() in s.lower() or 'within' in s.lower()), None))
    inter_src   = next((s for s in src_names if '*' in s or 'interaction' in s.lower() or 'inter' in s.lower()), None)

    n_groups = desc[between_col].nunique()
    n_times  = desc[within_col].nunique()
    lines.append(f"**Study Design:** {n_groups}-group × {n_times}-time-point Mixed-Design ANOVA (N = {n_total})")
    lines.append("")

    def _row_info(src):
        if src and src in aov_dict:
            r = aov_dict[src]
            p = r.get('p_unc', r.get('p-unc', r.get('p-GG-corr', np.nan)))
            F = r.get('F', np.nan)
            eta = r.get('np2', r.get('eta_sq', np.nan))
            return F, p, eta
        return np.nan, np.nan, np.nan

    # Between-subjects
    F_b, p_b, eta_b = _row_info(between_src)
    lines.append("**① Between-Subjects Effect (Group)**")
    if not np.isnan(p_b):
        sig = "statistically significant" if p_b < .05 else "not statistically significant"
        mag = "large" if (eta_b or 0)>.14 else ("medium" if (eta_b or 0)>.06 else "small")
        lines.append(f"The main effect of Group was {sig} "
                     f"(F = {fmt(F_b,2)}, p = {fmt_p_val(p_b)}, η²ₚ = {fmt(eta_b,3)}). "
                     f"This represents a {mag} effect size. "
                     + ("Groups differed significantly on the outcome, independent of time." if p_b < .05
                        else "No significant overall difference was found between groups."))
    lines.append("")

    # Within-subjects
    F_w, p_w, eta_w = _row_info(within_src)
    lines.append("**② Within-Subjects Effect (Time)**")
    if not np.isnan(p_w):
        sig = "statistically significant" if p_w < .05 else "not statistically significant"
        mag = "large" if (eta_w or 0)>.14 else ("medium" if (eta_w or 0)>.06 else "small")
        lines.append(f"The main effect of Time was {sig} "
                     f"(F = {fmt(F_w,2)}, p = {fmt_p_val(p_w)}, η²ₚ = {fmt(eta_w,3)}). "
                     f"This is a {mag} effect. "
                     + ("Scores changed significantly across time points, collapsed across groups." if p_w < .05
                        else "Scores did not change significantly over time when collapsed across groups."))
    lines.append("")

    # Interaction
    F_i, p_i, eta_i = _row_info(inter_src)
    lines.append("**③ Interaction Effect (Group × Time)**")
    if not np.isnan(p_i):
        mag = "large" if (eta_i or 0)>.14 else ("medium" if (eta_i or 0)>.06 else "small")
        if p_i < .05:
            lines.append(f"🔴 **SIGNIFICANT INTERACTION DETECTED** "
                         f"(F = {fmt(F_i,2)}, p = {fmt_p_val(p_i)}, η²ₚ = {fmt(eta_i,3)}, effect = {mag}).")
            lines.append("The effect of time differed significantly between groups — the groups followed "
                         "different trajectories across time points. This is the key finding: "
                         "**the main effects should be interpreted with caution** because the relationship "
                         "between group and outcome changes across time. Conduct post-hoc pairwise comparisons "
                         "to identify precisely where the groups diverged.")
        elif p_i < .10:
            lines.append(f"⚠️ **MARGINAL INTERACTION** "
                         f"(F = {fmt(F_i,2)}, p = {fmt_p_val(p_i)}, η²ₚ = {fmt(eta_i,3)}).")
            lines.append("A trend toward a Group × Time interaction is observed (p < .10). "
                         "Interpret with caution and consider the practical significance.")
        else:
            lines.append(f"✅ **NO SIGNIFICANT INTERACTION** "
                         f"(F = {fmt(F_i,2)}, p = {fmt_p_val(p_i)}, η²ₚ = {fmt(eta_i,3)}).")
            lines.append("The groups followed similar patterns of change across time. "
                         "The main effects can be interpreted independently.")
    lines.append("")
    lines.append("**Effect Size Benchmarks (Cohen, 1988):** Small: η²ₚ ≥ .01 · Medium: η²ₚ ≥ .06 · Large: η²ₚ ≥ .14")
    return "\n\n".join(lines)


# ════════════════════════════════════════════════════════════════════════════
# RENDER ANOVA TABLE (HTML)
# ════════════════════════════════════════════════════════════════════════════

def render_anova_table(aov):
    rows = ""
    for _, r in aov.iterrows():
        src = str(r.get('Source', r.get('source', '')))
        ss  = r.get('SS', np.nan)
        df  = r.get('DF1', r.get('DF', r.get('ddof1', np.nan)))
        ms  = r.get('MS', ss/df if (not np.isnan(ss) and not np.isnan(df) and df > 0) else np.nan)
        F   = r.get('F', np.nan)
        p   = r.get('p_unc', r.get('p-unc', r.get('p-GG-corr', np.nan)))
        eta = r.get('np2', r.get('eta_sq', np.nan))
        eps = r.get('eps', np.nan)

        cls = sig_color(p) if not np.isnan(p) else ''
        star= sig_star(p)  if not np.isnan(p) else ''
        rows += f"""
        <tr>
            <td><b>{src}</b></td>
            <td>{fmt(ss,3)}</td>
            <td>{fmt(df,0) if not np.isnan(df) else '—'}</td>
            <td>{fmt(ms,3)}</td>
            <td>{fmt(F,3)}</td>
            <td class="{cls}">{fmt_p_val(p)} {star}</td>
            <td>{fmt(eta,3)}</td>
            <td>{fmt(eps,3) if not np.isnan(eps) else '—'}</td>
        </tr>"""

    html = f"""
    <table class="stat-table">
        <thead><tr>
            <th>Source</th><th>SS</th><th>df</th><th>MS</th>
            <th>F</th><th>p-value</th><th>η²ₚ</th><th>ε (GG)</th>
        </tr></thead>
        <tbody>{rows}</tbody>
    </table>
    <p style="font-size:0.75rem;color:#8b949e;margin-top:6px;">
    *** p &lt; .001 &nbsp;** p &lt; .01 &nbsp;* p &lt; .05 &nbsp;† p &lt; .10 &nbsp;ns p ≥ .10
    </p>"""
    return html


# ════════════════════════════════════════════════════════════════════════════
# PDF REPORT GENERATOR
# ════════════════════════════════════════════════════════════════════════════

def generate_pdf_report(desc, norm_df, lev_df, aov, interp_text, between_col, within_col, dv_col):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             rightMargin=40, leftMargin=40,
                             topMargin=50, bottomMargin=40)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Title'],
                                  fontSize=16, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#1a1a2e'),
                                  spaceAfter=6)
    h2_style = ParagraphStyle('H2', parent=styles['Heading2'],
                               fontSize=12, fontName='Helvetica-Bold',
                               textColor=colors.HexColor('#1f6feb'),
                               spaceAfter=4, spaceBefore=12)
    body_style = ParagraphStyle('Body', parent=styles['Normal'],
                                 fontSize=9, fontName='Helvetica',
                                 leading=14, spaceAfter=4)
    note_style = ParagraphStyle('Note', parent=styles['Normal'],
                                 fontSize=8, fontName='Helvetica-Oblique',
                                 textColor=colors.grey, leading=12)

    def make_table(df, title=None):
        items = []
        if title:
            items.append(Paragraph(title, h2_style))
        data = [list(df.columns)]
        for _, row in df.iterrows():
            data.append([str(round(v, 3)) if isinstance(v, float) else str(v) for v in row])
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f6feb')),
            ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
            ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 7.5),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f5f7fa')]),
            ('GRID',       (0,0), (-1,-1), 0.3, colors.HexColor('#dddddd')),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        items.append(t)
        return items

    story = []
    story.append(Paragraph("Mixed-Design ANOVA (Split-Plot) Report", title_style))
    story.append(Paragraph(f"Between-Subject Factor: <b>{between_col}</b> &nbsp;|&nbsp; "
                            f"Within-Subject Factor: <b>{within_col}</b> &nbsp;|&nbsp; "
                            f"Dependent Variable: <b>{dv_col}</b>", body_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#1f6feb')))
    story.append(Spacer(1, 10))

    # Descriptives
    story += make_table(desc.round(3), "1. Descriptive Statistics")
    story.append(Spacer(1, 8))

    # Normality
    story += make_table(norm_df.round(3), "2. Normality Tests")
    story.append(Paragraph("Note: Shapiro-Wilk is used for n ≤ 50; Kolmogorov-Smirnov for n > 50.", note_style))
    story.append(Spacer(1, 8))

    # Levene
    if not lev_df.empty:
        story += make_table(lev_df.round(3), "3. Levene's Test of Equality of Variances")
        story.append(Spacer(1, 8))

    # ANOVA
    story += make_table(aov.round(4), "4. Mixed-Design ANOVA Results")
    story.append(Paragraph("η²ₚ: Partial Eta Squared · ε: Greenhouse-Geisser epsilon", note_style))
    story.append(Spacer(1, 8))

    # Interpretation
    story.append(Paragraph("5. Statistical Interpretation", h2_style))
    for line in interp_text.split('\n\n'):
        clean = line.replace('**', '').replace('①', '①').replace('②', '②').replace('③', '③')
        story.append(Paragraph(clean, body_style))
        story.append(Spacer(1, 4))

    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
    story.append(Paragraph("Generated by Mixed-Design ANOVA Analyzer · Powered by pingouin & Streamlit", note_style))

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="padding:1rem 0;">
            <div style="font-family:'IBM Plex Mono',monospace;font-size:1.1rem;
                        font-weight:600;color:#58a6ff;margin-bottom:4px;">
                📊 Mixed ANOVA
            </div>
            <div style="font-size:0.75rem;color:#8b949e;">SPSS-Equivalent Analyzer</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        with st.expander("📖 User Guide", expanded=False):
            st.markdown("""
**What is Mixed-Design ANOVA?**
Also called Split-Plot ANOVA, it analyzes designs with:
- One **between-subjects** factor (independent groups)
- One **within-subjects** factor (repeated measures / time)

**Data Format (Wide)**
Each row = one participant.
| ID | Group | Pre | Post | Follow-up |
|----|-------|-----|------|-----------|
| S1 | A     | 45  | 60   | 55        |

**Step-by-Step**
1. Upload your CSV file
2. Map columns (ID, Group, Time points)
3. Click **Run Analysis**

**Interpreting Results**
- **p < .05** → statistically significant
- **η²ₚ ≥ .14** → large effect
- **Significant interaction** → groups differ in their change patterns over time

**Assumption Checks**
- *Normality*: SW (n≤50), KS (n>50)
- *Homogeneity*: Levene's test
- *Sphericity*: Mauchly's (time > 2 points); use GG/HF corrections if violated

**Post-Hoc Tests**
- Bonferroni: Conservative, controls family-wise error
- Tukey HSD: Balanced, good for equal group sizes
            """)

        st.markdown("---")
        st.markdown("<div style='font-size:0.8rem;color:#8b949e;'>📥 Sample Datasets</div>",
                    unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            df2 = generate_sample_2x2()
            st.download_button("2×2 Data", df2.to_csv(index=False),
                                file_name="sample_2x2.csv", mime="text/csv",
                                use_container_width=True, key="dl2x2")
        with col2:
            df3 = generate_sample_3x3()
            st.download_button("3×3 Data", df3.to_csv(index=False),
                                file_name="sample_3x3.csv", mime="text/csv",
                                use_container_width=True, key="dl3x3")

        st.markdown("---")
        st.markdown("""
        <div style="font-size:0.72rem;color:#484f58;line-height:1.6;">
        <b style="color:#8b949e;">Statistical Engine</b><br>
        pingouin · scipy · statsmodels<br><br>
        <b style="color:#8b949e;">Sum of Squares</b><br>
        Type III (SPSS-compatible)<br><br>
        <b style="color:#8b949e;">Sphericity Correction</b><br>
        Greenhouse-Geisser & Huynh-Feldt
        </div>
        """, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ════════════════════════════════════════════════════════════════════════════

def main():
    render_sidebar()

    # ── Header ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <h1>Mixed-Design ANOVA Analyzer</h1>
        <p>Split-Plot ANOVA · SPSS-Level Statistical Analysis · Pingouin Engine</p>
        <div style="margin-top:1rem;">
            <span class="badge badge-green">Type III SS</span>
            <span class="badge badge-blue">Sphericity Corrections</span>
            <span class="badge badge-purple">Post-Hoc Tests</span>
            <span class="badge badge-green">PDF Export</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">① Data Input</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Upload Wide-Format CSV",
        type=['csv'],
        help="Each row = one participant. Columns: Subject ID, Group, and multiple time-point measurements."
    )

    if uploaded is None:
        st.markdown("""
        <div class="interp-box">
        <b>No file uploaded yet.</b> Please upload a wide-format CSV, or download a sample dataset from the sidebar.<br><br>
        <b>Expected format:</b><br>
        <code>ID, Group, Time1, Time2, Time3, ...</code><br>
        Each row represents one participant. Time columns contain numeric measurements.
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        return

    try:
        df = pd.read_csv(uploaded)
    except Exception as e:
        st.error(f"❌ Could not read CSV: {e}")
        st.markdown('</div>', unsafe_allow_html=True)
        return

    # Show preview
    with st.expander(f"📄 Data Preview — {df.shape[0]} rows × {df.shape[1]} columns", expanded=True):
        st.dataframe(df.head(10), use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Column Mapping ────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">② Column Mapping</div>', unsafe_allow_html=True)

    all_cols = list(df.columns)
    num_cols = list(df.select_dtypes(include=[np.number]).columns)

    col1, col2, col3 = st.columns(3)
    with col1:
        subject_col = st.selectbox("👤 Subject ID Column", all_cols,
                                    index=0, key="subj")
    with col2:
        between_col = st.selectbox("👥 Between-Subjects Factor (Group)", all_cols,
                                    index=min(1, len(all_cols)-1), key="between")
    with col3:
        within_cols = st.multiselect("🕐 Within-Subjects Columns (Time Points)",
                                      [c for c in num_cols if c not in [subject_col, between_col]],
                                      default=[c for c in num_cols if c not in [subject_col, between_col]][:3],
                                      key="within")

    # Options
    col4, col5, col6 = st.columns(3)
    with col4:
        dv_name  = st.text_input("📏 Dependent Variable Label", value="Score", key="dv")
    with col5:
        posthoc_method = st.selectbox("🔬 Post-Hoc Correction", ['bonf','tukey','holm','fdr_bh'],
                                       format_func=lambda x: {'bonf':'Bonferroni','tukey':'Tukey HSD',
                                                               'holm':'Holm','fdr_bh':'FDR (BH)'}[x],
                                       key="ph_method")
    with col6:
        alpha_level = st.number_input("α Level", value=0.05, min_value=0.001,
                                       max_value=0.20, step=0.01, key="alpha")

    run_btn = st.button("▶  Run Mixed-Design ANOVA", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if not run_btn:
        return

    # ── Validation ────────────────────────────────────────────────────────────
    if len(within_cols) < 2:
        st.error("❌ Please select at least 2 time-point columns for the within-subjects factor.")
        return

    required = [subject_col, between_col] + within_cols
    missing_cols = [c for c in required if c not in df.columns]
    if missing_cols:
        st.error(f"❌ Columns not found: {missing_cols}")
        return

    # Check numeric
    for c in within_cols:
        if not pd.api.types.is_numeric_dtype(df[c]):
            st.error(f"❌ Column '{c}' is not numeric. Please check your data.")
            return

    # Drop missing
    df_clean = df[required].dropna()
    n_dropped = len(df) - len(df_clean)
    if n_dropped > 0:
        st.markdown(f'<div class="warn-box">⚠️ {n_dropped} rows with missing values were removed. N = {len(df_clean)}</div>',
                    unsafe_allow_html=True)

    if len(df_clean) < 6:
        st.error("❌ Too few complete observations. Need at least 6.")
        return

    # ── Reshape to Long ───────────────────────────────────────────────────────
    within_col = 'Time'
    df_long = df_clean.melt(id_vars=[subject_col, between_col],
                             value_vars=within_cols,
                             var_name=within_col, value_name=dv_name)
    df_long[within_col] = pd.Categorical(df_long[within_col], categories=within_cols, ordered=True)

    n_total  = df_clean[subject_col].nunique()
    n_groups = df_long[between_col].nunique()
    n_times  = len(within_cols)

    # ── Quick Stats ───────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card"><div class="metric-val">{n_total}</div><div class="metric-label">Participants</div></div>
        <div class="metric-card"><div class="metric-val">{n_groups}</div><div class="metric-label">Groups</div></div>
        <div class="metric-card"><div class="metric-val">{n_times}</div><div class="metric-label">Time Points</div></div>
        <div class="metric-card"><div class="metric-val">{n_total*n_times}</div><div class="metric-label">Total Observations</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════════
    # COMPUTE ALL STATISTICS
    # ════════════════════════════════════════════════════════════════════════
    with st.spinner("⚙️  Running statistical analysis…"):
        try:
            desc     = compute_descriptives(df_long, between_col, within_col, dv_name)
            norm_df  = test_normality(df_long, between_col, within_col, dv_name)
            lev_df   = test_levene(df_long, between_col, within_col, dv_name)
            aov      = run_mixed_anova(df_long, subject_col, between_col, within_col, dv_name)
            sph_res  = run_mauchly(df_long, subject_col, within_col, dv_name) if n_times > 2 else None
            interp   = generate_interpretation(aov, desc, between_col, within_col, n_total)

            # Check if interaction is significant
            aov_src = [str(r).strip() for r in aov.get('Source', aov.get('source', [])).tolist()]
            inter_row = aov[aov['Source'].astype(str).str.contains(r'\*|inter', case=False, na=False)]
            int_p = inter_row['p_unc'].values[0] if not inter_row.empty else 1.0
            run_ph = int_p < alpha_level

            posthoc_df = run_posthoc(df_long, between_col, within_col,
                                      dv_name, subject_col, posthoc_method) if run_ph else pd.DataFrame()

        except Exception as e:
            st.error(f"❌ Analysis failed: {e}")
            st.exception(e)
            return

    st.success("✅ Analysis complete!")

    # ════════════════════════════════════════════════════════════════════════
    # TABS
    # ════════════════════════════════════════════════════════════════════════
    tabs = st.tabs([
        "📋 Descriptives",
        "🔍 Assumptions",
        "📊 ANOVA Results",
        "🔬 Post-Hoc",
        "📈 Visualization",
        "📝 Interpretation",
        "⬇️ Export"
    ])

    # ─── Tab 1: Descriptives ─────────────────────────────────────────────────
    with tabs[0]:
        st.markdown('<div class="section-title">Descriptive Statistics by Group × Time</div>',
                    unsafe_allow_html=True)
        disp = desc.copy()
        disp.columns = [c.replace(between_col, 'Group').replace(within_col, 'Time') for c in disp.columns]
        for col in ['Mean','SD','SE','Min','Max','Median','95% CI Lower','95% CI Upper']:
            if col in disp.columns:
                disp[col] = disp[col].round(3)
        st.dataframe(disp, use_container_width=True)

        # Marginal means
        st.markdown('<div class="section-title" style="margin-top:1.5rem;">Marginal Means</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**By Group**")
            mg = df_long.groupby(between_col)[dv_name].agg(['mean','std','count']).round(3)
            mg.columns = ['Mean','SD','N']
            st.dataframe(mg, use_container_width=True)
        with c2:
            st.markdown("**By Time**")
            mt = df_long.groupby(within_col)[dv_name].agg(['mean','std','count']).round(3)
            mt.columns = ['Mean','SD','N']
            st.dataframe(mt, use_container_width=True)

    # ─── Tab 2: Assumptions ──────────────────────────────────────────────────
    with tabs[1]:
        st.markdown('<div class="section-title">A. Normality Tests</div>', unsafe_allow_html=True)

        # Recommendation note
        sw_n  = (norm_df['N'] <= 50).sum()
        ks_n  = (norm_df['N'] > 50).sum()
        st.markdown(f"""
        <div class="interp-box">
        <b>Automatic Test Selection:</b><br>
        • <b>Shapiro-Wilk</b> (n ≤ 50): Applied to <b>{sw_n}</b> cells — preferred for small samples (highest power).<br>
        • <b>Kolmogorov-Smirnov</b> (n > 50): Applied to <b>{ks_n}</b> cells — preferred for large samples.<br>
        H₀: The data are normally distributed. Reject H₀ if <i>p</i> &lt; {alpha_level}.
        </div>
        """, unsafe_allow_html=True)

        norm_show = norm_df.copy()
        norm_show['Statistic'] = norm_show['Statistic'].round(4)
        norm_show['p-value']   = norm_show['p-value'].apply(lambda x: fmt_p_val(x) if not pd.isna(x) else '—')
        norm_show['Result']    = norm_show['Normal'].map({True:'✅ Normal', False:'⚠️ Non-Normal', None:'—'})
        st.dataframe(norm_show[['Group','Time','N','Test','Statistic','p-value','Result','Recommendation']],
                     use_container_width=True)

        all_normal = norm_df['Normal'].dropna().all()
        any_nonnormal = (~norm_df['Normal'].dropna()).any()
        if all_normal:
            st.markdown('<div class="ok-box">✅ All cells satisfy normality assumption (p > .05).</div>',
                        unsafe_allow_html=True)
        elif any_nonnormal:
            st.markdown('<div class="warn-box">⚠️ Some cells show non-normality. '
                        'ANOVA is generally robust to moderate violations (especially n > 15 per cell). '
                        'Consider non-parametric alternatives (e.g., Friedman + Kruskal-Wallis) for severe violations.</div>',
                        unsafe_allow_html=True)

        st.markdown('<div class="section-title" style="margin-top:1.5rem;">B. Homogeneity of Variance (Levene\'s Test)</div>',
                    unsafe_allow_html=True)
        st.markdown(f"""
        <div class="interp-box">
        Tests whether variances are equal across groups at each time point.<br>
        H₀: Variances are equal. Reject H₀ if <i>p</i> &lt; {alpha_level}.
        </div>
        """, unsafe_allow_html=True)

        if not lev_df.empty:
            lev_show = lev_df.copy()
            lev_show['Levene F'] = lev_show['Levene F'].round(3)
            lev_show['p-value']  = lev_show['p-value'].apply(fmt_p_val)
            lev_show['Result']   = lev_show['Equal Variances'].map({True:'✅ Equal', False:'⚠️ Unequal'})
            st.dataframe(lev_show, use_container_width=True)

            if lev_df['Equal Variances'].all():
                st.markdown('<div class="ok-box">✅ Homogeneity of variance is satisfied at all time points.</div>',
                            unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ Unequal variances detected at one or more time points. '
                            'Interpret F-statistics with caution; consider Welch corrections.</div>',
                            unsafe_allow_html=True)

        if n_times > 2:
            st.markdown('<div class="section-title" style="margin-top:1.5rem;">C. Sphericity (Mauchly\'s Test)</div>',
                        unsafe_allow_html=True)
            st.markdown("""
            <div class="interp-box">
            Sphericity requires equal variances of the differences between all pairs of time points.<br>
            H₀: Sphericity is met. Reject H₀ if <i>p</i> &lt; .05.<br>
            If violated, use <b>Greenhouse-Geisser (ε &lt; .75)</b> or <b>Huynh-Feldt (ε ≥ .75)</b> correction.
            </div>
            """, unsafe_allow_html=True)

            if sph_res is not None:
                try:
                    sph_stat = sph_res[1] if isinstance(sph_res, tuple) else None
                    gg_eps   = sph_res[2] if isinstance(sph_res, tuple) else None
                    hf_eps   = sph_res[3] if isinstance(sph_res, tuple) else None
                    sph_p    = sph_res[4] if isinstance(sph_res, tuple) else None

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Mauchly's W", fmt(sph_stat, 4) if sph_stat else "—")
                    with col2:
                        st.metric("p-value", fmt_p_val(sph_p) if sph_p else "—")
                    with col3:
                        st.metric("GG ε", fmt(gg_eps, 4) if gg_eps else "—")

                    if sph_p is not None:
                        if sph_p < .05:
                            rec = "Greenhouse-Geisser" if (gg_eps or 1) < .75 else "Huynh-Feldt"
                            st.markdown(f'<div class="warn-box">⚠️ Sphericity violated (p = {fmt_p_val(sph_p)}). '
                                        f'<b>{rec} correction applied</b> (ε = {fmt(gg_eps,3)}).</div>',
                                        unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="ok-box">✅ Sphericity assumption satisfied (p > .05). '
                                        'No correction needed.</div>', unsafe_allow_html=True)
                except Exception:
                    st.info("Mauchly's test result format could not be parsed. "
                            "Greenhouse-Geisser correction is applied by default in pingouin.")
        else:
            st.info("ℹ️ Sphericity test not applicable for 2 time points.")

    # ─── Tab 3: ANOVA Results ─────────────────────────────────────────────────
    with tabs[2]:
        st.markdown('<div class="section-title">Mixed-Design ANOVA Summary Table</div>',
                    unsafe_allow_html=True)
        st.markdown("""
        <div class="interp-box">
        <b>Algorithm:</b> pingouin <code>mixed_anova()</code> · <b>Sum of Squares:</b> Type III (SPSS-compatible) · 
        <b>Sphericity correction:</b> Greenhouse-Geisser (applied when ε &lt; 1.0)
        </div>
        """, unsafe_allow_html=True)

        st.markdown(render_anova_table(aov), unsafe_allow_html=True)

        # Effect size legend
        st.markdown("""
        <div style="margin-top:1rem;padding:0.75rem 1rem;background:#161b22;border:1px solid #30363d;
                    border-radius:8px;font-size:0.8rem;color:#8b949e;font-family:'IBM Plex Mono',monospace;">
        <b style="color:#e6edf3;">Partial η² benchmarks (Cohen, 1988):</b> &nbsp;
        <span style="color:#3fb950">Small: ≥ .01</span> &nbsp;|&nbsp;
        <span style="color:#e3b341">Medium: ≥ .06</span> &nbsp;|&nbsp;
        <span style="color:#f78166">Large: ≥ .14</span>
        </div>
        """, unsafe_allow_html=True)

        # Full pingouin output
        with st.expander("🔍 Full pingouin Output (Raw)"):
            st.dataframe(aov, use_container_width=True)

    # ─── Tab 4: Post-Hoc ─────────────────────────────────────────────────────
    with tabs[3]:
        if not run_ph:
            st.markdown(f"""
            <div class="interp-box">
            ℹ️ Post-hoc tests were <b>not conducted</b> because the Group × Time interaction was 
            not statistically significant (p = {fmt_p_val(int_p)}, α = {alpha_level}).<br><br>
            Post-hoc tests are typically conducted only when the interaction or a main effect 
            of interest is significant.
            </div>
            """, unsafe_allow_html=True)
        elif posthoc_df.empty:
            st.warning("Post-hoc comparisons could not be computed.")
        else:
            method_name = {'bonf':'Bonferroni','tukey':'Tukey HSD',
                           'holm':'Holm','fdr_bh':'FDR (Benjamini-Hochberg)'}[posthoc_method]
            st.markdown(f'<div class="section-title">Pairwise Comparisons ({method_name} Corrected)</div>',
                        unsafe_allow_html=True)

            # Between-group comparisons
            ph_between = posthoc_df[posthoc_df['Comparison Type']=='Between-Groups'] if 'Comparison Type' in posthoc_df.columns else pd.DataFrame()
            ph_within  = posthoc_df[posthoc_df['Comparison Type']=='Within-Time'] if 'Comparison Type' in posthoc_df.columns else pd.DataFrame()

            if not ph_between.empty:
                st.markdown("**Between-Group Comparisons at Each Time Point**")
                show_cols = [c for c in ['Time Point','Group A','Group B','Mean Diff','SE','t','df','p (unadj)','p (adj)','Cohen d','Sig'] if c in ph_between.columns]
                ph_b_show = ph_between[show_cols].copy()
                for col in ['Mean Diff','SE','t','Cohen d']:
                    if col in ph_b_show.columns:
                        ph_b_show[col] = ph_b_show[col].round(3)
                for col in ['p (unadj)','p (adj)']:
                    if col in ph_b_show.columns:
                        ph_b_show[col] = ph_b_show[col].apply(lambda x: fmt_p_val(x) if not pd.isna(x) else '—')
                st.dataframe(ph_b_show, use_container_width=True)

            if not ph_within.empty:
                st.markdown("**Within-Subject Time Comparisons per Group**")
                show_cols = [c for c in ['Group','Time A','Time B','Mean Diff','SE','t','df','p (unadj)','p (adj)','Cohen d','Sig'] if c in ph_within.columns]
                ph_w_show = ph_within[show_cols].copy()
                for col in ['Mean Diff','SE','t','Cohen d']:
                    if col in ph_w_show.columns:
                        ph_w_show[col] = ph_w_show[col].round(3)
                for col in ['p (unadj)','p (adj)']:
                    if col in ph_w_show.columns:
                        ph_w_show[col] = ph_w_show[col].apply(lambda x: fmt_p_val(x) if not pd.isna(x) else '—')
                st.dataframe(ph_w_show, use_container_width=True)

            st.markdown("""
            <div style="font-size:0.78rem;color:#8b949e;margin-top:0.5rem;">
            *** p &lt; .001 · ** p &lt; .01 · * p &lt; .05 · † p &lt; .10 · ns p ≥ .10
            </div>
            """, unsafe_allow_html=True)

    # ─── Tab 5: Visualization ─────────────────────────────────────────────────
    with tabs[4]:
        st.markdown('<div class="section-title">Profile Plot (Interaction Plot)</div>',
                    unsafe_allow_html=True)
        fig = make_interaction_plot(desc, between_col, within_col,
                                     title=f"Profile Plot: {dv_name} by {between_col} × Time")
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

        # Distribution plots
        st.markdown('<div class="section-title" style="margin-top:1.5rem;">Distribution by Group × Time</div>',
                    unsafe_allow_html=True)
        plt.style.use('dark_background')
        n_t = len(within_cols)
        fig2, axes2 = plt.subplots(1, n_t, figsize=(5*n_t, 4), sharey=True)
        fig2.patch.set_facecolor('#0d1117')
        if n_t == 1:
            axes2 = [axes2]
        palette = ['#58a6ff','#3fb950','#bc8cff','#e3b341','#f78166']
        groups  = df_long[between_col].unique()

        for ax, t in zip(axes2, within_cols):
            sub = df_long[df_long[within_col]==t]
            ax.set_facecolor('#161b22')
            ax.spines['bottom'].set_color('#30363d')
            ax.spines['left'].set_color('#30363d')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.tick_params(colors='#8b949e', labelsize=8)
            for i, g in enumerate(groups):
                d = sub[sub[between_col]==g][dv_name].dropna()
                ax.hist(d, bins=12, alpha=0.55, color=palette[i % len(palette)],
                        label=str(g), edgecolor='#0d1117', linewidth=0.4)
            ax.set_title(str(t), color='#e6edf3', fontsize=10)
            ax.set_xlabel(dv_name, color='#8b949e', fontsize=9)
            ax.grid(axis='y', color='#21262d', alpha=0.6, linewidth=0.5)
        axes2[0].legend(title=between_col, fontsize=8, facecolor='#21262d',
                         edgecolor='#30363d', labelcolor='#c9d1d9')
        plt.tight_layout()
        st.pyplot(fig2, use_container_width=True)
        plt.close(fig2)

    # ─── Tab 6: Interpretation ────────────────────────────────────────────────
    with tabs[5]:
        st.markdown('<div class="section-title">Automated Statistical Interpretation</div>',
                    unsafe_allow_html=True)
        st.markdown(f'<div class="interp-box">{interp.replace(chr(10), "<br>")}</div>',
                    unsafe_allow_html=True)

        # APA-style report
        st.markdown('<div class="section-title" style="margin-top:1.5rem;">APA-Style Results Write-Up</div>',
                    unsafe_allow_html=True)
        src_names = aov['Source'].astype(str).tolist() if 'Source' in aov.columns else []
        inter_src = next((s for s in src_names if '*' in s or 'inter' in s.lower()), None)
        if inter_src:
            ir = aov[aov['Source']==inter_src].iloc[0]
            F_i  = ir.get('F', np.nan)
            df1  = ir.get('ddof1', ir.get('DF', np.nan))
            df2  = ir.get('ddof2', np.nan)
            p_i  = ir.get('p_unc', ir.get('p-unc', ir.get('p-GG-corr', np.nan)))
            eta_i= ir.get('np2', np.nan)
            apa  = (f"A {n_groups}-group (between) × {n_times}-time (within) mixed-design ANOVA "
                    f"was conducted with {n_total} participants. "
                    f"The Group × Time interaction was "
                    f"{'significant' if p_i < alpha_level else 'not significant'}, "
                    f"F({fmt(df1,0)}, {fmt(df2,0)}) = {fmt(F_i,2)}, p = {fmt_p_val(p_i)}, "
                    f"η²ₚ = {fmt(eta_i,3)}.")
            st.info(apa)

    # ─── Tab 7: Export ────────────────────────────────────────────────────────
    with tabs[6]:
        st.markdown('<div class="section-title">Download Statistical Report</div>',
                    unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        # CSV Export
        with col1:
            st.markdown("**📄 CSV Export**")
            csv_buf = io.StringIO()
            csv_buf.write("Mixed-Design ANOVA Report\n\n")
            csv_buf.write("DESCRIPTIVE STATISTICS\n")
            desc.to_csv(csv_buf, index=False)
            csv_buf.write("\n\nNORMALITY TESTS\n")
            norm_df.to_csv(csv_buf, index=False)
            csv_buf.write("\n\nLEVENE'S TEST\n")
            lev_df.to_csv(csv_buf, index=False)
            csv_buf.write("\n\nANOVA RESULTS\n")
            aov.to_csv(csv_buf, index=False)
            if not posthoc_df.empty:
                csv_buf.write("\n\nPOST-HOC TESTS\n")
                posthoc_df.to_csv(csv_buf, index=False)
            csv_buf.write("\n\nINTERPRETATION\n")
            csv_buf.write(interp.replace('**',''))
            csv_data = csv_buf.getvalue().encode()

            st.download_button(
                label="⬇️ Download CSV Report",
                data=csv_data,
                file_name="mixed_anova_report.csv",
                mime="text/csv",
                use_container_width=True
            )

        # PDF Export
        with col2:
            st.markdown("**📑 PDF Export**")
            try:
                pdf_data = generate_pdf_report(
                    desc, norm_df, lev_df, aov, interp, between_col, within_col, dv_name
                )
                st.download_button(
                    label="⬇️ Download PDF Report",
                    data=pdf_data,
                    file_name="mixed_anova_report.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"PDF generation failed: {e}")

        # Long-format data
        st.markdown("**📊 Processed Long-Format Data**")
        st.download_button(
            label="⬇️ Download Long-Format CSV",
            data=df_long.to_csv(index=False).encode(),
            file_name="long_format_data.csv",
            mime="text/csv",
            use_container_width=True
        )


# ════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    main()
