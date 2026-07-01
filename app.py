"""
Internal Pay Equity & Rewards Intelligence Report
--------------------------------------------------
Upload a compensation workbook (.xlsx) to generate a fully interactive
internal equity analysis across pay equity, compa-ratio positioning,
performance-reward alignment, range outliers, retention risk, and
individual employee benchmarking.

Run locally:    streamlit run app.py
Deploy:         push to GitHub → Streamlit Community Cloud
"""

import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Rewards Equity Dashboard",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# DESIGN TOKENS
# ============================================================
INDIGO   = "#3B3F8C"
TEAL     = "#0EA5A8"
EMERALD  = "#1FA67A"
AMBER    = "#E0A82E"
ROSE     = "#D6536D"
VIOLET   = "#7B5EA7"
NAVY     = "#10243E"
SLATE    = "#5B6B7B"
CYAN     = "#0284C7"
ORANGE   = "#EA580C"

BG_PAGE  = "#EEF2F7"
BG_CARD  = "#FFFFFF"
BG_SECT  = "#F8FAFD"
BORDER   = "#DDE3EC"

PALETTE  = [TEAL, INDIGO, AMBER, ROSE, EMERALD, VIOLET, CYAN, ORANGE]

TAB_COLORS = {
    "Overview":         ("#EFF6FF", "#1D4ED8", "#DBEAFE"),
    "Pay Equity":       ("#FFF0F3", "#BE123C", "#FFE4E6"),
    "Pay & Performance":("#F0FDF4", "#15803D", "#DCFCE7"),
    "Range Flags":      ("#FFF7ED", "#C2410C", "#FFEDD5"),
    "Retention & Risk": ("#FDF4FF", "#7E22CE", "#F3E8FF"),
    "Critical Talent":  ("#ECFDF5", "#065F46", "#D1FAE5"),
    "Employee Lookup":  ("#F0F9FF", "#0369A1", "#E0F2FE"),
    "Year-over-Year":   ("#FAFAF9", "#44403C", "#F5F5F4"),
    "Data Explorer":    ("#F7FEE7", "#365314", "#ECFCCB"),
}

TAB_NAMES = list(TAB_COLORS.keys())

PLOTLY_TEMPLATE = "simple_white"
QUARTILE_ORDER  = ["Q1", "Q2", "Q3", "Q4"]
PERF_ORDER      = ["NI", "SP", "HP", "EP"]
RISK_ORDER      = ["Low", "Medium", "High"]
CRIT_ORDER      = ["Low", "Medium", "High", "Critical"]

# ============================================================
# GLOBAL CSS
# ============================================================
st.markdown(f"""
<style>
/* ── Page background ── */
.stApp {{ background-color: {BG_PAGE}; }}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg,{NAVY} 0%,#1a3558 100%);
    border-right: none;
}}
section[data-testid="stSidebar"] * {{ color: #E2E8F0 !important; }}
section[data-testid="stSidebar"] h3 {{ color: #94B8D0 !important; font-size:0.78rem; letter-spacing:1.5px; text-transform:uppercase; }}
section[data-testid="stSidebar"] .stFileUploader label {{ color:#CBD5E1 !important; }}
section[data-testid="stSidebar"] .stCaption {{ color:#64748B !important; }}

/* ── KPI cards ── */
.kpi-card {{
    border-radius:16px; padding:20px 22px; color:white;
    box-shadow:0 4px 18px rgba(16,36,62,0.15);
    min-height:130px; display:flex; flex-direction:column; justify-content:space-between;
}}
.kpi-label {{ font-size:0.72rem; font-weight:700; letter-spacing:1px; text-transform:uppercase; opacity:0.88; }}
.kpi-value {{ font-size:1.75rem; font-weight:900; line-height:1.1; margin-top:6px; }}
.kpi-sub   {{ font-size:0.72rem; opacity:0.80; margin-top:4px; }}

/* ── Section card wrapper ── */
.section-card {{
    background:{BG_CARD}; border-radius:16px; padding:24px 26px;
    border:1px solid {BORDER}; margin-bottom:18px;
    box-shadow:0 1px 6px rgba(16,36,62,0.06);
}}

/* ── Tab strip ── */
div[data-baseweb="tab-list"] {{
    gap:6px; background:{BG_PAGE}; padding:8px 4px 0 4px;
    border-bottom:2px solid {BORDER};
    flex-wrap: nowrap;
    overflow-x: auto;
}}
button[data-baseweb="tab"] {{
    border-radius:10px 10px 0 0 !important;
    padding:8px 14px !important;
    font-size:0.78rem !important;
    font-weight:700 !important;
    border:1.5px solid transparent !important;
    border-bottom:none !important;
    transition: all 0.15s ease;
    white-space: nowrap !important;
    min-width: fit-content !important;
}}
button[data-baseweb="tab"]:nth-child(1)  {{ background:#DBEAFE !important; color:#1D4ED8 !important; border-color:#BFDBFE !important; }}
button[data-baseweb="tab"]:nth-child(2)  {{ background:#FFE4E6 !important; color:#BE123C !important; border-color:#FECDD3 !important; }}
button[data-baseweb="tab"]:nth-child(3)  {{ background:#DCFCE7 !important; color:#15803D !important; border-color:#BBF7D0 !important; }}
button[data-baseweb="tab"]:nth-child(4)  {{ background:#FFEDD5 !important; color:#C2410C !important; border-color:#FED7AA !important; }}
button[data-baseweb="tab"]:nth-child(5)  {{ background:#F3E8FF !important; color:#7E22CE !important; border-color:#E9D5FF !important; }}
button[data-baseweb="tab"]:nth-child(6)  {{ background:#D1FAE5 !important; color:#065F46 !important; border-color:#A7F3D0 !important; }}
button[data-baseweb="tab"]:nth-child(7)  {{ background:#E0F2FE !important; color:#0369A1 !important; border-color:#BAE6FD !important; }}
button[data-baseweb="tab"]:nth-child(8)  {{ background:#F5F5F4 !important; color:#44403C !important; border-color:#E7E5E4 !important; }}
button[data-baseweb="tab"]:nth-child(9)  {{ background:#ECFCCB !important; color:#365314 !important; border-color:#D9F99D !important; }}
button[data-baseweb="tab"][aria-selected="true"] {{
    opacity:1 !important; box-shadow:0 -2px 8px rgba(16,36,62,0.12) !important;
    transform:translateY(-2px); filter: brightness(0.93);
}}

/* ── Section subheaders ── */
h2, h3 {{ color:{NAVY} !important; font-family:'Segoe UI',sans-serif; }}
h2 {{ font-size:1.15rem !important; font-weight:800 !important; border-left:4px solid {TEAL}; padding-left:10px; margin-bottom:14px !important; }}

/* ── Badge ── */
.badge {{
    display:inline-block; padding:2px 10px; border-radius:99px;
    font-size:0.72rem; font-weight:700; letter-spacing:0.5px;
}}
.badge-red   {{ background:#FFE4E6; color:#BE123C; }}
.badge-amber {{ background:#FFEDD5; color:#C2410C; }}
.badge-green {{ background:#DCFCE7; color:#15803D; }}
.badge-blue  {{ background:#DBEAFE; color:#1D4ED8; }}
.badge-violet{{ background:#F3E8FF; color:#7E22CE; }}

/* ── Note boxes ── */
.note-box {{
    background:#EFF6FF; border-radius:10px; padding:12px 16px;
    font-size:0.88rem; color:{NAVY}; margin:6px 0 16px 0;
    border-left:4px solid {TEAL};
}}
.warn-box {{
    background:#FFF7ED; border-radius:10px; padding:12px 16px;
    font-size:0.88rem; color:{NAVY}; margin:6px 0 16px 0;
    border-left:4px solid {AMBER};
}}
.danger-box {{
    background:#FFF0F3; border-radius:10px; padding:12px 16px;
    font-size:0.88rem; color:{NAVY}; margin:6px 0 16px 0;
    border-left:4px solid {ROSE};
}}

/* ── Profile table ── */
.profile-table {{ width:100%; border-collapse:collapse; font-size:0.88rem; }}
.profile-table tr:nth-child(even) {{ background:#F8FAFD; }}
.profile-table td {{ padding:7px 10px; border-bottom:1px solid #EEF2F7; color:{NAVY}; }}
.profile-table td:first-child {{ color:{SLATE}; font-weight:600; width:45%; }}

/* ── Section label ── */
.section-note {{ color:{SLATE}; font-size:0.88rem; margin-top:-10px; margin-bottom:10px; }}

/* ── Divider ── */
hr {{ border:none; border-top:1px solid {BORDER}; margin:20px 0; }}

/* ── Dataframe tweaks ── */
.stDataFrame {{ border-radius:12px; overflow:hidden; border:1px solid {BORDER}; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# HELPERS
# ============================================================
def kpi_card(label, value, color, sub="", icon=""):
    return f"""<div class="kpi-card" style="background:{color};">
        <div class="kpi-label">{icon} {label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""

def render_kpi_row(cards):
    cols = st.columns(len(cards))
    for col, item in zip(cols, cards):
        label, value, color, sub = item[0], item[1], item[2], item[3]
        icon = item[4] if len(item) > 4 else ""
        with col:
            st.markdown(kpi_card(label, value, color, sub, icon), unsafe_allow_html=True)

def section(title, color=TEAL):
    st.markdown(f"<h2 style='border-left-color:{color};'>{title}</h2>", unsafe_allow_html=True)

def fmt_pct(x, decimals=1):
    return "—" if pd.isna(x) else f"{x*100:.{decimals}f}%"

def fmt_num(x):
    return "—" if pd.isna(x) else f"{x:,.0f}"

def fmt_currency(x):
    return "—" if pd.isna(x) else f"Rs. {x:,.0f}"

def safe_cols(cols, frame):
    return [c for c in cols if c in frame.columns]

def badge(text, style="blue"):
    return f'<span class="badge badge-{style}">{text}</span>'

def perf_badge(p):
    m = {"EP": "green", "HP": "blue", "SP": "amber", "NI": "red"}
    return badge(p, m.get(str(p).upper(), "blue"))

def risk_badge(r):
    m = {"High": "red", "Medium": "amber", "Low": "green"}
    return badge(r, m.get(str(r).title(), "blue"))

def compa_badge(c):
    if pd.isna(c): return badge("—", "blue")
    if c < 0.80: return badge(fmt_pct(c), "red")
    if c > 1.20: return badge(fmt_pct(c), "amber")
    return badge(fmt_pct(c), "green")

# ============================================================
# COLUMN MAPPING
# ============================================================
COLUMN_CANDIDATES = {
    "employee_id":        ["Employee ID", "EmpID", "Employee Id"],
    "employee_name":      ["Employee Name", "Name"],
    "department":         ["Department"],
    "designation":        ["Designation", "Title"],
    "grade":              ["Existing Grade", "Grade"],
    "new_grade":          ["New Grade"],
    "promotion_flag":     ["Promotion"],
    "gender":             ["Gender"],
    "doj":                ["DOJ", "Date of Joining"],
    "tenure":             ["Tenure (2025)", "Tenure", "Tenure (Years)"],
    "years_in_role":      ["Years in Current Role"],
    "job_family":         ["Job Family"],
    "career_level":       ["Career Level"],
    "position_criticality":["Position Criticality"],
    "supervisory":        ["Supervisory Responsibility (Y/N)", "Supervisory Responsibility"],
    "basic_salary":       ["Basic Salary (Rs.)", "Basic Salary"],
    "gross_salary":       ["Gross Salary (Table) (Rs.)", "Gross Salary (Rs.)", "Gross Salary"],
    "salary_min":         ["Salary Range Minimum ", "Salary Range Minimum"],
    "salary_mid":         ["Salary Range Midpoint"],
    "salary_max":         ["Salary Range Maximum"],
    "compa_ratio":        ["Compa Ratio"],
    "range_penetration":  ["Range Penetration"],
    "quartile":           ["Quartile in Range"],
    "performance":        ["Performance Rating"],
    "last_performance":   ["Last Performance Rating"],
    "increment_pct":      ["Increment %", "Increment Percentage"],
    "num_promotions":     ["Number of Promotions"],
    "last_promotion_date":["Last Promotion Date"],
    "qualification":      ["Highest Qualification"],
    "critical_skills":    ["Critical Skills"],
    "skill_rating":       ["Skill Rating"],
    "retention_risk":     ["Retention Risk"],
    "disciplinary":       ["Disciplinary Action in Last Year"],
    "employment_type":    ["Employment Type"],
    "employment_status":  ["Employment Status"],
}

def _norm(s):
    return str(s).strip().lower()

def map_columns(df):
    lookup = {_norm(c): c for c in df.columns}
    found, missing = {}, []
    for std_name, candidates in COLUMN_CANDIDATES.items():
        match = next((lookup[_norm(c)] for c in candidates if _norm(c) in lookup), None)
        if match:
            found[std_name] = match
        else:
            missing.append(std_name)
    return found, missing

@st.cache_data(show_spinner=False)
def load_workbook(file_bytes):
    df_raw = pd.read_excel(io.BytesIO(file_bytes))
    colmap, missing = map_columns(df_raw)
    df = pd.DataFrame()
    for std_name, original in colmap.items():
        df[std_name] = df_raw[original]

    for c in ["compa_ratio", "range_penetration", "increment_pct"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            if df[c].max(skipna=True) is not None and df[c].max(skipna=True) > 3:
                df[c] = df[c] / 100.0

    for c in ["gross_salary", "basic_salary", "salary_min", "salary_mid", "salary_max",
              "tenure", "years_in_role", "num_promotions", "skill_rating"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "gender" in df.columns:
        df["gender"] = df["gender"].astype(str).str.strip().str.title()
    if "quartile" in df.columns:
        df["quartile"] = df["quartile"].astype(str).str.strip().str.upper()
    if "performance" in df.columns:
        df["performance"] = df["performance"].astype(str).str.strip().str.upper()
    if "last_performance" in df.columns:
        df["last_performance"] = df["last_performance"].astype(str).str.strip().str.upper()
    if "retention_risk" in df.columns:
        df["retention_risk"] = df["retention_risk"].astype(str).str.strip().str.title()

    if "compa_ratio" in df.columns:
        df["compa_flag"] = np.select(
            [df["compa_ratio"] < 0.80, df["compa_ratio"] > 1.20],
            ["Below 80%", "Above 120%"],
            default="Within Range",
        )
        if "grade" in df.columns:
            grp = df.groupby("grade")["compa_ratio"]
            df["compa_zscore"] = (df["compa_ratio"] - grp.transform("mean")) / grp.transform("std")

    return df, df_raw, missing

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.markdown("### 📂 DATA SOURCE")
uploaded = st.sidebar.file_uploader("Upload compensation workbook (.xlsx)", type=["xlsx"])
prior_uploaded = st.sidebar.file_uploader("Prior period (optional — for trend)", type=["xlsx"], key="prior")

if not uploaded:
    st.markdown(f"""
    <div style="text-align:center; padding:80px 20px 40px 20px;">
        <div style="font-size:4rem; margin-bottom:16px;">🧭</div>
        <h1 style="color:{NAVY}; font-size:1.9rem; margin-bottom:6px;">
            Internal Pay Equity & Rewards Intelligence
        </h1>
        <div style="color:{SLATE}; font-size:0.9rem; margin-bottom:24px;">
            <b>Developed by Hamza Aslam</b> · Compensation &amp; Benefits Professional
        </div>
        <div style="background:white; border-radius:16px; padding:28px 32px; max-width:620px;
                    margin:0 auto; border:1px solid {BORDER}; box-shadow:0 4px 20px rgba(16,36,62,0.08);">
            <p style="color:{SLATE}; font-size:1rem; margin:0;">
                Upload a compensation workbook using the sidebar to instantly generate
                an interactive pay equity analysis — from workforce-level trends down
                to individual employee benchmarking.
            </p>
        </div>
        <div style="margin-top:28px; display:flex; justify-content:center; gap:12px; flex-wrap:wrap;">
            {"".join(f'<span style="background:{c[2]};color:{c[1]};border-radius:8px;padding:5px 14px;font-size:0.75rem;font-weight:700;">{n}</span>'
                     for n, c in TAB_COLORS.items())}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

df, df_raw, missing_cols = load_workbook(uploaded.getvalue())

if missing_cols:
    st.sidebar.warning("Columns not mapped (views skipped): " + ", ".join(missing_cols[:8]))

df_prior = None
if prior_uploaded:
    df_prior, _, _ = load_workbook(prior_uploaded.getvalue())

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔽 FILTERS")

def multiselect_filter(label, col, icon=""):
    if col not in df.columns:
        return None
    opts = sorted(df[col].dropna().unique().tolist())
    return st.sidebar.multiselect(f"{icon} {label}", opts, default=opts)

f_dept   = multiselect_filter("Department",   "department",   "🏢")
f_jf     = multiselect_filter("Job Family",   "job_family",   "💼")
f_grade  = multiselect_filter("Grade",        "grade",        "📊")
f_gender = multiselect_filter("Gender",       "gender",       "👤")
f_career = multiselect_filter("Career Level", "career_level", "🎯")

fdf = df.copy()
for col, sel in [("department", f_dept), ("job_family", f_jf), ("grade", f_grade),
                  ("gender", f_gender), ("career_level", f_career)]:
    if sel is not None:
        fdf = fdf[fdf[col].isin(sel)]

st.sidebar.markdown("---")
st.sidebar.caption(f"🔎 Showing **{len(fdf):,}** of **{len(df):,}** employees")

# ============================================================
# PAGE HEADER
# ============================================================
st.markdown(f"""
<div style="background:linear-gradient(135deg,{NAVY} 0%,{INDIGO} 100%);
            border-radius:18px; padding:24px 30px; margin-bottom:22px;
            display:flex; align-items:center; justify-content:space-between;
            box-shadow:0 4px 20px rgba(16,36,62,0.2);">
    <div>
        <h1 style="color:white; margin:0; font-size:1.55rem; font-weight:900;">
            🧭 Internal Pay Equity & Rewards Intelligence
        </h1>
        <div style="color:#94B8D0; font-size:0.85rem; margin-top:4px;">
            Compensation structure · Pay equity · Reward effectiveness
        </div>
    </div>
    <div style="text-align:right;">
        <div style="color:#CBD5E1; font-size:0.78rem; font-weight:700;">DEVELOPED BY</div>
        <div style="color:white; font-size:0.95rem; font-weight:800;">Hamza Aslam</div>
        <div style="color:#94B8D0; font-size:0.72rem;">Compensation & Benefits Professional</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# KPI ROWS
# ============================================================
headcount = len(fdf)
avg_compa = fdf["compa_ratio"].mean() if "compa_ratio" in fdf.columns else np.nan
median_pen= fdf["range_penetration"].median() if "range_penetration" in fdf.columns else np.nan

gap_pct = np.nan
if {"gender", "gross_salary"}.issubset(fdf.columns):
    m_sal = fdf.loc[fdf["gender"]=="Male", "gross_salary"].mean()
    f_sal = fdf.loc[fdf["gender"]=="Female", "gross_salary"].mean()
    if not pd.isna(m_sal) and m_sal != 0:
        gap_pct = (m_sal - f_sal) / m_sal

pct_high_risk = (fdf["retention_risk"]=="High").mean() if "retention_risk" in fdf.columns else np.nan
avg_incr      = fdf["increment_pct"].mean() if "increment_pct" in fdf.columns else np.nan
n_below = int((fdf["compa_flag"]=="Below 80%").sum()) if "compa_flag" in fdf.columns else 0
n_above = int((fdf["compa_flag"]=="Above 120%").sum()) if "compa_flag" in fdf.columns else 0

render_kpi_row([
    ("HEADCOUNT IN VIEW",        f"{headcount:,}",   NAVY,         "",                       "👥"),
    ("AVG COMPA-RATIO",          fmt_pct(avg_compa), TEAL,         "Target: 100%",           "⚖️"),
    ("MEDIAN RANGE PENETRATION", fmt_pct(median_pen),EMERALD,      "Position in pay range",  "📍"),
    ("GENDER PAY GAP",           fmt_pct(gap_pct),   ROSE,         "Unadjusted male vs female","♀"),
])
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
render_kpi_row([
    ("AVG INCREMENT %",      fmt_pct(avg_incr),         VIOLET,    "This cycle",             "📈"),
    ("HIGH RETENTION RISK",  fmt_pct(pct_high_risk),    AMBER,     "Share of workforce",     "⚠️"),
    ("BELOW 80% COMPA",      f"{n_below:,}",            "#B23A48", fmt_pct(n_below/headcount) if headcount else "", "🔻"),
    ("ABOVE 120% COMPA",     f"{n_above:,}",            "#B45309", fmt_pct(n_above/headcount) if headcount else "", "🔺"),
])

st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# ============================================================
# TABS
# ============================================================
tabs = st.tabs(TAB_NAMES)
(tab_overview, tab_payequity, tab_perf, tab_flags,
 tab_retention, tab_talent, tab_micro, tab_trend, tab_data) = tabs

# ─────────────────────────────────────────────────────────────
# 1. OVERVIEW
# ─────────────────────────────────────────────────────────────
with tab_overview:
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section("Headcount & Average Pay by Department", TEAL)
    if {"department","gross_salary"}.issubset(fdf.columns):
        agg = (fdf.groupby("department")
               .agg(Headcount=("department","size"), Avg_Gross=("gross_salary","mean"))
               .reset_index().sort_values("Headcount", ascending=False))
        fig = go.Figure()
        fig.add_bar(x=agg["department"], y=agg["Headcount"], name="Headcount",
                    marker_color=TEAL, marker_line_width=0)
        fig.add_trace(go.Scatter(x=agg["department"], y=agg["Avg_Gross"],
                                  name="Avg Gross Salary", mode="lines+markers",
                                  marker=dict(color=ROSE, size=9), line=dict(color=ROSE, width=2),
                                  yaxis="y2"))
        fig.update_layout(template=PLOTLY_TEMPLATE, height=400,
                           paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                           yaxis=dict(title="Headcount", gridcolor="#F0F4FA"),
                           yaxis2=dict(title="Avg Gross Salary (Rs.)", overlaying="y",
                                       side="right", showgrid=False),
                           legend=dict(orientation="h", y=1.12), margin=dict(t=20,b=10))
        st.plotly_chart(fig, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Gender Composition", ROSE)
        if "gender" in fdf.columns:
            gc = fdf["gender"].value_counts().reset_index()
            gc.columns = ["Gender","Count"]
            fig = px.pie(gc, names="Gender", values="Count", hole=0.58,
                         color_discrete_sequence=[INDIGO, TEAL, AMBER, ROSE])
            fig.update_traces(textfont_size=13, pull=[0.03]*len(gc))
            fig.update_layout(template=PLOTLY_TEMPLATE, height=360, margin=dict(t=10,b=10),
                               paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Performance Distribution", EMERALD)
        if "performance" in fdf.columns:
            pr = fdf["performance"].value_counts().reindex(PERF_ORDER).dropna().reset_index()
            pr.columns = ["Rating","Count"]
            PERF_COLORS = {"NI": ROSE, "SP": AMBER, "HP": TEAL, "EP": EMERALD}
            fig = px.bar(pr, x="Rating", y="Count", color="Rating",
                         category_orders={"Rating": PERF_ORDER},
                         color_discrete_map=PERF_COLORS)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=360, showlegend=False,
                               margin=dict(t=10,b=10), paper_bgcolor="rgba(0,0,0,0)",
                               plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Headcount by Career Level", VIOLET)
        if "career_level" in fdf.columns:
            cl = fdf["career_level"].value_counts().reset_index()
            cl.columns = ["Career Level","Count"]
            fig = px.bar(cl, x="Count", y="Career Level", orientation="h", color="Count",
                         color_continuous_scale=[TEAL, INDIGO])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                               coloraxis_showscale=False, margin=dict(t=10,b=10),
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                               yaxis=dict(categoryorder="total ascending"))
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c4:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Compa-Ratio Distribution", AMBER)
        if "compa_ratio" in fdf.columns:
            fig = px.histogram(fdf, x="compa_ratio", nbins=45,
                                color_discrete_sequence=[TEAL], opacity=0.85)
            fig.add_vline(x=1.0, line_dash="dash", line_color=NAVY, line_width=2,
                          annotation_text="Midpoint", annotation_position="top right")
            fig.add_vline(x=0.80, line_dash="dot", line_color=ROSE, line_width=1.5,
                          annotation_text="80%", annotation_position="top")
            fig.add_vline(x=1.20, line_dash="dot", line_color=AMBER, line_width=1.5,
                          annotation_text="120%", annotation_position="top")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380,
                               xaxis_tickformat=".0%", xaxis_title="Compa-Ratio",
                               yaxis_title="Employees", margin=dict(t=10,b=10),
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 2. PAY EQUITY
# ─────────────────────────────────────────────────────────────
with tab_payequity:
    if {"gender","gross_salary","grade"}.issubset(fdf.columns):
        overall = fdf.groupby("gender")["gross_salary"].mean()
        unadj_gap, adj_gap = np.nan, np.nan
        if "Male" in overall.index and "Female" in overall.index:
            if not pd.isna(overall["Male"]) and overall["Male"] != 0:
                unadj_gap = (overall["Male"] - overall["Female"]) / overall["Male"]
                ggrp = fdf.groupby(["grade","gender"])["gross_salary"].mean().unstack()
                if {"Male","Female"}.issubset(ggrp.columns):
                    gcounts = fdf.groupby("grade").size()
                    ggrp["gap"] = (ggrp["Male"] - ggrp["Female"]) / ggrp["Male"]
                    valid = ggrp["gap"].dropna()
                    adj_gap = np.average(valid, weights=gcounts.reindex(valid.index))

        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Gender Pay Gap Summary", ROSE)
        render_kpi_row([
            ("UNADJUSTED PAY GAP",     fmt_pct(unadj_gap),                              ROSE,   "Raw average male vs female", "♀"),
            ("GRADE-ADJUSTED PAY GAP", fmt_pct(adj_gap) if not pd.isna(adj_gap) else "—", INDIGO,"Same-grade comparison",      "⚖️"),
            ("MALE HEADCOUNT",         f"{(fdf['gender']=='Male').sum():,}",              NAVY,   "", "👨"),
            ("FEMALE HEADCOUNT",       f"{(fdf['gender']=='Female').sum():,}",            VIOLET, "", "👩"),
        ])
        st.markdown("""<div class='note-box' style='margin-top:14px;'>
            The grade-adjusted figure compares men and women within the same grade, removing
            the effect of role-mix. When the adjusted gap is substantially smaller than the
            unadjusted gap, the raw difference is primarily driven by representation (fewer
            women in senior grades) rather than unequal pay for the same role.
        </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Avg Compa-Ratio by Grade & Gender", TEAL)
        if {"grade","gender","compa_ratio"}.issubset(fdf.columns):
            cg = fdf.groupby(["grade","gender"])["compa_ratio"].mean().reset_index()
            fig = px.bar(cg, x="grade", y="compa_ratio", color="gender", barmode="group",
                         color_discrete_sequence=[INDIGO, TEAL, AMBER])
            fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5,
                          annotation_text="Midpoint")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                               legend=dict(orientation="h", y=1.1), margin=dict(t=20),
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Pay Quartile Mix by Gender", VIOLET)
        if {"gender","quartile"}.issubset(fdf.columns):
            qc = fdf.groupby(["gender","quartile"]).size().reset_index(name="Count")
            qc["quartile"] = pd.Categorical(qc["quartile"], categories=QUARTILE_ORDER, ordered=True)
            fig = px.bar(qc.sort_values("quartile"), x="gender", y="Count", color="quartile",
                         category_orders={"quartile": QUARTILE_ORDER},
                         color_discrete_sequence=[ROSE, AMBER, TEAL, INDIGO], barmode="stack")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, legend_title="Quartile",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Compa-Ratio by Tenure Band", EMERALD)
        if {"tenure","compa_ratio"}.issubset(fdf.columns):
            bins   = [-1, 2, 5, 10, 100]
            labels = ["0–2 yrs", "3–5 yrs", "6–10 yrs", "10+ yrs"]
            tmp = fdf.copy()
            tmp["_tb"] = pd.cut(tmp["tenure"], bins=bins, labels=labels)
            tb = tmp.groupby("_tb", observed=True)["compa_ratio"].mean().reindex(labels).reset_index()
            fig = px.bar(tb, x="_tb", y="compa_ratio", color="_tb",
                         color_discrete_sequence=[EMERALD, TEAL, INDIGO, VIOLET])
            fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis_tickformat=".0%",
                               showlegend=False, xaxis_title="Tenure Band",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c4:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Compa-Ratio Spread by Department (Box Plot)", AMBER)
        if {"department","compa_ratio"}.issubset(fdf.columns):
            fig = px.box(fdf, x="department", y="compa_ratio", color="department",
                         color_discrete_sequence=PALETTE, notched=True)
            fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                               yaxis_tickformat=".0%", paper_bgcolor="rgba(0,0,0,0)",
                               plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 3. PAY & PERFORMANCE
# ─────────────────────────────────────────────────────────────
with tab_perf:
    if {"performance","compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Avg Compa-Ratio by Performance Rating", EMERALD)
            pr_compa = (fdf.groupby("performance")["compa_ratio"].mean()
                        .reindex(PERF_ORDER).dropna().reset_index())
            fig = px.bar(pr_compa, x="performance", y="compa_ratio",
                         category_orders={"performance": PERF_ORDER},
                         color="performance",
                         color_discrete_map={"NI": ROSE,"SP": AMBER,"HP": TEAL,"EP": EMERALD})
            fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis_tickformat=".0%",
                               showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                               plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Avg Increment % by Performance Rating", VIOLET)
            if "increment_pct" in fdf.columns:
                pr_incr = (fdf.groupby("performance")["increment_pct"].mean()
                           .reindex(PERF_ORDER).dropna().reset_index())
                fig = px.bar(pr_incr, x="performance", y="increment_pct",
                             category_orders={"performance": PERF_ORDER},
                             color="performance",
                             color_discrete_map={"NI": ROSE,"SP": AMBER,"HP": TEAL,"EP": EMERALD})
                fig.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis_tickformat=".0%",
                                   showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                                   plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("""<div class='note-box'>
            Both charts should step upward left-to-right (NI → SP → HP → EP) for pay
            to be genuinely differentiating performance. A flat or inverted pattern indicates
            that performance ratings are not translating into pay positioning.
        </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Increment % vs. Position in Range (Scatter)", TEAL)
        if "increment_pct" in fdf.columns:
            hover_c = safe_cols(["employee_name","department","designation","grade"], fdf)
            fig = px.scatter(fdf, x="compa_ratio", y="increment_pct", color="performance",
                              category_orders={"performance": PERF_ORDER},
                              color_discrete_map={"NI": ROSE,"SP": AMBER,"HP": TEAL,"EP": EMERALD},
                              hover_data=hover_c, opacity=0.75, size_max=8)
            fig.add_vline(x=1.0, line_dash="dot", line_color=NAVY, line_width=1.5)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=480,
                               xaxis_tickformat=".0%", yaxis_tickformat=".0%",
                               xaxis_title="Compa-Ratio (position in grade range)",
                               yaxis_title="Increment %",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 4. RANGE FLAGS
# ─────────────────────────────────────────────────────────────
with tab_flags:
    if "compa_flag" in fdf.columns:
        n_below = int((fdf["compa_flag"]=="Below 80%").sum())
        n_above = int((fdf["compa_flag"]=="Above 120%").sum())
        n_within= int((fdf["compa_flag"]=="Within Range").sum())

        render_kpi_row([
            ("BELOW 80% COMPA",  f"{n_below:,}",             "#B23A48", fmt_pct(n_below/headcount) if headcount else "", "🔻"),
            ("ABOVE 120% COMPA", f"{n_above:,}",             "#B45309", fmt_pct(n_above/headcount) if headcount else "", "🔺"),
            ("WITHIN RANGE",     f"{n_within:,}",            EMERALD,   fmt_pct(n_within/headcount)if headcount else "", "✅"),
            ("TOTAL FLAGGED",    f"{n_below+n_above:,}",     INDIGO,    fmt_pct((n_below+n_above)/headcount) if headcount else "", "🚩"),
        ])

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Flags by Department", ROSE)
            fc = (fdf[fdf["compa_flag"]!="Within Range"]
                  .groupby(["department","compa_flag"]).size().reset_index(name="Count"))
            if not fc.empty and "department" in fdf.columns:
                fig = px.bar(fc, x="department", y="Count", color="compa_flag", barmode="group",
                             color_discrete_map={"Below 80%":"#B23A48","Above 120%":"#B45309"})
                fig.update_layout(template=PLOTLY_TEMPLATE, height=380, legend_title="Flag",
                                   paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Flags by Grade", AMBER)
            if "grade" in fdf.columns:
                fg = (fdf[fdf["compa_flag"]!="Within Range"]
                      .groupby(["grade","compa_flag"]).size().reset_index(name="Count"))
                if not fg.empty:
                    fig = px.bar(fg, x="grade", y="Count", color="compa_flag", barmode="group",
                                 color_discrete_map={"Below 80%":"#B23A48","Above 120%":"#B45309"})
                    fig.update_layout(template=PLOTLY_TEMPLATE, height=380, legend_title="Flag",
                                       paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Compa-Ratio Map per Grade (Flagged Highlighted)", INDIGO)
        if "grade" in fdf.columns:
            fig = px.strip(fdf, x="grade", y="compa_ratio", color="compa_flag",
                            color_discrete_map={"Below 80%":"#B23A48","Above 120%":"#B45309",
                                                "Within Range":"#C7D0DA"},
                            stripmode="overlay",
                            hover_data=safe_cols(["employee_name","designation","department"], fdf))
            fig.add_hline(y=0.80, line_dash="dot", line_color="#B23A48", line_width=1.5,
                          annotation_text="80% floor")
            fig.add_hline(y=1.20, line_dash="dot", line_color="#B45309", line_width=1.5,
                          annotation_text="120% ceiling")
            fig.add_hline(y=1.00, line_dash="dash", line_color=NAVY, line_width=1.5,
                          annotation_text="Midpoint")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=480, yaxis_tickformat=".0%",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if "compa_zscore" in fdf.columns:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Statistical Outliers (±2 Standard Deviations within Grade)", VIOLET)
            outliers = fdf[fdf["compa_zscore"].abs() > 2].copy()
            st.caption(f"{len(outliers):,} employees are 2+ SDs from their grade's average compa-ratio.")
            show = safe_cols(["employee_id","employee_name","department","designation","grade",
                               "gender","gross_salary","compa_ratio","compa_zscore"], outliers)
            if show:
                st.dataframe(
                    outliers[show].sort_values("compa_zscore")
                    .style.background_gradient(subset=["compa_zscore"], cmap="RdYlGn")
                    .format({c: "{:.0f}" for c in ["gross_salary"] if c in show}),
                    use_container_width=True, height=280
                )
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Full Flagged Employee List", ROSE)
        flagged = fdf[fdf["compa_flag"]!="Within Range"].copy()
        show2 = safe_cols(["employee_id","employee_name","department","designation","grade",
                            "gender","gross_salary","salary_mid","compa_ratio","compa_flag"], flagged)
        if show2:
            st.dataframe(
                flagged[show2].sort_values("compa_ratio")
                .style.map(
                    lambda v: f"color:{'#B23A48' if v=='Below 80%' else ('#B45309' if v=='Above 120%' else 'inherit')}",
                    subset=["compa_flag"] if "compa_flag" in show2 else []
                ),
                use_container_width=True, height=380
            )
            csv = flagged[show2].to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download flagged list (CSV)", csv,
                               "compa_ratio_flags.csv", "text/csv")
        st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 5. RETENTION & RISK
# ─────────────────────────────────────────────────────────────
with tab_retention:
    if {"retention_risk","compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Headcount by Retention Risk", ROSE)
            rr = (fdf["retention_risk"].value_counts()
                  .reindex(RISK_ORDER).dropna().reset_index())
            rr.columns = ["Risk","Count"]
            fig = px.bar(rr, x="Risk", y="Count", color="Risk",
                         category_orders={"Risk": RISK_ORDER},
                         color_discrete_map={"Low": EMERALD,"Medium": AMBER,"High": ROSE})
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Avg Compa-Ratio by Retention Risk", AMBER)
            rc = (fdf.groupby("retention_risk")["compa_ratio"].mean()
                  .reindex(RISK_ORDER).dropna().reset_index())
            fig = px.bar(rc, x="retention_risk", y="compa_ratio", color="retention_risk",
                         category_orders={"retention_risk": RISK_ORDER},
                         color_discrete_map={"Low": EMERALD,"Medium": AMBER,"High": ROSE})
            fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                               yaxis_tickformat=".0%",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("""<div class='danger-box'>
            🚩 If <b>High</b> retention risk employees also show a <b>lower</b> average
            compa-ratio than Low risk employees, pay positioning is contributing to attrition
            exposure — a compelling case for targeted retention budget allocation.
        </div>""", unsafe_allow_html=True)

        if "performance" in fdf.columns:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("High-Performer Flight Risk Heat Map", INDIGO)
            hp = fdf[fdf["performance"].isin(["HP","EP"])].copy()
            fig = px.density_heatmap(hp, x="performance", y="retention_risk",
                                      category_orders={"performance":["HP","EP"],
                                                       "retention_risk": RISK_ORDER},
                                      color_continuous_scale=["#EAF7F6", TEAL, INDIGO],
                                      text_auto=True)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=420,
                               paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        if "department" in fdf.columns:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Retention Risk × Compa-Ratio by Department (Bubble)", EMERALD)
            bub = (fdf.groupby(["department","retention_risk"])
                   .agg(Count=("compa_ratio","size"), Avg_Compa=("compa_ratio","mean"))
                   .reset_index())
            fig = px.scatter(bub, x="department", y="Avg_Compa",
                              size="Count", color="retention_risk",
                              category_orders={"retention_risk": RISK_ORDER},
                              color_discrete_map={"Low": EMERALD,"Medium": AMBER,"High": ROSE},
                              size_max=55)
            fig.add_hline(y=1.0, line_dash="dot", line_color=NAVY)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=460, yaxis_tickformat=".0%",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 6. CRITICAL TALENT
# ─────────────────────────────────────────────────────────────
with tab_talent:
    if {"position_criticality","compa_ratio"}.issubset(fdf.columns):
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Avg Compa-Ratio by Position Criticality", EMERALD)
        order = [c for c in CRIT_ORDER if c in fdf["position_criticality"].unique()]
        crit = fdf.groupby("position_criticality")["compa_ratio"].mean().reindex(order).reset_index()
        fig = px.bar(crit, x="position_criticality", y="compa_ratio", color="position_criticality",
                     color_discrete_sequence=[TEAL, AMBER, ROSE, INDIGO])
        fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, line_width=1.5)
        fig.update_layout(template=PLOTLY_TEMPLATE, height=380, yaxis_tickformat=".0%",
                           showlegend=False,
                           paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("""<div class='warn-box'>
            Critical positions sitting below Medium/Low criticality roles in compa-ratio
            is a market risk — the org may be under-investing in the positions it can least
            afford to lose.
        </div>""", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if {"skill_rating","compa_ratio"}.issubset(fdf.columns):
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Skill Rating vs. Pay Positioning", VIOLET)
        hover_c = safe_cols(["employee_name","department","designation","grade","retention_risk"], fdf)
        fig = px.scatter(fdf, x="skill_rating", y="compa_ratio",
                          color="job_family" if "job_family" in fdf.columns else "performance",
                          color_discrete_sequence=PALETTE, opacity=0.75, hover_data=hover_c)
        fig.add_hline(y=1.0, line_dash="dot", line_color=NAVY)
        fig.update_layout(template=PLOTLY_TEMPLATE, height=480, yaxis_tickformat=".0%",
                           xaxis_title="Skill Rating (1–5)", yaxis_title="Compa-Ratio",
                           paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    if {"num_promotions","tenure"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Promotion Velocity (Avg Tenure per Promo Level)", TEAL)
            pv = fdf.groupby("num_promotions")["tenure"].mean().reset_index()
            fig = px.line(pv, x="num_promotions", y="tenure", markers=True,
                          color_discrete_sequence=[INDIGO])
            fig.update_traces(line=dict(width=3), marker=dict(size=9))
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380,
                               xaxis_title="Number of Promotions", yaxis_title="Avg Tenure (yrs)",
                               paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown("<div class='section-card'>", unsafe_allow_html=True)
            section("Avg Promotions by Gender", ROSE)
            if "gender" in fdf.columns:
                pg = fdf.groupby("gender")["num_promotions"].mean().reset_index()
                fig = px.bar(pg, x="gender", y="num_promotions", color="gender",
                             color_discrete_sequence=[INDIGO, TEAL, AMBER])
                fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                                   paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 7. EMPLOYEE LOOKUP
# ─────────────────────────────────────────────────────────────
with tab_micro:
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#0369A1,#0EA5A8);
                border-radius:14px; padding:18px 24px; margin-bottom:18px; color:white;">
        <div style="font-size:1.1rem; font-weight:800;">🔍 Employee Reward Profile</div>
        <div style="font-size:0.85rem; opacity:0.85; margin-top:4px;">
            Search by Employee ID or Name to generate a complete reward profile,
            pay positioning analysis, and peer comparison.
        </div>
    </div>
    """, unsafe_allow_html=True)

    id_col   = "employee_id"   if "employee_id"   in fdf.columns else None
    name_col = "employee_name" if "employee_name" in fdf.columns else None
    label_col= name_col or id_col

    if not label_col:
        st.info("Employee Name or Employee ID column not found in workbook.")
    else:
        s1, s2 = st.columns([1, 3])
        with s1:
            search = st.text_input("🔎 Search by name or ID", placeholder="Type to search…")
        all_opts = fdf[label_col].dropna().astype(str).unique().tolist()
        filtered_opts = [o for o in all_opts if search.lower() in o.lower()] if search else all_opts
        with s2:
            selected = st.selectbox("Select employee", sorted(filtered_opts) if filtered_opts else ["—"])

        if selected and selected != "—":
            emp_rows = fdf[fdf[label_col].astype(str) == selected]
            if emp_rows.empty:
                st.warning("Employee not found.")
            else:
                emp = emp_rows.iloc[0]

                # peer group = same grade + department
                peer_mask = pd.Series(True, index=fdf.index)
                if "grade"      in fdf.columns: peer_mask &= fdf["grade"]      == emp.get("grade")
                if "department" in fdf.columns: peer_mask &= fdf["department"] == emp.get("department")
                peers = fdf[peer_mask]
                n_peers = len(peers)

                compa  = emp.get("compa_ratio",     np.nan)
                pen    = emp.get("range_penetration",np.nan)
                gross  = emp.get("gross_salary",    np.nan)
                sal_mid= emp.get("salary_mid",      np.nan)
                sal_min= emp.get("salary_min",      np.nan)
                sal_max= emp.get("salary_max",      np.nan)
                incr   = emp.get("increment_pct",   np.nan)

                peer_pctile = np.nan
                if "compa_ratio" in peers.columns and n_peers > 1:
                    peer_pctile = (peers["compa_ratio"] < compa).mean()

                dept_avg_compa = np.nan
                if "compa_ratio" in fdf.columns and "department" in fdf.columns:
                    dept_avg_compa = fdf.loc[fdf["department"]==emp.get("department"), "compa_ratio"].mean()

                # ── KPI STRIP ──
                render_kpi_row([
                    ("GROSS SALARY",      fmt_currency(gross),  NAVY,    f"Grade: {emp.get('grade','—')}",       "💰"),
                    ("COMPA-RATIO",        fmt_pct(compa),       TEAL,    "vs grade midpoint",                    "⚖️"),
                    ("RANGE PENETRATION",  fmt_pct(pen),         EMERALD, "position in pay band",                 "📍"),
                    ("PEER PERCENTILE",
                     fmt_pct(peer_pctile) if not pd.isna(peer_pctile) else "—",
                     VIOLET,  f"within {n_peers} grade/dept peers",            "🏆"),
                ])
                st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

                # ── PROFILE + GAUGES ──
                pc1, pc2 = st.columns([1.1, 1.9])

                with pc1:
                    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                    section("Complete Reward Profile", INDIGO)
                    perf_val = emp.get("performance","—")
                    risk_val = emp.get("retention_risk","—")
                    compa_flag_val = emp.get("compa_flag","—")

                    profile_rows = [
                        ("Employee ID",       str(emp.get("employee_id","—"))),
                        ("Department",        str(emp.get("department","—"))),
                        ("Designation",       str(emp.get("designation","—"))),
                        ("Grade",             str(emp.get("grade","—"))),
                        ("Job Family",        str(emp.get("job_family","—"))),
                        ("Career Level",      str(emp.get("career_level","—"))),
                        ("Gender",            str(emp.get("gender","—"))),
                        ("Tenure (yrs)",      str(emp.get("tenure","—"))),
                        ("Years in Role",     str(emp.get("years_in_role","—"))),
                        ("Performance",       perf_val),
                        ("Last Performance",  str(emp.get("last_performance","—"))),
                        ("Retention Risk",    risk_val),
                        ("Position Criticality", str(emp.get("position_criticality","—"))),
                        ("Supervisory",       str(emp.get("supervisory","—"))),
                        ("Skill Rating",      str(emp.get("skill_rating","—"))),
                        ("Qualification",     str(emp.get("qualification","—"))),
                        ("Num Promotions",    str(emp.get("num_promotions","—"))),
                        ("Increment %",       fmt_pct(incr)),
                        ("Basic Salary",      fmt_currency(emp.get("basic_salary",np.nan))),
                        ("Gross Salary",      fmt_currency(gross)),
                        ("Salary Min",        fmt_currency(sal_min)),
                        ("Salary Midpoint",   fmt_currency(sal_mid)),
                        ("Salary Max",        fmt_currency(sal_max)),
                        ("Compa-Ratio",       fmt_pct(compa)),
                        ("Range Penetration", fmt_pct(pen)),
                        ("Quartile",          str(emp.get("quartile","—"))),
                        ("Compa Flag",        compa_flag_val),
                        ("Disciplinary Action", str(emp.get("disciplinary","—"))),
                    ]

                    badge_fields = {
                        "Performance": perf_badge(perf_val),
                        "Retention Risk": risk_badge(risk_val),
                        "Compa-Ratio": compa_badge(compa),
                        "Compa Flag": (badge(compa_flag_val, "red") if compa_flag_val=="Below 80%"
                                       else badge(compa_flag_val,"amber") if compa_flag_val=="Above 120%"
                                       else badge(compa_flag_val,"green")),
                    }

                    rows_html = ""
                    for field, val in profile_rows:
                        display = badge_fields.get(field, val)
                        rows_html += f"<tr><td>{field}</td><td>{display}</td></tr>"

                    st.markdown(f"""
                    <table class="profile-table">
                        <tbody>{rows_html}</tbody>
                    </table>
                    """, unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

                with pc2:
                    # ── GAUGE: Compa-Ratio ──
                    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                    section("Pay Position in Band — Compa-Ratio Gauge", TEAL)
                    if not pd.isna(compa):
                        needle = min(max(compa, 0.5), 1.5)
                        fig = go.Figure(go.Indicator(
                            mode="gauge+number+delta",
                            value=compa * 100,
                            delta={"reference": 100, "valueformat": ".1f",
                                   "suffix": "%", "relative": False},
                            number={"suffix": "%", "valueformat": ".1f",
                                    "font": {"size": 36, "color": NAVY}},
                            gauge={
                                "axis": {"range": [60, 140], "ticksuffix": "%",
                                         "tickcolor": SLATE},
                                "bar": {"color": TEAL, "thickness": 0.3},
                                "bgcolor": "white",
                                "steps": [
                                    {"range": [60,  80], "color": "#FECDD3"},
                                    {"range": [80, 100], "color": "#FEF3C7"},
                                    {"range": [100,120], "color": "#D1FAE5"},
                                    {"range": [120,140], "color": "#FFEDD5"},
                                ],
                                "threshold": {"line": {"color": NAVY, "width": 4},
                                              "thickness": 0.85, "value": 100},
                            },
                            title={"text": "Compa-Ratio (100% = Midpoint)", "font": {"color": SLATE, "size": 13}},
                        ))
                        fig.update_layout(height=280, margin=dict(t=30,b=10,l=20,r=20),
                                          paper_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig, use_container_width=True)

                    # ── SALARY RANGE BAR ──
                    if not any(pd.isna(x) for x in [sal_min, sal_mid, sal_max, gross]):
                        pct_in_range = (gross - sal_min) / (sal_max - sal_min) if sal_max != sal_min else 0.5
                        pct_in_range = min(max(pct_in_range, 0), 1)

                        st.markdown(f"""
                        <div style="margin-top:8px;">
                            <div style="display:flex; justify-content:space-between; font-size:0.75rem;
                                        color:{SLATE}; font-weight:700; margin-bottom:4px;">
                                <span>Min<br/>{fmt_currency(sal_min)}</span>
                                <span style="text-align:center;">Midpoint<br/>{fmt_currency(sal_mid)}</span>
                                <span style="text-align:right;">Max<br/>{fmt_currency(sal_max)}</span>
                            </div>
                            <div style="position:relative; background:#EEF2F7; border-radius:8px; height:22px;">
                                <div style="position:absolute; left:50%; top:0; width:2px; height:100%;
                                            background:{SLATE}; opacity:0.4;"></div>
                                <div style="position:absolute; left:0; top:0;
                                            width:{pct_in_range*100:.1f}%; height:100%;
                                            background:linear-gradient(90deg,{TEAL},{EMERALD});
                                            border-radius:8px; opacity:0.85;"></div>
                                <div style="position:absolute; left:{pct_in_range*100:.1f}%; top:-4px;
                                            transform:translateX(-50%);
                                            width:12px; height:30px; background:{NAVY};
                                            border-radius:3px;"></div>
                            </div>
                            <div style="text-align:center; margin-top:6px; font-size:0.82rem; color:{NAVY}; font-weight:700;">
                                Gross Salary: {fmt_currency(gross)} &nbsp;|&nbsp; {fmt_pct(pct_in_range)} in range
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

                    st.markdown("</div>", unsafe_allow_html=True)

                    # ── PEER GROUP CHART ──
                    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                    section(f"Peer Group Distribution  ({n_peers} peers — same grade & dept)", VIOLET)
                    if "compa_ratio" in peers.columns and n_peers > 2:
                        fig = px.histogram(peers, x="compa_ratio", nbins=20,
                                            color_discrete_sequence=[TEAL], opacity=0.8)
                        fig.add_vline(x=compa, line_color=ROSE, line_width=3,
                                      annotation_text=f"← This employee ({fmt_pct(compa)})",
                                      annotation_position="top right",
                                      annotation_font_color=ROSE)
                        fig.add_vline(x=1.0, line_dash="dot", line_color=NAVY, line_width=1.5,
                                      annotation_text="Midpoint")
                        fig.update_layout(template=PLOTLY_TEMPLATE, height=280,
                                           xaxis_tickformat=".0%", xaxis_title="Compa-Ratio",
                                           yaxis_title="Peers", margin=dict(t=20,b=10),
                                           paper_bgcolor="rgba(0,0,0,0)",
                                           plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Not enough peers for distribution chart.")
                    st.markdown("</div>", unsafe_allow_html=True)

                # ── PEER COMPARISON TABLE ──
                st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                section("Peer Comparison — Same Grade & Department", EMERALD)
                compare_opts = peers[label_col].astype(str).tolist()
                # 1. Create a filtered list of options that excludes the selected individual
filtered_options = [o for o in compare_opts if o != selected]

# 2. Safely pass it into the widget
compare_sel = st.multiselect(
    "Add specific peers to compare side-by-side",
    options=filtered_options,
    default=filtered_options[:min(5, len(filtered_options))]
)
                if compare_sel:
                compare_df = fdf[fdf[label_col].astype(str).isin([selected] + compare_sel)].copy()
                show_c = safe_cols([label_col,"designation","grade","department",
                                        "gross_salary","salary_mid","compa_ratio",
                                        "range_penetration","quartile","performance",
                                        "increment_pct","tenure","retention_risk",
                                        "skill_rating"], compare_df)
                    if show_c:
                        styled = compare_df[show_c].style
                        if "compa_ratio" in show_c:
                            styled = styled.background_gradient(subset=["compa_ratio"],
                                                                 cmap="RdYlGn", vmin=0.7, vmax=1.3)
                        if "increment_pct" in show_c:
                            styled = styled.background_gradient(subset=["increment_pct"],
                                                                 cmap="YlGn")
                        st.dataframe(styled, use_container_width=True, hide_index=True)
                st.markdown("</div>", unsafe_allow_html=True)

                # ── DEPT CONTEXT ──
                st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                section("Department Pay Context", AMBER)
                if "department" in fdf.columns and "compa_ratio" in fdf.columns:
                    dept_df = fdf[fdf["department"] == emp.get("department")].copy()
                    ca, cb = st.columns(2)
                    with ca:
                        d_stats = dept_df["compa_ratio"].describe().reset_index()
                        d_stats.columns = ["Statistic","Compa-Ratio"]
                        d_stats["Compa-Ratio"] = d_stats["Compa-Ratio"].apply(
                            lambda x: fmt_pct(x) if isinstance(x, float) else str(x))
                        st.markdown("**Department Compa-Ratio Statistics**")
                        st.dataframe(d_stats, use_container_width=True, hide_index=True)
                    with cb:
                        if "performance" in dept_df.columns:
                            pc = dept_df["performance"].value_counts().reset_index()
                            pc.columns = ["Rating","Count"]
                            fig = px.pie(pc, names="Rating", values="Count", hole=0.5,
                                         color_discrete_sequence=PALETTE, title="Dept Performance Mix")
                            fig.update_layout(height=280, margin=dict(t=30),
                                               paper_bgcolor="rgba(0,0,0,0)")
                            st.plotly_chart(fig, use_container_width=True)
                st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 8. YEAR-OVER-YEAR
# ─────────────────────────────────────────────────────────────
with tab_trend:
    if df_prior is None:
        st.markdown(f"""
        <div style="background:{BG_CARD}; border-radius:16px; padding:40px;
                    text-align:center; border:2px dashed {BORDER};">
            <div style="font-size:2.5rem;">📁</div>
            <h3 style="color:{SLATE};">Upload a prior period file in the sidebar to unlock year-over-year comparison.</h3>
        </div>
        """, unsafe_allow_html=True)
    else:
        rows_t = []
        for lbl, frame in [("Prior Period", df_prior), ("Current Period", df)]:
            row_t = {"Period": lbl, "Headcount": len(frame)}
            if "compa_ratio"   in frame.columns: row_t["Avg Compa-Ratio"] = frame["compa_ratio"].mean()
            if "increment_pct" in frame.columns: row_t["Avg Increment %"] = frame["increment_pct"].mean()
            if {"gender","gross_salary"}.issubset(frame.columns):
                mm = frame.loc[frame["gender"]=="Male", "gross_salary"].mean()
                ff = frame.loc[frame["gender"]=="Female", "gross_salary"].mean()
                row_t["Gender Pay Gap"] = (mm-ff)/mm if not pd.isna(mm) and mm != 0 else np.nan
            rows_t.append(row_t)
        trend_df = pd.DataFrame(rows_t)

        cards_t = []
        for col_t, color_t, icon_t in [("Avg Compa-Ratio", TEAL, "⚖️"),
                                        ("Gender Pay Gap",  ROSE, "♀"),
                                        ("Avg Increment %", VIOLET, "📈")]:
            if col_t in trend_df.columns:
                curr = trend_df[col_t].iloc[1]
                prev = trend_df[col_t].iloc[0]
                delta = curr - prev
                cards_t.append((col_t.upper(), fmt_pct(curr), color_t,
                                 f"{'+' if delta >= 0 else ''}{fmt_pct(delta)} vs prior", icon_t))
        cards_t.append(("HEADCOUNT", f"{trend_df['Headcount'].iloc[1]:,}", NAVY,
                         f"Prior: {trend_df['Headcount'].iloc[0]:,}", "👥"))
        if cards_t:
            render_kpi_row(cards_t)

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        section("Headcount & Compa-Ratio Trend", TEAL)
        fig = go.Figure()
        fig.add_bar(x=trend_df["Period"], y=trend_df["Headcount"],
                    name="Headcount", marker_color=TEAL, marker_line_width=0)
        if "Avg Compa-Ratio" in trend_df.columns:
            fig.add_trace(go.Scatter(
                x=trend_df["Period"], y=trend_df["Avg Compa-Ratio"],
                name="Avg Compa-Ratio", mode="lines+markers+text",
                marker=dict(color=ROSE, size=12), line=dict(color=ROSE, width=3),
                text=[fmt_pct(v) for v in trend_df["Avg Compa-Ratio"]],
                textposition="top center", yaxis="y2"
            ))
        fig.update_layout(template=PLOTLY_TEMPLATE, height=440,
                           yaxis=dict(title="Headcount", gridcolor="#F0F4FA"),
                           yaxis2=dict(title="Avg Compa-Ratio", overlaying="y",
                                       side="right", tickformat=".0%", showgrid=False),
                           legend=dict(orientation="h", y=1.1),
                           paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# 9. DATA EXPLORER
# ─────────────────────────────────────────────────────────────
with tab_data:
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#365314,#4D7C0F);
                border-radius:14px; padding:18px 24px; margin-bottom:18px; color:white;">
        <div style="font-size:1.1rem; font-weight:800;">🗂️ Interactive Data Explorer</div>
        <div style="font-size:0.85rem; opacity:0.85; margin-top:4px;">
            Filter, sort, and export the full enriched workforce dataset.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # inline quick filters
    xc1, xc2, xc3, xc4 = st.columns(4)
    with xc1:
        compa_filter = st.selectbox("Compa Flag", ["All","Below 80%","Within Range","Above 120%"])
    with xc2:
        perf_filter = st.multiselect("Performance", PERF_ORDER, default=PERF_ORDER) if "performance" in fdf.columns else PERF_ORDER
    with xc3:
        risk_filter = st.multiselect("Retention Risk", RISK_ORDER, default=RISK_ORDER) if "retention_risk" in fdf.columns else RISK_ORDER
    with xc4:
        if "grade" in fdf.columns:
            grade_filter = st.multiselect("Grade (quick)", sorted(fdf["grade"].dropna().unique()), default=sorted(fdf["grade"].dropna().unique()))
        else:
            grade_filter = None

    view = fdf.copy()
    if compa_filter != "All" and "compa_flag" in view.columns:
        view = view[view["compa_flag"] == compa_filter]
    if "performance" in view.columns:
        view = view[view["performance"].isin(perf_filter)]
    if "retention_risk" in view.columns:
        view = view[view["retention_risk"].isin(risk_filter)]
    if grade_filter is not None and "grade" in view.columns:
        view = view[view["grade"].isin(grade_filter)]

    st.caption(f"Showing **{len(view):,}** employees after quick filters")

    # column visibility
    all_cols = view.columns.tolist()
    with st.expander("⚙️ Show / hide columns"):
        show_cols = st.multiselect("Columns to display", all_cols, default=all_cols[:min(18, len(all_cols))])

    if show_cols:
        display_view = view[show_cols].copy()

        # conditional formatting
        style = display_view.style

        numeric_cols = display_view.select_dtypes("number").columns.tolist()
        if "compa_ratio" in numeric_cols:
            style = style.background_gradient(subset=["compa_ratio"], cmap="RdYlGn", vmin=0.7, vmax=1.3)
        if "increment_pct" in numeric_cols:
            style = style.background_gradient(subset=["increment_pct"], cmap="YlGn")
        if "skill_rating" in numeric_cols:
            style = style.background_gradient(subset=["skill_rating"], cmap="Blues")
        if "tenure" in numeric_cols:
            style = style.background_gradient(subset=["tenure"], cmap="Purples")

        fmt_dict = {}
        for c in numeric_cols:
            if c in ["compa_ratio","range_penetration","increment_pct"]:
                fmt_dict[c] = "{:.1%}"
            elif c in ["gross_salary","basic_salary","salary_min","salary_mid","salary_max"]:
                fmt_dict[c] = "{:,.0f}"
        if fmt_dict:
            style = style.format(fmt_dict, na_rep="—")

        st.dataframe(style, use_container_width=True, height=520)

    # downloads
    dc1, dc2 = st.columns(2)
    with dc1:
        csv_all = fdf.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download full dataset (CSV)", csv_all,
                            "equity_full_export.csv", "text/csv")
    with dc2:
        if show_cols:
            csv_view = view[show_cols].to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download current view (CSV)", csv_view,
                                "equity_filtered_view.csv", "text/csv")

    # ── Summary stats ──
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    section("Descriptive Statistics — Numeric Columns", EMERALD)
    num_view = view.select_dtypes("number")
    if not num_view.empty:
        desc = num_view.describe().T
        desc.index.name = "Column"
        st.dataframe(desc.style.background_gradient(cmap="Blues", subset=["mean"]),
                     use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)
