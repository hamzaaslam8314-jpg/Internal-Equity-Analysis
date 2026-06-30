"""
Internal Pay Equity & Rewards Intelligence Report
----------------------------------------
Upload a compensation workbook to generate an interactive internal equity
analysis: pay equity, compa-ratio positioning, pay-for-performance alignment,
range outliers, retention risk, critical talent, and employee-level
benchmarking.

Run locally:    streamlit run app.py
Deploy:         push this repo to GitHub, then deploy via streamlit.io
"""

import io
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# PAGE CONFIG & THEME
# ============================================================
st.set_page_config(
    page_title="Rewards Equity Dashboard",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

INDIGO = "#3B3F8C"
TEAL = "#0EA5A8"
EMERALD = "#1FA67A"
AMBER = "#E0A82E"
ROSE = "#D6536D"
VIOLET = "#7B5EA7"
NAVY = "#10243E"
SLATE = "#5B6B7B"
BG = "#F4F6FA"
CARD_BG = "#FFFFFF"
GRID = "#E7ECF1"

PALETTE = [TEAL, INDIGO, AMBER, ROSE, EMERALD, VIOLET, "#7D8FA8"]
PLOTLY_TEMPLATE = "simple_white"

QUARTILE_ORDER = ["Q1", "Q2", "Q3", "Q4"]
PERF_ORDER = ["NI", "SP", "HP", "EP"]
RISK_ORDER = ["Low", "Medium", "High"]
CRITICALITY_ORDER = ["Low", "Medium", "High", "Critical"]

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {BG}; }}
    h1, h2, h3 {{ color: {NAVY}; font-family: 'Segoe UI', sans-serif; }}
    section[data-testid="stSidebar"] {{ background-color: #FFFFFF; border-right: 1px solid {GRID}; }}
    .section-note {{ color: {SLATE}; font-size: 0.9rem; }}
    .note-box {{
        background: #F0F4FF; border-radius: 8px; padding: 10px 14px;
        font-size: 0.9rem; color: {NAVY}; margin: 6px 0 16px 0; border: 1px solid #DCE3F7;
    }}
    .flag-underpaid {{
        background: #FCEEF0; border-radius: 8px; padding: 10px 14px;
        font-size: 0.9rem; color: {NAVY}; margin: 6px 0 16px 0; border: 1px solid #F3D2D9;
    }}
    .kpi-card {{
        border-radius: 14px; padding: 16px 18px; color: white;
        box-shadow: 0 2px 10px rgba(16,36,62,0.12);
    }}
    .kpi-label {{ font-size: 0.80rem; opacity: 0.92; font-weight: 600; letter-spacing: 0.2px; }}
    .kpi-value {{ font-size: 1.65rem; font-weight: 800; margin-top: 2px; }}
    .kpi-sub {{ font-size: 0.74rem; opacity: 0.85; margin-top: 2px; }}
    div[data-baseweb="tab-list"] {{ gap: 4px; }}
    button[data-baseweb="tab"] {{ font-weight: 600; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def kpi_card(label, value, color, sub=""):
    return f"""
    <div class="kpi-card" style="background:{color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """


def render_kpi_row(cards):
    cols = st.columns(len(cards))
    for col, (label, value, color, sub) in zip(cols, cards):
        with col:
            st.markdown(kpi_card(label, value, color, sub), unsafe_allow_html=True)


# ============================================================
# COLUMN MAPPING (robust to minor header variations year-to-year)
# ============================================================
COLUMN_CANDIDATES = {
    "employee_id": ["Employee ID", "EmpID", "Employee Id"],
    "employee_name": ["Employee Name", "Name"],
    "department": ["Department"],
    "designation": ["Designation", "Title"],
    "grade": ["Existing Grade", "Grade"],
    "new_grade": ["New Grade"],
    "promotion_flag": ["Promotion"],
    "gender": ["Gender"],
    "doj": ["DOJ", "Date of Joining"],
    "tenure": ["Tenure (2025)", "Tenure", "Tenure (Years)"],
    "years_in_role": ["Years in Current Role"],
    "job_family": ["Job Family"],
    "career_level": ["Career Level"],
    "position_criticality": ["Position Criticality"],
    "supervisory": ["Supervisory Responsibility (Y/N)", "Supervisory Responsibility"],
    "basic_salary": ["Basic Salary (Rs.)", "Basic Salary"],
    "gross_salary": ["Gross Salary (Table) (Rs.)", "Gross Salary (Rs.)", "Gross Salary"],
    "salary_min": ["Salary Range Minimum ", "Salary Range Minimum"],
    "salary_mid": ["Salary Range Midpoint"],
    "salary_max": ["Salary Range Maximum"],
    "compa_ratio": ["Compa Ratio"],
    "range_penetration": ["Range Penetration"],
    "quartile": ["Quartile in Range"],
    "performance": ["Performance Rating"],
    "last_performance": ["Last Performance Rating"],
    "increment_pct": ["Increment %", "Increment Percentage"],
    "num_promotions": ["Number of Promotions"],
    "last_promotion_date": ["Last Promotion Date"],
    "qualification": ["Highest Qualification"],
    "critical_skills": ["Critical Skills"],
    "skill_rating": ["Skill Rating"],
    "retention_risk": ["Retention Risk"],
    "disciplinary": ["Disciplinary Action in Last Year"],
    "employment_type": ["Employment Type"],
    "employment_status": ["Employment Status"],
}


def _norm(s):
    return str(s).strip().lower()


def map_columns(df):
    lookup = {_norm(c): c for c in df.columns}
    found, missing = {}, []
    for std_name, candidates in COLUMN_CANDIDATES.items():
        match = None
        for cand in candidates:
            if _norm(cand) in lookup:
                match = lookup[_norm(cand)]
                break
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
            if df[c].max() is not None and df[c].max() > 3:
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


def fmt_pct(x, decimals=1):
    if pd.isna(x):
        return "—"
    return f"{x*100:.{decimals}f}%"


def fmt_currency(x):
    if pd.isna(x):
        return "—"
    return f"Rs. {x:,.0f}"


def safe_cols(cols, frame):
    return [c for c in cols if c in frame.columns]


# ============================================================
# SIDEBAR — UPLOAD & FILTERS
# ============================================================
st.sidebar.markdown("### Data Source")
uploaded = st.sidebar.file_uploader("Upload compensation workbook (.xlsx)", type=["xlsx"])
prior_uploaded = st.sidebar.file_uploader(
    "Optional — prior period file for trend comparison", type=["xlsx"], key="prior"
)

if not uploaded:
    st.markdown(
        f"""
        <div style="text-align:center; padding-top:70px;">
        <h1 style="color:{NAVY};">🧭 Internal Pay Equity & Rewards Intelligence</h1>

        <div style="margin-top:-8px; margin-bottom:18px; color:#6b7280; font-size:16px;">
            <b>Developed by Hamza Aslam</b> | Executive Compensation & Benefits
        </div>

        <p style="color:{SLATE}; font-size:1.05rem; max-width:620px; margin:auto;">
        Upload a compensation workbook in the sidebar to generate pay equity,
        compa-ratio, performance-reward alignment, and retention analysis —
        down to the individual employee level.
        </p>

        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

df, df_raw, missing_cols = load_workbook(uploaded.getvalue())

if missing_cols:
    st.sidebar.warning("Columns not found (related views skipped): " + ", ".join(missing_cols))

df_prior = None
if prior_uploaded:
    df_prior, _, _ = load_workbook(prior_uploaded.getvalue())

st.sidebar.markdown("---")
st.sidebar.markdown("### Filters")


def multiselect_filter(label, col):
    if col not in df.columns:
        return None
    options = sorted(df[col].dropna().unique().tolist())
    return st.sidebar.multiselect(label, options, default=options)


f_dept = multiselect_filter("Department", "department")
f_jobfam = multiselect_filter("Job Family", "job_family")
f_grade = multiselect_filter("Grade", "grade")
f_gender = multiselect_filter("Gender", "gender")
f_career = multiselect_filter("Career Level", "career_level")

fdf = df.copy()
for col, sel in [("department", f_dept), ("job_family", f_jobfam), ("grade", f_grade),
                  ("gender", f_gender), ("career_level", f_career)]:
    if sel is not None:
        fdf = fdf[fdf[col].isin(sel)]

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(fdf):,} of {len(df):,} employees in view")

# ============================================================
# HEADER
# ============================================================
st.markdown("<h1 style='margin-bottom:2px;'>🧭 Internal Pay Equity & Rewards Dashboard</h1>", unsafe_allow_html=True)
 <div style="margin-top:-8px; margin-bottom:18px; color:#6b7280; font-size:16px;">
            <b>Developed by Hamza Aslam</b> | Executive Compensation & Benefits
        </div>
st.markdown("<p class='section-note'>Compensation structure, pay equity, and reward effectiveness across the workforce.</p>", unsafe_allow_html=True)
st.markdown("")

# ============================================================
# TOP KPI ROW
# ============================================================
headcount = len(fdf)
avg_compa = fdf["compa_ratio"].mean() if "compa_ratio" in fdf.columns else np.nan
median_pen = fdf["range_penetration"].median() if "range_penetration" in fdf.columns else np.nan

gap_pct = np.nan
if {"gender", "gross_salary"}.issubset(fdf.columns):
    m = fdf.loc[fdf["gender"] == "Male", "gross_salary"].mean()
    f = fdf.loc[fdf["gender"] == "Female", "gross_salary"].mean()
    if m and not pd.isna(m) and m != 0:
        gap_pct = (m - f) / m

pct_high_risk = (fdf["retention_risk"] == "High").mean() if "retention_risk" in fdf.columns and headcount else np.nan
avg_incr = fdf["increment_pct"].mean() if "increment_pct" in fdf.columns else np.nan
n_below = (fdf["compa_flag"] == "Below 80%").sum() if "compa_flag" in fdf.columns else 0
n_above = (fdf["compa_flag"] == "Above 120%").sum() if "compa_flag" in fdf.columns else 0

render_kpi_row([
    ("HEADCOUNT IN VIEW", f"{headcount:,}", INDIGO, ""),
    ("AVG COMPA-RATIO", fmt_pct(avg_compa) if not pd.isna(avg_compa) else "—", TEAL, "Target: 100%"),
    ("MEDIAN RANGE PENETRATION", fmt_pct(median_pen) if not pd.isna(median_pen) else "—", EMERALD, ""),
    ("GENDER PAY GAP (UNADJUSTED)", fmt_pct(gap_pct) if not pd.isna(gap_pct) else "—", ROSE, "Male vs Female avg"),
])
st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
render_kpi_row([
    ("AVG INCREMENT %", fmt_pct(avg_incr) if not pd.isna(avg_incr) else "—", VIOLET, ""),
    ("HIGH RETENTION RISK", fmt_pct(pct_high_risk) if not pd.isna(pct_high_risk) else "—", AMBER, "Share of workforce"),
    ("BELOW 80% COMPA", f"{n_below:,}", "#B23A48", "Underpaid vs. range"),
    ("ABOVE 120% COMPA", f"{n_above:,}", "#C9831A", "Overpaid vs. range"),
])

st.markdown("---")

# ============================================================
# TABS
# ============================================================
tab_overview, tab_payequity, tab_perf, tab_flags, tab_retention, tab_talent, tab_micro, tab_trend, tab_data = st.tabs(
    ["Overview", "Pay Equity", "Pay & Performance", "Range Flags",
     "Retention & Risk", "Critical Talent", "Employee Lookup", "Year-over-Year", "Data Explorer"]
)

# ------------------------------------------------------------
# OVERVIEW
# ------------------------------------------------------------
with tab_overview:
    c1, c2 = st.columns([1.3, 1])

    with c1:
        st.subheader("Headcount & Pay by Department")
        if {"department", "gross_salary"}.issubset(fdf.columns):
            agg = fdf.groupby("department").agg(
                Headcount=("department", "size"), Avg_Gross=("gross_salary", "mean")
            ).reset_index().sort_values("Headcount", ascending=False)
            fig = go.Figure()
            fig.add_bar(x=agg["department"], y=agg["Headcount"], name="Headcount", marker_color=TEAL)
            fig.add_trace(go.Scatter(x=agg["department"], y=agg["Avg_Gross"], name="Avg Gross Salary",
                                      mode="lines+markers", marker_color=INDIGO, yaxis="y2"))
            fig.update_layout(template=PLOTLY_TEMPLATE, height=420,
                               yaxis=dict(title="Headcount"),
                               yaxis2=dict(title="Avg Gross Salary", overlaying="y", side="right", showgrid=False),
                               legend=dict(orientation="h", y=1.12), margin=dict(t=30))
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Gender Composition")
        if "gender" in fdf.columns:
            gc = fdf["gender"].value_counts().reset_index()
            gc.columns = ["Gender", "Count"]
            fig = px.pie(gc, names="Gender", values="Count", hole=0.55,
                         color_discrete_sequence=[INDIGO, TEAL, AMBER])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=420, margin=dict(t=30))
            st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Headcount by Career Level")
        if "career_level" in fdf.columns:
            cl = fdf["career_level"].value_counts().reset_index()
            cl.columns = ["Career Level", "Count"]
            fig = px.bar(cl, x="Count", y="Career Level", orientation="h", color="Career Level",
                         color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=10), showlegend=False,
                               yaxis=dict(categoryorder="total ascending"))
            st.plotly_chart(fig, use_container_width=True)
    with c4:
        st.subheader("Performance Rating Distribution")
        if "performance" in fdf.columns:
            pr = fdf["performance"].value_counts().reindex(PERF_ORDER).dropna().reset_index()
            pr.columns = ["Rating", "Count"]
            fig = px.bar(pr, x="Rating", y="Count", color="Rating",
                         category_orders={"Rating": PERF_ORDER}, color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=10), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    st.subheader("Compa-Ratio Distribution")
    if "compa_ratio" in fdf.columns:
        fig = px.histogram(fdf, x="compa_ratio", nbins=40, color_discrete_sequence=[TEAL])
        fig.add_vline(x=1.0, line_dash="dash", line_color=NAVY, annotation_text="Midpoint")
        fig.add_vline(x=0.80, line_dash="dot", line_color=ROSE, annotation_text="80%")
        fig.add_vline(x=1.20, line_dash="dot", line_color=AMBER, annotation_text="120%")
        fig.update_layout(template=PLOTLY_TEMPLATE, height=380, xaxis_tickformat=".0%",
                           xaxis_title="Compa-Ratio", yaxis_title="Employees")
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# PAY EQUITY
# ------------------------------------------------------------
with tab_payequity:
    st.subheader("Gender Pay Equity")
    if {"gender", "gross_salary", "grade"}.issubset(fdf.columns):
        overall = fdf.groupby("gender")["gross_salary"].mean()
        unadj_gap, adj_gap = np.nan, np.nan
        if "Male" in overall.index and "Female" in overall.index and overall["Male"]:
            unadj_gap = (overall["Male"] - overall["Female"]) / overall["Male"]
            grade_grp = fdf.groupby(["grade", "gender"])["gross_salary"].mean().unstack()
            if {"Male", "Female"}.issubset(grade_grp.columns):
                grade_counts = fdf.groupby("grade").size()
                grade_grp["gap"] = (grade_grp["Male"] - grade_grp["Female"]) / grade_grp["Male"]
                valid = grade_grp["gap"].dropna()
                adj_gap = np.average(valid, weights=grade_counts.reindex(valid.index))

        render_kpi_row([
            ("UNADJUSTED PAY GAP", fmt_pct(unadj_gap), ROSE, "Raw average difference"),
            ("GRADE-ADJUSTED PAY GAP", fmt_pct(adj_gap) if not pd.isna(adj_gap) else "—", INDIGO, "Same-grade comparison"),
        ])
        st.markdown(
            "<div class='note-box'>The grade-adjusted figure compares men and women within the "
            "same grade, removing the effect of role mix. A gap that shrinks substantially after "
            "adjustment usually points to representation differences across grades rather than "
            "unequal pay for the same role.</div>", unsafe_allow_html=True
        )

        st.markdown("#### Average Compa-Ratio by Grade and Gender")
        cg = fdf.groupby(["grade", "gender"])["compa_ratio"].mean().reset_index()
        fig = px.bar(cg, x="grade", y="compa_ratio", color="gender", barmode="group",
                     color_discrete_sequence=[INDIGO, TEAL, AMBER])
        fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE)
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420, yaxis_tickformat=".0%",
                           legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Pay Quartile Mix by Gender")
        if {"gender", "quartile"}.issubset(fdf.columns):
            qc = fdf.groupby(["gender", "quartile"]).size().reset_index(name="Count")
            qc["quartile"] = pd.Categorical(qc["quartile"], categories=QUARTILE_ORDER, ordered=True)
            fig = px.bar(qc.sort_values("quartile"), x="gender", y="Count", color="quartile",
                         category_orders={"quartile": QUARTILE_ORDER},
                         color_discrete_sequence=[ROSE, AMBER, TEAL, INDIGO], barmode="stack")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, legend_title="Quartile")
            st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown("#### Compa-Ratio by Tenure Band")
        if {"tenure", "compa_ratio"}.issubset(fdf.columns):
            bins = [-1, 2, 5, 10, 100]
            labels = ["0-2 yrs", "3-5 yrs", "6-10 yrs", "10+ yrs"]
            fdf["_tenure_band"] = pd.cut(fdf["tenure"], bins=bins, labels=labels)
            tb = fdf.groupby("_tenure_band", observed=True)["compa_ratio"].mean().reindex(labels).reset_index()
            fig = px.bar(tb, x="_tenure_band", y="compa_ratio", color="_tenure_band",
                         color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                               showlegend=False, xaxis_title="Tenure Band")
            st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Compa-Ratio Spread by Department")
    if {"department", "compa_ratio"}.issubset(fdf.columns):
        fig = px.box(fdf, x="department", y="compa_ratio", color="department",
                     color_discrete_sequence=PALETTE)
        fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE)
        fig.update_layout(template=PLOTLY_TEMPLATE, height=440, showlegend=False, yaxis_tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# PAY & PERFORMANCE
# ------------------------------------------------------------
with tab_perf:
    st.subheader("Pay-for-Performance Alignment")
    if {"performance", "compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            pr_compa = fdf.groupby("performance")["compa_ratio"].mean().reindex(PERF_ORDER).dropna().reset_index()
            fig = px.bar(pr_compa, x="performance", y="compa_ratio", category_orders={"performance": PERF_ORDER},
                         color="performance", color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                               showlegend=False, title="Avg Compa-Ratio by Performance Rating")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "increment_pct" in fdf.columns:
                pr_incr = fdf.groupby("performance")["increment_pct"].mean().reindex(PERF_ORDER).dropna().reset_index()
                fig = px.bar(pr_incr, x="performance", y="increment_pct", category_orders={"performance": PERF_ORDER},
                             color="performance", color_discrete_sequence=PALETTE)
                fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                                   showlegend=False, title="Avg Increment % by Performance Rating")
                st.plotly_chart(fig, use_container_width=True)
        st.markdown(
            "<div class='note-box'>In a well-functioning pay-for-performance structure, both "
            "charts should step upward from left to right. A flat or inverted pattern suggests "
            "pay positioning is not tracking performance ratings.</div>", unsafe_allow_html=True
        )

        st.markdown("#### Increment % vs. Position in Range")
        if "increment_pct" in fdf.columns:
            fig = px.scatter(fdf, x="compa_ratio", y="increment_pct", color="performance",
                              category_orders={"performance": PERF_ORDER}, color_discrete_sequence=PALETTE,
                              hover_data=safe_cols(["department", "designation", "grade"], fdf), opacity=0.7)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=460, xaxis_tickformat=".0%", yaxis_tickformat=".0%",
                               xaxis_title="Compa-Ratio", yaxis_title="Increment %")
            st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# RANGE FLAGS (compa < 80% and > 120%)
# ------------------------------------------------------------
with tab_flags:
    st.subheader("Compa-Ratio Range Flags")
    st.markdown(
        "<p class='section-note'>Employees positioned below 80% (underpaid relative to range) "
        "or above 120% (overpaid relative to range) of their grade midpoint.</p>",
        unsafe_allow_html=True,
    )

    if "compa_flag" in fdf.columns:
        n_below = (fdf["compa_flag"] == "Below 80%").sum()
        n_above = (fdf["compa_flag"] == "Above 120%").sum()
        n_within = (fdf["compa_flag"] == "Within Range").sum()
        render_kpi_row([
            ("BELOW 80% COMPA", f"{n_below:,}", "#B23A48", fmt_pct(n_below / headcount) if headcount else ""),
            ("ABOVE 120% COMPA", f"{n_above:,}", "#C9831A", fmt_pct(n_above / headcount) if headcount else ""),
            ("WITHIN RANGE", f"{n_within:,}", EMERALD, fmt_pct(n_within / headcount) if headcount else ""),
            ("TOTAL FLAGGED", f"{n_below + n_above:,}", INDIGO, fmt_pct((n_below + n_above) / headcount) if headcount else ""),
        ])

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Flags by Department")
            fc = fdf[fdf["compa_flag"] != "Within Range"].groupby(["department", "compa_flag"]).size().reset_index(name="Count")
            if not fc.empty:
                fig = px.bar(fc, x="department", y="Count", color="compa_flag", barmode="group",
                             color_discrete_map={"Below 80%": "#B23A48", "Above 120%": "#C9831A"})
                fig.update_layout(template=PLOTLY_TEMPLATE, height=400, legend_title="Flag")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("#### Flags by Grade")
            fg = fdf[fdf["compa_flag"] != "Within Range"].groupby(["grade", "compa_flag"]).size().reset_index(name="Count")
            if not fg.empty:
                fig = px.bar(fg, x="grade", y="Count", color="compa_flag", barmode="group",
                             color_discrete_map={"Below 80%": "#B23A48", "Above 120%": "#C9831A"})
                fig.update_layout(template=PLOTLY_TEMPLATE, height=400, legend_title="Flag")
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Compa-Ratio Map (flagged employees highlighted)")
        if "grade" in fdf.columns:
            fig = px.strip(fdf, x="grade", y="compa_ratio", color="compa_flag",
                            color_discrete_map={"Below 80%": "#B23A48", "Above 120%": "#C9831A", "Within Range": "#C7D0DA"},
                            stripmode="overlay")
            fig.add_hline(y=0.80, line_dash="dot", line_color="#B23A48")
            fig.add_hline(y=1.20, line_dash="dot", line_color="#C9831A")
            fig.update_layout(template=PLOTLY_TEMPLATE, height=460, yaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Statistical Outliers (z-score beyond ±2 within grade)")
        if "compa_zscore" in fdf.columns:
            outliers = fdf[fdf["compa_zscore"].abs() > 2]
            st.caption(f"{len(outliers):,} employees sit more than 2 standard deviations from their grade's average compa-ratio.")
            cols = safe_cols(["employee_id", "employee_name", "department", "designation", "grade",
                               "gender", "gross_salary", "compa_ratio", "compa_zscore"], outliers)
            st.dataframe(outliers[cols].sort_values("compa_zscore"), use_container_width=True, height=260)

        st.markdown("#### Full Flagged List")
        flagged = fdf[fdf["compa_flag"] != "Within Range"]
        cols = safe_cols(["employee_id", "employee_name", "department", "designation", "grade",
                           "gender", "gross_salary", "salary_mid", "compa_ratio", "compa_flag"], flagged)
        st.dataframe(flagged[cols].sort_values("compa_ratio"), use_container_width=True, height=320)
        st.download_button("Download flagged list (CSV)", flagged[cols].to_csv(index=False).encode("utf-8"),
                            "compa_ratio_flags.csv", "text/csv")

# ------------------------------------------------------------
# RETENTION & RISK
# ------------------------------------------------------------
with tab_retention:
    st.subheader("Retention Risk and Pay Positioning")
    if {"retention_risk", "compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            rr = fdf["retention_risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
            rr.columns = ["Risk", "Count"]
            fig = px.bar(rr, x="Risk", y="Count", color="Risk", category_orders={"Risk": RISK_ORDER},
                         color_discrete_sequence=[EMERALD, AMBER, ROSE])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, showlegend=False, title="Headcount by Retention Risk")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            rc = fdf.groupby("retention_risk")["compa_ratio"].mean().reindex(RISK_ORDER).dropna().reset_index()
            fig = px.bar(rc, x="retention_risk", y="compa_ratio", color="retention_risk",
                         category_orders={"retention_risk": RISK_ORDER}, color_discrete_sequence=[EMERALD, AMBER, ROSE])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, showlegend=False, yaxis_tickformat=".0%",
                               title="Avg Compa-Ratio by Retention Risk")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown(
            "<div class='flag-underpaid'>If High retention risk employees also carry a lower "
            "average compa-ratio than Low risk employees, pay positioning is likely contributing "
            "to attrition exposure in that group.</div>", unsafe_allow_html=True
        )

        if "performance" in fdf.columns:
            st.markdown("#### High-Performer Risk Concentration")
            hp_risk = fdf[fdf["performance"].isin(["HP", "EP"])]
            fig = px.density_heatmap(hp_risk, x="performance", y="retention_risk",
                                      category_orders={"performance": ["HP", "EP"], "retention_risk": RISK_ORDER},
                                      color_continuous_scale=["#EAF7F6", TEAL, INDIGO])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380)
            st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# CRITICAL TALENT
# ------------------------------------------------------------
with tab_talent:
    st.subheader("Critical Talent and Skill-Based Pay")
    if {"position_criticality", "compa_ratio"}.issubset(fdf.columns):
        order = [c for c in CRITICALITY_ORDER if c in fdf["position_criticality"].unique()]
        crit = fdf.groupby("position_criticality")["compa_ratio"].mean().reindex(order).reset_index()
        fig = px.bar(crit, x="position_criticality", y="compa_ratio", color="position_criticality",
                     color_discrete_sequence=[TEAL, AMBER, ROSE, INDIGO])
        fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                           showlegend=False, title="Avg Compa-Ratio by Position Criticality")
        st.plotly_chart(fig, use_container_width=True)

    if {"skill_rating", "compa_ratio"}.issubset(fdf.columns):
        st.markdown("#### Skill Rating vs. Pay Positioning")
        fig = px.scatter(fdf, x="skill_rating", y="compa_ratio",
                          color="job_family" if "job_family" in fdf.columns else None,
                          color_discrete_sequence=PALETTE, opacity=0.7,
                          hover_data=safe_cols(["employee_name", "department", "designation"], fdf))
        fig.update_layout(template=PLOTLY_TEMPLATE, height=460, yaxis_tickformat=".0%",
                           xaxis_title="Skill Rating (1-5)", yaxis_title="Compa-Ratio")
        st.plotly_chart(fig, use_container_width=True)

    if {"num_promotions", "tenure"}.issubset(fdf.columns):
        st.markdown("#### Promotion Velocity")
        c1, c2 = st.columns(2)
        with c1:
            avg_tenure_per_promo = fdf.groupby("num_promotions")["tenure"].mean().reset_index()
            fig = px.line(avg_tenure_per_promo, x="num_promotions", y="tenure", markers=True,
                          color_discrete_sequence=[INDIGO])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, title="Avg Tenure by Number of Promotions",
                               xaxis_title="Number of Promotions", yaxis_title="Avg Tenure (yrs)")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "gender" in fdf.columns:
                promo_gender = fdf.groupby("gender")["num_promotions"].mean().reset_index()
                fig = px.bar(promo_gender, x="gender", y="num_promotions", color="gender",
                             color_discrete_sequence=[INDIGO, TEAL, AMBER])
                fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                                   title="Avg Number of Promotions by Gender")
                st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# EMPLOYEE LOOKUP (MICRO-LEVEL ANALYSIS)
# ------------------------------------------------------------
with tab_micro:
    st.subheader("Employee-Level Benchmarking")
    st.markdown(
        "<p class='section-note'>Look up an individual employee and benchmark them against "
        "their direct peer group (same grade and department).</p>", unsafe_allow_html=True
    )

    id_col = "employee_id" if "employee_id" in fdf.columns else None
    name_col = "employee_name" if "employee_name" in fdf.columns else None

    if name_col or id_col:
        label_col = name_col or id_col
        search = st.text_input("Search by name or ID", "")
        options = fdf[label_col].astype(str)
        if search:
            options = options[options.str.contains(search, case=False, na=False)]
        selected = st.selectbox("Select employee", options.unique().tolist() if len(options) else [])

        if selected:
            emp = fdf[fdf[label_col].astype(str) == selected].iloc[0]

            peer_mask = pd.Series(True, index=fdf.index)
            if "grade" in fdf.columns:
                peer_mask &= fdf["grade"] == emp.get("grade")
            if "department" in fdf.columns:
                peer_mask &= fdf["department"] == emp.get("department")
            peers = fdf[peer_mask]

            percentile = np.nan
            if "compa_ratio" in peers.columns and len(peers) > 1:
                percentile = (peers["compa_ratio"] < emp["compa_ratio"]).mean()

            render_kpi_row([
                ("GROSS SALARY", fmt_currency(emp.get("gross_salary", np.nan)), INDIGO, ""),
                ("COMPA-RATIO", fmt_pct(emp.get("compa_ratio", np.nan)), TEAL, "vs. grade midpoint"),
                ("RANGE PENETRATION", fmt_pct(emp.get("range_penetration", np.nan)), EMERALD, ""),
                ("PEER PERCENTILE", fmt_pct(percentile) if not pd.isna(percentile) else "—", VIOLET,
                 f"within {len(peers)} grade/dept peers"),
            ])

            c1, c2 = st.columns([1, 1.2])
            with c1:
                st.markdown("#### Profile")
                profile_fields = ["department", "designation", "grade", "job_family", "career_level",
                                   "gender", "tenure", "performance", "retention_risk",
                                   "position_criticality", "qualification", "skill_rating"]
                profile_fields = safe_cols(profile_fields, fdf)
                profile_df = pd.DataFrame({"Field": profile_fields, "Value": [emp.get(f, "—") for f in profile_fields]})
                st.dataframe(profile_df, use_container_width=True, hide_index=True, height=380)
            with c2:
                st.markdown("#### Peer Group Distribution")
                if "compa_ratio" in peers.columns and len(peers) > 1:
                    fig = px.histogram(peers, x="compa_ratio", nbins=25, color_discrete_sequence=[TEAL])
                    fig.add_vline(x=emp["compa_ratio"], line_color=ROSE, line_width=3,
                                  annotation_text="Selected employee", annotation_position="top")
                    fig.update_layout(template=PLOTLY_TEMPLATE, height=420, xaxis_tickformat=".0%",
                                       xaxis_title="Compa-Ratio", yaxis_title="Peers")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Not enough peers in the same grade and department to plot a distribution.")

            st.markdown("#### Compare with Specific Peers")
            compare_options = peers[label_col].astype(str).tolist()
            compare_sel = st.multiselect("Add peers to compare", [o for o in compare_options if o != selected])
            if compare_sel:
                compare_df = fdf[fdf[label_col].astype(str).isin([selected] + compare_sel)]
                cols = safe_cols(["employee_name", "employee_id", "department", "grade", "gross_salary",
                                   "compa_ratio", "range_penetration", "performance", "tenure"], compare_df)
                st.dataframe(compare_df[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Employee Name or Employee ID column not found.")

# ------------------------------------------------------------
# YEAR-OVER-YEAR
# ------------------------------------------------------------
with tab_trend:
    st.subheader("Year-over-Year Trend")
    if df_prior is None:
        st.info("Upload a prior period file in the sidebar to unlock trend comparisons.")
    else:
        rows = []
        for label, frame in [("Prior Period", df_prior), ("Current Period", df)]:
            row = {"Period": label, "Headcount": len(frame)}
            if "compa_ratio" in frame.columns:
                row["Avg Compa-Ratio"] = frame["compa_ratio"].mean()
            if {"gender", "gross_salary"}.issubset(frame.columns):
                m = frame.loc[frame["gender"] == "Male", "gross_salary"].mean()
                f = frame.loc[frame["gender"] == "Female", "gross_salary"].mean()
                row["Gender Pay Gap"] = (m - f) / m if m else np.nan
            if "increment_pct" in frame.columns:
                row["Avg Increment %"] = frame["increment_pct"].mean()
            rows.append(row)
        trend_df = pd.DataFrame(rows)

        cards = []
        if "Avg Compa-Ratio" in trend_df.columns:
            delta = trend_df["Avg Compa-Ratio"].iloc[1] - trend_df["Avg Compa-Ratio"].iloc[0]
            cards.append(("AVG COMPA-RATIO", fmt_pct(trend_df["Avg Compa-Ratio"].iloc[1]), TEAL, f"{fmt_pct(delta)} vs prior"))
        if "Gender Pay Gap" in trend_df.columns:
            delta = trend_df["Gender Pay Gap"].iloc[1] - trend_df["Gender Pay Gap"].iloc[0]
            cards.append(("GENDER PAY GAP", fmt_pct(trend_df["Gender Pay Gap"].iloc[1]), ROSE, f"{fmt_pct(delta)} vs prior"))
        if "Avg Increment %" in trend_df.columns:
            delta = trend_df["Avg Increment %"].iloc[1] - trend_df["Avg Increment %"].iloc[0]
            cards.append(("AVG INCREMENT %", fmt_pct(trend_df["Avg Increment %"].iloc[1]), VIOLET, f"{fmt_pct(delta)} vs prior"))
        if cards:
            render_kpi_row(cards)

        st.markdown("#### Headcount and Compa-Ratio Trend")
        fig = go.Figure()
        fig.add_bar(x=trend_df["Period"], y=trend_df["Headcount"], name="Headcount", marker_color=TEAL)
        if "Avg Compa-Ratio" in trend_df.columns:
            fig.add_trace(go.Scatter(x=trend_df["Period"], y=trend_df["Avg Compa-Ratio"], name="Avg Compa-Ratio",
                                      mode="lines+markers", marker_color=INDIGO, yaxis="y2"))
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420, yaxis=dict(title="Headcount"),
                           yaxis2=dict(title="Avg Compa-Ratio", overlaying="y", side="right",
                                       tickformat=".0%", showgrid=False))
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# DATA EXPLORER
# ------------------------------------------------------------
with tab_data:
    st.subheader("Filtered Data Explorer")
    st.dataframe(fdf, use_container_width=True, height=520)
    csv = fdf.to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered data (CSV)", csv, "internal_equity_filtered.csv", "text/csv")
