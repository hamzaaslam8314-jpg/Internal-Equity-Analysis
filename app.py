"""
Total Rewards & Internal Equity Intelligence Dashboard
-------------------------------------------------------
Upload your annual increment / compensation workbook and get an
interactive, presentation-ready internal pay equity analysis
(built for CHRO / CEO readouts).

Run locally:    streamlit run app.py
Deploy:         push this repo to GitHub, then deploy on streamlit.io
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
    page_title="Total Rewards Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#10243E"
TEAL = "#0EA5A8"
AMBER = "#F2A93B"
CORAL = "#E8604C"
SLATE = "#5B6B7B"
BG = "#F7F9FB"
GRID = "#E7ECF1"

PALETTE = [TEAL, NAVY, AMBER, CORAL, "#7D8FA8", "#9ED6D8"]
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
    [data-testid="stMetricValue"] {{ color: {NAVY}; font-weight: 700; }}
    [data-testid="stMetricLabel"] {{ color: {SLATE}; }}
    .kpi-card {{
        background: white; border-radius: 14px; padding: 18px 20px;
        box-shadow: 0 1px 6px rgba(16,36,62,0.08); border: 1px solid {GRID};
    }}
    .section-note {{ color: {SLATE}; font-size: 0.92rem; margin-top: -8px; }}
    .insight-box {{
        background: #EFFAFA; border-left: 4px solid {TEAL}; border-radius: 6px;
        padding: 10px 14px; font-size: 0.92rem; color: {NAVY}; margin: 8px 0 18px 0;
    }}
    .flag-box {{
        background: #FDF1EA; border-left: 4px solid {CORAL}; border-radius: 6px;
        padding: 10px 14px; font-size: 0.92rem; color: {NAVY}; margin: 8px 0 18px 0;
    }}
    div[data-baseweb="tab-list"] {{ gap: 4px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

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

    # type cleanup
    for c in ["compa_ratio", "range_penetration", "increment_pct"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            if df[c].max() is not None and df[c].max() > 3:  # stored as % not fraction
                df[c] = df[c] / 100.0
    for c in ["gross_salary", "basic_salary", "salary_min", "salary_mid", "salary_max",
              "tenure", "years_in_role", "num_promotions", "skill_rating"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["gender"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip().str.title()
    if "quartile" in df.columns:
        df["quartile"] = df["quartile"].astype(str).str.strip().str.upper()
    if "performance" in df.columns:
        df["performance"] = df["performance"].astype(str).str.strip().str.upper()
    if "retention_risk" in df.columns:
        df["retention_risk"] = df["retention_risk"].astype(str).str.strip().str.title()

    return df, df_raw, missing


def ordered_categorical(series, order):
    present = [o for o in order if o in series.unique()]
    extras = [o for o in series.dropna().unique() if o not in present]
    return present + extras


def fmt_pct(x, decimals=1):
    if pd.isna(x):
        return "—"
    return f"{x*100:.{decimals}f}%"


def fmt_currency(x):
    if pd.isna(x):
        return "—"
    return f"Rs. {x:,.0f}"


# ============================================================
# SIDEBAR — UPLOAD & FILTERS
# ============================================================
st.sidebar.markdown("## 📁 Data Source")
uploaded = st.sidebar.file_uploader(
    "Upload this year's compensation workbook (.xlsx)", type=["xlsx"]
)
prior_uploaded = st.sidebar.file_uploader(
    "Optional: upload prior year's file for trend comparison", type=["xlsx"], key="prior"
)

if not uploaded:
    st.markdown(
        f"""
        <div style="text-align:center; padding-top:80px;">
        <h1 style="color:{NAVY};">📊 Total Rewards & Internal Equity Intelligence</h1>
        <p style="color:{SLATE}; font-size:1.05rem; max-width:640px; margin:auto;">
        Upload your annual compensation / increment workbook in the sidebar to instantly
        generate a presentation-ready internal equity analysis — pay equity, compa-ratio
        positioning, performance-reward alignment, compression, and retention risk —
        built for CHRO and CEO readouts.
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

df, df_raw, missing_cols = load_workbook(uploaded.getvalue())

if missing_cols:
    st.sidebar.warning(
        "Some expected columns were not found and related views will be skipped:\n\n"
        + ", ".join(missing_cols)
    )

df_prior = None
if prior_uploaded:
    df_prior, _, _ = load_workbook(prior_uploaded.getvalue())

st.sidebar.markdown("---")
st.sidebar.markdown("## 🔍 Filters")


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
if f_dept is not None:
    fdf = fdf[fdf["department"].isin(f_dept)]
if f_jobfam is not None:
    fdf = fdf[fdf["job_family"].isin(f_jobfam)]
if f_grade is not None:
    fdf = fdf[fdf["grade"].isin(f_grade)]
if f_gender is not None:
    fdf = fdf[fdf["gender"].isin(f_gender)]
if f_career is not None:
    fdf = fdf[fdf["career_level"].isin(f_career)]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(fdf):,}** of **{len(df):,}** employees")

# ============================================================
# HEADER
# ============================================================
st.markdown(
    f"<h1 style='margin-bottom:0;'>📊 Total Rewards & Internal Equity Intelligence</h1>"
    f"<p class='section-note'>Confidential — prepared for executive leadership review</p>",
    unsafe_allow_html=True,
)

# ============================================================
# TOP KPI ROW
# ============================================================
k1, k2, k3, k4, k5, k6 = st.columns(6)

headcount = len(fdf)
avg_compa = fdf["compa_ratio"].mean() if "compa_ratio" in fdf.columns else np.nan
median_pen = fdf["range_penetration"].median() if "range_penetration" in fdf.columns else np.nan

gap_pct = np.nan
if {"gender", "gross_salary"}.issubset(fdf.columns):
    m = fdf.loc[fdf["gender"] == "Male", "gross_salary"].mean()
    f = fdf.loc[fdf["gender"] == "Female", "gross_salary"].mean()
    if m and not pd.isna(m) and m != 0:
        gap_pct = (m - f) / m

pct_high_risk = np.nan
if "retention_risk" in fdf.columns and headcount:
    pct_high_risk = (fdf["retention_risk"] == "High").mean()

avg_incr = fdf["increment_pct"].mean() if "increment_pct" in fdf.columns else np.nan

pct_below_min = np.nan
if {"compa_ratio"}.issubset(fdf.columns):
    pct_below_min = (fdf["compa_ratio"] < 0.80).mean()

with k1:
    st.metric("Headcount in View", f"{headcount:,}")
with k2:
    st.metric("Avg Compa-Ratio", fmt_pct(avg_compa) if not pd.isna(avg_compa) else "—")
with k3:
    st.metric("Median Range Penetration", fmt_pct(median_pen) if not pd.isna(median_pen) else "—")
with k4:
    st.metric("Unadjusted Gender Pay Gap", fmt_pct(gap_pct) if not pd.isna(gap_pct) else "—")
with k5:
    st.metric("High Retention Risk", fmt_pct(pct_high_risk) if not pd.isna(pct_high_risk) else "—")
with k6:
    st.metric("Avg Increment %", fmt_pct(avg_incr) if not pd.isna(avg_incr) else "—")

st.markdown("---")

# ============================================================
# TABS
# ============================================================
tab_overview, tab_payequity, tab_perf, tab_compression, tab_retention, tab_talent, tab_trend, tab_data = st.tabs(
    ["🏠 Overview", "⚖️ Pay Equity", "🎯 Performance & Reward", "📐 Compression & Range",
     "🛡️ Retention & Risk", "🌟 Critical Talent", "📈 Year-over-Year", "🗂️ Data Explorer"]
)

# ------------------------------------------------------------
# OVERVIEW
# ------------------------------------------------------------
with tab_overview:
    c1, c2 = st.columns([1.3, 1])

    with c1:
        st.subheader("Headcount & Pay Mix by Department")
        if {"department", "gross_salary"}.issubset(fdf.columns):
            agg = fdf.groupby("department").agg(
                Headcount=("department", "size"),
                Avg_Gross=("gross_salary", "mean"),
            ).reset_index().sort_values("Headcount", ascending=False)
            fig = go.Figure()
            fig.add_bar(x=agg["department"], y=agg["Headcount"], name="Headcount",
                        marker_color=TEAL, yaxis="y1")
            fig.add_trace(go.Scatter(x=agg["department"], y=agg["Avg_Gross"], name="Avg Gross Salary",
                                      mode="lines+markers", marker_color=NAVY, yaxis="y2"))
            fig.update_layout(
                template=PLOTLY_TEMPLATE, height=420,
                yaxis=dict(title="Headcount"),
                yaxis2=dict(title="Avg Gross Salary", overlaying="y", side="right", showgrid=False),
                legend=dict(orientation="h", y=1.12), margin=dict(t=30),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Department / salary columns not found.")

    with c2:
        st.subheader("Gender Composition")
        if "gender" in fdf.columns:
            gc = fdf["gender"].value_counts().reset_index()
            gc.columns = ["Gender", "Count"]
            fig = px.pie(gc, names="Gender", values="Count", hole=0.55,
                         color_discrete_sequence=[NAVY, TEAL, AMBER])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=420, margin=dict(t=30))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Gender column not found.")

    c3, c4 = st.columns(2)
    with c3:
        st.subheader("Headcount by Career Level")
        if "career_level" in fdf.columns:
            cl = fdf["career_level"].value_counts().reset_index()
            cl.columns = ["Career Level", "Count"]
            fig = px.bar(cl, x="Count", y="Career Level", orientation="h",
                         color_discrete_sequence=[TEAL])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=10),
                               yaxis=dict(categoryorder="total ascending"))
            st.plotly_chart(fig, use_container_width=True)
    with c4:
        st.subheader("Performance Rating Distribution")
        if "performance" in fdf.columns:
            pr = fdf["performance"].value_counts().reindex(PERF_ORDER).dropna().reset_index()
            pr.columns = ["Rating", "Count"]
            fig = px.bar(pr, x="Rating", y="Count", color="Rating",
                         category_orders={"Rating": PERF_ORDER},
                         color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=10), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# PAY EQUITY
# ------------------------------------------------------------
with tab_payequity:
    st.subheader("Gender Pay Equity")
    if {"gender", "gross_salary", "grade"}.issubset(fdf.columns):
        overall = fdf.groupby("gender")["gross_salary"].mean()
        if "Male" in overall.index and "Female" in overall.index and overall["Male"]:
            unadj_gap = (overall["Male"] - overall["Female"]) / overall["Male"]
            grade_grp = fdf.groupby(["grade", "gender"])["gross_salary"].mean().unstack()
            if {"Male", "Female"}.issubset(grade_grp.columns):
                grade_counts = fdf.groupby("grade").size()
                grade_grp["gap"] = (grade_grp["Male"] - grade_grp["Female"]) / grade_grp["Male"]
                adj_gap = np.average(grade_grp["gap"].dropna(),
                                      weights=grade_counts.reindex(grade_grp["gap"].dropna().index))
            else:
                adj_gap = np.nan

            ic1, ic2 = st.columns(2)
            ic1.metric("Unadjusted Gender Pay Gap", fmt_pct(unadj_gap))
            ic2.metric("Grade-Adjusted Gender Pay Gap", fmt_pct(adj_gap) if not pd.isna(adj_gap) else "—")

            st.markdown(
                f"<div class='insight-box'>📌 The grade-adjusted gap compares men and women "
                f"<b>within the same grade</b>, removing the effect of role mix. A meaningfully "
                f"smaller adjusted gap usually means the raw gap is driven by representation "
                f"(fewer women in senior grades) rather than unequal pay for equal work — and "
                f"vice versa.</div>", unsafe_allow_html=True
            )

        st.markdown("#### Average Compa-Ratio by Grade and Gender")
        cg = fdf.groupby(["grade", "gender"])["compa_ratio"].mean().reset_index()
        fig = px.bar(cg, x="grade", y="compa_ratio", color="gender", barmode="group",
                     color_discrete_sequence=[NAVY, TEAL, AMBER])
        fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE, annotation_text="Range Midpoint (1.00)")
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420, yaxis_tickformat=".0%",
                           legend=dict(orientation="h", y=1.1))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Gender / Gross Salary / Grade columns not found.")

    st.markdown("#### Pay Quartile Mix by Gender")
    if {"gender", "quartile"}.issubset(fdf.columns):
        qc = fdf.groupby(["gender", "quartile"]).size().reset_index(name="Count")
        qc["quartile"] = pd.Categorical(qc["quartile"], categories=QUARTILE_ORDER, ordered=True)
        fig = px.bar(qc.sort_values("quartile"), x="gender", y="Count", color="quartile",
                     category_orders={"quartile": QUARTILE_ORDER},
                     color_discrete_sequence=[CORAL, AMBER, TEAL, NAVY],
                     barmode="stack")
        fig.update_layout(template=PLOTLY_TEMPLATE, height=400, legend_title="Quartile in Range")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown(
            "<div class='section-note'>Q1 = bottom of salary range, Q4 = top. A disproportionate "
            "share of one gender sitting in Q1 across grades is a classic equity red flag.</div>",
            unsafe_allow_html=True,
        )

    st.markdown("#### Compa-Ratio Spread by Department")
    if {"department", "compa_ratio"}.issubset(fdf.columns):
        fig = px.box(fdf, x="department", y="compa_ratio", color="department",
                     color_discrete_sequence=PALETTE)
        fig.add_hline(y=1.0, line_dash="dot", line_color=SLATE)
        fig.update_layout(template=PLOTLY_TEMPLATE, height=440, showlegend=False,
                           yaxis_tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# PERFORMANCE & REWARD
# ------------------------------------------------------------
with tab_perf:
    st.subheader("Is Pay Rewarding Performance?")
    if {"performance", "compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            pr_compa = fdf.groupby("performance")["compa_ratio"].mean().reindex(PERF_ORDER).dropna().reset_index()
            fig = px.bar(pr_compa, x="performance", y="compa_ratio",
                         category_orders={"performance": PERF_ORDER},
                         color="performance", color_discrete_sequence=PALETTE)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                               showlegend=False, title="Avg Compa-Ratio by Performance Rating")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown(
                "<div class='insight-box'>📌 In a healthy pay-for-performance structure, "
                "this should step upward left-to-right (NI &lt; SP &lt; HP &lt; EP). A flat "
                "or inverted line means high performers are not positioned ahead of low "
                "performers in their pay range.</div>", unsafe_allow_html=True
            )
        with c2:
            if "increment_pct" in fdf.columns:
                pr_incr = fdf.groupby("performance")["increment_pct"].mean().reindex(PERF_ORDER).dropna().reset_index()
                fig = px.bar(pr_incr, x="performance", y="increment_pct",
                             category_orders={"performance": PERF_ORDER},
                             color="performance", color_discrete_sequence=PALETTE)
                fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                                   showlegend=False, title="Avg Increment % by Performance Rating")
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Increment % vs. Position in Range")
        if "increment_pct" in fdf.columns and "compa_ratio" in fdf.columns:
            fig = px.scatter(fdf, x="compa_ratio", y="increment_pct", color="performance",
                              category_orders={"performance": PERF_ORDER},
                              color_discrete_sequence=PALETTE,
                              hover_data=[c for c in ["department", "designation", "grade"] if c in fdf.columns],
                              opacity=0.7)
            fig.update_layout(template=PLOTLY_TEMPLATE, height=460,
                               xaxis_tickformat=".0%", yaxis_tickformat=".0%",
                               xaxis_title="Compa-Ratio (position in range)",
                               yaxis_title="Increment %")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown(
                "<div class='section-note'>Look for high performers (lighter colors) clustered "
                "at low compa-ratio with low increments — these are your most under-rewarded "
                "top talent.</div>", unsafe_allow_html=True
            )
    else:
        st.info("Performance Rating / Compa-Ratio columns not found.")

# ------------------------------------------------------------
# COMPRESSION & RANGE
# ------------------------------------------------------------
with tab_compression:
    st.subheader("Salary Structure Health")
    if {"grade", "gross_salary"}.issubset(fdf.columns):
        disp = fdf.groupby("grade")["gross_salary"].agg(["mean", "std", "count"]).reset_index()
        disp["cv"] = disp["std"] / disp["mean"]
        disp = disp.sort_values("grade")
        fig = px.bar(disp, x="grade", y="cv", color="cv", color_continuous_scale=["#D9EFEF", TEAL, NAVY])
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420, yaxis_tickformat=".0%",
                           title="Pay Dispersion (Coefficient of Variation) by Grade",
                           coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
        st.markdown(
            "<div class='insight-box'>📌 Coefficient of Variation shows how spread out pay is "
            "<b>within the same grade</b>. Grades above ~15-20% dispersion usually warrant a "
            "closer look — it may signal inconsistent offer decisions or legacy pay history "
            "that hasn't been corrected.</div>", unsafe_allow_html=True
        )

    if {"grade", "compa_ratio"}.issubset(fdf.columns):
        st.markdown("#### Range Penetration Spread by Grade")
        fig = px.box(fdf, x="grade", y="range_penetration" if "range_penetration" in fdf.columns else "compa_ratio",
                     color_discrete_sequence=[TEAL])
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420, yaxis_tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Red-Circled (Above Max) & Green-Circled (Below Min) Employees")
    if {"compa_ratio"}.issubset(fdf.columns):
        below = fdf[fdf["range_penetration"] < 0] if "range_penetration" in fdf.columns else fdf[fdf["compa_ratio"] < 0.75]
        above = fdf[fdf["range_penetration"] > 1] if "range_penetration" in fdf.columns else fdf[fdf["compa_ratio"] > 1.25]
        c1, c2 = st.columns(2)
        c1.metric("Below Range Minimum", f"{len(below):,}")
        c2.metric("Above Range Maximum", f"{len(above):,}")
        show_cols = [c for c in ["employee_name", "department", "designation", "grade",
                                  "gross_salary", "compa_ratio", "range_penetration"] if c in fdf.columns]
        with st.expander("View employees outside range"):
            st.dataframe(pd.concat([below, above])[show_cols], use_container_width=True)

# ------------------------------------------------------------
# RETENTION & RISK
# ------------------------------------------------------------
with tab_retention:
    st.subheader("Retention Risk Through a Pay Lens")
    if {"retention_risk", "compa_ratio"}.issubset(fdf.columns):
        c1, c2 = st.columns(2)
        with c1:
            rr = fdf["retention_risk"].value_counts().reindex(RISK_ORDER).dropna().reset_index()
            rr.columns = ["Risk", "Count"]
            fig = px.bar(rr, x="Risk", y="Count", color="Risk",
                         category_orders={"Risk": RISK_ORDER},
                         color_discrete_sequence=[TEAL, AMBER, CORAL])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, showlegend=False,
                               title="Headcount by Retention Risk")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            rc = fdf.groupby("retention_risk")["compa_ratio"].mean().reindex(RISK_ORDER).dropna().reset_index()
            fig = px.bar(rc, x="retention_risk", y="compa_ratio", color="retention_risk",
                         category_orders={"retention_risk": RISK_ORDER},
                         color_discrete_sequence=[TEAL, AMBER, CORAL])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=400, showlegend=False,
                               yaxis_tickformat=".0%", title="Avg Compa-Ratio by Retention Risk")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown(
            "<div class='flag-box'>🚩 If <b>High</b> retention risk employees also show a "
            "<b>lower</b> average compa-ratio than Low risk employees, pay positioning is "
            "likely contributing to flight risk — a strong, board-ready talking point for "
            "targeted retention budget.</div>", unsafe_allow_html=True
        )

        if "performance" in fdf.columns:
            st.markdown("#### High-Performer Flight Risk Matrix")
            hp_risk = fdf[fdf["performance"].isin(["HP", "EP"])]
            matrix = hp_risk.groupby(["retention_risk", "performance"]).size().reset_index(name="Count")
            fig = px.density_heatmap(hp_risk, x="performance", y="retention_risk",
                                      category_orders={"performance": ["HP", "EP"], "retention_risk": RISK_ORDER},
                                      color_continuous_scale=["#EFFAFA", TEAL, NAVY])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380, title="High & Exceptional Performers — Risk Concentration")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Retention Risk / Compa-Ratio columns not found.")

# ------------------------------------------------------------
# CRITICAL TALENT
# ------------------------------------------------------------
with tab_talent:
    st.subheader("Critical Talent & Skill-Based Pay Positioning")
    if {"position_criticality", "compa_ratio"}.issubset(fdf.columns):
        crit = fdf.groupby("position_criticality")["compa_ratio"].mean().reindex(
            [c for c in CRITICALITY_ORDER if c in fdf["position_criticality"].unique()]
        ).reset_index()
        fig = px.bar(crit, x="position_criticality", y="compa_ratio", color="position_criticality",
                     color_discrete_sequence=[TEAL, AMBER, CORAL, NAVY])
        fig.update_layout(template=PLOTLY_TEMPLATE, height=400, yaxis_tickformat=".0%",
                           showlegend=False, title="Avg Compa-Ratio by Position Criticality")
        st.plotly_chart(fig, use_container_width=True)
        st.markdown(
            "<div class='insight-box'>📌 'Critical' roles sitting below 'Medium' or 'Low' "
            "criticality roles in compa-ratio is a market-competitiveness risk for the "
            "positions the business can least afford to lose.</div>", unsafe_allow_html=True
        )

    if {"skill_rating", "compa_ratio"}.issubset(fdf.columns):
        st.markdown("#### Skill Rating vs. Pay Positioning")
        fig = px.scatter(fdf, x="skill_rating", y="compa_ratio",
                          color="job_family" if "job_family" in fdf.columns else None,
                          color_discrete_sequence=PALETTE, opacity=0.7,
                          hover_data=[c for c in ["employee_name", "department", "designation"] if c in fdf.columns])
        fig.update_layout(template=PLOTLY_TEMPLATE, height=460, yaxis_tickformat=".0%",
                           xaxis_title="Skill Rating (1-5)", yaxis_title="Compa-Ratio")
        st.plotly_chart(fig, use_container_width=True)

    if {"num_promotions", "tenure"}.issubset(fdf.columns):
        st.markdown("#### Promotion Velocity")
        c1, c2 = st.columns(2)
        with c1:
            avg_tenure_per_promo = fdf.groupby("num_promotions")["tenure"].mean().reset_index()
            fig = px.line(avg_tenure_per_promo, x="num_promotions", y="tenure", markers=True,
                          color_discrete_sequence=[NAVY])
            fig.update_layout(template=PLOTLY_TEMPLATE, height=380,
                               title="Avg Tenure by Number of Promotions",
                               xaxis_title="Number of Promotions", yaxis_title="Avg Tenure (yrs)")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            if "gender" in fdf.columns:
                promo_gender = fdf.groupby("gender")["num_promotions"].mean().reset_index()
                fig = px.bar(promo_gender, x="gender", y="num_promotions", color="gender",
                             color_discrete_sequence=[NAVY, TEAL, AMBER])
                fig.update_layout(template=PLOTLY_TEMPLATE, height=380, showlegend=False,
                                   title="Avg Number of Promotions by Gender")
                st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------
# YEAR-OVER-YEAR
# ------------------------------------------------------------
with tab_trend:
    st.subheader("Year-over-Year Trend")
    if df_prior is None:
        st.info(
            "Upload last year's workbook in the sidebar ('Optional: upload prior year's file') "
            "to unlock trend comparisons for compa-ratio, pay gap, and headcount."
        )
    else:
        rows = []
        for label, frame in [("Prior Year", df_prior), ("Current Year", df)]:
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

        c1, c2, c3 = st.columns(3)
        if "Avg Compa-Ratio" in trend_df.columns:
            delta = trend_df["Avg Compa-Ratio"].iloc[1] - trend_df["Avg Compa-Ratio"].iloc[0]
            c1.metric("Avg Compa-Ratio", fmt_pct(trend_df["Avg Compa-Ratio"].iloc[1]), fmt_pct(delta))
        if "Gender Pay Gap" in trend_df.columns:
            delta = trend_df["Gender Pay Gap"].iloc[1] - trend_df["Gender Pay Gap"].iloc[0]
            c2.metric("Gender Pay Gap", fmt_pct(trend_df["Gender Pay Gap"].iloc[1]), fmt_pct(-delta) + " improvement" if delta < 0 else fmt_pct(delta))
        if "Avg Increment %" in trend_df.columns:
            delta = trend_df["Avg Increment %"].iloc[1] - trend_df["Avg Increment %"].iloc[0]
            c3.metric("Avg Increment %", fmt_pct(trend_df["Avg Increment %"].iloc[1]), fmt_pct(delta))

        st.markdown("#### Headcount & Compa-Ratio Trend")
        fig = go.Figure()
        fig.add_bar(x=trend_df["Period"], y=trend_df["Headcount"], name="Headcount", marker_color=TEAL)
        if "Avg Compa-Ratio" in trend_df.columns:
            fig.add_trace(go.Scatter(x=trend_df["Period"], y=trend_df["Avg Compa-Ratio"],
                                      name="Avg Compa-Ratio", mode="lines+markers",
                                      marker_color=NAVY, yaxis="y2"))
        fig.update_layout(template=PLOTLY_TEMPLATE, height=420,
                           yaxis=dict(title="Headcount"),
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
    st.download_button("⬇️ Download filtered data as CSV", csv, "internal_equity_filtered.csv", "text/csv")

st.markdown("---")
st.caption(
    "This dashboard is generated entirely from the uploaded file in your browser session; "
    "no data is stored or transmitted elsewhere."
)
