# Total Rewards & Internal Equity Intelligence Dashboard

An interactive Streamlit dashboard that turns an annual compensation /
increment workbook into a presentation-ready internal pay equity analysis —
built for CHRO and CEO readouts.

Upload a new workbook every year and the dashboard instantly regenerates:

- **Overview** — headcount, pay mix, gender composition, performance distribution
- **Pay Equity** — unadjusted & grade-adjusted gender pay gap, compa-ratio by grade/gender, quartile mix
- **Performance & Reward** — pay-for-performance alignment, increment vs. position-in-range
- **Compression & Range** — pay dispersion by grade, range penetration spread, red/green-circled employees
- **Retention & Risk** — retention risk vs. compa-ratio, high-performer flight-risk matrix
- **Critical Talent** — pay positioning for critical roles, skill rating vs. pay, promotion velocity
- **Year-over-Year** — optional second upload to trend key metrics against last year
- **Data Explorer** — filterable raw table with CSV export

No company-specific branding is hardcoded — safe to reuse across organizations or years.

## Expected columns

The app looks for these columns (case-insensitive, minor header variations are
auto-matched — see `COLUMN_CANDIDATES` in `app.py` to add more aliases):

Employee ID, Employee Name, Department, Designation, Existing Grade, New Grade,
Promotion, Gender, DOJ, Tenure, Years in Current Role, Job Family, Career Level,
Position Criticality, Supervisory Responsibility (Y/N), Basic Salary, Gross
Salary, Salary Range Minimum/Midpoint/Maximum, Compa Ratio, Range Penetration,
Quartile in Range, Performance Rating, Last Performance Rating, Increment %,
Number of Promotions, Last Promotion Date, Highest Qualification, Critical
Skills, Skill Rating, Retention Risk, Disciplinary Action in Last Year,
Employment Type, Employment Status.

Missing columns are handled gracefully — related charts are simply skipped
with a note in the sidebar.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this folder to a GitHub repository (keep `app.py`, `requirements.txt`,
   and `.streamlit/config.toml`).
2. Go to [share.streamlit.io](https://share.streamlit.io), connect your GitHub
   account, and select the repo.
3. Set the main file path to `app.py` and deploy.
4. Each year, open the live app and upload the new workbook — nothing else
   to configure.

## Privacy

All processing happens in-memory for the active session. The uploaded file is
not written to disk or persisted between sessions.
