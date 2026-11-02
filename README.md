# US Tech Job Demand Survey

**Which tech roles are large US companies actually hiring for right now?**

This project reads a list of 500 large US companies, finds the job system behind each careers page, pulls every live US job opening, and groups thousands of job titles into role families, then ranks them by how many companies are hiring for each one.

I built it to answer a practical question for my own job search: *where is real demand, and where is the competition lower?*

---

## How it works

```mermaid
flowchart LR
    A[500 companies<br/>career page URLs] --> B[Headless browser<br/>detects job system]
    B --> C[Pull live postings<br/>from public job feeds]
    C --> D[Keep US<br/>openings only]
    D --> E[Group titles into<br/>48 role families]
    E --> F[Excel report<br/>ranked by demand]
```

1. **Detect.** Playwright opens each careers page and watches the links and network requests to identify the job system (Workday, Greenhouse, Lever, Ashby, SmartRecruiters).
2. **Fetch.** The script pulls postings from each system's public job feed. For Workday, it finds the "United States" filter automatically and applies it on the server side.
3. **Filter.** Locations are checked against US states, cities and country fields, and non-US openings are dropped.
4. **Classify.** Keyword rules group titles into role families. For example, "Senior Data Analyst", "BI Developer" and "Reporting Analyst" all count as **Data Analyst / BI**.
5. **Report.** Results are written to an Excel workbook with a ranking, an industry breakdown, coverage per company, and every job counted.

Every step is cached, so a long run can be stopped and resumed.

---

## Results (November 2026 run)

| Metric | Value |
|---|---|
| Companies in list | 500 |
| Companies with jobs pulled | 144 |
| US openings counted | 50,178 |
| Role families | 48 |

### Top tech roles by number of companies hiring

| Rank | Role | Companies hiring (of 144) | US openings |
|---|---|---|---|
| 1 | Software Engineer / Developer | 88 | 1,921 |
| 2 | Cybersecurity | 83 | 1,078 |
| 3 | Machine Learning / AI Engineer | 83 | 984 |
| 4 | Product Manager / Owner | 76 | 385 |
| 5 | DevOps / SRE / Cloud | 54 | 373 |
| 6 | Data Scientist | 51 | 258 |
| 7 | Solutions / Enterprise Architect | 50 | 363 |
| 8 | Data Engineer | 50 | 197 |
| 9 | Data Analyst / BI | 46 | 123 |
| 10 | QA / Test Engineer | 45 | 308 |

**Key takeaways**

- **AI/ML demand has spread beyond big tech.** More than half of the companies covered, including banks, defense contractors and retailers, have open AI/ML roles.
- **Data Analyst roles are wide but thin.** They appear at 46 companies but average under 3 openings each, which means more applicants competing per seat.
- **Systems Administrator and Full Stack** roles have many openings per company (about 22 and 13), so each application targets a larger pool of seats.

---

## Run it yourself

```bash
pip3 install pandas openpyxl requests playwright
python3 -m playwright install chromium

# quick test on the first 20 companies
python3 role_demand_survey.py fortune500_careers.xlsx --limit 20

# full run
python3 role_demand_survey.py fortune500_careers.xlsx
```

The output is `role_demand_report.xlsx`.

**Input format:** an Excel sheet with the columns `Company`, `Industry` and `Career Page URL`. You can optionally add an `ATS Override` column with a direct job-board link for any company the script fails to detect.

**Options:**

| Flag | Meaning |
|---|---|
| `--limit N` | Only process the first N companies |
| `--concurrency N` | Browser tabs open at once (default 5) |
| `--max-jobs N` | Maximum postings per company (default 5000) |
| `--skip-detect` | Reuse cached job-system detection |
| `--out FILE` | Output file name |

---

## Limitations

- **Coverage is 144 of 500 companies.** Companies with custom career sites (Amazon, Google, Microsoft, Apple) and job systems that aren't supported yet (Oracle, SuccessFactors, iCIMS, Phenom, Eightfold) are not included. The sample therefore leans toward large enterprises that use Workday.
- **Workday returns at most about 2,000 postings per company**, so the opening counts for the biggest employers are undercounted. The "companies hiring" ranking is not affected.
- **One posting can represent several hires**, and some postings stay open all the time to collect applications.
- **Classification uses keyword rules**, so some titles land in broad or unclassified buckets. The `Unclassified Titles` sheet shows what to refine.

## Roadmap

- [ ] Support Oracle, SuccessFactors, iCIMS, Phenom and Eightfold
- [ ] Custom handlers for big-tech career sites
- [ ] Track demand over time with scheduled runs
- [ ] Dashboard for the results

---

## Tech stack

Python · Playwright · Requests · pandas · openpyxl

## Author

**Nageswararao Pathi**: AI engineering, workflow automation, and controls engineering.
