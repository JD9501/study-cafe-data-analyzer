# Study Café Survey Analyzer

Should a study café open near campus, and how should it be run? This tool takes a student survey and answers that with data: what students want, what they would pay, when they would come, and what the café would earn.

**Try it in your browser: https://jd9501.github.io/study-cafe-data-analyzer/**

Nothing to install. The analysis runs inside your browser, so any dataset you upload never leaves your computer. The first visit takes about a minute while Python loads; later visits are faster.

## What it does

- **Cleans the survey.** It removes duplicates, impossible values, contradictory and careless answers, and reports every step.
- **Shows what students want.** It covers times, seats, food and drink, noise, customer groups (k-means) and patterns of answers that go together (association rules, with significance tests).
- **Prices.** Demand curves for a membership and a visit, price elasticity with 95% ranges, room to raise prices, and who would pay more.
- **Forecasts.** Peak hours, footfall, revenue, seating demand and membership, with each model compared against simple guessing.
- **Membership predictor.** It estimates how likely one student is to buy a membership.
- **What-if scenarios.** Change seats, prices, costs and catchment to see revenue, profit, break-even and staffing, with a price sensitivity table.
- **Two views.** Surface view shows what an owner needs; Deep view shows the full analysis.
- **Reports.** Surface and Deep reports as PDF or interactive HTML, written in plain words, opening with the key findings.

It comes with the survey of 6,870 students it was built for. Drag the Excel icon into the box to start, or upload your own survey in the same format.

## Get it

| | |
|---|---|
| **Web** | https://jd9501.github.io/study-cafe-data-analyzer/ |
| **Google Colab** | Paste [`source/study_cafe_v57.py`](source/study_cafe_v57.py) into a Colab cell and run it. |

## How it works

The analysis is a Python app (Flask, pandas, scikit-learn, SciPy). In the web version, the same Python code runs in the browser through [Pyodide](https://pyodide.org), which is Python compiled to WebAssembly, inside a Web Worker. The page's requests to the "server" are answered by that worker, so the site is just static files on GitHub Pages. The charts use Chart.js, and the reports use jsPDF.

| Folder | What it holds |
|---|---|
| `index.html`, `app/`, `vendor/`, `wheels/`, `data/` | the website |
| `source/` | the Colab script and the script that builds the website |
| `docs/` | the design notes |
| `sample-reports/` | example Surface, Deep and What-if reports |

## About

Built by Jithin as part of the Global MBA at SP Jain School of Global Management (DAIDM course).

The survey answers are what students *said* they would do, so the figures are intentions, not sales. Read them as a guide to which options work better, not as a forecast.
