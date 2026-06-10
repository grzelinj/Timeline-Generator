# Timeline Generator

A Streamlit-based web app that generates customized project timelines from Excel templates. Built to automate timeline creation for equity compensation and vesting workflows.

## What it does

- Uploads an Excel task list and automatically calculates dates
- Generates a clean, ready-to-use project timeline
- Outputs structured timelines for project planning, vesting schedules, or operational tracking
- Handles date logic, task dependencies, and formatting automatically

## Tech stack

- **Python** + **Streamlit** — web UI
- **Pandas** / **openpyxl** — Excel processing
- **Deployed on DigitalOcean**

## How to run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open `http://localhost:8501` in your browser.

## Background

Built to solve a real problem at Computershare — manually creating vesting timelines for equity compensation plans (RSU, ESPP, SAYE) across S&P 500 clients took hours. This tool reduced that to minutes.

## Author

Jan Grzelinski — finance professional with background in equity compensation (Computershare) and commercial real estate (CBRE, Knight Frank). Building AI/automation tools as a long-term investment in a data & analytics career.
