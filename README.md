# UNZA Health Insurance & NHIMA Survey Dashboard

An Excel dashboard summarising the Ridgeway Campus student survey on health-insurance knowledge and NHIMA awareness (**358 responses**; August–September 2026).

## Files

- `UNZA_Health_Insurance_Dashboard.xlsx` — the finished dashboard and supporting analysis workbook.
- `previews/dashboard.pdf` — printable dashboard preview.
- `previews/dashboard.png` — dashboard screenshot.
- `analyze.py`, `build_dashboard.py` — Python analysis and workbook-generation scripts.

## Use the dashboard

Open the workbook at **Overview**. Use the four yellow dropdowns (sex, year, programme, and insurance-registration status) to explore a cohort. Supporting tabs contain cohort statistics, item-level results, full-sample association tests, filterable responses, the scoring key, and methodology notes.

## Confidentiality

This repository is **private**. The Excel workbook contains respondent-level records, demographic categories, survey answers, and timestamps. Do not make the repository public or distribute the workbook without appropriate data-protection approval. The separate original response workbook is intentionally not included.

## Rebuild locally (optional)

Requires Python 3, `openpyxl`, `numpy`, `scipy`, and `XlsxWriter`. Obtain the source responses file through an authorised channel and keep it outside Git. In the project directory, run:

```bash
python -m pip install openpyxl numpy scipy XlsxWriter
UNZA_SURVEY_FILE=/path/to/authorised_responses.xlsx python analyze.py
UNZA_SURVEY_FILE=/path/to/authorised_responses.xlsx python build_dashboard.py
```

The provisional instrument-derived answer key, especially policy-specific NHIMA items, must be validated against the study protocol and current benefit package before academic or policy use. See the workbook's **Notes** and **Codebook** sheets for scoring, denominator decisions, caveats, and details on the reported statistical tests.

## Sources and rights

The workbook was prepared from a survey response export and an accompanying analysis report. The source files are not republished here. No licence is asserted for the underlying survey data; confirm applicable permissions before reuse.
