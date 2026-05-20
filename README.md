# cwm_monthly_report.skill

Codex skill source for generating CyweeMotion monthly performance appraisal drafts from Raven project tracking spreadsheets.

## Skill

- Skill folder: `cwm-monthly-report/`
- Main script: `cwm-monthly-report/scripts/generate_monthly_appraisal.py`

## Basic Workflow

```powershell
python cwm-monthly-report/scripts/generate_monthly_appraisal.py extract --workbook <project-tracking.xlsx> --template <appraisal-template.xlsx> --output-dir <output-dir>
python cwm-monthly-report/scripts/generate_monthly_appraisal.py fill --template <appraisal-template.xlsx> --draft-json <output-dir>/draft.json --output-dir <output-dir>
```

The script writes `summary.json`, `summary.md`, and a filled workbook copy. It never overwrites the original appraisal template.
