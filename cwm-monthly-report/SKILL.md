---
name: cwm-monthly-report
description: Generate CyweeMotion monthly performance appraisal drafts from Raven project tracking Excel workbooks. Use when Codex is given a project statistics workbook with a Raven worksheet and a CyweeMotion appraisal workbook, and needs to extract the appraisal month work items, summarize them, score them, and save a filled appraisal copy.
---

# CWM Monthly Report

## Workflow

Use this skill to fill a CyweeMotion monthly performance appraisal from Raven's project tracking workbook.

1. Run extraction:

```powershell
python cwm-monthly-report/scripts/generate_monthly_appraisal.py extract --workbook <project-tracking.xlsx> --template <appraisal-template.xlsx> --output-dir <output-dir>
```

For a custom appraisal window, pass explicit dates:

```powershell
python cwm-monthly-report/scripts/generate_monthly_appraisal.py extract --workbook <project-tracking.xlsx> --template <appraisal-template.xlsx> --output-dir <output-dir> --start-date YYYY-MM-DD --end-date YYYY-MM-DD
```

2. Read `<output-dir>/summary.json` and `<output-dir>/summary.md`.

3. Draft concise Chinese appraisal text from the extracted work items:

- `main_description`: 4-6 numbered paragraphs for `F5`, grouped by customer/project coverage, sleep issues, sport features, sensor/power/configuration, UI-vs-algorithm output, cross-team follow-up, and unresolved work.
- `execution_description`: 3-4 numbered paragraphs for `F9`, focused on ownership, follow-through, problem closure, and multi-project execution.
- `collaboration_description`: 3-4 numbered paragraphs for `F10`, focused on customer communication, algorithm/PM/QA coordination, information handoff, and improvement points.
- Start each paragraph with two full-width spaces followed by `1.`/`2.`/`3.` style numbering, such as `　　1.`. Put each numbered paragraph on its own line.
- Keep `OPEN` or blank-status items as "持续跟进/推动验证/等待反馈"; do not describe them as completed or closed.
- Use the score recommendations from `summary.json` unless the user explicitly asks to adjust them.

4. Save a `draft.json` with this exact shape:

```json
{
  "main_description": "...",
  "main_score": 88,
  "execution_description": "...",
  "execution_score": 5,
  "collaboration_description": "...",
  "collaboration_score": 4
}
```

5. Fill a copy of the appraisal workbook:

```powershell
python cwm-monthly-report/scripts/generate_monthly_appraisal.py fill --template <appraisal-template.xlsx> --draft-json <output-dir>/draft.json --output-dir <output-dir>
```

Return links to the filled `.xlsx` and `summary.md`.

## Extraction Rules

- Read only the `Raven` worksheet from the project tracking workbook.
- If the user gives an explicit date window, use `--start-date YYYY-MM-DD --end-date YYYY-MM-DD`.
- Otherwise infer the appraisal month from the template's `D2:F2` merged title text, such as `考核时间：2026 年 3 月`. If inference fails, rerun extraction with `--period YYYY-MM`.
- Use the natural month as the default period, from the first day through the last day of that month.
- Recognize these columns by normalized header text: `Project Name`, `Function`, `Algo Ver.`, `Quantity`, `Issues`, `Status`, `Owners`, `Completion Date`, `Notes`.
- Carry down blank project context columns so issue rows inherit the project name, function, and algorithm version from merged or grouped rows above.
- Include an item when either `Completion Date` falls inside the month or any date mentioned in `Issues` falls inside the month.
- Do not filter by owner. Preserve `Owners` in the summary for review.
- Include `OPEN` and blank-status items, but mark them as in progress in the summary.
- Date parsing should tolerate Excel dates, `YYYY.M.D`, `YYYY/M/D`, `YYYY-MM-DD`, `YYYYMMDD`, and recoverable text such as `2026.4.15=7` or `2026402`.

## Writing Rules

The fill command must save a new workbook and never overwrite the template.

Write only these appraisal cells:

- `F5`: personal completion description
- `G5`: personal score
- `F9`: execution description
- `G9`: execution score
- `F10`: collaboration description
- `G10`: collaboration score

For `F5`, `F9`, and `F10`, keep text wrapped and apply visible first-line indentation by preserving the leading two full-width spaces before each numbered paragraph. Use numbered paragraphs so the appraisal is easier to scan in Excel.

Preserve existing formulas such as `G6`, `G11`, and `G12`.

## Scoring Rules

Use the script's computed recommendations:

```text
weighted_points = issue_count + 1.5 * project_count + 0.5 * closed_count + 0.75 * open_count
<15 => 85
15-24.99 => 86
25-34.99 => 87
35-44.99 => 88
>=45 => 89

execution_score = clamp(3 + floor(weighted_points / 18), 3, 6)
collaboration_score = clamp(3 + floor(project_count / 6), 3, 4)
```

## Validation

After implementation or changes, validate with:

```powershell
python C:\Users\ywwq0\.codex\skills\.system\skill-creator\scripts\quick_validate.py cwm-monthly-report
```

For an end-to-end smoke test, run extraction on the provided Raven project workbook and March 2026 appraisal workbook, create a draft JSON from the summary, run fill, then inspect `F5/G5/F9/G9/F10/G10` and confirm the source template is unchanged.
