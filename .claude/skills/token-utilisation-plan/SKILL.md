---
name: token-utilisation-plan
description: Build the customer-facing "Token Utilisation Plan" one-slide PowerPoint (plan vs actual, traffic light, use-case roadmap) for a given Genesys org from manually exported Tableau data (TokenUsage and AI Dashboard - Meters). Use when the user asks for a token utilisation plan, token consumption vs plan report, or an EBR token slide for a customer/org.
---

# Token Utilisation Plan

Generates one native, editable 16:9 PowerPoint slide following the layout of `assets/template.pptx`
(reference picture: `assets/reference-layout.png`) and the business rules of the "Token Utilisation
Plan | Master Prompt v3". Data comes from Tableau exports filtered by the org the user names.

Tableau views (the user's, need login):
- TokenUsage / TokenUsageTrend: https://10az.online.tableau.com/#/site/onegenesys/views/TokenUsage/TokenUsageTrend?:iid=1
- AIDashboard-Meters / ExecutiveSummary: https://10az.online.tableau.com/#/site/onegenesys/views/AIDashboard-Meters/ExecutiveSummary?%3Aiid=1

## Hard rules (from the Master Prompt - never break)
1. Commitment is **monthly**. Never calculate or show "Tokens Remaining".
2. Status compares **Actual vs Planned** of the current reporting month, never Actual vs commitment.
   GREEN: variance >= -5% | YELLOW: -10% <= variance < -5% | RED: variance < -10% (override only if the user gives other thresholds, via `thresholds`).
3. Utilisation % = Actual current month / Monthly commitment. Variance = Actual - Planned; % = (Actual - Planned) / Planned.
4. Never cumulative. Future months are planned only, never actual.
5. **Do not invent** customer facts, milestones, use cases or numbers. If a value is not in the exports or given by the user, ask, or leave the optional block out (the script then simplifies the visual).

## Workflow
1. **Ask the user** for: org name, current reporting month, and the manual exports of both Tableau views
   filtered by that org (Download > Data or Crosstab, CSV or XLSX; they can also paste data). Do not ask for Tableau passwords or tokens.
2. **Inspect the files** to see their real structure:
   `python .claude/skills/token-utilisation-plan/scripts/inspect_exports.py <files> --org "<org>"`
   Confirm the org appears and the filter really isolates it. Which column holds which metric is NOT known in advance:
   map columns by what the data shows, and ask the user whenever a mapping is ambiguous (e.g. plan vs commitment vs actual).
3. **Write the input JSON** (see schema below) in the scratchpad/outputs folder. Required: customer, monthly_commitment,
   contract_end, reporting_month, and actual + planned for the reporting month. Everything else is optional and must come from the data
   or the user. If the exports lack the planned consumption, the contract end date, the commitment, etc., stop and ask.
4. **Write `focus_areas` (3-4)** only from evidence in the data (e.g. the size of the gap, which months are below plan). Keep each
   `text` under ~90 characters. Do not promise customer actions that were not provided.
5. **Build**: `python .claude/skills/token-utilisation-plan/scripts/build_report.py input.json output.pptx`
   (needs `pip install python-pptx lxml`). The script prints the calculated variance, status, utilisation and months remaining.
6. **Check** the 12-point quality list below, ideally by rendering the slide (LibreOffice Impress -> PDF -> PNG) and looking at it.
7. **Deliver** the .pptx saved inside the repository (e.g. `reports/<org>/<YYYY-MM>_token_utilisation_plan.pptx`) so it can be shared/exported,
   plus a short summary of the numbers and any assumptions or missing data.

## Input JSON schema
```json
{
  "customer": "Org name",
  "monthly_commitment": 8000000,
  "contract_end": "2027-11-30",
  "reporting_month": "2026-09",
  "actual":  {"2026-08": 4900000, "2026-09": 5300000},
  "planned": {"2026-09": 6500000, "2026-10": 6800000},
  "current_actual": null,
  "current_planned": null,
  "thresholds": {"green": -0.05, "yellow": -0.10},
  "use_cases": [{"name": "STA", "description": "optional", "milestones": [{"label": "Pilot", "start": "2026-07", "end": "2026-09"}]}],
  "focus_areas": [{"title": "Close the short-term gap", "text": "..."}],
  "source_note": "optional footer text"
}
```
Values are raw token counts. Months are `YYYY-MM`. `current_actual/current_planned` are optional overrides; otherwise they are read from the
series at `reporting_month`. Up to 4 use cases are drawn. Without `use_cases` or `focus_areas` that block is omitted.
A worked example (illustrative data, not real) is in `examples/acme_sample.json` -> `examples/acme_sample.pptx`.

## Final quality check (from the Master Prompt)
Automatic (the script guarantees it): no "Tokens Remaining"; actual is never cumulative; future months are planned only; current month
highlighted; variance, status, utilisation and months remaining computed from the rules above; status based on Actual vs Plan.
Manual - verify each time: monthly commitment clearly shown as per month; historical actual vs planned visually distinct; contract
end date and months remaining correct for the data; use cases/milestones are exactly those supplied (not invented); the visual links the plan to the use cases.

## Pending / roadmap
- Data is read from manual exports for now. A Tableau Cloud MCP connection is expected later; when available, replace step 1-2 with queries
  of the two views filtered by org, keeping the same JSON schema and build script.
- The template slide is a flat picture, so the layout is rebuilt natively in `scripts/build_report.py` (theme from `assets/template.pptx`).
  Icons from the reference picture are not reproduced.
