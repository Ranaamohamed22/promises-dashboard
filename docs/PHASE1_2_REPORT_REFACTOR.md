# Phase 1.2: low-risk analytics refactoring

## Change rationale
The first safe refactor extracts three unchanged functions from
snb_performance.py into analytics_core/snb_performance_utils.py:
- _account_key preserves account IDs and leading zeros.
- _amounts validates payment values and normalizes separators.
- _creation_dates parses upload dates, including Excel serials.

snb_performance.py re-imports these under the SAME names, so existing
internal callers and public report entry points remain available. The original
function bodies were moved without any formula or KPI changes.

app.py, snb_comprehensive.py, run.py, requirements.txt, and existing
encrypted data files were NOT modified.

## Test strategy
Run existing SNB performance + comprehensive regression tests and new:
- parity/import tests for the moved helper functions,
- leading-zero account identifiers,
- amount formatting, missing and invalid inputs,
- Excel serial/date parsing,
- a synthetic Excel end-to-end check of the daily promises report.

These are synthetic-only; no real customer files are used.

## Verification
GitHub Actions in .github/workflows/analytics-quality.yml runs tests for
branches matching phase1/**. Do not merge if any check fails.

## Remaining work / limits
- Full app.py UI behavior and performance has NOT been refactored or tested.
- SNB and STC report logic embedded inside app.py needs targeted tests and
  reference fixtures before any extraction.
- Real-life workbook parity and performance benchmarks have not been proven.
- Any future main merge should follow review and staging verification.
