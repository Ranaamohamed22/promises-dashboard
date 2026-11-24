# Phase 1: Analytics quality foundation

This change adds a safety net before restructuring the dashboard.

## Scope
- Automated syntax checks for existing Python entry points.
- Synthetic-workbook regression tests for SNB comprehensive treatment categories,
  Writeoff display, invalid buckets, and repeat-upload column stability.
- Synthetic-data tests for SNB performance payment aggregation, account ID
  matching, missing columns, and nonnumeric amounts.
- GitHub Actions checks on the feature branch and pull requests.

## Explicit non-goals
- No changes to existing calculations or KPI business definitions.
- No large rewrite of app.py in this foundational step.
- No access to or publication of real banking records.
- No automatic changes to the deployed main branch.
- These tests do not yet cover every module or data shape.

## Local verification (Python 3.11)
1. python -m pip install -r requirements.txt
2. python -m compileall -q app.py run.py snb_comprehensive.py snb_performance.py
3. python -m unittest discover -s tests -p "test_*.py" -v

Tests create workbooks in memory and do not read encrypted files from data/.

## Next safe steps
1. Add verified sanitized reference-output test fixtures for other reports.
2. Extract cohesive functions from app.py with tests before and after each change.
3. Benchmark latency and memory before optimizing.
4. Audit upload handling, credentials and deployment security before adding AI.
