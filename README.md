# E2B_R3 XML Triage Application

Baseline: **v1.5.0-listedness-clarified**

## What changed (non-logic)
- Added a comment notifier **"Molecule name differ"** when a non-Celix company tag is present in the drug name (e.g., `Abiraterone [JANSSEN]`). This does **not** alter validity, reportability, or listedness logic; it only appends to the **Comment** field.
- Added `requirements.txt` for reproducible environment.

## How the notifier works
- Scans bracketed tags `[...]`, trailing tags after `--`, or phrases like `by <company>` in the raw drug name.
- Ignores common formulation/strength words (e.g., `5 mg`, `tablet`, `capsule`).
- If the tag mentions **Celix**, it's ignored; if it mentions a competitor or looks like a company string, it flags the comment.

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- The core logic for parsing, validity assessment, listedness, and reportability remains **unchanged**.
