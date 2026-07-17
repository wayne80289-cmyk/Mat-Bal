# Mat-Bal

Material balance / stock allocation tooling. A set of standalone Python CLI scripts
that read Excel (`.xls`/`.xlsx`) and PDF source files and write styled Excel reports.
There is **no web server, service, or UI** — every entrypoint is a script run with
`python <script>.py`.

## Cursor Cloud specific instructions

### Environment / running
- Python deps live in a virtualenv at `.venv` (the system Python is PEP 668
  "externally managed", so packages must be installed into a venv, not globally).
  The startup update script creates `.venv` and installs `requirements.txt`.
  Creating the venv requires the `python3-venv` apt package (part of the VM
  snapshot, not the update script).
- Activate before running anything: `source .venv/bin/activate`.
- Outputs are written to `Output/` (gitignored) or back into the data folders.

### Runnable entrypoints and their data needs
- `extract_allocated_stock.py` — reads `History Balance/<month>/*.xlsx` (committed)
  and writes `History Balance/Stock-Allocated-Material-Code-All-Months.xlsx`.
  Runs out of the box.
- `export_matching_rules_excel.py` — self-contained; writes
  `History Balance/Stock-Material-Code-Matching-Rules.xlsx`. Runs out of the box.
- `material_balance.py` — needs a `Sample Balance/` folder containing the
  "All Customer review ... .xlsx" workbook. The repo ships that workbook at the
  repo root; `material_balance.py` hardcodes the name
  `Sample Balance/All Customer review Jun '2026 review 20.06.2026.xlsx`, so stage it:
  `mkdir -p "Sample Balance" && cp "All Customer review Jun '2026 review 23.06.2026.xlsx" "Sample Balance/All Customer review Jun '2026 review 20.06.2026.xlsx"`.
  Also reads `Stock/MS004-*.xls` and `Froecast/*` (both committed). `Sample Balance/`
  is gitignored (it is monthly input data, like `Output/`).
- `material_order_balance.py` (uses `data_loaders.py`) — same `Sample Balance/`
  staging as above, plus it needs an **MP008 open-PO** source that is not shipped in
  a compatible form: `data_loaders.load_mp008()` looks for a standalone `*MP008*`
  export in `Sample Balance/`, else falls back to a workbook sheet literally named
  `MP008 xxxxxxxx`. The shipped workbook's sheet is dated (e.g. `MP008 22.06.26`), so
  this pipeline needs a real MP008 export dropped into `Sample Balance/` to run.

### Lint / test
- No configured linter or automated test suite exists. Use a compile check as a
  smoke test: `python -m py_compile *.py`.
