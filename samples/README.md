# samples/ (developer dataset)

Reference inputs and outputs used to develop and check the CPT extraction.
Not needed by end users and not part of a deployment.

| File | Role |
|---|---|
| `CPT MAPPING TEMPLATE (CSV FILE) - Sheet1.csv` | Target Maxicare template (header only) |
| `* MSMC *.pdf`, `Radiology Prices_MSMC *.xlsx` | MSMC provider price lists (inputs) |
| `PRICELIST 2026 (APPROVED as of June222026) *.csv` | Raw OCR-style export (noisy, inputs) |
| `msmc_merged_cpt.xlsx` | MSMC files merged in CPT mode |
| `test_cpt_mapping_output.csv` | Template-format output for MSMC Radiology (older run) |
| `880000000012226_ST. LUKE'S ... MATCHED FILE.xlsx` | Finished St. Luke's BGC mapping (expected end result) |
| `st_lukes_cpt_cleaned.xlsx` | St. Luke's after CPT cleaning |
