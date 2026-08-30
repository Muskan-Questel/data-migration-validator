# Python Migration Data Comparison Tool

FastAPI app for validating migration results by comparing a single-sheet Excel upload against a configured MySQL query result. The uploaded Excel file is the source of truth.

Uploaded source files may be `.csv`, `.xlsx`, or `.xlsm`. Legacy `.xls` files should be opened in Excel and saved as `.xlsx` before upload.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Store region-specific DB connection settings in a local `.env` file. Do not commit or share that file.

Run the app:

```powershell
python -m uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000.

## Configuration

Modules are loaded from `validation_config.yaml`. The current config points to
`validation_modules/`, where each module has its own YAML file. Each module
defines:

- `name`
- `query_template` using `:org_id`
- `key_columns`
- optional `column_aliases` mapping DB column names to Excel column names
- optional `header_aliases` mapping each query/header name to accepted uploaded sheet names
- `export_format`, either `xlsx` or `csv`

Users cannot enter SQL in the UI. Only configured module queries run.

## Region Configuration

Regions live in `db_regions.yaml`. Each region defines:

- display `name`
- `env_prefix`, used to resolve connection details from `.env`

DB hosts, ports, usernames, database names, and passwords are intentionally not stored in source files.

The app connects directly to the selected region's MySQL host. Connect to the required VPN before running validation.

## Outputs

Each validation run creates a folder under `runs/` with:

- folder name in `YYYYMMDD_HHMMSS_orgid` format
- DB extract as `.xlsx` or `.csv`
- final `comparison_report.xlsx`

The report includes `Summary`, `Column_Mapping`, `Mismatched_Records`, `Extra_In_Database`, and `Extra_In_Excel`.
`Column_Mapping` lists every uploaded Excel column in Column A and every DB column in Column B, pairing mapped columns on the same row and leaving the opposite cell blank for unmapped columns.
Configured `header_aliases` are matched case-insensitively and ignore spaces, hyphens, and punctuation, so `matter code`, `Matter-Code`, and `mattercode` can all map to the same query header.
`Mismatched_Records` contains complete Excel and database row pairs for matched keys,
with only the mismatched value cells highlighted.
