from __future__ import annotations

import re
import shutil
import pickle
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from app.comparison import _comparison_equal, _formatting_only_difference, compare_dataframes, write_report
from app.config import load_modules, load_regions
from app.custom_fields import (
    add_custom_field_values,
    custom_field_column_aliases,
    custom_field_headers,
)
from app.db import (
    execute_custom_field_queries,
    execute_history_queries,
    execute_module_query,
    execute_organisation_queries,
    execute_party_queries,
    execute_site_queries,
    execute_update_lookup_queries,
)
from app.exporters import export_dataframe
from app.history_fields import (
    add_history_values,
    history_column_aliases,
    history_headers,
    importer_history_field,
)
from app.update_sql import (
    UpdatePlan,
    plan_direct_matter_update,
    plan_history_update,
    plan_client_update,
    plan_lookup_matter_update,
    plan_custom_field_update,
    plan_site_update,
    plan_party_relationships,
)
from app.uploads import read_source_file


BASE_DIR = Path(__file__).resolve().parent.parent
RUNS_DIR = BASE_DIR / "runs"
RUN_ID_PATTERN = re.compile(r"^[A-Za-z0-9_]+$")
MODULES = load_modules(BASE_DIR / "validation_config.yaml")
REGIONS = load_regions(BASE_DIR / "db_regions.yaml")

app = FastAPI(title="Migration Lens")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
templates.env.cache = None


def create_run_dir(org_id: int) -> tuple[str, Path]:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_run_id = f"{timestamp}_{org_id}"

    for suffix in range(100):
        run_id = base_run_id if suffix == 0 else f"{base_run_id}_{suffix:02d}"
        run_dir = RUNS_DIR / run_id
        try:
            run_dir.mkdir(parents=True, exist_ok=False)
            return run_id, run_dir
        except FileExistsError:
            continue

    raise RuntimeError(f"Could not create unique run folder for {base_run_id}")


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "index.html",
        {"modules": MODULES, "regions": REGIONS, "result": None, "error": None},
    )


@app.post("/validate", response_class=HTMLResponse)
async def validate(
    request: Request,
    source_file: UploadFile = File(...),
    org_id: int = Form(...),
    region_id: str = Form(...),
    module_id: str = Form(...),
    export_format: str | None = Form(None),
) -> HTMLResponse:
    if module_id not in MODULES:
        raise HTTPException(status_code=400, detail=f"Unknown module: {module_id}")
    if region_id not in REGIONS:
        raise HTTPException(status_code=400, detail=f"Unknown region: {region_id}")

    module = MODULES[module_id]
    region = REGIONS[region_id]
    selected_export_format = export_format or module.export_format
    run_id, run_dir = create_run_dir(org_id)

    upload_path = run_dir / source_file.filename
    with upload_path.open("wb") as handle:
        shutil.copyfileobj(source_file.file, handle)

    try:
        excel_df = read_source_file(upload_path)
        db_df = execute_module_query(module, region, org_id)
        custom_metadata = pd.DataFrame()
        custom_values = pd.DataFrame()
        custom_options = pd.DataFrame()
        history_definitions = pd.DataFrame()
        history_values = pd.DataFrame()
        custom_headers = custom_field_headers(list(excel_df.columns))
        comparison_aliases = dict(module.column_aliases)
        if custom_headers:
            field_ids = [int(column.removeprefix("cf_")) for column in custom_headers.values()]
            custom_metadata, custom_values, custom_options = execute_custom_field_queries(region, org_id, field_ids)
            db_df = add_custom_field_values(excel_df, db_df, custom_metadata, custom_values, custom_options)
            comparison_aliases.update(custom_field_column_aliases(list(excel_df.columns), custom_metadata))
        if module.identifier == "matters":
            history_headers_found = history_headers(list(excel_df.columns))
            importer_history_ids = {
                history_id
                for column in excel_df.columns
                if (history_mapping := importer_history_field(column))
                for history_id, _ in [history_mapping]
            }
            if history_headers_found:
                history_ids = sorted(
                    {int(column.removeprefix("hid_").split(" ", 1)[0]) for column in history_headers_found.values()}
                    | importer_history_ids
                )
            elif importer_history_ids:
                history_ids = sorted(importer_history_ids)
            else:
                history_ids = []
            if history_ids:
                history_definitions, history_values = execute_history_queries(region, org_id, history_ids)
                db_df = add_history_values(excel_df, db_df, history_definitions, history_values)
                comparison_aliases.update(history_column_aliases(list(excel_df.columns), history_definitions))
        output_name_prefix = f"{module.identifier}_{org_id}"
        db_export = export_dataframe(
            db_df,
            run_dir / f"{output_name_prefix}_db_extract",
            selected_export_format,
        )
        comparison = compare_dataframes(
            excel_df=excel_df,
            db_df=db_df,
            key_columns=module.key_columns,
            optional_key_columns=module.optional_key_columns,
            column_aliases=comparison_aliases,
            header_aliases=module.header_aliases,
            summary_subject=module.summary_subject,
        )
        report_path = run_dir / f"{output_name_prefix}_comparison_report.xlsx"
        write_report(comparison, str(report_path))
        with (run_dir / "update_context.pkl").open("wb") as handle:
            mismatch_records = comparison.sheets["Mismatched_Records"]
            canonical_excel_df = excel_df.copy()
            canonical_excel_df.rename(columns=_header_alias_mapping(excel_df.columns, module.header_aliases), inplace=True)
            pickle.dump(
                {
                    "excel_df": canonical_excel_df,
                    "db_df": db_df,
                    "org_id": org_id,
                    "region_id": region_id,
                    "key_columns": module.key_columns,
                    "comparison_aliases": comparison_aliases,
                    "history_values": history_values,
                    "custom_metadata": custom_metadata,
                    "custom_values": custom_values,
                    "custom_options": custom_options,
                    "mismatch_columns": _mismatch_field_names(comparison),
                    "mismatch_records": mismatch_records,
                    "mismatch_options": _mismatch_options(
                        mismatch_records,
                        module.key_columns,
                        comparison.highlighted_cells["Mismatched_Records"],
                    ),
                },
                handle,
            )
    except Exception as exc:
        return templates.TemplateResponse(
            request,
            "index.html",
            {"modules": MODULES, "regions": REGIONS, "result": None, "error": str(exc)},
            status_code=400,
        )

    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "modules": MODULES,
            "regions": REGIONS,
            "error": None,
            "result": {
                "run_id": run_id,
                "module_id": module.identifier,
                "region_name": region.name,
                "summary": comparison.summary,
                "summary_rows": comparison.summary_rows,
                "db_export_name": db_export.name,
                "report_name": report_path.name,
                "update_fields": _mismatch_fields(comparison.sheets["Mismatched_Records"], comparison.highlighted_cells["Mismatched_Records"]),
                "update_options": _mismatch_options(
                    comparison.sheets["Mismatched_Records"],
                    module.key_columns,
                    comparison.highlighted_cells["Mismatched_Records"],
                ),
            },
        },
    )


@app.post("/prepare-update-plan", response_class=HTMLResponse)
async def prepare_update_plan(
    request: Request,
    run_id: str = Form(...),
    fields: list[str] = Form(default=[]),
) -> HTMLResponse:
    if not RUN_ID_PATTERN.fullmatch(run_id):
        raise HTTPException(status_code=400, detail="Invalid run id")
    context_path = RUNS_DIR / run_id / "update_context.pkl"
    if not context_path.is_file():
        raise HTTPException(status_code=404, detail="Validation context not found")

    with context_path.open("rb") as handle:
        context = pickle.load(handle)
    selected_fields = {value.strip() for value in fields if value.strip()}
    selected_fields = {_update_field_name(value) for value in selected_fields}
    if not selected_fields:
        return templates.TemplateResponse(
            request,
            "index.html",
            {"modules": MODULES, "regions": REGIONS, "result": None, "error": "Select at least one field."},
            status_code=400,
        )

    excel_df = context["excel_df"]
    db_df = context["db_df"]
    org_id = context["org_id"]
    comparison_aliases = context.get("comparison_aliases", {})
    history_values = context.get("history_values", pd.DataFrame())
    custom_metadata = context.get("custom_metadata", pd.DataFrame())
    custom_values = context.get("custom_values", pd.DataFrame())
    custom_options = context.get("custom_options", pd.DataFrame())
    db_field_by_selected_field = {
        selected_field: db_field
        for db_field, selected_field in comparison_aliases.items()
    }
    key_column = context["key_columns"][0]
    party_values = pd.DataFrame()
    organisation_values = pd.DataFrame()
    site_values = pd.DataFrame()
    lookup_values = {}
    if "client" in selected_fields:
        organisation_values = execute_organisation_queries(
            REGIONS[context["region_id"]],
            org_id,
        )
    if "who pays the bill" in selected_fields:
        site_values = execute_site_queries(
            REGIONS[context["region_id"]],
            org_id,
        )
    lookup_fields = {
        "parent mattercode", "family", "status", "category", "type of mark",
        "user", "user2", "attorney", "paralegal",
    }
    if selected_fields & lookup_fields:
        lookup_values = execute_update_lookup_queries(REGIONS[context["region_id"]], org_id)
    selected_party_fields = selected_fields & {"applicant", "applicants", "inventor", "inventors"}
    if selected_party_fields:
        matter_ids = [
            int(value)
            for value in db_df["Case ID"].dropna().tolist()
        ]
        party_values = execute_party_queries(
            REGIONS[context["region_id"]],
            org_id,
            matter_ids,
        )
    db_by_code = db_df.set_index(key_column, drop=False)
    selected_options = _selected_field_options(
        excel_df,
        db_df,
        key_column,
        selected_fields,
        db_field_by_selected_field,
    )
    rows = []
    plan_rows = []
    for option in selected_options:
        matter_code = option["matter_code"]
        field = option["field"]
        excel_row = excel_df[excel_df[key_column].astype(str).str.strip() == matter_code].iloc[0]
        if matter_code not in db_by_code.index:
            continue
        db_row = db_by_code.loc[matter_code]
        matter_id = db_row.get("Case ID")
        if pd.isna(matter_id):
            continue
        excel_field = option.get("excel_column", field)
        db_field = option.get("db_column", db_field_by_selected_field.get(field, field))
        if excel_field not in excel_df.columns or db_field not in db_df.columns:
            rows.append({"Matter": matter_code, "Field": field, "Excel Value": option["excel_value"], "Database Value": option["db_value"], "Action": "Skipped", "Warning": "Field is unavailable in the DB extract."})
            continue
        if field.lower() in {"applicant", "applicants", "inventor", "inventors"}:
            category = "applicant" if field.lower().startswith("applicant") else "inventor"
            party_rows = party_values[
                (party_values["matterid"] == int(matter_id))
                & (party_values["category"].str.lower() == category)
            ]
            current_names = [str(value) for value in party_rows["name"].tolist()]
            excel_names = [
                value.strip()
                for value in str(excel_row[excel_field]).split(";")
                if value.strip()
            ]
            contact_ids = {
                str(row.name).strip().casefold(): int(row.cid)
                for row in party_rows.itertuples(index=False)
                if pd.notna(row.name) and pd.notna(row.cid)
            }
            contact_ids.update({
                str(row.name).strip().casefold(): int(row.contact_id)
                for row in party_values.itertuples(index=False)
                if pd.notna(row.name) and pd.notna(row.contact_id)
            })
            party_plan_rows = plan_party_relationships(
                org_id,
                int(matter_id),
                matter_code,
                "applicants" if category == "applicant" else "inventors",
                current_names,
                excel_names,
                contact_ids,
            )
            plan_rows.extend(party_plan_rows)
            for party_plan_row in party_plan_rows:
                rows.append({
                    "Matter": matter_code,
                    "Field": field,
                    "Excel Value": option["excel_value"],
                    "Database Value": option["db_value"],
                    "Action": party_plan_row.action,
                    "Warning": party_plan_row.warning or "Ready to generate SQL.",
                })
            continue
        if field.lower() == "client":
            excel_client = str(excel_row[field]).strip()
            matching_organisations = organisation_values[
                organisation_values["organisation"].fillna("").astype(str).str.strip().str.casefold()
                == excel_client.casefold()
            ]
            organisation_id = (
                int(matching_organisations.iloc[0]["id"])
                if not matching_organisations.empty
                else None
            )
            plan_row = plan_client_update(
                org_id,
                int(matter_id),
                matter_code,
                db_row[db_field],
                excel_row[field],
                organisation_id,
            )
            plan_rows.append(plan_row)
            rows.append({
                "Matter": matter_code,
                "Field": field,
                "Excel Value": option["excel_value"],
                "Database Value": option["db_value"],
                "Action": plan_row.action,
                "Warning": plan_row.warning or "Ready to generate SQL.",
            })
            continue
        if field.lower() == "who pays the bill":
            excel_site = str(excel_row[field]).strip()
            matching_sites = site_values[
                site_values["sitename"].fillna("").astype(str).str.strip().str.casefold()
                == excel_site.casefold()
            ]
            site_id = int(matching_sites.iloc[0]["id"]) if not matching_sites.empty else None
            plan_row = plan_site_update(
                org_id,
                int(matter_id),
                matter_code,
                db_row[db_field],
                excel_row[field],
                site_id,
            )
            plan_rows.append(plan_row)
            rows.append({
                "Matter": matter_code,
                "Field": field,
                "Excel Value": option["excel_value"],
                "Database Value": option["db_value"],
                "Action": plan_row.action,
                "Warning": plan_row.warning or "Ready to generate SQL.",
            })
            continue
        lookup_field = field.lower()
        lookup_config = {
            "parent mattercode": ("parent_matter_id", "mattercode", "matters", "Matter code"),
            "family": ("family", "family", "families", "Family"),
            "status": ("status", "description", "dropdowns", "Status"),
            "category": ("category", "description", "dropdowns", "Category"),
            "type of mark": ("typeofmark", "description", "dropdowns", "Type of mark"),
            "user": ("user", "username", "users", "User"),
            "attorney": ("user", "username", "users", "User"),
            "user2": ("user2", "username", "users", "User"),
            "paralegal": ("user2", "username", "users", "User"),
        }.get(lookup_field)
        if lookup_config is not None:
            target_column, lookup_column, source_name, label = lookup_config
            lookup_df = lookup_values.get(source_name, pd.DataFrame())
            if source_name == "dropdowns":
                type_filter = {
                    "status": "matter_status",
                    "category": "matter_categories",
                    "type of mark": "typeofmark",
                }.get(lookup_field, "")
                lookup_df = lookup_df[
                    lookup_df["type"].fillna("").str.lower().str.contains(type_filter)
                ]
            matching = lookup_df[
                lookup_df[lookup_column].fillna("").astype(str).str.strip().str.casefold()
                == str(excel_row[field]).strip().casefold()
            ]
            lookup_id = int(matching.iloc[0]["oldid" if source_name == "dropdowns" else "id"]) if not matching.empty else None
            plan_row = plan_lookup_matter_update(
                org_id, int(matter_id), matter_code, field, target_column,
                db_row[db_field], excel_row[field], lookup_id, label,
            )
            plan_rows.append(plan_row)
            rows.append({
                "Matter": matter_code, "Field": field,
                "Excel Value": option["excel_value"],
                "Database Value": option["db_value"],
                "Action": plan_row.action,
                "Warning": plan_row.warning or "Ready to generate SQL.",
            })
            continue
        if field.lower().startswith("cf_"):
            field_id = int(field.lower().removeprefix("cf_"))
            metadata_rows = custom_metadata[custom_metadata["id"] == field_id]
            if metadata_rows.empty:
                plan_row = plan_custom_field_update(
                    org_id, int(matter_id), matter_code, field_id, "", db_row[db_field], excel_row[field], False,
                    missing_options=[f"Custom field {field_id} not found"],
                )
            else:
                field_type = str(metadata_rows.iloc[0]["field_type"])
                value_rows = custom_values[
                    (custom_values["field_id"] == field_id)
                    & (custom_values["entity_id"] == int(matter_id))
                ]
                option_ids = []
                missing_options = []
                if field_type.lower() in {"select", "multiselect"}:
                    labels = [part.strip() for part in str(excel_row[field]).split(",") if part.strip()]
                    for label_value in labels:
                        matching_option = custom_options[
                            (custom_options["field_id"] == field_id)
                            & (custom_options["text"].fillna("").astype(str).str.strip().str.casefold() == label_value.casefold())
                        ]
                        if matching_option.empty:
                            missing_options.append(label_value)
                        else:
                            option_ids.append(int(matching_option.iloc[0]["id"]))
                plan_row = plan_custom_field_update(
                    org_id, int(matter_id), matter_code, field_id, field_type,
                    db_row[db_field], excel_row[field], not value_rows.empty,
                    option_ids, missing_options,
                )
            plan_rows.append(plan_row)
            rows.append({
                "Matter": matter_code, "Field": field,
                "Excel Value": option["excel_value"],
                "Database Value": option["db_value"],
                "Action": plan_row.action,
                "Warning": plan_row.warning or "Ready to generate SQL.",
            })
            continue
        history_match = re.fullmatch(r"hid_(\d+) (date|country|number)", field)
        importer_history_match = importer_history_field(field)
        if history_match is not None:
            history_id = int(history_match.group(1))
            value_type = history_match.group(2)
        elif importer_history_match is not None:
            history_id, value_type = importer_history_match
        else:
            history_id = None
            value_type = ""
        if history_id is not None:
            existing_history = history_values[
                (history_values["matterid"] == int(matter_id))
                & (history_values["hid"] == history_id)
            ]
            plan_row = plan_history_update(
                org_id, int(matter_id), matter_code, history_id, value_type,
                    db_row[db_field], excel_row[excel_field], not existing_history.empty,
            )
        else:
            plan_row = plan_direct_matter_update(
                org_id, int(matter_id), matter_code, field,
                db_row[db_field], excel_row[excel_field],
            )
        plan_rows.append(plan_row)
        rows.append({"Matter": matter_code, "Field": field, "Excel Value": option["excel_value"], "Database Value": option["db_value"], "Action": plan_row.action, "Warning": plan_row.warning or "Ready to generate SQL."})

    plan = UpdatePlan(plan_rows)
    sql_path = RUNS_DIR / run_id / "selected_update_plan.sql"
    sql_path.write_text(plan.sql, encoding="utf-8")
    report_path = next((RUNS_DIR / run_id).glob("*_comparison_report.xlsx"), None)
    if report_path is None:
        raise HTTPException(status_code=404, detail="Comparison report not found")
    return templates.TemplateResponse(
        request,
        "update_preview.html",
        {
            "update_preview": rows,
            "run_id": run_id,
            "region_name": REGIONS[context["region_id"]].name,
            "selected_fields": sorted(selected_fields),
            "sql_name": sql_path.name,
            "report_name": report_path.name,
        },
    )


def _update_cases(mismatch_records: pd.DataFrame, key_columns: list[str]) -> list[dict[str, str]]:
    if mismatch_records.empty:
        return []
    return [
        {
            "value": str(row[key_columns[0]]),
            "label": " / ".join(str(row[column]) for column in key_columns),
        }
        for _, row in mismatch_records.drop_duplicates(subset=key_columns).iterrows()
    ]


def _mismatch_options(
    mismatch_records: pd.DataFrame,
    key_columns: list[str],
    highlights: dict[tuple[int, int], str],
) -> list[dict[str, str]]:
    options: list[dict[str, str]] = []
    option_number = 0
    for row_index, (_, row) in enumerate(mismatch_records.iterrows(), start=2):
        matter_code = str(row[key_columns[0]])
        for column in mismatch_records.columns:
            if not column.endswith(" (Excel)"):
                continue
            report_field = column.removesuffix(" (Excel)")
            field = _update_field_name(report_field)
            database_value = row[f"{report_field} (Database)"]
            excel_value = row[column]
            if _formatting_only_difference(excel_value, database_value):
                continue
            options.append({
                "id": f"change-{option_number}",
                "matter_code": matter_code,
                "field": field,
                "excel_value": str(excel_value),
                "db_value": str(database_value),
            })
            option_number += 1
    return options


def _mismatch_fields(
    mismatch_records: pd.DataFrame,
    highlights: dict[tuple[int, int], str],
) -> list[str]:
    fields: list[str] = []
    for column in mismatch_records.columns:
        if not column.endswith(" (Excel)"):
            continue
        field = column.removesuffix(" (Excel)")
        if any(
            not _formatting_only_difference(
                row[column], row[f"{field} (Database)"]
            )
            for _, row in mismatch_records.iterrows()
        ):
            normalized_field = _update_field_name(field)
            if normalized_field not in fields:
                fields.append(normalized_field)
    return fields


def _selected_field_options(
    excel_df: pd.DataFrame,
    db_df: pd.DataFrame,
    key_column: str,
    selected_fields: set[str],
    db_field_by_selected_field: dict[str, str],
) -> list[dict[str, str]]:
    db_by_code = db_df.set_index(key_column, drop=False)
    options: list[dict[str, str]] = []
    option_number = 0
    for _, excel_row in excel_df.iterrows():
        matter_code = str(excel_row[key_column]).strip()
        if matter_code not in db_by_code.index:
            continue
        db_row = db_by_code.loc[matter_code]
        for field in selected_fields:
            excel_field = next(
                (column for column in excel_df.columns if _update_field_name(column) == field),
                None,
            )
            db_field = db_field_by_selected_field.get(field, field)
            if db_field not in db_df.columns:
                db_field = next(
                    (column for column in db_df.columns if _update_field_name(column) == field),
                    None,
                )
            if excel_field is None or db_field is None:
                continue
            excel_value = excel_row[excel_field]
            db_value = db_row[db_field]
            if _comparison_equal(excel_value, db_value, field):
                continue
            if _formatting_only_difference(excel_value, db_value):
                continue
            options.append({
                "id": f"selected-{option_number}",
                "matter_code": matter_code,
                "field": field,
                "excel_column": excel_field,
                "db_column": db_field,
                "excel_value": str(excel_value),
                "db_value": str(db_value),
            })
            option_number += 1
    return options


def _whitespace_only_update_difference(left: object, right: object) -> bool:
    return (
        isinstance(left, str)
        and isinstance(right, str)
        and left != right
        and left.strip() == right.strip()
    )


def _update_field_name(field: str) -> str:
    normalized = str(field).strip().lower()
    if normalized in {"applicant", "applicants"}:
        return "applicants"
    if normalized in {"inventor", "inventors"}:
        return "inventors"
    return str(field).strip()


def _header_alias_mapping(columns, header_aliases):
    aliases = {}
    for canonical, configured_aliases in header_aliases.items():
        for value in [canonical, *configured_aliases]:
            normalized = "".join(character.lower() for character in str(value) if character.isalnum())
            aliases[normalized] = canonical
    return {
        column: aliases["".join(character.lower() for character in str(column) if character.isalnum())]
        for column in columns
        if "".join(character.lower() for character in str(column) if character.isalnum()) in aliases
    }


@app.get("/download/{run_id}/{filename}")
async def download(run_id: str, filename: str) -> FileResponse:
    if not RUN_ID_PATTERN.fullmatch(run_id):
        raise HTTPException(status_code=400, detail="Invalid run id")
    path = RUNS_DIR / run_id / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, filename=filename)


def _mismatch_field_names(comparison) -> list[str]:
    mismatch_df = comparison.sheets["Mismatched_Records"]
    fields = []
    for column in mismatch_df.columns:
        if column.endswith(" (Excel)"):
            fields.append(column.removesuffix(" (Excel)"))
    return fields


@app.post("/cleanup")
async def cleanup() -> dict[str, int]:
    removed = 0
    if RUNS_DIR.exists():
        for child in RUNS_DIR.iterdir():
            if child.is_dir():
                shutil.rmtree(child)
                removed += 1
    return {"removed_run_folders": removed}
