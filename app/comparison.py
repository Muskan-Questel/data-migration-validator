from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import PatternFill


@dataclass(frozen=True)
class ComparisonResult:
    sheets: dict[str, pd.DataFrame]
    summary: dict[str, Any]
    highlighted_cells: dict[str, set[tuple[int, int]]]


def compare_dataframes(
    excel_df: pd.DataFrame,
    db_df: pd.DataFrame,
    key_columns: list[str],
    column_aliases: dict[str, str] | None = None,
    header_aliases: dict[str, list[str]] | None = None,
) -> ComparisonResult:
    """Compare uploaded Excel data to a DB extract using Excel columns as contract."""
    aliases = column_aliases or {}
    excel_df = excel_df.copy()
    original_db_columns = list(db_df.columns)
    db_df = db_df.rename(columns=aliases).copy()
    excel_header_mapping = _excel_header_mapping(list(excel_df.columns), header_aliases or {})
    excel_df = excel_df.rename(columns=excel_header_mapping)

    excel_columns = list(excel_df.columns)
    db_columns = list(db_df.columns)

    missing_keys = [
        key
        for key in key_columns
        if key not in excel_columns or key not in db_columns
    ]
    if missing_keys:
        raise ValueError(
            "Configured key columns must exist in both Excel and DB extract: "
            + ", ".join(missing_keys)
        )

    missing_db_columns = [column for column in excel_columns if column not in db_columns]
    compare_columns = [column for column in excel_columns if column not in missing_db_columns]

    excel_dup_mask = excel_df.duplicated(subset=key_columns, keep=False)
    db_dup_mask = db_df.duplicated(subset=key_columns, keep=False)

    excel_unique = excel_df.loc[~excel_dup_mask].copy()
    db_unique = db_df.loc[~db_dup_mask].copy()

    excel_by_key = _rows_by_key(excel_unique, key_columns)
    db_by_key = _rows_by_key(db_unique, key_columns)

    excel_keys = set(excel_by_key)
    db_keys = set(db_by_key)
    matched_keys = sorted(excel_keys & db_keys, key=str)

    extra_in_excel = _keys_for_records(excel_by_key, excel_keys - db_keys, key_columns)
    extra_in_database = _keys_for_records(db_by_key, db_keys - excel_keys, key_columns)
    mismatched_records, highlighted_cells = _mismatched_records(
        excel_by_key,
        db_by_key,
        matched_keys,
        key_columns,
        compare_columns,
    )

    summary = {
        "excel_rows": len(excel_df),
        "db_rows": len(db_df),
        "compared_columns": len(compare_columns),
        "matched_rows": len(matched_keys),
        "mismatched_mattercodes": mismatched_records[key_columns].drop_duplicates().shape[0]
        if not mismatched_records.empty
        else 0,
        "extra_mattercodes_in_database": len(extra_in_database),
        "extra_mattercodes_in_excel": len(extra_in_excel),
        "duplicate_excel_rows": int(excel_dup_mask.sum()),
        "duplicate_db_rows": int(db_dup_mask.sum()),
        "missing_db_columns": len(missing_db_columns),
    }
    summary_sheet = pd.DataFrame(
        [{"Metric": _label(key), "Value": value} for key, value in summary.items()]
    )
    column_mapping_sheet = _column_mapping_sheet(
        excel_columns,
        original_db_columns,
        aliases,
        excel_header_mapping,
    )

    return ComparisonResult(
        sheets={
            "Summary": summary_sheet,
            "Column_Mapping": column_mapping_sheet,
            "Mismatched_Records": mismatched_records,
            "Extra_In_Database": extra_in_database,
            "Extra_In_Excel": extra_in_excel,
        },
        summary=summary,
        highlighted_cells={"Mismatched_Records": highlighted_cells},
    )


def write_report(result: ComparisonResult, output_path: str) -> None:
    mismatch_fill = PatternFill(
        fill_type="solid",
        start_color="FFF2CC",
        end_color="FFF2CC",
    )
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df in result.sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.book[sheet_name]
            worksheet.freeze_panes = "A2"
            for column_cells in worksheet.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 12),
                    45,
                )
        for sheet_name, cells in result.highlighted_cells.items():
            worksheet = writer.book[sheet_name]
            for row_number, column_number in cells:
                worksheet.cell(row=row_number, column=column_number).fill = mismatch_fill


def _mismatched_records(
    excel_by_key: dict[tuple[Any, ...], pd.Series],
    db_by_key: dict[tuple[Any, ...], pd.Series],
    matched_keys: list[tuple[Any, ...]],
    key_columns: list[str],
    compare_columns: list[str],
) -> tuple[pd.DataFrame, set[tuple[int, int]]]:
    rows: list[dict[str, Any]] = []
    highlighted_cells: set[tuple[int, int]] = set()
    columns = [
        *key_columns,
        "Source",
        *[column for column in compare_columns if column not in key_columns],
    ]
    column_numbers = {column: index + 1 for index, column in enumerate(columns)}
    value_columns = [column for column in compare_columns if column not in key_columns]
    for key in matched_keys:
        excel_row = excel_by_key[key]
        db_row = db_by_key[key]
        mismatched_columns = [
            column
            for column in value_columns
            if not _exact_equal(excel_row[column], db_row[column])
        ]
        if not mismatched_columns:
            continue

        excel_output_row = {column: excel_row[column] for column in compare_columns}
        excel_output_row["Source"] = "Excel"
        db_output_row = {column: db_row[column] for column in compare_columns}
        db_output_row["Source"] = "Database"
        rows.extend([excel_output_row, db_output_row])

        excel_sheet_row = len(rows)
        db_sheet_row = len(rows) + 1
        for column in mismatched_columns:
            highlighted_cells.add((excel_sheet_row, column_numbers[column]))
            highlighted_cells.add((db_sheet_row, column_numbers[column]))

    return pd.DataFrame(rows, columns=columns), highlighted_cells


def _keys_for_records(
    rows_by_key: dict[tuple[Any, ...], pd.Series],
    keys: set[tuple[Any, ...]],
    key_columns: list[str],
) -> pd.DataFrame:
    rows = [
        dict(zip(key_columns, key, strict=True))
        for key in sorted(keys, key=str)
    ]
    return pd.DataFrame(rows, columns=key_columns)


def _column_mapping_sheet(
    excel_columns: list[str],
    db_columns: list[str],
    aliases: dict[str, str],
    excel_header_mapping: dict[str, str] | None = None,
) -> pd.DataFrame:
    canonical_by_uploaded_column = {
        canonical_column: uploaded_column
        for uploaded_column, canonical_column in (excel_header_mapping or {}).items()
    }
    db_column_by_excel_column = {
        aliases.get(db_column, db_column): db_column
        for db_column in db_columns
        if aliases.get(db_column, db_column) in excel_columns
    }
    mapped_db_columns = set(db_column_by_excel_column.values())
    rows = [
        {
            "Excel Column": canonical_by_uploaded_column.get(excel_column, excel_column),
            "DB Column": db_column_by_excel_column.get(excel_column, ""),
        }
        for excel_column in excel_columns
    ]
    rows.extend(
        {"Excel Column": "", "DB Column": db_column}
        for db_column in db_columns
        if db_column not in mapped_db_columns
    )
    return pd.DataFrame(rows, columns=["Excel Column", "DB Column"])


def _excel_header_mapping(
    excel_columns: list[str],
    header_aliases: dict[str, list[str]],
) -> dict[str, str]:
    alias_to_canonical = _alias_to_canonical_column(header_aliases)
    mapping: dict[str, str] = {}
    mapped_canonical_columns: dict[str, str] = {}

    for excel_column in excel_columns:
        canonical_column = alias_to_canonical.get(_normalize_header(excel_column))
        if not canonical_column or canonical_column == excel_column:
            continue
        if canonical_column in mapped_canonical_columns:
            raise ValueError(
                "Multiple uploaded headers map to the same configured header "
                f"{canonical_column}: {mapped_canonical_columns[canonical_column]}, "
                f"{excel_column}"
            )
        mapping[excel_column] = canonical_column
        mapped_canonical_columns[canonical_column] = excel_column

    return mapping


def _alias_to_canonical_column(header_aliases: dict[str, list[str]]) -> dict[str, str]:
    alias_to_canonical: dict[str, str] = {}
    for canonical_column, aliases in header_aliases.items():
        for alias in [canonical_column, *aliases]:
            normalized_alias = _normalize_header(alias)
            existing_column = alias_to_canonical.get(normalized_alias)
            if existing_column and existing_column != canonical_column:
                raise ValueError(
                    f"Header alias {alias!r} is configured for both "
                    f"{existing_column!r} and {canonical_column!r}"
                )
            alias_to_canonical[normalized_alias] = canonical_column
    return alias_to_canonical


def _normalize_header(value: Any) -> str:
    return "".join(
        character.lower()
        for character in str(value).replace("\u00a0", " ")
        if character.isalnum()
    )


def _exact_equal(left: Any, right: Any) -> bool:
    if pd.isna(left) and pd.isna(right):
        return True
    if type(left) is not type(right):
        return False
    return bool(left == right)


def _rows_by_key(df: pd.DataFrame, key_columns: list[str]) -> dict[tuple[Any, ...], pd.Series]:
    rows: dict[tuple[Any, ...], pd.Series] = {}
    for _, row in df.iterrows():
        rows[tuple(row[column] for column in key_columns)] = row
    return rows


def _label(value: str) -> str:
    return value.replace("_", " ").title()
