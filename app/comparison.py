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
    summary_rows: list[dict[str, Any]]
    highlighted_cells: dict[str, dict[tuple[int, int], str]]


def compare_dataframes(
    excel_df: pd.DataFrame,
    db_df: pd.DataFrame,
    key_columns: list[str],
    optional_key_columns: list[str] | None = None,
    column_aliases: dict[str, str] | None = None,
    header_aliases: dict[str, list[str]] | None = None,
    summary_subject: str = "Records",
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
    resolved_key_columns = _resolve_key_columns(
        key_columns,
        optional_key_columns or [],
        excel_columns,
        db_columns,
    )

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
    if not resolved_key_columns:
        raise ValueError("At least one configured key column must exist in both Excel and DB extract.")

    missing_db_columns = [column for column in excel_columns if column not in db_columns]
    compare_columns = [column for column in excel_columns if column not in missing_db_columns]

    excel_dup_mask = excel_df.duplicated(subset=resolved_key_columns, keep=False)
    db_dup_mask = db_df.duplicated(subset=resolved_key_columns, keep=False)

    excel_unique = excel_df.loc[~excel_dup_mask].copy()
    db_unique = db_df.loc[~db_dup_mask].copy()

    excel_by_key = _rows_by_key(excel_unique, resolved_key_columns)
    db_by_key = _rows_by_key(db_unique, resolved_key_columns)

    excel_keys = set(excel_by_key)
    db_keys = set(db_by_key)
    matched_keys = sorted(excel_keys & db_keys, key=str)

    extra_in_excel = _keys_for_records(excel_by_key, excel_keys - db_keys, resolved_key_columns)
    extra_in_database = _keys_for_records(db_by_key, db_keys - excel_keys, resolved_key_columns)
    mismatched_records, highlighted_cells = _mismatched_records(
        excel_by_key,
        db_by_key,
        matched_keys,
        resolved_key_columns,
        compare_columns,
    )

    summary = {
        "excel_rows": len(excel_df),
        "db_rows": len(db_df),
        "compared_columns": len(compare_columns),
        "matched_rows": len(matched_keys),
        "mismatched_records": mismatched_records[
            resolved_key_columns
        ].drop_duplicates().shape[0]
        if not mismatched_records.empty
        else 0,
        "extra_records_in_database": len(extra_in_database),
        "extra_records_in_excel": len(extra_in_excel),
        "duplicate_excel_rows": int(excel_dup_mask.sum()),
        "duplicate_db_rows": int(db_dup_mask.sum()),
        "missing_db_columns": len(missing_db_columns),
    }
    summary_rows = _summary_rows(summary, summary_subject)
    summary_sheet = pd.DataFrame(summary_rows, columns=["Metric", "Value"])
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
        summary_rows=summary_rows,
        highlighted_cells={"Mismatched_Records": highlighted_cells},
    )


def _resolve_key_columns(
    required_key_columns: list[str],
    optional_key_columns: list[str],
    excel_columns: list[str],
    db_columns: list[str],
) -> list[str]:
    resolved_key_columns = list(required_key_columns)
    for column in optional_key_columns:
        if column in resolved_key_columns:
            continue
        if column in excel_columns and column in db_columns:
            resolved_key_columns.append(column)
    return resolved_key_columns


def write_report(result: ComparisonResult, output_path: str) -> None:
    mismatch_fills = {
        "mismatch": PatternFill(
            fill_type="solid",
            start_color="FFF2CC",
            end_color="FFF2CC",
        ),
        "whitespace": PatternFill(
            fill_type="solid",
            start_color="FFC6EFCE",
            end_color="FFC6EFCE",
        ),
    }
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
            for (row_number, column_number), highlight_type in cells.items():
                worksheet.cell(row=row_number, column=column_number).fill = mismatch_fills[
                    highlight_type
                ]


def _mismatched_records(
    excel_by_key: dict[tuple[Any, ...], pd.Series],
    db_by_key: dict[tuple[Any, ...], pd.Series],
    matched_keys: list[tuple[Any, ...]],
    key_columns: list[str],
    compare_columns: list[str],
) -> tuple[pd.DataFrame, dict[tuple[int, int], str]]:
    mismatch_rows: list[tuple[pd.Series, pd.Series, list[str]]] = []
    included_value_columns: set[str] = set()
    for key in matched_keys:
        excel_row = excel_by_key[key]
        db_row = db_by_key[key]
        mismatched_columns = [
            column
            for column in compare_columns
            if column not in key_columns
            and not _exact_equal(excel_row[column], db_row[column])
        ]
        if mismatched_columns:
            mismatch_rows.append((excel_row, db_row, mismatched_columns))
            included_value_columns.update(mismatched_columns)

    value_columns = [
        column
        for column in compare_columns
        if column not in key_columns and column in included_value_columns
    ]
    columns = [
        *key_columns,
        *[
            source_column
            for column in value_columns
            for source_column in (f"{column} (Excel)", f"{column} (Database)")
        ],
    ]
    column_numbers = {column: index + 1 for index, column in enumerate(columns)}
    rows: list[dict[str, Any]] = []
    highlighted_cells: dict[tuple[int, int], str] = {}
    for excel_row, db_row, mismatched_columns in mismatch_rows:
        output_row = {
            **{column: excel_row[column] for column in key_columns},
            **{
                source_column: value
                for column in value_columns
                for source_column, value in (
                    (f"{column} (Excel)", excel_row[column]),
                    (f"{column} (Database)", db_row[column]),
                )
            },
        }
        rows.append(output_row)

        sheet_row = len(rows) + 1
        for column in mismatched_columns:
            highlight_type = (
                "whitespace"
                if _whitespace_only_difference(excel_row[column], db_row[column])
                else "mismatch"
            )
            highlighted_cells[(sheet_row, column_numbers[f"{column} (Excel)"])] = highlight_type
            highlighted_cells[(sheet_row, column_numbers[f"{column} (Database)"])] = highlight_type

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

    if isinstance(left, str) and isinstance(right, str):
        return left == right

    if type(left) is not type(right):
        return False
    return bool(left == right)


def _whitespace_only_difference(left: Any, right: Any) -> bool:
    return (
        isinstance(left, str)
        and isinstance(right, str)
        and left != right
        and left.strip() == right.strip()
    )


def _normalize_key_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def _rows_by_key(df: pd.DataFrame, key_columns: list[str]) -> dict[tuple[Any, ...], pd.Series]:
    rows: dict[tuple[Any, ...], pd.Series] = {}
    for _, row in df.iterrows():
        key = tuple(_normalize_key_value(row[column]) for column in key_columns)
        rows[key] = row
    return rows


def _label(value: str) -> str:
    return value.replace("_", " ").title()


def _summary_rows(summary: dict[str, Any], subject: str) -> list[dict[str, Any]]:
    subject = subject.strip() or "Records"
    metric_labels = {
        "excel_rows": f"{subject} In Excel",
        "db_rows": f"{subject} In Database",
        "compared_columns": "Compared Columns",
        "matched_rows": f"Matched {subject}",
        "mismatched_records": f"Mismatched {subject}",
        "extra_records_in_database": f"Extra {subject} In Database",
        "extra_records_in_excel": f"Extra {subject} In Excel",
        "duplicate_excel_rows": f"Duplicate {subject} In Excel",
        "duplicate_db_rows": f"Duplicate {subject} In Database",
        "missing_db_columns": "Missing DB Columns",
    }
    return [
        {"Metric": metric_labels.get(key, _label(key)), "Value": value}
        for key, value in summary.items()
    ]
