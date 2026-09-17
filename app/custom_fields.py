from __future__ import annotations

import re
from typing import Any

import pandas as pd


CUSTOM_FIELD_HEADER = re.compile(r"^cf_\s*(\d+)$", re.IGNORECASE)


def custom_field_headers(columns: list[Any]) -> dict[Any, str]:
    mapping: dict[Any, str] = {}
    seen: set[str] = set()
    for column in columns:
        match = CUSTOM_FIELD_HEADER.fullmatch(str(column).strip())
        if not match:
            continue
        canonical = f"cf_{int(match.group(1))}"
        if canonical in seen:
            raise ValueError(f"Multiple uploaded headers identify custom field {canonical}")
        mapping[column] = canonical
        seen.add(canonical)
    return mapping


def add_custom_field_values(
    excel_df: pd.DataFrame,
    db_df: pd.DataFrame,
    metadata: pd.DataFrame,
    values: pd.DataFrame,
    options: pd.DataFrame,
) -> pd.DataFrame:
    header_mapping = custom_field_headers(list(excel_df.columns))
    if not header_mapping:
        return db_df
    excel_df.rename(columns=header_mapping, inplace=True)

    field_ids = {int(column.removeprefix("cf_")) for column in header_mapping.values()}
    metadata_by_id = {
        int(row.id): row
        for row in metadata.itertuples(index=False)
        if pd.notna(row.id)
    }
    option_labels = {
        (int(row.field_id), _key(row.id)): str(row.text)
        for row in options.itertuples(index=False)
        if pd.notna(row.field_id) and pd.notna(row.id)
    }
    values_by_field_and_entity: dict[tuple[int, str], list[Any]] = {}
    for row in values.itertuples(index=False):
        if pd.isna(row.field_id) or pd.isna(row.entity_id):
            continue
        key = (int(row.field_id), _key(row.entity_id))
        values_by_field_and_entity.setdefault(key, []).append(row.value)

    entity_column = _entity_column(db_df)
    if entity_column is None:
        raise ValueError("Custom fields require the DB query to include a Case ID column")

    db_df = db_df.copy()
    for field_id in sorted(field_ids):
        field = metadata_by_id.get(field_id)
        column = _field_column_name(field_id, field)
        field_type = str(field.field_type).lower() if field is not None else ""
        db_df[column] = [
            _resolved_value(
                values_by_field_and_entity.get((field_id, _key(entity_id)), []),
                field_id,
                field_type,
                option_labels,
            )
            for entity_id in db_df[entity_column]
        ]
    return db_df


def custom_field_column_aliases(
    excel_columns: list[Any],
    metadata: pd.DataFrame,
) -> dict[str, str]:
    header_mapping = custom_field_headers(excel_columns)
    metadata_by_id = {
        int(row.id): row
        for row in metadata.itertuples(index=False)
        if pd.notna(row.id)
    }
    aliases: dict[str, str] = {}
    for canonical in header_mapping.values():
        field_id = int(canonical.removeprefix("cf_"))
        field = metadata_by_id.get(field_id)
        aliases[_field_column_name(field_id, field)] = canonical
    return aliases


def _entity_column(db_df: pd.DataFrame) -> str | None:
    for column in ("Case ID", "case id", "matter id", "id"):
        if column in db_df.columns:
            return column
    return None


def _field_column_name(field_id: int, field: Any) -> str:
    if field is not None and pd.notna(field.field_name) and str(field.field_name).strip():
        return str(field.field_name).strip()
    return f"cf_{field_id}"


def _resolved_value(
    raw_values: list[Any],
    field_id: int,
    field_type: str,
    option_labels: dict[tuple[int, str], str],
) -> str:
    if not raw_values:
        return ""
    if field_type in {"select", "multiselect"}:
        resolved: list[str] = []
        for raw_value in raw_values:
            for option_id in str(raw_value).split(","):
                option_id = option_id.strip()
                if not option_id:
                    continue
                resolved.append(option_labels.get((field_id, option_id), option_id))
        return ",".join(resolved)
    return str(raw_values[0]) if pd.notna(raw_values[0]) else ""


def _key(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()