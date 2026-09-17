from __future__ import annotations

import json
import re
from typing import Any

import pandas as pd


HISTORY_HEADER = re.compile(r"^hid_\s*(\d+)\s+(date|country|number)$", re.IGNORECASE)
IMPORTER_HISTORY_FIELDS = {
    "priority date": (1, "date"),
    "priority country": (1, "country"),
    "priority number": (1, "number"),
    "priority code": (1, "number"),
    "application date": (2, "date"),
    "filing date": (2, "date"),
    "application country": (2, "country"),
    "filing country": (2, "country"),
    "application number": (2, "number"),
    "application code": (2, "number"),
    "filing number": (2, "number"),
    "filing code": (2, "number"),
    "publication date": (3, "date"),
    "pub/adv date": (3, "date"),
    "publication country": (3, "country"),
    "pub/adv country": (3, "country"),
    "publication number": (3, "number"),
    "publication code": (3, "number"),
    "pub/adv number": (3, "number"),
    "pub/adv code": (3, "number"),
    "grant date": (4, "date"),
    "grant country": (4, "country"),
    "grant number": (4, "number"),
    "grant code": (4, "number"),
    "nat/reg entry date": (5, "date"),
    "lodged date": (5, "date"),
    "nat/reg entry country": (5, "country"),
    "lodged country": (5, "country"),
    "nat/reg entry number": (5, "number"),
    "nat/reg entry code": (5, "number"),
    "lodged number": (5, "number"),
    "lodged code": (5, "number"),
    "nat/reg publication date": (6, "date"),
    "nat/reg publication country": (6, "country"),
    "nat/reg publication number": (6, "number"),
    "nat/reg publication code": (6, "number"),
}


def importer_history_field(field: Any) -> tuple[int, str] | None:
    normalized = " ".join(str(field).strip().lower().split())
    return IMPORTER_HISTORY_FIELDS.get(normalized)


def history_headers(columns: list[Any]) -> dict[Any, str]:
    mapping: dict[Any, str] = {}
    seen: set[str] = set()
    for column in columns:
        match = HISTORY_HEADER.fullmatch(str(column).strip())
        if not match:
            continue
        canonical = f"hid_{int(match.group(1))} {match.group(2).lower()}"
        if canonical in seen:
            raise ValueError(f"Multiple uploaded headers identify history field {canonical}")
        mapping[column] = canonical
        seen.add(canonical)
    return mapping


def add_history_values(
    excel_df: pd.DataFrame,
    db_df: pd.DataFrame,
    definitions: pd.DataFrame,
    values: pd.DataFrame,
) -> pd.DataFrame:
    header_mapping = history_headers(list(excel_df.columns))
    if not header_mapping:
        return db_df
    excel_df.rename(columns=header_mapping, inplace=True)

    requested = {
        (int(match.group(1)), match.group(2).lower())
        for canonical in header_mapping.values()
        if (match := HISTORY_HEADER.fullmatch(canonical))
    }
    names_by_id = _history_names(definitions)
    values_by_key = {
        (int(row.hid), _key(row.matterid)): row
        for row in values.itertuples(index=False)
        if pd.notna(row.hid) and pd.notna(row.matterid)
    }
    entity_column = _entity_column(db_df)
    if entity_column is None:
        raise ValueError("History fields require the DB query to include a Case ID column")

    db_df = db_df.copy()
    for history_id, value_type in sorted(requested):
        db_column = _history_column_name(history_id, value_type, names_by_id)
        db_df[db_column] = [
            _history_value(
                values_by_key.get((history_id, _key(entity_id))),
                value_type,
            )
            for entity_id in db_df[entity_column]
        ]
    return db_df


def history_column_aliases(
    excel_columns: list[Any],
    definitions: pd.DataFrame,
) -> dict[str, str]:
    names_by_id = _history_names(definitions)
    aliases: dict[str, str] = {}
    for canonical in history_headers(excel_columns).values():
        match = HISTORY_HEADER.fullmatch(canonical)
        if match is None:
            continue
        history_id = int(match.group(1))
        value_type = match.group(2).lower()
        aliases[_history_column_name(history_id, value_type, names_by_id)] = canonical
    return aliases


def _history_names(definitions: pd.DataFrame) -> dict[int, str]:
    names_by_id: dict[int, str] = {}
    for row in definitions.itertuples(index=False):
        if pd.isna(row.history):
            continue
        try:
            names = json.loads(row.names) if isinstance(row.names, str) else row.names
        except (TypeError, json.JSONDecodeError):
            names = None
        if isinstance(names, list) and names:
            names_by_id[int(row.history)] = str(names[0]).strip()
        elif names:
            names_by_id[int(row.history)] = str(names).strip()
    return names_by_id


def _history_column_name(history_id: int, value_type: str, names_by_id: dict[int, str]) -> str:
    return f"{names_by_id.get(history_id, f'hid_{history_id}')} {value_type}"


def _history_value(row: Any, value_type: str) -> str:
    if row is None:
        return ""
    value = getattr(row, {"date": "date", "country": "country", "number": "code"}[value_type])
    if pd.isna(value):
        return ""
    if value_type == "date":
        try:
            return pd.Timestamp(value).strftime("%Y-%m-%d")
        except (TypeError, ValueError):
            return str(value)
    return str(value)


def _entity_column(db_df: pd.DataFrame) -> str | None:
    for column in ("Case ID", "case id", "matter id", "id"):
        if column in db_df.columns:
            return column
    return None


def _key(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()