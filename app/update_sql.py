from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import re
import unicodedata


DIRECT_MATTER_COLUMNS = {
    "shorttitle": "shorttitle",
    "longtitle": "longtitle",
    "country": "country",
    "clientref": "clientref",
    "Parentage": "parentage",
    "goods": "extra1",
    "classes": "extra3",
    "notes": "notes",
}


@dataclass(frozen=True)
class UpdatePlanRow:
    matter_id: int
    matter_code: str
    field: str
    current_value: Any
    excel_value: Any
    action: str
    sql: str
    warning: str = ""


@dataclass(frozen=True)
class UpdatePlan:
    rows: list[UpdatePlanRow]

    @property
    def sql(self) -> str:
        statements = [row.sql for row in self.rows if row.sql]
        if not statements:
            return "-- No safe update statements were generated.\n"
        return "START TRANSACTION;\n\n" + "\n\n".join(statements) + "\n\nCOMMIT;\n"


def plan_direct_matter_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    field: str,
    current_value: Any,
    excel_value: Any,
) -> UpdatePlanRow:
    column = DIRECT_MATTER_COLUMNS.get(field)
    if column is None:
        return _warning_row(
            matter_id,
            matter_code,
            field,
            current_value,
            excel_value,
            f"Field {field!r} is not enabled for direct matter updates.",
        )
    sql = (
        "UPDATE myprompts_matters\n"
        f"SET {column} = {_sql_value(excel_value)}\n"
        f"WHERE orgid = {int(org_id)} AND id = {int(matter_id)};"
    )
    return _update_row(matter_id, matter_code, field, current_value, excel_value, sql)


def plan_lookup_matter_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    field: str,
    column: str,
    current_value: Any,
    excel_value: Any,
    lookup_id: int | None,
    lookup_label: str,
) -> UpdatePlanRow:
    if lookup_id is None:
        return _warning_row(
            matter_id, matter_code, field, current_value, excel_value,
            f"{lookup_label} not found; {field} update skipped.",
        )
    sql = (
        "UPDATE myprompts_matters\n"
        f"SET `{column}` = {int(lookup_id)}\n"
        f"WHERE orgid = {int(org_id)} AND id = {int(matter_id)};"
    )
    return _update_row(matter_id, matter_code, field, current_value, excel_value, sql)


def plan_client_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    current_value: Any,
    excel_value: Any,
    organisation_id: int | None = None,
) -> UpdatePlanRow:
    if organisation_id is None:
        return _warning_row(
            matter_id,
            matter_code,
            "client",
            current_value,
            excel_value,
            "Organisation not found in contacts_organisations; client update skipped.",
        )
    sql = (
        "UPDATE myprompts_matters\n"
        f"SET oid = {int(organisation_id)}\n"
        f"WHERE orgid = {int(org_id)} AND id = {int(matter_id)};"
    )
    return _update_row(matter_id, matter_code, "client", current_value, excel_value, sql)


def plan_site_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    current_value: Any,
    excel_value: Any,
    site_id: int | None = None,
) -> UpdatePlanRow:
    if site_id is None:
        return _warning_row(
            matter_id,
            matter_code,
            "who pays the bill",
            current_value,
            excel_value,
            "Site not found in contacts_sites; who pays the bill update skipped.",
        )
    sql = (
        "UPDATE myprompts_matters\n"
        f"SET sid = {int(site_id)}\n"
        f"WHERE orgid = {int(org_id)} AND id = {int(matter_id)};"
    )
    return _update_row(matter_id, matter_code, "who pays the bill", current_value, excel_value, sql)


def plan_history_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    history_id: int,
    value_type: str,
    current_value: Any,
    excel_value: Any,
    row_exists: bool,
) -> UpdatePlanRow:
    column = {"date": "date", "country": "country", "number": "code"}.get(value_type)
    field = f"hid_{history_id} {value_type}"
    if column is None:
        return _warning_row(matter_id, matter_code, field, current_value, excel_value, "Unsupported history value type.")
    if row_exists:
        sql = (
            "UPDATE myprompts_matterhistories\n"
            f"SET {column} = {_sql_value(excel_value)}\n"
            f"WHERE orgid = {int(org_id)} AND matterid = {int(matter_id)} AND hid = {int(history_id)};"
        )
        return _update_row(matter_id, matter_code, field, current_value, excel_value, sql)
    sql = (
        "INSERT INTO myprompts_matterhistories (orgid, matterid, hid, date, country, code)\n"
        f"VALUES ({int(org_id)}, {int(matter_id)}, {int(history_id)}, "
        f"{_sql_value(excel_value) if value_type == 'date' else 'NULL'}, "
        f"{_sql_value(excel_value) if value_type == 'country' else 'NULL'}, "
        f"{_sql_value(excel_value) if value_type == 'number' else 'NULL'});"
    )
    return _update_row(matter_id, matter_code, field, current_value, excel_value, sql)


def plan_custom_field_update(
    org_id: int,
    matter_id: int,
    matter_code: str,
    field_id: int,
    field_type: str,
    current_value: Any,
    excel_value: Any,
    value_row_exists: bool,
    option_ids: list[int] | None = None,
    missing_options: list[str] | None = None,
) -> UpdatePlanRow:
    field = f"cf_{field_id}"
    missing_options = missing_options or []
    if missing_options:
        return _warning_row(
            matter_id,
            matter_code,
            field,
            current_value,
            excel_value,
            "Custom-field options do not exist: " + ", ".join(missing_options),
        )
    if field_type.lower() in {"select", "multiselect"}:
        stored_value = ",".join(str(option_id) for option_id in (option_ids or []))
    else:
        stored_value = excel_value
    if value_row_exists:
        sql = (
            "UPDATE myprompts_custom_fields_values\n"
            f"SET value = {_sql_value(stored_value)}\n"
            f"WHERE field_id = {int(field_id)} AND entity_id = {int(matter_id)};"
        )
    else:
        sql = (
            "INSERT INTO myprompts_custom_fields_values (field_id, entity_id, value)\n"
            f"VALUES ({int(field_id)}, {int(matter_id)}, {_sql_value(stored_value)});"
        )
    return _update_row(matter_id, matter_code, field, current_value, excel_value, sql)


def plan_party_relationships(
    org_id: int,
    matter_id: int,
    matter_code: str,
    field: str,
    current_names: list[str],
    excel_names: list[str],
    existing_contact_ids: dict[str, int],
) -> list[UpdatePlanRow]:
    category = {"applicants": "applicant", "inventors": "inventor"}.get(field.lower())
    if category is None:
        return [_warning_row(matter_id, matter_code, field, current_names, excel_names, "Unsupported party field.")]
    current_by_key = {_name_key(name): name for name in current_names if name.strip()}
    excel_by_key = {_name_key(name): name for name in excel_names if name.strip()}
    rows: list[UpdatePlanRow] = []
    contact_ids_by_key = {
        _name_key(name): contact_id
        for name, contact_id in existing_contact_ids.items()
    }
    missing_names = [
        name for name_key, name in excel_by_key.items()
        if name_key not in contact_ids_by_key
    ]
    if missing_names:
        return [
            _warning_row(
                matter_id,
                matter_code,
                field,
                current_names,
                excel_names,
                "Update skipped completely; contact(s) do not exist: "
                + ", ".join(missing_names),
            )
        ]
    removed_keys = sorted(current_by_key.keys() - excel_by_key.keys())
    added_keys = sorted(excel_by_key.keys() - current_by_key.keys())
    for old_key, new_key in zip(removed_keys, added_keys, strict=False):
        old_contact_id = contact_ids_by_key.get(old_key)
        new_contact_id = contact_ids_by_key.get(new_key)
        if old_contact_id is None or new_contact_id is None:
            continue
        sql = (
            "UPDATE myprompts_contacts\n"
            f"SET cid = {int(new_contact_id)}\n"
            f"WHERE orgid = {int(org_id)} AND matterid = {int(matter_id)}\n"
            f"  AND category = {_sql_value(category)} AND cid = {int(old_contact_id)};"
        )
        rows.append(_update_row(matter_id, matter_code, field, current_names, excel_names, sql, "Update relationship cid only"))
    for new_key in added_keys[len(removed_keys):]:
        new_contact_id = contact_ids_by_key.get(new_key)
        if new_contact_id is not None:
            sql = (
                "INSERT INTO myprompts_contacts (orgid, matterid, cid, category, sequence)\n"
                f"VALUES ({int(org_id)}, {int(matter_id)}, {int(new_contact_id)}, {_sql_value(category)}, 0);"
            )
            rows.append(_update_row(matter_id, matter_code, field, current_names, excel_names, sql, "Add relationship row"))
    for old_key in removed_keys[len(added_keys):]:
        rows.append(
            _warning_row(
                matter_id,
                matter_code,
                field,
                current_names,
                excel_names,
                "Existing relationship retained; no replacement Excel contact was available.",
            )
        )
    if not rows:
        rows.append(
            _warning_row(
                matter_id,
                matter_code,
                field,
                current_names,
                excel_names,
                "No relationship change required; existing contacts match ignoring case and accents.",
            )
        )
    return rows


def _update_row(
    matter_id: int,
    matter_code: str,
    field: str,
    current_value: Any,
    excel_value: Any,
    sql: str,
    warning: str = "",
) -> UpdatePlanRow:
    return UpdatePlanRow(matter_id, matter_code, field, current_value, excel_value, "Update", sql, warning)


def _warning_row(
    matter_id: int,
    matter_code: str,
    field: str,
    current_value: Any,
    excel_value: Any,
    warning: str,
) -> UpdatePlanRow:
    return UpdatePlanRow(matter_id, matter_code, field, current_value, excel_value, "Skipped", "", warning)


def _sql_value(value: Any) -> str:
    if value is None:
        return "NULL"
    text = str(value)
    if not text.strip():
        return "''"
    escaped = text.replace("\\", "\\\\").replace("'", "''")
    return f"'{escaped}'"


def _name_key(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value))
    without_accents = "".join(
        character for character in normalized
        if not unicodedata.combining(character)
    )
    return " ".join(without_accents.casefold().split())
