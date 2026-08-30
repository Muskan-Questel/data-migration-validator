from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class ModuleConfig:
    identifier: str
    name: str
    summary_subject: str
    query_template: str
    key_columns: list[str]
    column_aliases: dict[str, str]
    header_aliases: dict[str, list[str]]
    export_format: str


@dataclass(frozen=True)
class RegionConfig:
    identifier: str
    name: str
    env_prefix: str


def load_modules(config_path: str | Path = "validation_config.yaml") -> dict[str, ModuleConfig]:
    path = Path(config_path)
    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}

    modules: dict[str, ModuleConfig] = {}
    for identifier, payload in (raw.get("modules") or {}).items():
        module = _module_from_payload(identifier, payload)
        modules[identifier] = module

    modules_dir = raw.get("modules_dir")
    if modules_dir:
        module_path = Path(modules_dir)
        if not module_path.is_absolute():
            module_path = path.parent / module_path
        for module_file in sorted(module_path.glob("*.yml")) + sorted(module_path.glob("*.yaml")):
            for identifier, payload in _module_payloads_from_file(module_file).items():
                module = _module_from_payload(identifier, payload)
                modules[identifier] = module
    return modules


def load_regions(config_path: str | Path = "db_regions.yaml") -> dict[str, RegionConfig]:
    with Path(config_path).open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}

    regions: dict[str, RegionConfig] = {}
    for identifier, payload in (raw.get("regions") or {}).items():
        region = RegionConfig(
            identifier=identifier,
            name=payload["name"],
            env_prefix=payload["env_prefix"],
        )
        regions[identifier] = region
    return regions


def _module_from_payload(identifier: str, payload: dict[str, Any]) -> ModuleConfig:
    export_format = payload.get("export_format", "xlsx")
    if export_format not in {"xlsx", "csv"}:
        raise ValueError(f"Unsupported export format for module {identifier}: {export_format}")

    return ModuleConfig(
        identifier=identifier,
        name=payload["name"],
        summary_subject=payload.get("summary_subject") or _default_summary_subject(payload["name"]),
        query_template=payload["query_template"],
        key_columns=list(payload["key_columns"]),
        column_aliases=dict(payload.get("column_aliases") or {}),
        header_aliases={
            column: list(aliases)
            for column, aliases in (payload.get("header_aliases") or {}).items()
        },
        export_format=export_format,
    )


def _module_payloads_from_file(path: Path) -> dict[str, dict[str, Any]]:
    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}

    if "modules" in raw:
        return dict(raw["modules"] or {})
    return dict(raw)


def _default_summary_subject(module_name: str) -> str:
    if module_name.lower() == "address book":
        return "Address Book Records"
    if module_name.endswith("s"):
        return module_name
    return f"{module_name} Records"
