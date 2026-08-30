from __future__ import annotations

from pathlib import Path

import pandas as pd


def export_dataframe(df: pd.DataFrame, path: Path, export_format: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    if export_format == "csv":
        output = path.with_suffix(".csv")
        df.to_csv(output, index=False)
        return output
    if export_format == "xlsx":
        output = path.with_suffix(".xlsx")
        df.to_excel(output, index=False, engine="openpyxl")
        return output
    raise ValueError(f"Unsupported export format: {export_format}")
