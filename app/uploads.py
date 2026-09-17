from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile, ZipFile
import re

import pandas as pd


SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
SUPPORTED_CSV_EXTENSIONS = {".csv"}


def read_source_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in SUPPORTED_CSV_EXTENSIONS:
        source_df = pd.read_csv(path, keep_default_na=False)
        raw_df = pd.read_csv(path, dtype=str, keep_default_na=False)
        source_df.attrs["scientific_number_values"] = [
            (row_index, column, raw_df.iat[row_index, column_index])
            for row_index in range(len(raw_df))
            for column_index, column in enumerate(raw_df.columns)
            if _is_scientific_number(raw_df.iat[row_index, column_index])
        ]
        return source_df
    if suffix not in SUPPORTED_EXCEL_EXTENSIONS:
        raise ValueError(
            "Unsupported source file type. Please upload a .csv, .xlsx, or .xlsm file. "
            "If your file is legacy .xls, open it in Excel and save it as .xlsx first."
        )

    try:
        with ZipFile(path) as archive:
            if "[Content_Types].xml" not in archive.namelist():
                raise ValueError
    except (BadZipFile, ValueError) as exc:
        raise ValueError(
            "The uploaded file is not a valid .xlsx/.xlsm workbook. "
            "This often happens when a .csv or old .xls file was renamed to .xlsx. "
            "Open the file in Excel and use Save As -> Excel Workbook (*.xlsx), then upload that file."
        ) from exc

    return pd.read_excel(path, sheet_name=0, engine="openpyxl")


_SCIENTIFIC_NUMBER = re.compile(
    r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)[eE][+-]?\d+$"
)


def _is_scientific_number(value: object) -> bool:
    return bool(_SCIENTIFIC_NUMBER.fullmatch(str(value).strip()))
