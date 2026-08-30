from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile, ZipFile

import pandas as pd


SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
SUPPORTED_CSV_EXTENSIONS = {".csv"}


def read_source_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in SUPPORTED_CSV_EXTENSIONS:
        return pd.read_csv(path, keep_default_na=False)
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
