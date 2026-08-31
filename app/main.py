from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from app.comparison import compare_dataframes, write_report
from app.config import load_modules, load_regions
from app.db import execute_module_query
from app.exporters import export_dataframe
from app.uploads import read_source_file


BASE_DIR = Path(__file__).resolve().parent.parent
RUNS_DIR = BASE_DIR / "runs"
RUN_ID_PATTERN = re.compile(r"^[A-Za-z0-9_]+$")
MODULES = load_modules(BASE_DIR / "validation_config.yaml")
REGIONS = load_regions(BASE_DIR / "db_regions.yaml")

app = FastAPI(title="Migration Data Comparison Tool")
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
            column_aliases=module.column_aliases,
            header_aliases=module.header_aliases,
            summary_subject=module.summary_subject,
        )
        report_path = run_dir / f"{output_name_prefix}_comparison_report.xlsx"
        write_report(comparison, str(report_path))
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
                "region_name": region.name,
                "summary": comparison.summary,
                "summary_rows": comparison.summary_rows,
                "db_export_name": db_export.name,
                "report_name": report_path.name,
            },
        },
    )


@app.get("/download/{run_id}/{filename}")
async def download(run_id: str, filename: str) -> FileResponse:
    if not RUN_ID_PATTERN.fullmatch(run_id):
        raise HTTPException(status_code=400, detail="Invalid run id")
    path = RUNS_DIR / run_id / filename
    if not path.is_file():
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(path, filename=filename)


@app.post("/cleanup")
async def cleanup() -> dict[str, int]:
    removed = 0
    if RUNS_DIR.exists():
        for child in RUNS_DIR.iterdir():
            if child.is_dir():
                shutil.rmtree(child)
                removed += 1
    return {"removed_run_folders": removed}
