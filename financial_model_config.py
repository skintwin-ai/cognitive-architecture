from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Iterable

import openpyxl

OUTPUT_ENV_VAR = 'SKINTWIN_FINANCIAL_MODEL_PATH'
DEFAULT_OUTPUT_PATH = Path(__file__).resolve().parent / 'skin_care_financial_model.xlsx'


def parse_output_args(description: str) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        '--output',
        default=os.environ.get(OUTPUT_ENV_VAR),
        help=f'Output workbook path. Defaults to {OUTPUT_ENV_VAR} or {DEFAULT_OUTPUT_PATH}.',
    )
    return parser.parse_args()


def resolve_output_path(output_path: str | os.PathLike[str] | None = None) -> Path:
    if output_path:
        return Path(output_path).expanduser().resolve()
    return DEFAULT_OUTPUT_PATH


def save_workbook(workbook: openpyxl.Workbook, output_path: str | os.PathLike[str]) -> Path:
    path = resolve_output_path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def load_existing_workbook(output_path: str | os.PathLike[str], required_sheets: Iterable[str] | None = None) -> openpyxl.Workbook:
    path = resolve_output_path(output_path)
    if not path.exists():
        raise FileNotFoundError(f'Workbook not found at {path}. Run create_assumptions_sheet.py first or provide --output.')

    workbook = openpyxl.load_workbook(path)
    if required_sheets:
        missing = [sheet for sheet in required_sheets if sheet not in workbook.sheetnames]
        if missing:
            raise ValueError(f'Missing required sheet(s): {", ".join(missing)} in {path}.')
    return workbook
