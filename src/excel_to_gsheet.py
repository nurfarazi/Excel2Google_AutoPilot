"""
Automates clearing a Google Sheet tab and uploading new data from an Excel workbook.

Environment configuration is read from a `.env` file (see `.env.example`).
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, WorksheetNotFound
from gspread.utils import rowcol_to_a1

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CHUNK_ROW_COUNT = 1000  # limit payload size when sending to Sheets


@dataclass(frozen=True)
class Settings:
    excel_file_path: Path
    service_account_file: Path
    spreadsheet_id: str
    worksheet_name: str
    columns: Sequence[str]

    @classmethod
    def from_env(cls) -> "Settings":
        project_root = Path(__file__).resolve().parents[1]
        # Load the project-level .env, fall back to default search if not found.
        env_path = project_root / ".env"
        if env_path.exists():
            load_dotenv(env_path)
        else:
            load_dotenv()

        excel_file = os.getenv("EXCEL_FILE_PATH")
        service_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
        spreadsheet_id = os.getenv("GOOGLE_SPREADSHEET_ID")
        worksheet_name = os.getenv("GOOGLE_WORKSHEET_NAME")
        column_env = os.getenv("GOOGLE_COLUMNS", "")

        missing = [
            name
            for name, value in [
                ("EXCEL_FILE_PATH", excel_file),
                ("GOOGLE_SERVICE_ACCOUNT_FILE", service_file),
                ("GOOGLE_SPREADSHEET_ID", spreadsheet_id),
                ("GOOGLE_WORKSHEET_NAME", worksheet_name),
            ]
            if not value
        ]
        if missing:
            raise RuntimeError(
                "Missing required environment variables: "
                f"{', '.join(missing)}. Copy .env.example to .env and fill in values."
            )

        columns = tuple(col.strip() for col in column_env.split(",") if col.strip())
        return cls(
            excel_file_path=(project_root / excel_file).resolve()
            if not Path(excel_file).is_absolute()
            else Path(excel_file),
            service_account_file=(project_root / service_file).resolve()
            if not Path(service_file).is_absolute()
            else Path(service_file),
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
            columns=columns,
        )


def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Clear a Google Sheet and upload data from an Excel file."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Load and validate configuration, but do not modify the Google Sheet.",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable debug logs.",
    )
    return parser.parse_args(argv)


def authenticate(settings: Settings) -> gspread.Client:
    logging.debug("Authenticating with service account credentials %s", settings.service_account_file)
    credentials = Credentials.from_service_account_file(
        settings.service_account_file, scopes=SCOPES
    )
    return gspread.authorize(credentials)


def get_worksheet(client: gspread.Client, spreadsheet_id: str, worksheet_name: str) -> gspread.Worksheet:
    logging.debug("Opening spreadsheet %s worksheet %s", spreadsheet_id, worksheet_name)
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except APIError as exc:
        raise RuntimeError(f"Failed to open spreadsheet: {exc}") from exc

    try:
        return spreadsheet.worksheet(worksheet_name)
    except WorksheetNotFound as exc:
        raise RuntimeError(
            f"Worksheet '{worksheet_name}' not found in spreadsheet {spreadsheet_id}."
        ) from exc


def read_excel_data(settings: Settings) -> pd.DataFrame:
    logging.debug("Reading Excel data from %s", settings.excel_file_path)
    if not settings.excel_file_path.exists():
        raise FileNotFoundError(f"Excel file not found: {settings.excel_file_path}")

    df = pd.read_excel(settings.excel_file_path)
    if df.empty:
        logging.warning("Excel file %s is empty.", settings.excel_file_path)

    if settings.columns:
        missing_cols = [col for col in settings.columns if col not in df.columns]
        if missing_cols:
            raise RuntimeError(
                "Configured columns are missing from Excel file: "
                f"{', '.join(missing_cols)}"
            )
        df = df[[col for col in settings.columns]]

    df = df.fillna("")
    logging.info("Loaded %d rows and %d columns from Excel.", len(df), len(df.columns))
    return df


def clear_worksheet(worksheet: gspread.Worksheet) -> None:
    logging.info("Clearing worksheet %s", worksheet.title)
    worksheet.clear()


def build_value_matrix(df: pd.DataFrame) -> List[List[str]]:
    headers = list(df.columns)
    rows = df.astype(str).values.tolist()
    return [headers, *rows]


def chunk_rows(rows: Sequence[Sequence[str]], chunk_size: int) -> Iterable[Sequence[Sequence[str]]]:
    for index in range(0, len(rows), chunk_size):
        yield rows[index : index + chunk_size]


def upload_data(worksheet: gspread.Worksheet, rows: Sequence[Sequence[str]]) -> None:
    logging.info("Uploading %d rows (including header) to worksheet %s", len(rows), worksheet.title)
    start_row = 1
    for chunk in chunk_rows(rows, CHUNK_ROW_COUNT):
        start_cell = rowcol_to_a1(start_row, 1)
        logging.debug("Updating range starting at %s with %d rows", start_cell, len(chunk))
        worksheet.update(start_cell, chunk, value_input_option="USER_ENTERED")
        start_row += len(chunk)


def main(argv: Sequence[str]) -> int:
    args = parse_args(argv)
    configure_logging(args.verbose)

    try:
        settings = Settings.from_env()
        logging.debug("Loaded settings: %s", settings)
        df = read_excel_data(settings)
        value_matrix = build_value_matrix(df)

        if args.dry_run:
            logging.info("Dry run enabled; skipping Google Sheet updates.")
            return 0

        client = authenticate(settings)
        worksheet = get_worksheet(client, settings.spreadsheet_id, settings.worksheet_name)

        clear_worksheet(worksheet)
        upload_data(worksheet, value_matrix)
        logging.info("Data transfer completed successfully.")
        return 0
    except Exception as exc:  # noqa: BLE001
        logging.error("Automation failed: %s", exc, exc_info=args.verbose)
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
