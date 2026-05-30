"""Batch processing entrypoint for PDF invoice extraction and Excel export."""
from __future__ import annotations

import logging
from pathlib import Path
from typing import List

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from pdf_extractor import PDFInvoiceExtractor

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


def process_pdfs(input_dir: Path) -> pd.DataFrame:
    """Scan the input directory for PDF files, extract invoice data, and return a cleaned DataFrame."""
    input_dir = input_dir.expanduser().resolve()
    pdf_files = sorted(input_dir.glob("*.pdf"))

    if not pdf_files:
        logger.warning("No PDF files found in %s", input_dir)
        return pd.DataFrame(columns=["Filename", "Date", "Amount", "Vendor"])

    records: List[dict] = []
    total_files = len(pdf_files)

    for index, pdf_path in enumerate(pdf_files, start=1):
        logger.info("Processing %d/%d: %s", index, total_files, pdf_path.name)
        extractor = PDFInvoiceExtractor(str(pdf_path))
        extracted = extractor.extract_data()

        records.append(
            {
                "Filename": pdf_path.name,
                "Date": extracted.get("date"),
                "Amount": extracted.get("amount"),
                "Vendor": extracted.get("vendor"),
            }
        )

    df = pd.DataFrame(records)
    return normalize_dataframe(df)


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize DataFrame types and handle missing values gracefully."""
    if df.empty:
        return df

    df = df.copy()

    # Convert amounts to float. Invalid values become NaN.
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    # Convert dates to datetime and leave invalid or missing values as NaT.
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # Keep Vendor text as a stripped string, use blank for missing vendor names.
    df["Vendor"] = df["Vendor"].astype("string").str.strip().replace({"<NA>": ""})

    return df


def export_to_excel(df: pd.DataFrame, output_path: Path) -> None:
    """Export the DataFrame to Excel and apply professional formatting with openpyxl."""
    output_path = output_path.expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Save initial DataFrame to Excel with pandas using openpyxl engine.
    df.to_excel(output_path, index=False, engine="openpyxl")

    # Reopen the workbook with openpyxl to apply styling.
    workbook = load_workbook(output_path)
    sheet = workbook.active

    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="FFDDDDDD", end_color="FFDDDDDD", fill_type="solid")
    center_alignment = Alignment(horizontal="center", vertical="center")

    # Style the header row: bold and light gray fill.
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment

    # Determine maximum width per column for auto-sizing.
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value is None:
                continue
            cell_value = str(cell.value)
            max_length = max(max_length, len(cell_value))
        adjusted_width = max_length + 2
        sheet.column_dimensions[column].width = adjusted_width

    # Apply number formatting to the Amount column if it exists.
    amount_column = None
    for cell in sheet[1]:
        if str(cell.value).strip().lower() == "amount":
            amount_column = cell.column_letter
            break

    if amount_column:
        for row in range(2, sheet.max_row + 1):
            amount_cell = sheet[f"{amount_column}{row}"]
            amount_cell.number_format = "$#,##0.00"

    # Center-align the Date column.
    date_column = None
    for cell in sheet[1]:
        if str(cell.value).strip().lower() == "date":
            date_column = cell.column_letter
            break

    if date_column:
        for row in range(2, sheet.max_row + 1):
            date_cell = sheet[f"{date_column}{row}"]
            date_cell.alignment = center_alignment

    workbook.save(output_path)
    logger.info("Exported styled Excel report to %s", output_path)


def main() -> None:
    input_dir = Path("./data/input_pdfs")
    output_path = Path("./data/output/extracted_report.xlsx")

    logger.info("Starting batch PDF extraction")
    df = process_pdfs(input_dir)

    if df.empty:
        logger.info("No extracted records to export. Exiting.")
        return

    export_to_excel(df, output_path)
    logger.info("Batch processing completed successfully.")


if __name__ == "__main__":
    main()
