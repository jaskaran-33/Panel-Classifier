"""
Generates diagnostic logs and exception reports.
File is unconditionally generated to mirror BoM-Python schema.
"""
from pathlib import Path
import logging
import openpyxl

from config import OUTPUT_DIR, EXCEPTIONS_FILE


def export_exceptions(exceptions: list[dict[str, str]], logger: logging.Logger) -> None:
    """
    Compiles operational anomalies into a standardized format.
    Mirrors BoM-Python structure. Unconditionally writes the file.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Exceptions"

    static_fields = ["DESCRIPTION", "SPEC", "MAKE", "UNIT", "CAT NO."]
    header = ["Sheet", "Row", "Issue", "Description", ""] + static_fields
    ws.append(header)

    if not exceptions:
        logger.info("No exceptions generated.")
    else:
        for exc in exceptions:
            row_data = [
                           exc.get("sheet", ""),
                           exc.get("row", ""),
                           exc.get("issue", ""),
                           exc.get("description", ""),
                           ""
                       ] + [""] * len(static_fields)
            ws.append(row_data)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path: Path = OUTPUT_DIR / EXCEPTIONS_FILE
    wb.save(out_path)
    logger.info(f"Exceptions log written to {out_path}")