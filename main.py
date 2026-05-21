# main.py
"""
Execution orchestrator. Initializes the environment, enforces path
security, and triggers the processing pipeline.
"""
import logging
import sys
from pathlib import Path

from config import INPUT_DIR, LOG_FILE, LOGS_DIR, OUTPUT_DIR
from utils.excel_reader import process_workbook

def _setup_logger() -> logging.Logger:
    """Configures systemic logging with standard output and disk handlers."""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path: Path = LOGS_DIR / LOG_FILE

    fmt: str = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("panel_classifier")

def main() -> None:
    """Main execution sequence."""
    logger: logging.Logger = _setup_logger()

    for directory in (INPUT_DIR, OUTPUT_DIR, LOGS_DIR):
        directory.mkdir(parents=True, exist_ok=True)

    candidates: list[Path] = [
        p for p in INPUT_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in (".xlsx", ".xls") and not p.name.startswith("~$")
    ]

    if not candidates:
        logger.error("No valid Excel workbook found in the input directory. Terminating.")
        sys.exit(1)

    # Path traversal mitigation: Resolve absolute path strictly within INPUT_DIR
    target_file: Path = candidates[0].resolve()
    if INPUT_DIR.resolve() not in target_file.parents:
        logger.error("Path traversal anomaly detected. Terminating.")
        sys.exit(1)

    logger.info(f"Targeting workbook: {target_file.name}")
    success: bool = process_workbook(target_file, logger)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()