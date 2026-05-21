"""
Configuration and operational constants for the panel-classifier.
Establishes structural boundaries and security limits.
"""
from pathlib import Path
import decimal

# Directory Mapping
BASE_DIR: Path = Path(__file__).resolve().parent
INPUT_DIR: Path = BASE_DIR / "data"
OUTPUT_DIR: Path = BASE_DIR / "output"
LOGS_DIR: Path = BASE_DIR / "logs"

# File Assignments
LOG_FILE: str = "run_log.log"
OUTPUT_FILE: str = "classified_panels.xlsx"
EXCEPTIONS_FILE: str = "exceptions.xlsx"

# Excel Structural Constants
PANEL_ROW: int = 2
HEADER_ROW: int = 4
PANEL_START_COL: int = 6  # F Column mapping
END_MARKER: str = "EXISTING COST"

# Security Boundaries (Resource Exhaustion Mitigation)
MAX_ROWS: int = 50000
MAX_COLS: int = 200

# Financial Precision Context
decimal.getcontext().prec = 10