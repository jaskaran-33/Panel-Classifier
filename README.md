# Panel Classifier

## Overview

Panel Classifier is a Python-based utility designed to parse master Bill of Materials (BoM) or panel schedules in Excel format. It consolidates and categorizes panel series into individual worksheets while strictly preserving existing structural elements, static columns, formulas, and formatting. The architecture adheres to SOLID principles, specifically the Open-Closed Principle (OCP), isolating data extraction, classification logic, and validation into independent, extensible modules.

## Features

* **Dynamic Series Detection:** Automatically identifies panel columns starting from Column F (Row 1) up to the explicit "EXISTING COST" boundary marker.
* **Structural Integrity Retention:** Preserves all master formulas, cell colors (e.g., green panel headers), and static category rows (Rows 1–4).
* **Automated Row/Column Pruning:** Deletes rows with zero or null values for the active panel series and strips unrelated panel columns without mutating the source schema.
* **Triple-Output Generation:**
* `classified_panels.xlsx`: The final workbook containing isolated sheets for each panel series.
* `exceptions.xlsx`: A diagnostic report of sheets lacking panels or encountering structural errors.
* `run_log.log`: A comprehensive execution trace.


* **Cryptographic-Style Validation:** Implements a post-processing verification check to ensure the numeric sum of all panel quantities in the output strictly matches the master source.

## Project Structure

```text
panel-classifier/
├── config.py                 # Core constants and directory mappings
├── main.py                   # Entry point and runtime orchestration
├── requirements.txt          # Dependency manifest
└── utils/
    ├── __init__.py
    ├── classifier.py         # OCP-compliant boundary detection and row/col pruning logic
    ├── excel_reader.py       # Workbook ingestion and routing
    ├── exporter.py           # Exception data formatting and disk writing
    └── validator.py          # Post-execution quantity verification

```

## Prerequisites

* Python 3.9+
* `openpyxl`

## Installation

1. Clone the repository.
2. Install dependencies:
```bash
pip install -r requirements.txt

```



## Execution

1. Place the target Excel workbook (`.xlsx`) in the `input/` directory. Ensure Row 1 contains panel headers starting at Column F, terminating with an "EXISTING COST" merged cell.
2. Execute the orchestration script:
```bash
python main.py

```


3. Retrieve outputs from the `output/` directory and execution traces from the `logs/` directory.