# Automated CSV Data Cleaner & Excel Report Pipeline

A robust, fault-tolerant Python pipeline designed to ingest unformatted raw CSV files, resolve common data anomalies (whitespace, duplicate entries, malformed currency strings, inconsistent dates), and export a multi-sheet, executive-ready Excel workbook (.xlsx).

## Business Value

Raw data exported from CRMs, ERPs, or web platforms often suffers from syntax errors, missing fields, and arbitrary formatting. This automated pipeline removes manual spreadsheet intervention by:

Processing uncleaned raw data in seconds with built-in fault tolerance against corrupted rows.

Standardizing data types (converting currency strings into numeric floats, normalizing dates to ISO formats).

Generating high-level business KPIs and summary reports automatically.

## Features & Technical Implementation

Header Normalization: Strips irregular spacing, converts headers to lowercase, and standardizes names with underscores.

Whitespace & Enclosure Resilience: Handles leading/trailing spaces and quoted string delimiters using skipinitialspace.

Currency & Numeric Parsing: Strips currency symbols ($, ,, quotes), casting string representations into numeric floats for financial analysis.

Date Normalization: Parses mixed date representations into a uniform ISO YYYY-MM-DD format.

Fault Tolerance: Uses on_bad_lines='skip' to ensure pipeline continuity even when encountering malformed rows.

Multi-Sheet Excel Export: Uses pandas.ExcelWriter with openpyxl to separate row-level cleaned records from aggregated metrics.

# Quickstart

### 1. Prerequisites & Installation

Ensure you have Python 3 installed, then install the required dependencies:
```bash
pip install pandas openpyxl
```

### 2. Usage

Standard Execution (looks for raw_sales_data.csv in script directory):
```bash
python data_cleaner.py
```

Custom File Target via CLI:
```bash
python data_cleaner.py path/to/your_raw_file.csv
```

## Output Structure

The generated cleaned_sales_report.xlsx workbook contains two dedicated sheets:

Cleaned_Data: The standardized dataset containing clean customer names, ISO formatted dates, parsed quantities, numeric unit prices, and calculated total_amount values per row.

Customer_Summary: An aggregated executive summary grouping total orders, total units purchased, and total revenue per customer, sorted by revenue descending.

### Author: Computer & Automation Engineering Student (Politecnico di Bari). Focused on backend software development, data pipelines, and workflow automation.
