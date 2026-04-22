# Fire Inspection PDF Automation and Dashboarding

An end-to-end data automation and reporting project that extracts structured information from fire inspection PDF reports, transforms it into analysis-ready tabular data, and supports dashboard-based monitoring for operational and compliance review.

This project was designed to reduce manual data entry, improve reporting consistency, and make semi-structured inspection data easier to analyze at scale.

---

## Overview

This project automates the extraction of fire inspection data from PDF reports and prepares it for dashboard reporting and analysis.

---

## Key Capabilities

- Extracts inspection findings from PDF reports automatically
- Converts semi-structured report content into clean Excel output
- Standardizes inconsistent fields across reports
- Handles missing values and formatting issues during transformation
- Produces dashboard-ready data for Power BI reporting
- Supports tracking of violations, due dates, hazard levels, and abatement status
- Improves visibility into inspection trends across locations

---

## Problem Statement

Fire inspection reports often exist as PDF files, which makes large-scale analysis difficult. Important details such as violations, descriptions, code references, due dates, and hazard classifications are usually embedded in semi-structured text rather than stored in a database.

Manual extraction of this information is:

- Time-consuming
- Repetitive
- Difficult to scale
- Vulnerable to human error

This project automates the process of converting PDF-based inspection reports into structured datasets that can be used for monitoring, reporting, and decision-making.

---

## Project Goal

The goal of this project is to build a repeatable workflow that:

- Reads inspection reports in PDF format
- Extracts the important compliance-related fields
- Cleans and standardizes the output
- Stores the results in Excel format
- Feeds the cleaned data into a reporting dashboard

This creates a more efficient pipeline for turning document-based inspection records into usable operational insights.

---

## Workflow Overview

```text
PDF Inspection Reports
        ↓
Python Extraction Script
        ↓
Data Cleaning and Standardization
        ↓
Structured Excel Output
        ↓
Shared Storage / Document Platform
        ↓
Power BI Dashboard
```

---

## Extracted Fields

The automation pipeline captures and organizes fields such as:

- Violation
- Description
- Code Reference
- Hazard Classification
- Due Date
- Facility
- Building
- Campus / Location
- Region Code
- Building Number
- Housing / Dorm Indicator
- Abatement / Status-related details

These fields are cleaned and aligned into a consistent structure for downstream reporting.

---

## Tools and Technologies

- **Python** for automation and data transformation
- **pandas** for cleaning and structuring extracted data
- **pdfplumber** for reading PDF report content
- **openpyxl** for writing formatted Excel outputs
- **Excel** for structured output delivery
- **Power BI** for dashboarding and visualization
- **SharePoint / Shared Storage** for hosting output files and supporting report access

---

## Data Processing Approach

The project follows a document-to-dashboard pipeline:

### 1. PDF Ingestion

Inspection reports are collected from a designated input location and processed through Python.

### 2. Text Extraction

The script reads report content and identifies relevant sections containing inspection findings.

### 3. Field Mapping

Important report labels are mapped into a standardized tabular structure.

### 4. Data Cleaning

The extracted values are cleaned to improve consistency across files.

### 5. Excel Output

The final output is written into a structured Excel file in table format so it can be directly used for reporting.

### 6. Dashboard Integration

The cleaned output serves as the source dataset for dashboard-based monitoring and analysis.

---

## Data Challenges Addressed

This project handles several practical issues commonly found in PDF-based reporting workflows, including:

- Inconsistent report layouts
- Blank or missing fields
- Mixed date formats
- Semi-structured text blocks
- Naming differences across reports
- Incomplete hazard classification values
- Inconsistent facility and building labeling
- Extraction noise from nearby sections in the PDF

Addressing these issues was important to make the final dataset reliable enough for dashboard use.

---

## Reporting and Dashboard Use Case

Once the data is cleaned and structured, it can be used in a dashboard to support analysis such as:

- Total number of violations
- Open vs. closed findings
- Overdue items
- Hazard classification breakdown
- Violations by campus or location
- Trends over time
- Top locations with repeated issues
- Summary views for leadership or operational review

This makes it easier for users to move from document review to data-driven monitoring.

---

## Why This Project Matters

This project demonstrates how semi-structured compliance documents can be transformed into operational reporting assets.

It adds value by:

- Reducing manual work
- Improving data consistency
- Supporting faster reporting
- Making inspection records easier to analyze
- Creating a scalable process for repeated reporting cycles

It also shows practical experience in data extraction, cleaning, automation, and dashboard integration.

---

## Sample Architecture

```text
Inspection PDF Reports
        │
        ▼
Python Extraction Script
(pdfplumber + pandas)
        │
        ▼
Data Cleaning and Standardization
        │
        ▼
Excel Output
(openpyxl)
        │
        ▼
Shared Storage / Document Platform
        │
        ▼
Power BI Dashboard
        │
        ▼
User Access through Dashboard Interface
```

---

## Output

The project generates a cleaned Excel dataset that is:

- Structured for analysis
- Consistent across processed reports
- Ready for Power BI ingestion
- Easier to review than raw PDFs

The output can be refreshed as new reports are processed, making it useful for ongoing monitoring.

---

## Example Project Structure

```text
fire-inspection-automation/
│
├── data/
│   ├── input/              # source PDF reports
│   └── output/             # cleaned Excel outputs
│
├── scripts/
│   └── fire_inspection_extractor.py
│
├── dashboard/
│   └── power_bi_notes/     # documentation or dashboard-related notes
│
├── requirements.txt
└── README.md
```

---

## Future Scope

Possible future improvements for this project include:

- Scheduled automatic script execution
- Cloud or server-based hosting instead of a single local machine
- More reliable refresh integration with reporting tools
- Automated alerts for overdue or high-risk findings
- Expanded data validation and logging
- A simple chatbot or search assistant for users to query inspection data
- Stronger dashboard self-service features for non-technical users

---

## What This Project Demonstrates

This project highlights skills in:

- Python automation
- PDF data extraction
- Data cleaning and transformation
- Excel output generation
- Workflow design
- Dashboard data preparation
- Reporting-oriented problem solving

It reflects a practical business use case where raw documents are converted into usable decision-support data.

---

## Confidentiality Note

This repository presents the project in a generalized and anonymized form.

To protect sensitive information, it does not include:

- Confidential inspection reports
- Internal file paths
- Restricted dashboard links
- Private organizational data
- Non-public operational details

The focus of this repository is the workflow, technical approach, and problem-solving process rather than the underlying confidential data.

---

## Author

**Karteek Abothula**

