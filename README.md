# Excel VBA Data Quality & Profiling Tool

An Excel VBA-based data quality and profiling solution designed to analyse source datasets before data migration, transformation or reporting.

The tool combines **column-level data quality profiling** with **row-level exception reporting**, helping identify data issues quickly and trace them back to the affected source records.

## Key Features

- Data completeness analysis and missing-value detection
- Duplicate identification and uniqueness profiling
- Formatting inconsistency checks
- Invalid data identification
- UK postcode validation using VBA and pattern matching
- Email address validation and malformed email detection
- Automated row-level exception reporting
- Refreshable Email and Postcode validation reports
- Dynamic profiling of columns within the source data table
- VBA-driven drill-down from summary metrics to affected records

## Data Quality Dashboard

The **Quality Report** provides a column-level overview of the dataset, including metrics such as:

- Data type
- Completeness %
- Missing count
- Uniqueness %
- Duplicate count
- Formatting inconsistencies
- Invalid data %

This provides a quick way to assess the overall quality of a dataset before migration or further processing.

## Email Validation

The VBA email validation engine analyses non-blank email addresses and identifies malformed values.

Invalid records are automatically copied to the **InvalidEmails** worksheet, allowing the user to review the complete source records associated with each validation failure.

A **Refresh Invalid Emails** button reruns the validation against the latest source data.

Blank email addresses are handled separately by the Missing Data check.

## UK Postcode Validation

The postcode validation engine uses VBA and regular-expression pattern matching to validate UK postcode structures.

The validator normalises postcode formatting before validation and identifies non-blank values that do not conform to recognised UK postcode patterns.

Invalid records are copied automatically to the **InvalidPostcodes** worksheet.

A **Refresh Invalid Postcodes** button allows the validation report to be regenerated after the source data changes.

Blank postcodes are handled separately by the Missing Data check.

## Row-Level Exception Reporting

The workbook provides dedicated worksheets for investigating:

- Missing Data
- Duplicates
- Formatting issues
- Invalid Data
- Invalid Emails
- Invalid Postcodes

This allows summary-level data quality problems to be traced back to the individual affected records.

## How to Use

1. Open `Data_Quality_Portfolio.xlsm`.
2. Enable macros when prompted.
3. Add or replace data within the `ReportData` table.
4. Review the **Quality Report** for column-level data quality metrics.
5. Use the supporting exception worksheets to investigate individual records.
6. Use the refresh buttons on **InvalidEmails** and **InvalidPostcodes** to rerun the VBA validation routines.

## VBA Modules

The VBA source code is included separately in the repository for easy review:

- `QualityReportEngine.bas` – drives quality-report processing and detailed exception reporting
- `EmailValidation.bas` – performs email validation and generates the invalid-email report
- `PostcodeValidation.bas` – performs UK postcode validation and generates the invalid-postcode report
- `ClearQualityReport.bas` – clears generated report output when required

## Technologies

**Microsoft Excel | VBA | Regular Expressions | Data Profiling | Data Validation | Data Quality | Data Migration**

## Portfolio Version

This repository contains a portfolio demonstration version of the solution using **fictional test data**.

No confidential, client or production data is included.

---

### Author

**Hasnaad Din**  
Senior Data Analyst & Excel VBA Automation Engineer  
19+ years of experience in data analysis, automation, ETL and data migration.
