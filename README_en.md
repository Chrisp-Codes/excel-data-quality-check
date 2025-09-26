# Excel Data Quality Check (POC)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Status](https://img.shields.io/badge/status-POC-orange)
[![README Deutsch](https://img.shields.io/badge/README-Deutsch-informational?style=flat-square)](README.md)


## Overview
This repository contains a **VBA macro** as a proof of concept for automated data quality checks in Excel tables.  
The macro scans rows, verifies mandatory fields, and creates a report highlighting missing data.

## Features
- Opens and processes an Excel sheet (`Data`)
- Checks mandatory fields (configurable in the code)
- Generates a structured `Report` sheet
- Marks missing fields with **X**
- Optional export as PDF for easy sharing
- Fully written in **VBA** (works in Excel without additional dependencies)

## Motivation
Manual checks in Excel can be time-consuming and error-prone.  
This macro demonstrates how simple automation can reduce effort and minimize mistakes.

## Usage
1. Open Excel and press `ALT + F11` to access the VBA editor.  
2. Insert a new module and paste the code from [`src/DataQualityCheck.vba`](src/DataQualityCheck.vba).  
3. Run the macro `DataQualityCheck`.  
4. The macro will generate a new sheet called **Report** and optionally export it as PDF.  

> **Note:** The file uses the `.vba` extension so GitHub correctly classifies it as VBA.  
> In Excel, simply copy the code into a module – the extension has no impact on execution.

## Example Data
To make testing easier, an example file is provided:

- **dummy_data.xlsx**  
  - Sheet name: `Data`  
  - Contains 5 rows with fake names and mixed empty fields  

Run the macro on this file to see how missing fields are marked with **X** in the generated report.

## Status
- Proof of Concept (POC)  
- Not intended for production use  
- Tested with dummy data only  

## Technologies
- VBA (Visual Basic for Applications)  
- Microsoft Excel  

## License
This project is licensed under the MIT License.  
See the [LICENSE](LICENSE) file for details.
