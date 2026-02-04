# Compliance Document Automation

## Overview

This project is a desktop automation tool designed to optimise the preparation of compliance datasheets for regulatory submissions within a supply chain environment.

The application automates the process of locating, validating, and organising material datasheets required for regulatory applications (e.g., FANR submissions), significantly reducing manual preparation time and minimising documentation errors.

---

## Business Problem

Preparing compliance submission folders manually required:

- Searching for datasheet PDFs across multiple directories
- Identifying missing documentation
- Renaming duplicate files
- Organising files per shipment

This process was time-consuming and error-prone.

---

## Solution

The FANR Datasheet Checker automates this workflow by:

- Identifying required datasheets
- Detecting missing documents
- Generating product links
- Copying files into structured submission folders

---

## How It Works

### 1. Select Excel File
Upload an Excel file (.xls) containing materials requiring compliance documentation.

Only relevant materials should be included.

### 2. Choose Datasheets Folder
Select the master directory containing all available datasheet PDFs.

### 3. Set Destination Folder
Choose the output folder where submission documents will be generated.

### 4. Check Missing Datasheets
The system scans materials and identifies missing PDFs.

Missing materials are listed in the output panel with generated product links.

### 5. Download Missing Files
Users can open links directly in a browser to retrieve required datasheets.

### 6. Copy Datasheets
Once all files are present, the tool automatically copies required PDFs into the destination submission folder.

---

## Features

- Automated folder generation
- Datasheet validation
- Missing document detection
- Product link generation
- Batch file copying
- Submission-ready folder structuring

---

## Tech Stack

- Python
- File system automation
- Excel processing
- GUI integration

---

## Use Case

Developed to support regulatory compliance workflows in a supply chain setting, improving efficiency and reducing manual documentation workload.

