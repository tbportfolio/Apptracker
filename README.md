# Apptracker v7.3

A Python-based automation system for tracking job applications, generating reports, validating output, maintaining audit logs, sending email notifications, and running automatically on a daily schedule using macOS launchd.

---

## Overview

Apptracker was built to automate the management of PDF job applications stored in two folders:

- Active Applications
- Rejected Applications

The application scans both folders, generates reports, validates all output, creates an audit log, optionally emails results, and can run automatically every day without user intervention.

Current Version: **v7.3**

---

# Note

Launcher/scheduler files are used locally and not included in the GitHub repository:

- Apptracker.app
- com.apptracker.plist
- Apptracker.command
- Install Scheduler.command
- Remove Scheduler.command

This project was originally developed for macOS, and the scheduling 
system still uses macOS launchd. PDF creation dates were originally 
pulled using a macOS only file attribute; as of v7.1, this is 
platform-aware, using the native creation-time attribute on macOS and 
Windows, and falling back to last-modified time on Linux, since Linux 
has no reliable creation timestamp. As of v7.2, email notification 
messages no longer reference macOS Keychain by name, since credential 
storage is handled by the `keyring` package and resolves to the 
correct backend (Keychain, Windows Credential Locker, or Secret 
Service) automatically.

---

# Quick Start

Already have Apptracker installed?

Open Terminal and run:

```bash
cd ~/Documents/Jobs/Apptracker

source .venv/bin/activate

python apptracker.py

deactivate
```

To run the production version, simply double-click:

```
Apptracker.app
```

To run manually with a visible Terminal window, double-click:

```
Apptracker.command
```

If the scheduler has been installed, Apptracker will run automatically at the scheduled time without any user interaction.

---

# Features

- Count Active and Rejected applications
- Read application metadata from PDF filenames
- Capture Finder Created Date
- Generate Apptracker.xlsx
- Generate Active.csv
- Generate Rejected.csv
- Generate Validation.xlsx
- Validation Dashboard
- Application Dashboard
- Audit logging
- User Report Email
- Admin Summary Email
- Demo Data Generator
- macOS Automator launcher
- launchd scheduler
- Global exception handling
- Configuration using Settings.txt
- Modular application architecture
- requirements.txt dependency management

---

# Requirements

macOS

Python 3.x

Required packages

- pandas
- openpyxl
- keyring

---

# Folder Structure

Example

```
Jobs/

├── Active/
│      PDF resumes
│
├── Rejected/
│      Rejected PDFs
│
└── Apptracker/
       # Application
       Apptracker.app

       # Comma-separated values
       Active.csv
       Audit.csv
       Rejected.csv

       # Microsoft Excel workbooks
       Apptracker_Demo.xlsx
       Apptracker.xlsx
       Validation.xlsx

       # Plain text documents
       requirements.txt
       Settings.txt

       # Property list
       com.apptracker.plist

       # Python scripts
       apptracker.py
       audit.py
       collector.py
       config.py
       dashboards.py
       demo.py
       make_demo_data.py
       notify.py
       report.py
       validation.py
       xlsx_utils.py

       # Terminal shell scripts
       Apptracker.command
       Install Scheduler.command
       Remove Scheduler.command
```

---

# Installation

Create a Python virtual environment.

```bash
python3 -m venv .venv
```

Activate the environment.

```bash
source .venv/bin/activate
```

Install required packages.

```bash
pip install pandas

pip install openpyxl

pip install keyring
```

or

```bash
pip install -r requirements.txt
```

---

# Configuration

Edit **Settings.txt**

```
active_folder=

rejected_folder=

apptracker_folder=
```

Specify the full folder paths.

Email settings are optional.

```
email_enabled=False

admin_email_enabled=False
```

If email is enabled, configure SMTP settings.

```
smtp_username=

smtp_server=smtp.gmail.com

smtp_port=465
```

---

# Running Apptracker

Open Terminal.

Navigate to the project folder.

```bash
cd ~/Documents/Jobs/Apptracker
```

Activate Python.

```bash
source .venv/bin/activate
```

Run Apptracker.

```bash
python apptracker.py
```

Deactivate Python.

```bash
deactivate
```

---

# Running from Finder

Double-click

```
Apptracker.app
```

Runs silently.

---

Double-click

```
Apptracker.command
```

Displays the Terminal window while executing.

---

# Scheduler

Install

Double-click

```
Install Scheduler.command
```

The scheduler installs the launchd service.

Apptracker will automatically execute at the scheduled time defined in

```
com.apptracker.plist
```

---

Remove Scheduler

Double-click

```
Remove Scheduler.command
```

or

```bash
launchctl unload ~/Library/LaunchAgents/com.apptracker.plist
```

---

# Generated Output

Each execution creates

```
Apptracker.xlsx
```

Contains

- Summary
- Active
- Rejected
- Dashboard

---

```
Validation.xlsx
```

Contains

- Summary
- History
- Dashboard

---

CSV Files

```
Active.csv

Rejected.csv
```

---

Audit Log

```
Audit.csv
```

Records every execution step including

- timestamp
- elapsed time
- status
- execution details

---

# PDF Filename Format

Apptracker reads application metadata from each PDF filename. Name files using this format:

```text
Work Type [Hybrid/Onsite/Remote] - [Company] - [Position].pdf
```

Example:

```text
Remote - Example Company - Senior Program Manager.pdf
```

The collector uses the filename to populate Work Type, Company, and Position. It also records the Finder Created Date and Created Timestamp from the PDF file metadata.

As of v7.3, any filename that doesn't match this format still generates an application record (with blank fields for the missing parts), but is also flagged. A console warning lists the affected filenames, and `Audit.csv` logs a `Filename Validation` step, `WARNING` with the file count and names if any are malformed, or `SUCCESS` if every filename matched.

---

# Email Notifications

Optional.

### User Report

Can send

- Apptracker.xlsx
- Validation.xlsx
- Active.csv
- Rejected.csv

---

### Admin Summary

Sends

- execution summary
- timing
- audit results
- Audit.csv

SMTP passwords are securely stored using the system credential store (macOS Keychain, Windows Credential Locker, or Secret Service on Linux).

---

# Validation

Apptracker validates that

```
PDF Count

=

Excel Rows

=

CSV Rows
```

If every value matches

```
Validation PASSED
```

Otherwise

```
Validation FAILED
```

---

# Demo Data Generator

Optional.

Creates

```
Apptracker_Demo.xlsx
```

using randomized company names and job titles while preserving reporting structure.

---

# v7.0 Modular Architecture

v7.0 is an architectural refactor of the existing production workflow, not a feature or logic rewrite. The established PDF inputs, reporting outputs, dashboards, validation controls, audit logging, notifications, demo process, and scheduling workflow are retained.

`apptracker.py` remains the production entry point and orchestrates the run in this order:

```text
config → collect → report → validate → notify → demo
```

The refactor separates the code into focused modules:

| Module | Responsibility |
| --- | --- |
| `audit.py` | Writes `Audit.csv`, tracks run state and timing, and installs the uncaught-exception handler. |
| `collector.py` | Scans Active and Rejected PDFs and builds application records from filenames and macOS creation metadata. |
| `config.py` | Reads `Settings.txt`, creates configuration, and validates required folders. |
| `dashboards.py` | Creates the Application and Validation Dashboard charts. |
| `demo.py` | Optionally launches the demo-data process when enabled. |
| `make_demo_data.py` | Creates the demo data and demo workbook. |
| `notify.py` | Sends the User Report and Admin Summary emails. |
| `report.py` | Creates `Apptracker.xlsx`, `Active.csv`, and `Rejected.csv`. |
| `validation.py` | Reconciles PDF, Excel, and CSV counts and updates `Validation.xlsx`. |
| `xlsx_utils.py` | Provides shared Excel helpers, including column sizing. |

The Application Dashboard continues to include a work-type pie chart, monthly bar chart, and newest-ten list. The Validation Dashboard continues to use `Validation.xlsx` History data to chart active and total trends across runs.

---

# v7.0 Reliability Fixes

## Exception and audit interaction

The application error boundary now records failure information in `audit.run_summary` before re-raising the original exception. Re-raising preserves the original traceback and allows Audit’s `sys.excepthook` to handle an uncaught failure correctly.

## Empty-folder handling

An empty Active and Rejected folder pair is a valid zero-application run. Percentage calculations now guard against division by zero, and report and dashboard generation complete normally when no PDF files are present.

## Verification

v7.0 was verified against real production data with matching v6.0 and v7.0 results. The normal execution path, a forced failure path, and the zero-file path were also tested.

---

# Troubleshooting

### Folder not found

Verify the folder paths in

```
Settings.txt
```

---

### Email not sending

Verify

- SMTP username
- SMTP password
- Keychain credentials
- Internet connection

---

### Scheduler not running

Verify

```
com.apptracker.plist
```

is installed.

Reload if necessary.

---

### Validation Failed

Compare

- PDF count
- Excel rows
- CSV rows

All values should match.

---

# launchd Scheduler Troubleshooting

If Apptracker does not run automatically at the scheduled time, use the following commands to verify the macOS launchd configuration.

## Step 1 – Verify the scheduler is loaded

Open Terminal and run:

```bash
launchctl list | grep apptracker
```

Expected output:

```text
-    0    com.apptracker
```

If nothing is returned, the scheduler is not currently loaded.

---

## Step 2 – Verify the plist file exists

```bash
ls ~/Library/LaunchAgents | grep apptracker
```

Expected output:

```text
com.apptracker.plist
```

---

## Step 3 – View the scheduler configuration

Display the actual launchd plist file.

```bash
cat ~/Library/LaunchAgents/com.apptracker.plist
```

This displays the complete XML configuration exactly as launchd reads it.

To view the parsed configuration in a more readable format, run:

```bash
plutil -p ~/Library/LaunchAgents/com.apptracker.plist
```

This displays the schedule, program path, `RunAtLoad` setting, and other launchd configuration values.

---

## Step 4 – Reload the scheduler

After making any changes to **com.apptracker.plist**, reload the job.

Unload:

```bash
launchctl unload ~/Library/LaunchAgents/com.apptracker.plist
```

Load:

```bash
launchctl load ~/Library/LaunchAgents/com.apptracker.plist
```

Verify:

```bash
launchctl list | grep apptracker
```

---

## Step 5 – Test immediately

To verify that Apptracker launches correctly without waiting for the scheduled time, temporarily change the plist to:

```xml
<key>RunAtLoad</key>
<true/>
```

Then reload the scheduler:

```bash
launchctl unload ~/Library/LaunchAgents/com.apptracker.plist

launchctl load ~/Library/LaunchAgents/com.apptracker.plist
```

If **Apptracker.app** launches immediately, the application is functioning correctly and the scheduler is installed. Any remaining issue is likely related to the scheduled trigger rather than the application itself.

---

# Version History

| Version | Description |
|----------|-------------|
| v1–v2 | PDF counter |
| v3 | CSV export |
| v4 | Program refactor |
| v5.1 | Excel reporting |
| v5.2 | Settings.txt, Validation |
| v5.2a | Folder validation, Email configuration |
| v5.2b | Email status console |
| v5.3 | Dashboard |
| v5.3a | Dashboard formatting |
| v5.4 | Finder Created Date & Timestamp |
| v5.5 | Validation Dashboard |
| v5.6 | Demo Data Generator automation |
| v5.7 | Audit Logging & Apptracker.command |
| v5.8 | Admin Summary Email & Global Exception Handling |
| v5.9 | Apptracker.app, Custom Icon, and Finder integration |
| v6.0 | launchd Scheduled Automation |
| v7.0 | Modular architecture, exception/audit fix, and empty-folder handling |
| v7.1 | Cross-platform PDF creation-date retrieval |
| v7.2 | Windows-safe notification wording and output |
| **v7.3** | Filename validation logging; Validation Dashboard column layout |

---

# Author

Thomas Barreto

Website: https://thomasbarreto.com/

Python ETL Reporting System

2026

# License
This repository is shared for portfolio purposes. All rights reserved, 
this code is not licensed for reuse or distribution.