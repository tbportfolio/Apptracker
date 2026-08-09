# Apptracker Changelog

## v7.3 — Filename Validation and Dashboard Layout

### Added: filename validation logging

- `collector.py` previously parsed each PDF filename with `pdf.stem.split(" - ", 2)` and silently produced blank Work Type, Company, or Position fields when a filename didn't follow the expected pattern.
- `load_applications()` now collects the names of any malformed filenames and returns them alongside the application records.
- `collect_applications()` prints a console warning listing any malformed filenames.
- `apptracker.py` logs the result to `Audit.csv` as a `Filename Validation` step, `WARNING` with the file count and names if any are malformed, or `SUCCESS` if every filename matched the expected pattern.

### Changed: Validation Dashboard column layout

- `dashboards.py` now hides columns N through Z on the `Validation.xlsx` Dashboard tab after the charts and helper data are created.
- The AA:AC helper columns that feed the Active and Total trend charts are unaffected and remain fully visible; hiding the empty columns between the charts and the helper data simply brings them closer together visually.

### Verification

- Ran in production on macOS with no errors after both changes.

---

## v7.2 — Windows-Compatible Notifications

### Fixed: notify.py wording and output

- `notify.py` previously referenced "macOS Keychain" directly in console output, even though credential storage is handled by the `keyring` package, which already resolves to the correct backend on macOS, Windows, or Linux. Messages now describe this generically as the "system credential store."
- Replaced the `✓` checkmark character in success messages with `[OK]`, since some Windows terminal configurations can raise a `UnicodeEncodeError` on non-ASCII characters.

### Verification

- Ran in production on macOS with no errors after the change.

---

## v7.1 — Cross-Platform Timestamp Fix

### Fixed: PDF creation date retrieval

- `collector.py` previously read PDF creation timestamps using `st_birthtime`, a macOS-only file attribute, which would raise an error on other platforms.
- Timestamp retrieval is now platform-aware: macOS uses `st_birthtime`, Windows uses `st_ctime` (its native creation-time attribute), and Linux falls back to `st_mtime` (last modified), since Linux does not expose a true creation timestamp.
- Scheduling (`launchd`) and the Finder-based launcher remain macOS-specific; this fix applies to the core reporting logic only.

### Verification

- Confirmed matching output against existing production data after the change.

---

## v7.0 — Modular Architecture Release

### Summary

v7.0 is an architectural refactor of the established Apptracker application. It preserves the existing application-tracking workflow and production outputs while reorganizing the code into smaller, purpose-specific modules.

### Architecture

- Retained `apptracker.py` as the production entry point and orchestration layer.
- Split responsibilities into `config.py`, `audit.py`, `collector.py`, `report.py`, `dashboards.py`, `validation.py`, `notify.py`, `demo.py`, and `xlsx_utils.py`.
- Added `_run()` beneath the main application boundary so setup, error handling, and workflow orchestration are easier to follow.
- Kept the existing launch and scheduling workflow compatible with the production entry point.

### Configuration and dependencies

- Centralized settings parsing and required-folder validation through `config.py`.
- Added `requirements.txt` for repeatable Python dependency installation.
- Kept runtime settings in `Settings.txt`, including folders, email/SMTP values, Keychain references, and demo options.

### Reporting, validation, and dashboards

- Preserved `Apptracker.xlsx`, `Active.csv`, `Rejected.csv`, `Validation.xlsx`, and `Audit.csv` as the core output set.
- Preserved the reporting workbook’s Summary, Active, Rejected, and Dashboard views.
- Preserved validation reconciliation and historical trend reporting.
- Preserved user-report and administrator-summary email delivery.

### Demo support

- Documented `make_demo_data.py` as the standalone demo-data generator.
- Kept `demo.py` responsible for optional integration of demo-data generation into an Apptracker run.

### Fixed: audit exception handling

- Corrected the conflict between top-level exception handling and Audit’s `sys.excepthook`.
- On an unexpected failure, the entry point now records `FAILED` state, finish time, elapsed time, and the error message in the audit summary before re-raising the original exception.
- The original traceback reaches Audit’s uncaught-exception handler, preserving the expected error output and failure exit code.

### Fixed: empty Active and Rejected folders

- Corrected the zero-file scenario in which both input folders contain no PDFs.
- Added guards around percentage calculations to prevent division-by-zero errors.
- Confirmed reporting and dashboard generation complete normally for a zero-application run.

### Verification

- Verified against real production data with matching v6.0 and v7.0 results.
- Verified a normal success path.
- Verified an intentional failure path, including failed audit state and uncaught-exception handling.
- Verified the empty-folder path completes normally without errors.

---

## v6.0 — Production Release

### Added

- Scheduled automation using macOS launchd
- `Install Scheduler.command`
- `Remove Scheduler.command`
- `com.apptracker.plist`
- Fully unattended daily execution

This release completed the original production-ready automation system.

## v5.9

### Added

- `Apptracker.app` silent Automator launcher
- Custom macOS application icon
- Improved Finder integration

## v5.8

### Added

- Admin Summary Email
- Global exception handler
- Improved execution reporting
- Runtime summary

## v5.7

### Added

- Audit logging
- `Audit.csv`
- Execution timing
- `Apptracker.command` launcher

## v5.6

### Added

- External script automation
- Demo Data Generator integration
- Automatic execution of `make_demo_data.py`
- Script validation
- Success and failure reporting

## v5.5

### Added

- Validation Dashboard
- Historical validation charts

## v5.4

### Added

- Finder Created Date
- Created Timestamp
- Timestamp sorting
- Dashboard improvements

## v5.3a

### Improved

- Dashboard formatting
- Automatic worksheet sizing

## v5.3

### Added

- Dashboard worksheet
- Charts
- Summary reporting

## v5.2b

### Added

- Email-status console output

## v5.2a

### Added

- Folder validation
- Email configuration
- macOS Keychain integration

## v5.2

### Added

- `Validation.xlsx`
- `Settings.txt`
- Date column
- Version logging
- Execution timing

## v5.1

### Added

- Excel reporting
- CSV export

## v4.0

### Refactored

- Reorganized the application into a modular Python program
- Improved readability and maintainability

## v3.0

### Added

- Data export functionality

## v1.0–v2.0 — Initial Release

Apptracker began as a Python utility that counted PDF job applications stored in Active and Rejected folders. That proof of concept established the foundation for the later reporting, validation, dashboard, audit, notification, and scheduling capabilities.

## Project evolution

```text
v1–v2  PDF Counter
  │
  ▼
v3     CSV Export
  │
  ▼
v4     Program Refactor
  │
  ▼
v5.1   Excel Reporting
  │
  ▼
v5.2   Settings and Validation
  │
  ▼
v5.3   Application Dashboard
  │
  ▼
v5.4   Metadata and Timestamps
  │
  ▼
v5.5   Validation Dashboard
  │
  ▼
v5.6   Demo Data Automation
  │
  ▼
v5.7   Audit Logging
  │
  ▼
v5.8   Email Automation
  │
  ▼
v5.9   macOS Application
  │
  ▼
v6.0   Scheduled Automation
  │
  ▼
v7.0   Modular Architecture and Reliability Fixes
  │
  ▼
v7.1   Cross-Platform Timestamp Fix
  │
  ▼
v7.2   Windows-Compatible Notifications
  │
  ▼
v7.3   Filename Validation and Dashboard Layout
```

## Current status

**Project:** Apptracker  
**Status:** Production Release

Apptracker has evolved from a simple PDF counter into a macOS Python ETL automation system with configuration management, PDF metadata collection, Excel and CSV reporting, application and validation dashboards, audit logging, email notifications, demo support, and unattended scheduled execution.
