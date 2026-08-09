# Apptracker v7.3 Architecture

## Architecture summary

Apptracker v7.0 separates the existing application workflow into focused Python modules while retaining `apptracker.py` as the stable production entry point. The refactor is organizational: it improves clarity and maintainability without redesigning the business workflow or changing the established outputs. v7.1 builds on that foundation with a platform-aware fix in `collector.py`, v7.2 follows with a Windows-compatibility cleanup in `notify.py`, and v7.3 adds filename validation logging and a Validation Dashboard layout improvement, all described below.

## System layers

```text
User and Scheduling Layer
├── Apptracker.app
├── Apptracker.command
└── launchd via com.apptracker.plist
             │
             ▼
Application Orchestration Layer
└── apptracker.py
             │
             ▼
Application Services
├── audit.py        Audit logging and exception hook
├── collector.py    PDF collection and record creation
├── config.py       Configuration and folder validation
├── dashboards.py   Excel dashboard creation
├── demo.py         Optional demo-workflow integration
├── notify.py       User and administrator emails
├── report.py       Workbook and CSV creation
└── validation.py   Output reconciliation and history

Shared and Supporting Components
├── make_demo_data.py   Demo-data generator
├── xlsx_utils.py        Shared worksheet helper functions
├── requirements.txt    Python dependency definition
└── Settings.txt        Runtime configuration
```

## Execution flow

```text
apptracker.py
  │
  ├─ initialize Audit
  ├─ read and validate configuration
  ├─ collect Active and Rejected PDF records
  ├─ create Apptracker.xlsx and CSV exports
  ├─ build reporting dashboard
  ├─ validate PDF, workbook, and CSV results
  ├─ update Validation.xlsx and validation dashboard
  ├─ send configured notifications
  └─ optionally run demo-data generation
```

The audit component supports the entire sequence rather than occupying one isolated workflow step. It records major events, elapsed time, execution status, and failure details throughout the run.

## System Workflow

The v7.0 modules implement the same operational sequence established in v6.0:

```text
                     Settings.txt
                          │
                          ▼
                  Load Configuration
                          │
                          ▼
               Validate Folder Paths
                          │
                          ▼
           Scan Active & Rejected Folders
                          │
                          ▼
             Extract PDF Metadata
       (Filename, Created Date, Timestamp)
                          │
                          ▼
                Build Application Lists
                          │
          ┌───────────────┴───────────────┐
          ▼                               ▼
 Generate Apptracker.xlsx            Generate CSV Files
          │                               │
          └───────────────┬───────────────┘
                          ▼
                 Validate All Output
                          │
                          ▼
               Generate Validation.xlsx
                          │
                          ▼
                  Write Audit.csv
                          │
                          ▼
            Send User Report Email (Optional)
                          │
                          ▼
           Generate Demo Workbook (Optional)
                          │
                          ▼
          Send Admin Summary Email (Optional)
                          │
                          ▼
                  Application Complete
```

## Entry point: `apptracker.py`

`apptracker.py` is the production application boundary. It prepares the working directory, initializes `Audit`, and calls `_run()` for the business workflow.

`main()` owns the top-level failure boundary. If an unexpected error occurs, it updates `audit.run_summary` with the failure status, completion time, elapsed time, and error message, then performs a bare `raise`. Re-raising preserves the original exception and traceback so the uncaught-exception hook installed by `Audit` can complete its normal error handling.

This structure keeps the orchestration readable while allowing the audit system to observe both successful and failed runs.

## Configuration: `config.py`

`config.py` reads `Settings.txt` and builds the `Config` object used by the other modules. It centralizes runtime values rather than making each module parse configuration independently.

The configuration object contains, as applicable:

- Active, Rejected, and Apptracker folder paths
- Output and audit file locations
- Email recipient and SMTP settings
- Credential store references (used by `keyring`)
- Demo-data flags and options
- Version information used by reporting and audit output

The module validates that the required folders exist before processing begins. If they do not, the workflow exits cleanly as a configuration correction rather than treating the condition as an application-processing failure.

## Audit and error handling: `audit.py`

The `Audit` class provides operational traceability for an Apptracker run. It:

- Tracks start time and elapsed time
- Writes step-by-step entries to `Audit.csv`
- Maintains the `run_summary` consumed by the administrator email
- Stores success and failure state
- Installs the uncaught-exception handler for the process

The exception design relies on cooperation between `apptracker.py` and `audit.py`. The entry point first records the failure state, then re-raises the exception. That leaves the original exception uncaught at process level, allowing the Audit `sys.excepthook` to display the expected error banner and traceback while Python returns a failure exit code.

## Collection: `collector.py`

`collector.py` reads the Active and Rejected PDF folders and constructs the application records used downstream. It obtains information from the filename and the file's creation timestamp, then assigns the application status from the source folder.

As of v7.1, timestamp retrieval is platform-aware: macOS uses the true file creation time, Windows uses its native creation-time attribute, and Linux falls back to last-modified time, since Linux does not expose a reliable creation timestamp.

The record set includes Work Type, Company, Position, Date, and Created Timestamp. The collector also produces the console summary used during a manual run.

As of v7.3, filenames that don't match the expected "Work Type - Company - Position" pattern still produce an application record (with blank fields for the missing parts), but their names are also collected and returned to the caller. `apptracker.py` logs the result to `Audit.csv`, so a naming problem surfaces as an audit warning instead of only appearing as unexplained blank cells in the generated workbook.

The zero-file scenario is explicitly supported. When both input folders are empty, the module reports zero values rather than calculating percentages against a zero total.

## Reporting: `report.py`

`report.py` transforms the collected record sets into the main reporting artifacts:

- `Apptracker.xlsx`
- `Active.csv`
- `Rejected.csv`

The workbook contains Summary, Active, and Rejected sheets. The Active and Rejected sheets use formatted Excel tables, while shared helpers provide consistent column sizing. `report.py` supplies the structured workbook content that `dashboards.py` enriches with dashboard visuals.

## Dashboard generation: `dashboards.py`

`dashboards.py` builds the Dashboard worksheet for both primary workbooks.

For `Apptracker.xlsx`, it creates:

- Work-type pie chart
- Monthly application bar chart
- Newest-ten application list

For `Validation.xlsx`, it creates:

- Active trend chart based on validation history
- Total trend chart based on validation history

Dashboard construction handles a zero-application dataset gracefully, so a valid empty-folder run still produces usable output.

### Chart data flow

The Application Dashboard uses the current run’s report data. Its work-type pie chart summarizes application categories, the monthly bar chart groups records by date, and the newest-ten list uses the timestamp-sorted application records.

The Validation Dashboard uses a different source: the persistent History worksheet in `Validation.xlsx`. Each successful validation contributes a point for active and total application counts, allowing the dashboard to show trend lines over multiple executions instead of only the current run.

As of v7.3, the Validation Dashboard hides columns N through Z after the charts and helper data are created. The AA:AC helper columns that feed the trend charts remain fully visible; hiding the empty columns between the charts and the helper data just brings them closer together visually.

## Validation: `validation.py`

`validation.py` is the reconciliation layer. It cross-checks PDF counts against Excel and CSV record counts, then writes or updates `Validation.xlsx`.

`Validation.xlsx` includes:

- **Summary** — results of the current output checks
- **History** — run-by-run validation history
- **Dashboard** — trend reporting built from History

This separation helps distinguish report creation from independent verification of the generated artifacts.

## Notifications: `notify.py`

`notify.py` delivers the configured email outputs using SMTP credentials retrieved through the `keyring` package, which resolves to the appropriate system credential store (macOS Keychain, Windows Credential Locker, or Secret Service on Linux). As of v7.2, console messages describe this generically rather than naming macOS Keychain specifically, and success messages avoid non-ASCII characters for Windows terminal compatibility.

The User Report email includes the primary workbook, validation workbook, and the two CSV exports. The Admin Summary email uses the Audit run summary and includes `Audit.csv` as its attachment.

## Demo components: `demo.py` and `make_demo_data.py`

The demo workflow has two distinct responsibilities:

- `make_demo_data.py` generates the sample data used for demonstrations and testing.
- `demo.py` integrates that generation into Apptracker by launching the generator as a subprocess when enabled in `Settings.txt`.

Keeping generation separate from integration prevents demo logic from becoming part of the core collection and reporting modules.

## Shared Excel helpers: `xlsx_utils.py`

`xlsx_utils.py` contains reusable worksheet utilities. The shared `autosize_columns()` helper is used by `report.py`, `validation.py`, and `dashboards.py`, keeping formatting behavior consistent and avoiding duplicated code.

## Launching and scheduling

`Apptracker.app` is the Finder-friendly production launcher. `Apptracker.command` supports manual execution with visible Terminal output. The launchd schedule is defined by `com.apptracker.plist` and managed through `Install Scheduler.command` and `Remove Scheduler.command`.

Because `apptracker.py` remained the entry point through the refactor, these launch and schedule mechanisms remain compatible with the production workflow.

## Dependency management

`requirements.txt` defines the Python dependencies needed to run the application. It makes environment setup repeatable and provides a single source of truth for required packages.

## Technology stack

| Area | Components |
| --- | --- |
| Language | Python 3 |
| Data and reporting | pandas, openpyxl, Excel, CSV |
| Credentials | keyring (macOS Keychain, Windows Credential Locker, or Secret Service) |
| Email | SMTP |
| macOS automation | Apptracker.app, Automator, launchd |
| Scheduling | `com.apptracker.plist` managed with the scheduler commands |

## Production status

| Item | Status |
| --- | --- |
| Release type | Production architectural refactor |
| Platform | macOS (primary), cross-platform for core reporting logic |
| Entry point | `apptracker.py` |

## Author

Thomas Barreto  
https://thomasbarreto.com/
