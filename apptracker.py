"""
Apptracker v7.3
Entry point. Wires config -> collection -> reporting -> validation ->
notifications -> demo data together. All business logic lives in the
other modules; this file is just the orchestration.
"""

import os
from datetime import datetime
from pathlib import Path

from config import build_config, validate_folders
from audit import Audit
from collector import collect_applications
from report import build_report, export_csv
from validation import validate_data, save_validation
from notify import send_user_report_email, send_admin_summary_email
from demo import run_demo_data_generator

VERSION = "v7.3"


def print_banner():
    print("\n=========================================")
    print("Apptracker")
    print(f"Version: {VERSION}")
    print("Python ETL Reporting System")
    print()
    print("Author: Thomas Barreto")
    print("Website: https://thomasbarreto.com/")
    print(f"Run Date: {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}")
    print("=========================================\n")


def main():

    # Ensure Apptracker always runs from its own directory.
    # Allows execution from VS Code, Terminal, or Finder launcher.
    app_dir = Path(__file__).resolve().parent
    os.chdir(app_dir)

    # Install the exception hook and start the run clock as early as
    # possible -- matches v6.0, which set sys.excepthook before Settings.txt
    # was even read.
    audit = Audit(VERSION)

    try:
        _run(app_dir, audit)

    except Exception as error:
        audit.run_summary["status"] = "FAILED"
        audit.run_summary["finish_time"] = datetime.now()
        audit.run_summary["elapsed"] = audit.elapsed_seconds()
        audit.run_summary["error"] = str(error)

        # Re-raise so it stays a genuinely uncaught exception at the top
        # level -- that's what lets Audit's sys.excepthook fire and print
        # the "UNEXPECTED ERROR" banner, and lets Python set exit code 1
        # on its own, same as v6.0 did.
        raise


def _run(app_dir, audit):

    print_banner()

    config = build_config(app_dir, version=VERSION)
    audit.set_audit_file(config.audit_file)

    # Create Apptracker folder if it doesn't exist
    config.apptracker_folder.mkdir(parents=True, exist_ok=True)

    # First audit entry
    audit.log("Application Started", "SUCCESS", f"Version {VERSION}")

    # Validate folder names and quit if ERROR
    if not validate_folders(config):
        return

    active_applications, rejected_applications, malformed_filenames = collect_applications(config)

    active = len(active_applications)
    rejected = len(rejected_applications)
    total = active + rejected

    if malformed_filenames:
        audit.log(
            "Filename Validation",
            "WARNING",
            f"{len(malformed_filenames)} file(s) did not match expected naming pattern: "
            + ", ".join(malformed_filenames)
        )
    else:
        audit.log(
            "Filename Validation",
            "SUCCESS",
            "All filenames matched expected pattern"
        )

    # -------------------------------------------------------
    # Excel report + CSV export
    # -------------------------------------------------------
    excel_file = build_report(active_applications, rejected_applications, config)

    audit.log(
        "Apptracker.xlsx Created",
        "SUCCESS",
        f"{active} Active, {rejected} Rejected"
    )

    active_csv, rejected_csv = export_csv(active_applications, rejected_applications, config)

    audit.log(
        "CSV Export Completed",
        "SUCCESS",
        "Active.csv, Rejected.csv"
    )

    # -------------------------------------------------------
    # Validation
    # -------------------------------------------------------
    validation = validate_data(active, rejected, excel_file, active_csv, rejected_csv)

    print("\nValidation")
    print("---------------------")

    print(f"Active PDFs     : {validation['active_pdf']}")
    print(f"Active Excel    : {validation['active_excel']}")
    print(f"Active CSV      : {validation['active_csv']}")

    print()

    print(f"Rejected PDFs   : {validation['rejected_pdf']}")
    print(f"Rejected Excel  : {validation['rejected_excel']}")
    print(f"Rejected CSV    : {validation['rejected_csv']}")

    print()

    if validation["passed"]:
        print("Validation PASSED")
    else:
        print("Validation FAILED")

    print()

    validation_file = config.apptracker_folder / "Validation.xlsx"

    save_validation(validation, validation_file, VERSION, audit.elapsed_seconds())

    print(f"Validation Excel report saved to:\n{validation_file}\n")

    audit.log(
        "Validation.xlsx Created",
        "SUCCESS" if validation["passed"] else "FAILED",
        "PASSED" if validation["passed"] else "FAILED"
    )

    # -------------------------------------------------------
    # Notifications
    # -------------------------------------------------------
    report_files = {
        "excel_file": excel_file,
        "active_csv": active_csv,
        "rejected_csv": rejected_csv,
        "validation_file": validation_file,
    }
    counts = {"active": active, "rejected": rejected, "total": total}

    send_user_report_email(validation["passed"], config, audit, report_files, counts)

    # -------------------------------------------------------
    # Demo Data
    # -------------------------------------------------------
    run_demo_data_generator(config, audit)

    # -------------------------------------------------------
    # Wrap up
    # -------------------------------------------------------
    elapsed_time = audit.finalize()

    print(f"\nFinal Elapsed Time: {elapsed_time:.2f} seconds\n")
    print("Run completed successfully.\n")

    send_admin_summary_email(config, audit)


if __name__ == "__main__":
    main()
