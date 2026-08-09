"""Run-lifetime tracking: elapsed time, Audit.csv logging, and the
uncaught-exception hook that used to be bare module-level code in v6.0.
"""

import csv
import sys
import time
import traceback
from datetime import datetime


class Audit:
    """
    Tracks a single Apptracker run: elapsed time, the step-by-step
    Audit.csv log, and the run_summary dict used by the Admin Summary email.

    Installing the exception hook happens in __init__ so it's active as
    early as possible in main() -- matching v6.0, which set sys.excepthook
    before Settings.txt was even read.
    """

    def __init__(self, version):
        self.start_time = time.perf_counter()
        self.run_started = datetime.now()
        self.version = version
        self.audit_file = None

        self.run_summary = {
            "status": "SUCCESS",
            "steps": [],
            "start_time": self.run_started,
            "finish_time": None,
            "elapsed": 0,
            "error": ""
        }

        sys.excepthook = self._global_exception_handler

    def set_audit_file(self, audit_file):
        """Must be called once Settings.txt has been read and the
        Apptracker folder is known, before the first call to log()."""
        self.audit_file = audit_file

    def elapsed_seconds(self):
        return time.perf_counter() - self.start_time

    def _global_exception_handler(self, exc_type, exc_value, exc_traceback):

        self.run_summary["status"] = "FAILED"
        self.run_summary["finish_time"] = datetime.now()
        self.run_summary["elapsed"] = self.elapsed_seconds()
        self.run_summary["error"] = str(exc_value)

        print("\n=========================================")
        print("UNEXPECTED ERROR")
        print("=========================================\n")

        traceback.print_exception(exc_type, exc_value, exc_traceback)

    def log(self, step, status, details=""):

        if self.audit_file is None:
            raise RuntimeError(
                "Audit.set_audit_file() must be called before log()."
            )

        elapsed = f"{self.elapsed_seconds():.2f}"
        file_exists = self.audit_file.exists()

        with open(self.audit_file, "a", newline="", encoding="utf-8") as file:

            writer = csv.writer(file)

            if not file_exists:
                writer.writerow([
                    "Timestamp",
                    "Elapsed (sec)",
                    "Step",
                    "Status",
                    "Details"
                ])

            writer.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                elapsed,
                step,
                status,
                details
            ])

        self.run_summary["steps"].append({
            "step": step,
            "status": status,
            "details": details
        })

    def finalize(self):
        """Log the final 'Application Finished' step and return elapsed seconds."""

        elapsed_time = self.elapsed_seconds()

        self.run_summary["finish_time"] = datetime.now()
        self.run_summary["elapsed"] = elapsed_time

        self.log(
            "Application Finished",
            "SUCCESS",
            f"Elapsed {elapsed_time:.2f} sec"
        )

        return elapsed_time
