"""PDF scanning and application record collection."""

import platform
from datetime import datetime


def get_pdf_created_datetime(pdf_path):
    """Return the PDF creation timestamp."""
    stat = pdf_path.stat()

    if platform.system() == "Darwin":
        timestamp = stat.st_birthtime
    elif platform.system() == "Windows":
        timestamp = stat.st_ctime
    else:
        # Linux has no true creation time; fall back to last modified
        timestamp = stat.st_mtime

    return datetime.fromtimestamp(timestamp)


def get_pdf_created_date(pdf_path):
    """Return the PDF creation date."""
    return get_pdf_created_datetime(pdf_path).strftime("%m/%d/%Y")


def load_applications(folder):
    """
    Read all PDF filenames from a folder and return a list of applications.

    Filenames are expected to follow "Work Type - Company - Position.pdf".
    A filename with fewer than two " - " separators still produces an
    application record (with blank fields for the missing parts), but its
    name is collected into malformed_filenames so the caller can warn about
    it instead of it silently showing up as blank cells in Excel.
    """

    applications = []
    malformed_filenames = []

    for pdf in sorted(folder.glob("*.pdf")):

        parts = pdf.stem.split(" - ", 2)

        if len(parts) < 3:
            malformed_filenames.append(pdf.name)

        created_timestamp = get_pdf_created_datetime(pdf)
        date_created = created_timestamp.strftime("%m/%d/%Y")

        application = {
            "work_type": parts[0] if len(parts) > 0 else "",
            "company": parts[1] if len(parts) > 1 else "",
            "position": parts[2] if len(parts) > 2 else "",
            "date": date_created,
            "created_timestamp": created_timestamp,
            "filename": pdf.stem
        }

        applications.append(application)

    return applications, malformed_filenames


def collect_applications(config):
    """
    Scan the Active and Rejected folders and print the same console
    summary v6.0 printed. Returns (active_applications, rejected_applications,
    malformed_filenames).
    """

    active_applications, active_malformed = load_applications(config.active_folder)
    rejected_applications, rejected_malformed = load_applications(config.rejected_folder)

    malformed_filenames = active_malformed + rejected_malformed

    active = len(active_applications)
    rejected = len(rejected_applications)
    total = active + rejected

    print("\nJob Application Summary")
    print("-----------------------")
    print(f"{'Active Applications':<20}: {active}")
    print(f"{'Rejected':<20}: {rejected}")
    print(f"{'Total Submitted':<20}: {total}")

    active_percent = (active / total * 100) if total > 0 else 0
    rejected_percent = (rejected / total * 100) if total > 0 else 0

    print()
    print(f"{'Active %':<20}: {active_percent:.1f}%")
    print(f"{'Rejected %':<20}: {rejected_percent:.1f}%")

    if active > rejected:
        print("\nKeep going! Most of your applications are still active.\n")
    else:
        print("\nTime to submit some more applications.\n")

    print(f"You have {len(active_applications)} Active Applications:\n")

    for app in active_applications:
        print(app["filename"])
    print()

    print(f"You have {len(rejected_applications)} Rejected Applications:\n")

    for app in rejected_applications:
        print(app["filename"])
    print()

    if malformed_filenames:
        print(f"WARNING: {len(malformed_filenames)} file(s) did not match the expected naming pattern:\n")
        for name in malformed_filenames:
            print(f"  {name}")
        print()

    return active_applications, rejected_applications, malformed_filenames
