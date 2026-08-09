"""Email notifications: User Report email and Admin Summary email."""

import smtplib
import ssl
import textwrap
from datetime import datetime
from email.message import EmailMessage
from getpass import getpass

import keyring


def get_email_password(smtp_username):
    print("Looking up password in system credential store...")

    password = keyring.get_password(
        "Apptracker",
        smtp_username
    )

    if password:
        return password

    print("\nERROR")
    print("----------------------------------")
    print("No password found in the system credential store.")
    print(f"Service : Apptracker")
    print(f"Account : {smtp_username}")

    password = getpass("\nSMTP Password (leave blank to cancel): ")

    if not password:
        return None

    save = input("Save password to system credential store? (Y/N): ").strip().lower()

    if save == "y":
        keyring.set_password(
            "Apptracker",
            smtp_username,
            password
        )

    return password


def send_user_report_email(validation_passed, config, audit, report_files, counts):
    """
    report_files: dict with keys excel_file, active_csv, rejected_csv, validation_file
    counts: dict with keys active, rejected, total
    """

    print("\n=========================================")
    print("USER REPORT EMAIL STATUS")
    print("=========================================\n")

    print(f"{'User Report Email Enabled':<28} {'Yes' if config.email_enabled else 'No'}")
    print(f"{'Validation Passed':<28} {'Yes' if validation_passed else 'No'}")
    print()

    if not config.email_enabled:

        print("No User Report Email sent.\n")

        audit.log(
            "User Report Email",
            "SKIPPED",
            "Email disabled"
        )

        return

    if config.email_only_on_validation_pass and not validation_passed:
        print("User Report Email cancelled.\n")
        print("Reason:")
        print("Validation failed and")
        print("Only send User Report Email if Validation passed\n")

        audit.log(
            "User Report Email",
            "SKIPPED",
            "Validation failed"
        )

        return

    password = get_email_password(config.smtp_username)
    if password:
        print("Password found.\n")

    if password is None:
        print("User Report Email cancelled.")
        return

    excel_file = report_files["excel_file"]
    active_csv = report_files["active_csv"]
    rejected_csv = report_files["rejected_csv"]
    validation_file = report_files["validation_file"]

    msg = EmailMessage()

    msg["Subject"] = (
        f"Apptracker User Report - "
        f"{datetime.now():%m/%d/%Y %I:%M:%S %p}"
    )

    msg["From"] = config.smtp_username

    email_recipients = [
        email.strip()
        for email in config.email_recipients.split(",")
        if email.strip()
    ]
    msg["To"] = ", ".join(email_recipients)

    msg.set_content(
        f"""\
Apptracker completed successfully.

Version: {config.version}

Active Applications: {counts['active']}
Rejected Applications: {counts['rejected']}
Total Applications: {counts['total']}

Validation:
{"PASSED" if validation_passed else "FAILED"}
"""
    )

    attachments = [
        excel_file,
        active_csv,
        rejected_csv,
        validation_file
    ]

    for file in attachments:

        with open(file, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="octet-stream",
                filename=file.name
            )

    context = ssl.create_default_context()

    print(f"Connecting to {config.smtp_server}...")
    print("Sending email...\n")

    try:

        with smtplib.SMTP_SSL(
            config.smtp_server,
            config.smtp_port,
            context=context
        ) as smtp:

            smtp.login(
                config.smtp_username,
                password
            )

            smtp.send_message(msg)

        print("[OK] User Report Email sent successfully.\n")

        audit.log(
            "User Report Email",
            "SUCCESS",
            f"{len(email_recipients)} recipient(s), {len(attachments)} attachment(s)"
        )

        print("Recipients:")

        for recipient in email_recipients:
            print(f"• {recipient}")

        print("\nAttachments:")
        print(f"• {excel_file.name}")
        print(f"• {active_csv.name}")
        print(f"• {rejected_csv.name}")
        print(f"• {validation_file.name}")
        print()

    except Exception as e:

        print("ERROR")
        print("-----------------------------------------")
        print("Unable to send User Report Email.\n")
        print(e)
        print()

        audit.log(
            "User Report Email",
            "FAILED",
            str(e)
        )


def send_admin_summary_email(config, audit):

    print("\n=========================================")
    print("ADMIN SUMMARY EMAIL STATUS")
    print("=========================================\n")

    print(f"{'Admin Email Enabled':<28} {'Yes' if config.admin_email_enabled else 'No'}")
    print()

    if not config.admin_email_enabled:

        print("No Admin Summary Email sent.\n")

        audit.log(
            "Admin Summary Email",
            "SKIPPED",
            "Email disabled"
        )

        return

    run_summary = audit.run_summary

    summary = textwrap.dedent(f"""\
    ===========================
    APPTRACKER ADMIN SUMMARY
    ===========================

    Status: {run_summary["status"]}
    Version: {config.version}

    Started: {run_summary["start_time"].strftime("%m/%d/%Y %I:%M:%S %p")}
    Finished: {run_summary["finish_time"].strftime("%m/%d/%Y %I:%M:%S %p")}
    Elapsed: {run_summary["elapsed"]:.2f} sec

    ---------------------------
    Execution Summary
    ---------------------------

    """)

    print(f"Status        : {run_summary['status']}")
    print(f"Version       : {config.version}")
    print()

    print(f"Started       : {run_summary['start_time'].strftime('%m/%d/%Y %I:%M:%S %p')}")
    print(f"Finished      : {run_summary['finish_time'].strftime('%m/%d/%Y %I:%M:%S %p')}")
    print(f"Elapsed       : {run_summary['elapsed']:.2f} sec")
    print()

    print("---------------------------")
    print("Execution Summary")
    print("---------------------------")

    for step in run_summary["steps"]:

        print(f"{step['status']:<8} {step['step']}")

        if step["details"]:
            print(f"         {step['details']}")

        print()

    for step in run_summary["steps"]:

        summary += f"{step['status']:<8} {step['step']}\n"

        if step["details"]:
            summary += f"         {step['details']}\n"

        summary += "\n"

    summary += "Audit.csv attached.\n"

    password = get_email_password(config.smtp_username)

    if password:
        print("Password found.\n")

    if password is None:
        print("Admin Summary Email cancelled.")
        return

    msg = EmailMessage()

    msg["Subject"] = (
        f"Apptracker Admin Summary - "
        f"{datetime.now():%m/%d/%Y %I:%M:%S %p}"
    )

    msg["From"] = config.smtp_username

    admin_recipients = [
        email.strip()
        for email in config.admin_email_recipients.split(",")
        if email.strip()
    ]

    msg["To"] = ", ".join(admin_recipients)

    msg.set_content(summary)

    with open(audit.audit_file, "rb") as f:

        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=audit.audit_file.name
        )

    context = ssl.create_default_context()

    print(f"Connecting to {config.smtp_server}...")
    print("Sending Admin Summary Email...\n")

    try:

        with smtplib.SMTP_SSL(
            config.smtp_server,
            config.smtp_port,
            context=context
        ) as smtp:

            smtp.login(
                config.smtp_username,
                password
            )

            smtp.send_message(msg)

        print("[OK] Admin Summary Email sent successfully.\n")

        audit.log(
            "Admin Summary Email",
            "SUCCESS",
            f"{len(admin_recipients)} recipient(s), Audit.csv"
        )

        print("Recipients:")

        for recipient in admin_recipients:
            print(f"• {recipient}")

        print("\nAttachments:")
        print(f"• {audit.audit_file.name}")
        print()

    except Exception as e:

        print("ERROR")
        print("---------------------------")
        print("Unable to send Admin Summary Email.\n")
        print(e)
        print()

        audit.log(
            "Admin Summary Email",
            "FAILED",
            str(e)
        )

    print()
