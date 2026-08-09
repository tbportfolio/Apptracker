"""Configuration loading and folder validation for Apptracker."""

from dataclasses import dataclass
from pathlib import Path


def load_settings(settings_file):
    """Parse Settings.txt into a flat dict of key=value pairs."""

    settings = {}

    with open(settings_file, "r") as file:
        for line in file:
            line = line.strip()

            # Ignore blank lines and comments
            if not line or line.startswith("#"):
                continue

            if "=" in line:
                key, value = line.split("=", 1)
                settings[key.strip()] = value.strip()

    return settings


@dataclass
class Config:
    script_folder: Path
    active_folder: Path
    rejected_folder: Path
    apptracker_folder: Path
    audit_file: Path

    email_enabled: bool
    email_only_on_validation_pass: bool
    email_recipients: str

    admin_email_enabled: bool
    admin_email_recipients: str

    smtp_username: str
    smtp_server: str
    smtp_port: int

    make_demo_data: bool
    make_demo_script: str

    version: str = "v7.3"


def build_config(script_folder, version="v7.3"):
    """Load Settings.txt and return a fully-populated Config object."""

    settings_file = script_folder / "Settings.txt"
    settings = load_settings(settings_file)

    active_folder = Path(settings["active_folder"])
    rejected_folder = Path(settings["rejected_folder"])
    apptracker_folder = Path(settings["apptracker_folder"])

    return Config(
        script_folder=script_folder,
        active_folder=active_folder,
        rejected_folder=rejected_folder,
        apptracker_folder=apptracker_folder,
        audit_file=apptracker_folder / "Audit.csv",

        email_enabled=settings["email_enabled"].lower() == "true",
        email_only_on_validation_pass=(
            settings["email_only_on_validation_pass"].lower() == "true"
        ),
        email_recipients=settings["email_recipients"],

        admin_email_enabled=settings["admin_email_enabled"].lower() == "true",
        admin_email_recipients=settings["admin_email_recipients"],

        smtp_username=settings["smtp_username"],
        smtp_server=settings["smtp_server"],
        smtp_port=int(settings["smtp_port"]),

        make_demo_data=settings["make_demo_data"].lower() == "true",
        make_demo_script=settings["make_demo_script"],

        version=version,
    )


def validate_folders(config):
    """Return True if the active, rejected, and apptracker folders all exist."""

    folders = [
        config.active_folder,
        config.rejected_folder,
        config.apptracker_folder
    ]

    for folder in folders:

        if not folder.exists():

            print("\nERROR:")
            print(folder)
            print("does not exist.\n")

            return False

    return True
