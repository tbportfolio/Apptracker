"""Optional demo data generator invocation."""

import subprocess
import sys


def run_demo_data_generator(config, audit):

    print("\n" + "=" * 60)
    print("POST PROCESSING - DEMO DATA")
    print("=" * 60)

    enabled_text = "Yes" if config.make_demo_data else "No"

    print(f"\nEnabled : {enabled_text}")
    print(f"Script  : {config.make_demo_script}")
    print()

    if not config.make_demo_data:

        audit.log(
            "Demo Data Generated",
            "SKIPPED",
            "Demo data disabled"
        )

        return

    demo_script_path = config.script_folder / config.make_demo_script

    if not demo_script_path.exists():

        print("Script Check : FAIL")
        print(f"ERROR: Cannot find {config.make_demo_script}")

        audit.log(
            "Demo Data Generated",
            "FAILED",
            f"Cannot find {config.make_demo_script}"
        )

        return

    print("Script Check : PASS")
    print("Running Demo Data Generator...")

    try:
        subprocess.run(
            [sys.executable, str(demo_script_path)],
            check=True
        )

        print("Demo Generator : SUCCESS")

        audit.log(
            "Demo Data Generated",
            "SUCCESS",
            "Apptracker_Demo.xlsx"
        )

    except subprocess.CalledProcessError as e:
        print("Demo Generator : FAILED")
        print(f"Exit Code: {e.returncode}")

        audit.log(
            "Demo Data Generated",
            "FAILED",
            f"Exit Code: {e.returncode}"
        )
