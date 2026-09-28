from pathlib import Path
import subprocess
import sys
import json


# ============================================================
# PROJECT PATHS
# ============================================================

BASE = Path(__file__).resolve().parent

INPUT_FOLDER = BASE / "Input_Data"

PROCESSED_FILE = BASE / "processed_files.json"

INPUT_FOLDER.mkdir(exist_ok=True)


# ============================================================
# LOAD PREVIOUSLY PROCESSED FILES
# ============================================================

if PROCESSED_FILE.exists():

    with open(PROCESSED_FILE, "r") as f:
        processed_files = set(json.load(f))

else:

    processed_files = set()


# ============================================================
# FIND EXCEL FILES
# ============================================================

excel_files = sorted(
    [
        file
        for file in INPUT_FOLDER.glob("*.xlsx")
        if not file.name.startswith("~$")
    ],
    key=lambda x: x.name.lower()
)


print("==============================================")
print("STUDENT REPORT AUTOMATION")
print("==============================================")

print("\nFiles currently in Input_Data:")

for file in excel_files:
    print(f"  - {file.name}")


# ============================================================
# FIND NEXT UNPROCESSED FILE
# ============================================================

next_file = None

for file in excel_files:

    if file.name not in processed_files:

        next_file = file
        break


# ============================================================
# NOTHING NEW
# ============================================================

if next_file is None:

    print("\nNo new Excel file to process.")
    print("All available files have already been processed.")

    sys.exit()


# ============================================================
# PROCESS ONLY ONE FILE
# ============================================================

print()
print("----------------------------------------------")
print(f"Reading: {next_file.name}")
print("----------------------------------------------")
print()


try:

    subprocess.run(
        [
            sys.executable,
            str(BASE / "generate_student_reports.py"),
            str(next_file)
        ],
        check=True
    )


    # --------------------------------------------------------
    # MARK FILE AS PROCESSED
    # --------------------------------------------------------

    processed_files.add(next_file.name)

    with open(PROCESSED_FILE, "w") as f:

        json.dump(
            sorted(processed_files),
            f,
            indent=4
        )


    print()
    print("----------------------------------------------")
    print(f"COMPLETED: {next_file.name}")
    print("----------------------------------------------")

    print()
    print("This run is finished.")
    print("The next Excel file will be processed")
    print("when you run watcher.py again.")


except subprocess.CalledProcessError as error:

    print()
    print("----------------------------------------------")
    print("ERROR: Report generation failed.")
    print(f"Error code: {error.returncode}")
    print("----------------------------------------------")

    sys.exit(1)