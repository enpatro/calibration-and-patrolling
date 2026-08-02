import pandas as pd
from pathlib import Path
INPUT_FILE = "Compliance_Master.xlsx"
OUTPUT_DIR = Path("data")
OUTPUT_DIR.mkdir(exist_ok=True)
cal = pd.read_excel(INPUT_FILE, sheet_name="Calibration", engine="openpyxl").fillna("")
pat = pd.read_excel(INPUT_FILE, sheet_name="Patrolling", engine="openpyxl").fillna("")
cal.to_json(OUTPUT_DIR / "calibration.json", orient="records", indent=2, force_ascii=False)
pat.to_json(OUTPUT_DIR / "patrolling.json", orient="records", indent=2, force_ascii=False)
print("Generated JSON files in data folder")
