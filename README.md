# Sandman EDF Conversion Automation

This project automates the conversion of Sandman `.edf` files using **AutoHotkey (AHK v2)** + **Python (win32com)**.  
It integrates with Excel metadata to fetch file paths and names, then drives Sandman software to perform batch conversion.

---

## Features
- Automatically fetches **Path** and **Name** from Excel metadata.
- Cleans invalid file names and re-groups output into subfolders.
- Handles Sandman UI automation:
  - Open **Configuration**
  - Select **Drives** and add new entries
  - Launch **Data Management** and click **Convert**
  - Set destination folder and auto-start conversion
- Logs progress back into Excel for tracking.

---

## Project Structure

Sandman_EDF_Automation/
├── get_next.py                  # Python script to fetch next task from Excel
├── update_status.py             # Python script to update Excel status
├── Sandman_edf_annotation.ahk   # AHK automation script
├── test_paths.xlsx              # PSG metadata
└── README.md                    # Project documentation

---

## Requirements

- Windows 10/11
- [Python 3.9+](https://www.python.org/downloads/)
- [AutoHotkey v2](https://www.autohotkey.com/)
- Microsoft Excel (with COM enabled)
- Sandman Elite software (Data Management + Analysis)
- Python dependencies: pip install pywin32

## Usage

	1.	Prepare Excel metadata
	•	Columns required: Path, Name, Status
	•	Set Status=1 for completed, leave empty or other values for pending.
	2.	Run the automation
	•	Start the AHK script
	•	The script will:
	•	Fetch next task from Excel
	•	Insert into Sandman configuration
	•	Trigger conversion
	•	Update Excel status

	3.	Batch processing
	•	Script loops until all rows are marked Status=1.

⸻

## Notes
	•	If Sandman deletes output due to same patient study, create a different folder to save the output before conversion.
	•	If Sandman pops up duplicated/additional/warning windows due to illegal filenames, clean file names before conversion.
	•	Add delays (Sleep) if Sandman UI is slow to respond.
	•	Recommended to test with a small batch first.

⸻

## 📌 To Do
	•	Add logging for failed conversions
	•	Add resume function for interrupted sessions
	•	Support multiple output folders

⸻

## 👤 Author

Han Wu