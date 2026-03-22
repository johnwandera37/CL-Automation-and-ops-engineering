# KRA Checker - JSON Recovery Tool

This tool is used to recover and resend KRA JSON transaction files.

It copies `.json` files from the `processedArchive` directory back to the `trnsSales` folder for resubmission.

---

## 📌 Purpose

Sometimes transactions fail to be submitted to KRA.  
This tool allows you to safely recover and resend those files based on a selected date range.

---

## ⚙️ How It Works

1. Prompts for the station PIN folder path
2. Accepts a start date (required)
3. Accepts an optional end date
4. Scans `processedArchive` folders within the date range
5. Copies eligible `.json` files to `trnsSales`
6. Skips:
   - Files outside the date range
   - Files that already exist in `trnsSales`
7. Logs all actions

---

## 📁 Required Folder Structure

<Station PIN Folder>/
└── Data/
├── processed/
│ └── processedArchive/
│ └── YYYYMMDD/
│ └── *.json
└── resend/
└── trnsSales/

---

## ▶️ Running the Script (Python)

```bash
python copy_to_trns_sales.py
```

---

## 🖥️ Running as Executable (.exe)

1. Build the executable
   - pyinstaller --onefile --name kra-checker copy_to_trns_sales.py
2. Output
   - dist/kra-checker.exe
3. Run
   - kra-checker.exe
4. Final Output
   - Files copied to trnsSales
   - Log file generated in the station folder: kra_copy_log_YYYYMMDD_HHMMSS.log
   

## ⚠️ Important Notes

- Existing files in trnsSales are NOT overwritten
- Only .json files are processed
- Ensure correct folder path is provided
- Date format must be: YYYYMMDD

