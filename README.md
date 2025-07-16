# Extract IQS Database

This repository contains small automation scripts to export data from the IQS software and to generate PDF reports from the exported Excel files.

The project was created for internal use at ENERCON to automatically gather action and problem tracking data. The automation relies on screenshots and mouse coordinates, so adjustments may be required for other environments.

## Requirements

The scripts require Python 3 and the packages listed in `requirements.txt`:

```
pyautogui
opencv-python
pandas
openpyxl
reportlab
```

Install them with pip:

```bash
pip install -r requirements.txt
```

## Usage

### 1. Exporting the database

Running `extract_iqs_actions.py` automates IQS to export two Excel spreadsheets (`actions_db.xlsx` and `problems_db.xlsx`). The paths and screen coordinates are specific to the original environment, so adapt them as needed.

```bash
python extract_iqs_actions.py
```

### 2. Generating PDF reports

`pdf_report_generator.py` reads the exported Excel file and builds a landscape A4 PDF with the filtered data. The first command-line argument is the input Excel file and the second is the output path (without the `.pdf` extension). When executed, the script asks which set of actions should be included in the report.

```bash
python pdf_report_generator.py actions_db.xlsx output/report
```

The script will create `output/report_all.pdf`, `output/report_kaizen.pdf` or `output/report_rkm.pdf` depending on your choice.

## License

This project is released under the [MIT License](LICENSE).
