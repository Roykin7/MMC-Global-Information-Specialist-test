# MMC Global Information Specialist - Analysis Package

This workspace contains a full Python analysis workflow for `IMS_Test_Data 1.xlsx` and a stakeholder-ready 5-slide PowerPoint.

## Files

- `excel_analysis.py`: data profiling and detailed markdown/csv outputs.
- `build_stakeholder_presentation.py`: full analysis + chart generation + 5-slide PPT creation.
- `requirements.txt`: Python dependencies.

## Run

1. Install dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

2. Run the profile report:

```powershell
.\.venv\Scripts\python.exe excel_analysis.py
```

3. Generate the stakeholder presentation:

```powershell
.\.venv\Scripts\python.exe build_stakeholder_presentation.py
```

## Outputs

- Profile report: `analysis_output/excel_analysis_report.md`
- Stakeholder deck: `presentation_output/MMC_Dataset_Stakeholder_Presentation.pptx`
- Slide charts + summary: `presentation_output/`

## Software used

- Language: Python
- Libraries: pandas, openpyxl, matplotlib, seaborn, python-pptx
