
# Excel Sales Report Automation

## Project Overview
A Python-based automation project that combines multiple Excel sales files, cleans the data, calculates key performance indicators (KPIs), and generates a consolidated Excel report.

## Tools and Technologies
- Python
- Pandas
- OpenPyXL
- Microsoft Excel

## Key Features
- Combines multiple Excel files automatically
- Removes duplicate records
- Handles missing and invalid data
- Calculates sales KPIs
- Generates regional, product, monthly, and salesperson summaries
- Exports results into a formatted Excel workbook
- Uses separate folders for input data, output reports, and logs

## Project Structure
- `data/raw/` — Input Excel files
- `src/` — Python scripts
- `output/` — Generated Excel reports
- `logs/` — Automation logs
- `requirements.txt` — Project dependencies

## How to Run
1. Install the required packages:
   `pip install -r requirements.txt`
2. Generate sample data:
   `python src/generate_sample_data.py`
3. Run the automation:
   `python src/automate_report.py`

## Output
The generated Excel workbook contains six worksheets:
1. KPI Summary
2. Region Summary
3. Product Summary
4. Monthly Trend
5. Salesperson Summary
6. Cleaned Data
