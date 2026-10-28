from pathlib import Path
from datetime import datetime
import logging
import pandas as pd

# Locate the main project folder
ROOT = Path(__file__).resolve().parents[1]

# Define input, output, and log folders
RAW_DIR = ROOT / "data" / "raw"
OUTPUT_DIR = ROOT / "output"
LOG_DIR = ROOT / "logs"

# Create output and log folders if they don't exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

# Create a unique report filename for each run
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
REPORT_PATH = OUTPUT_DIR / f"sales_report_{timestamp}.xlsx"
LOG_PATH = LOG_DIR / "automation.log"

def load_files():
    # Find all Excel files in the input folder
    files = sorted(RAW_DIR.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(
            f"No Excel files found in {RAW_DIR}"
        )

    frames = []

    for file in files:
        print(f"Reading {file.name}...")

        # Read each Excel file
        df = pd.read_excel(file)

        # Remember which file each row came from
        df["Source_File"] = file.name

        frames.append(df)

    # Combine all files into one DataFrame
    combined = pd.concat(frames, ignore_index=True)

    print(f"\nFiles processed: {len(files)}")
    print(f"Total rows loaded: {len(combined)}")

    return combined, len(files)

def clean_data(df):
    # Record the original row count
    original_rows = len(df)

    # Remove duplicate order IDs
    df = df.drop_duplicates(
        subset=["Order_ID"]
    ).copy()

    # Convert columns to appropriate data types
    df["Order_Date"] = pd.to_datetime(
        df["Order_Date"], errors="coerce"
    )

    for col in ["Quantity", "Unit_Price", "Sales_Amount"]:
        df[col] = pd.to_numeric(
            df[col], errors="coerce"
        )

    # Remove rows missing essential information
    df = df.dropna(
        subset=[
            "Order_ID",
            "Order_Date",
            "Quantity",
            "Unit_Price",
            "Sales_Amount"
        ]
    ).copy()

    # Remove invalid quantities or negative values
    df = df[
        (df["Quantity"] > 0)
        & (df["Unit_Price"] >= 0)
        & (df["Sales_Amount"] >= 0)
    ].copy()

    # Recalculate sales instead of trusting the source value
    df["Sales_Amount"] = (
        df["Quantity"] * df["Unit_Price"]
    )

    # Replace missing categories with "Unknown"
    for col in ["Region", "Product", "Salesperson"]:
        df[col] = df[col].fillna("Unknown")

    # Create a month column for trend analysis
    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    removed_rows = original_rows - len(df)

    print(f"Rows before cleaning: {original_rows}")
    print(f"Rows after cleaning: {len(df)}")
    print(f"Rows removed: {removed_rows}")

    return df, original_rows

def build_report(df, files_processed, original_rows):

    # 1. Sales summary by region
    regional = (
        df.groupby("Region", as_index=False)
        .agg(
            Total_Sales=("Sales_Amount", "sum"),
            Total_Orders=("Order_ID", "nunique"),
            Units_Sold=("Quantity", "sum")
        )
        .sort_values("Total_Sales", ascending=False)
    )

    # 2. Sales summary by product
    product = (
        df.groupby("Product", as_index=False)
        .agg(
            Total_Sales=("Sales_Amount", "sum"),
            Units_Sold=("Quantity", "sum")
        )
        .sort_values("Total_Sales", ascending=False)
    )

    # 3. Monthly sales trend
    monthly = (
        df.groupby("Month", as_index=False)
        .agg(
            Total_Sales=("Sales_Amount", "sum"),
            Total_Orders=("Order_ID", "nunique")
        )
        .sort_values("Month")
    )

    # 4. Sales summary by salesperson
    salesperson = (
        df.groupby("Salesperson", as_index=False)
        .agg(
            Total_Sales=("Sales_Amount", "sum"),
            Units_Sold=("Quantity", "sum")
        )
        .sort_values("Total_Sales", ascending=False)
    )

    # 5. Key performance indicators
    total_sales = df["Sales_Amount"].sum()
    total_orders = df["Order_ID"].nunique()

    kpis = pd.DataFrame({
        "Metric": [
            "Files Processed",
            "Rows Before Cleaning",
            "Rows After Cleaning",
            "Rows Removed",
            "Total Sales",
            "Unique Orders",
            "Units Sold",
            "Average Order Value"
        ],
        "Value": [
            files_processed,
            original_rows,
            len(df),
            original_rows - len(df),
            total_sales,
            total_orders,
            df["Quantity"].sum(),
            total_sales / total_orders if total_orders else 0
        ]
    })

    print("\nSummary tables created successfully!")

    # Export all summaries into one Excel workbook
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        kpis.to_excel(
            writer, sheet_name="KPI Summary", index=False
        )

        regional.to_excel(
            writer, sheet_name="Region Summary", index=False
        )

        product.to_excel(
            writer, sheet_name="Product Summary", index=False
        )

        monthly.to_excel(
            writer, sheet_name="Monthly Trend", index=False
        )

        salesperson.to_excel(
            writer, sheet_name="Salesperson Summary", index=False
        )

        df.to_excel(
            writer, sheet_name="Cleaned Data", index=False
        )

        # Format each worksheet for easier reading
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            for column_cells in sheet.columns:
                column_letter = column_cells[0].column_letter

                max_length = max(
                    len(str(cell.value))
                    if cell.value is not None else 0
                    for cell in column_cells
                )

                sheet.column_dimensions[
                    column_letter
                ].width = min(max_length + 2, 28)

    print(f"\nExcel report saved successfully!")
    print(f"Report location: {REPORT_PATH}")

def main():
    try:
        print("Starting Excel report automation...\n")

        # Step 1: Load all Excel files
        df, files_processed = load_files()

        # Step 2: Clean the combined data
        cleaned_df, original_rows = clean_data(df)

        # Step 3: Generate and save the report
        build_report(
            cleaned_df,
            files_processed,
            original_rows
        )

        print("\nAutomation completed successfully!")

    except Exception as error:
        logging.exception("Automation failed")
        print(f"\nAutomation failed: {error}")


if __name__ == "__main__":
    main()

