
from pathlib import Path
from datetime import datetime
import logging
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

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

# Configure logging
logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


def load_files():
    """Load and combine all Excel files from the raw data folder."""
    files = sorted(RAW_DIR.glob("*.xlsx"))

    if not files:
        raise FileNotFoundError(
            f"No Excel files found in {RAW_DIR}"
        )

    frames = []

    for file in files:
        print(f"Reading {file.name}...")
        logging.info("Reading file: %s", file.name)

        df = pd.read_excel(file)
        df["Source_File"] = file.name
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)

    print(f"\nFiles processed: {len(files)}")
    print(f"Total rows loaded: {len(combined)}")

    logging.info(
        "Loaded %s files and %s rows",
        len(files),
        len(combined)
    )

    return combined, len(files)


def clean_data(df):
    """Clean data and prepare it for reporting."""
    original_rows = len(df)

    # Remove duplicate order IDs
    df = df.drop_duplicates(subset=["Order_ID"]).copy()

    # Convert columns to appropriate data types
    df["Order_Date"] = pd.to_datetime(
        df["Order_Date"], errors="coerce"
    )

    for col in ["Quantity", "Unit_Price", "Sales_Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

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

    # Recalculate sales amount
    df["Sales_Amount"] = df["Quantity"] * df["Unit_Price"]

    # Replace missing categories
    for col in ["Region", "Product", "Salesperson"]:
        df[col] = df[col].fillna("Unknown")

    # Create month column
    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    removed_rows = original_rows - len(df)

    print(f"Rows before cleaning: {original_rows}")
    print(f"Rows after cleaning: {len(df)}")
    print(f"Rows removed: {removed_rows}")

    logging.info(
        "Data cleaning complete: before=%s, after=%s, removed=%s",
        original_rows,
        len(df),
        removed_rows
    )

    return df, original_rows


def build_report(df, files_processed, original_rows):
    """Create summary tables and export a formatted Excel report."""

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

    # Export summaries into one Excel workbook
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

        # Format each worksheet
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            # Format the header row
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="1F4E78"
                )
                cell.alignment = Alignment(horizontal="center")

            # Adjust column widths
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

    print("\nExcel report saved successfully!")
    print(f"Report location: {REPORT_PATH}")

    logging.info("Excel report created: %s", REPORT_PATH)


def main():
    try:
        print("Starting Excel report automation...\n")
        logging.info("Automation started")

        # Step 1: Load files
        df, files_processed = load_files()

        # Step 2: Clean data
        cleaned_df, original_rows = clean_data(df)

        # Step 3: Generate report
        build_report(
            cleaned_df,
            files_processed,
            original_rows
        )

        print("\nAutomation completed successfully!")
        logging.info("Automation completed successfully")

    except Exception as error:
        logging.exception("Automation failed")
        print(f"\nAutomation failed: {error}")


if __name__ == "__main__":
    main()
