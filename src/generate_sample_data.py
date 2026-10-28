from pathlib import Path
import pandas as pd
import random
from datetime import datetime, timedelta

# Locate the project folders
ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "raw"
RAW_DIR.mkdir(parents=True, exist_ok=True)

# Make the generated data reproducible
random.seed(42)

regions = ["North", "South", "East", "West"]
products = ["Laptop", "Monitor", "Keyboard", "Mouse", "Headphones"]
salespeople = ["Anu", "Rahul", "Meera", "Arjun", "Neha"]

start_date = datetime(2026, 1, 1)

# Generate seven Excel files
for file_number in range(1, 8):
    records = []

    for i in range(1000):
        quantity = random.randint(1, 10)
        price = random.choice([500, 1000, 1500, 2500, 5000])

        records.append({
            "Order_ID": f"ORD-{file_number}-{i + 1:04d}",
            "Order_Date": start_date + timedelta(
                days=random.randint(0, 299)
            ),
            "Region": random.choice(regions),
            "Product": random.choice(products),
            "Salesperson": random.choice(salespeople),
            "Quantity": quantity,
            "Unit_Price": price,
            "Sales_Amount": quantity * price
        })

    df = pd.DataFrame(records)

    # Introduce missing values for data-cleaning practice
    if file_number == 1:
        df.loc[0, "Region"] = None
        df.loc[1, "Quantity"] = None

    # Introduce a duplicate record for testing
    if file_number == 2:
        df = pd.concat([df, df.iloc[[2]]], ignore_index=True)

    file_path = RAW_DIR / f"sales_{file_number}.xlsx"
    df.to_excel(file_path, index=False)

    print(f"Created: {file_path.name}")

print("\nSample data generation completed!")