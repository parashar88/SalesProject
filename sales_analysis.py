import pandas as pd
import numpy as np  # ✅ New: Import NumPy

df = None  # Define the variable first

try:
    # Load the Excel file
    df = pd.read_excel('sales_data.xlsx')

    # Show the first 5 rows
    print("✅ File loaded successfully! Here's a preview of the data:")
    print(df.head())

except FileNotFoundError:
    print("❌ Error: Could not find the file.")
except Exception as e:
    print("⚠️ An error occurred:", e)

# --------------------------------------
# Clean data only if file was loaded
# --------------------------------------
if df is not None:
    print("\n🔍 Missing values before cleaning:")
    print(df.isnull().sum())

    # Remove rows with any missing values
    df_cleaned = df.dropna()

    print("\n🧹 Missing values after cleaning:")
    print(df_cleaned.isnull().sum())

    print("\n✅ Preview of cleaned data:")
    print(df_cleaned.head())

    # --------------------------------------
    # Analysis 1: Sum of Sales by Region
    # --------------------------------------
    print("\n📊 Total Sales by Region:")
    sales_by_region = df_cleaned.groupby('Region')['TotalSales'].sum()
    print(sales_by_region)

    # --------------------------------------
    # Analysis 2: Average Sales per Product
    # --------------------------------------
    print("\n📈 Average Sales per Product:")
    avg_sales_product = df_cleaned.groupby('Product')['TotalSales'].mean()
    print(avg_sales_product)

    # --------------------------------------
    # Analysis 3: Highest & Lowest Selling Products
    # --------------------------------------
    print("\n🥇 Highest Selling Product:")
    product_sales = df_cleaned.groupby('Product')['TotalSales'].sum()
    highest = product_sales.idxmax()
    print(f"{highest}: ₹{product_sales[highest]}")

    print("\n🥈 Lowest Selling Product:")
    lowest = product_sales.idxmin()
    print(f"{lowest}: ₹{product_sales[lowest]}")

    # --------------------------------------
    # Analysis 4: NumPy Stats for Numeric Fields
    # --------------------------------------
    print("\n🧠 NumPy Stats (mean, median, std):")
    numeric_cols = ['Quantity', 'UnitPrice', 'TotalSales']

    for col in numeric_cols:
        data = df_cleaned[col]
        print(f"\n🔢 {col}:")
        print(f"Mean: {np.mean(data):.2f}")
        print(f"Median: {np.median(data):.2f}")
        print(f"Standard Deviation: {np.std(data):.2f}")
