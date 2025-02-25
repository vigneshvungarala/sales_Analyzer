import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference
import os


def load_data(file_path):
    """Load the dataset and validate columns"""
    try:
        df = pd.read_csv(file_path)
        df.columns = df.columns.str.strip()  # Remove spaces from column names
        print("\n✅ Data Loaded Successfully!")
        print("📌 Available Columns:", df.columns.tolist())  # Show column names
        return df
    except Exception as e:
        print(f"❌ Error loading data: {e}")
        return None


def clean_data(df):
    """Perform data cleaning with error handling"""
    # Updated column names based on your dataset
    column_mapping = {
        "Total_Sales": "Sales",  # Renaming Total_Sales to Sales
        "Order_Date": "Date"  # Renaming Order_Date to Date
    }

    df.rename(columns=column_mapping, inplace=True)  # Apply renaming

    required_columns = ["Sales", "Date", "Product", "Region"]

    # Check if required columns exist after renaming
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"❌ Missing columns: {missing_columns}")
        print("📌 Available Columns:", df.columns.tolist())
        return None  # Stop execution if columns are missing

    # Clean data
    df.dropna(inplace=True)  # Remove missing values
    df["Sales"] = df["Sales"].astype(float)  # Convert Sales to float
    df["Date"] = pd.to_datetime(df["Date"])  # Convert Date to DateTime format

    print("✅ Data Cleaning Done!\n")
    return df


def analyze_data(df):
    """Perform automated sales analysis"""
    total_sales = df["Sales"].sum()
    best_product = df.groupby("Product")["Sales"].sum().idxmax()
    worst_product = df.groupby("Product")["Sales"].sum().idxmin()
    top_region = df.groupby("Region")["Sales"].sum().idxmax()

    insights = {
        "Total Sales": total_sales,
        "Best-Selling Product": best_product,
        "Worst-Selling Product": worst_product,
        "Top Region": top_region
    }

    print("✅ Data Analysis Done!\n")
    print("📊 Key Insights:")
    for key, value in insights.items():
        print(f"   {key}: {value}")

    return insights


def generate_reports(df, insights):
    """Generate reports and export them"""
    if df is None or insights is None:
        print("❌ Cannot generate reports due to missing data.")
        return

    # Summary Report
    summary_df = pd.DataFrame(list(insights.items()), columns=["Metric", "Value"])
    summary_csv = "Sales_Summary_Report.csv"
    summary_df.to_csv(summary_csv, index=False)
    print(f"📄 CSV Report Generated: {summary_csv}")

    # Excel Report with Charts
    excel_file = "Sales_Report.xlsx"
    writer = pd.ExcelWriter(excel_file, engine="openpyxl")
    df.to_excel(writer, sheet_name="Raw Data", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

    # Add Chart to Excel
    wb = writer.book
    ws = wb["Summary"]
    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.title = "Sales Insights"
    ws.add_chart(chart, "E2")
    writer.close()

    print(f"📊 Excel Report Generated: {excel_file}")

    # Insights Text Report
    insights_file = "Sales_Insights.txt"
    with open(insights_file, "w") as file:
        for key, value in insights.items():
            file.write(f"{key}: {value}\n")
    print(f"📝 Insights Report Generated: {insights_file}")


def main():
    """Main function to run the automation"""
    file_path = input("📂 Enter the path to your sales dataset: ")

    if not os.path.exists(file_path):
        print("❌ File not found! Please enter a valid path.")
        return

    df = load_data(file_path)
    if df is None:
        return

    df = clean_data(df)
    if df is None:
        return

    insights = analyze_data(df)
    generate_reports(df, insights)

    print("\n🎯 Sales Data Analysis Completed Successfully!")


if __name__ == "__main__":
    main()




