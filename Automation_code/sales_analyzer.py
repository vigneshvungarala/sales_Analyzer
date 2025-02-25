import pandas as pd
import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference
import os


def load_data(file_path):
    """Load the dataset with error handling"""
    try:
        df = pd.read_csv(file_path)
        print("Data Loaded Successfully!")
        return df
    except Exception as e:
        print(f" Error loading data: {e}")
        return None


def clean_data(df):
    """Perform data cleaning"""
    if df is None:
        print("No data to clean.")
        return None

    df.dropna(inplace=True)  # Remove missing values
    if "Sales" in df.columns:
        df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")  # Convert Sales column to float
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")  # Convert Date column

    print("Data Cleaning Done!")
    return df


def analyze_data(df):
    """Perform automated analysis with column checks"""
    if df is None:
        print("❌ No data to analyze.")
        return None

    total_sales = df["Sales"].sum() if "Sales" in df.columns else 0
    best_product = worst_product = top_region = "Data Unavailable"

    if "Product" in df.columns and "Sales" in df.columns:
        best_product = df.groupby("Product")["Sales"].sum().idxmax()
        worst_product = df.groupby("Product")["Sales"].sum().idxmin()

    if "Region" in df.columns and "Sales" in df.columns:
        top_region = df.groupby("Region")["Sales"].sum().idxmax()

    insights = {
        "Total Sales": total_sales,
        "Best-Selling Product": best_product,
        "Worst-Selling Product": worst_product,
        "Top Region": top_region
    }

    print("✅ Analysis Done!")
    return insights


def generate_reports(df, insights):
    """Generate reports and export data"""
    if df is None or insights is None:
        print(" No data to generate reports.")
        return

    # Summary Report
    summary_df = pd.DataFrame(list(insights.items()), columns=["Metric", "Value"])
    summary_df.to_csv("Sales_Summary_Report.csv", index=False)

    # Excel Report with Charts
    excel_file = "Sales_Report.xlsx"
    writer = pd.ExcelWriter(excel_file, engine="openpyxl")
    df.to_excel(writer, sheet_name="Raw Data", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

    # Adding a Chart
    wb = writer.book
    ws = wb["Summary"]
    chart = BarChart()
    data = Reference(ws, min_col=2, min_row=1, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.title = "Sales Insights"
    ws.add_chart(chart, "E2")

    writer.close()

    # Insights Text Report
    with open("Sales_Insights.txt", "w") as file:
        for key, value in insights.items():
            file.write(f"{key}: {value}\n")

    print("✅ Reports Generated Successfully!")


def main():
    """Main function to run the automation"""
    file_path = input("Enter the path to your sales dataset: ")

    if not os.path.exists(file_path):
        print("❌ File not found! Please enter a valid path.")
        return

    df = load_data(file_path)
    df = clean_data(df)
    insights = analyze_data(df)
    generate_reports(df, insights)

    print("🎉 Sales Data Analysis Completed Successfully!")


if __name__ == "__main__":
    main()



