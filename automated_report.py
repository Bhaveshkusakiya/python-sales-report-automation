import pandas as pd
import matplotlib.pyplot as plt
import os

# Load Excel file
df = pd.read_excel("cleaned_sales_data.xlsx", engine="openpyxl")


# Convert date column
df['Order Date'] = pd.to_datetime(df['Order Date'], dayfirst=True, errors='coerce')

# Create Year & Month
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month_name()

# KPI calculations
total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
profit_margin = total_profit / total_sales

# Grouped summaries
sales_by_region = df.groupby('Region')['Sales'].sum().reset_index()
profit_by_category = df.groupby('Category')['Profit'].sum().reset_index()
monthly_sales = df.groupby(df['Order Date'].dt.to_period('M'))['Sales'].sum().reset_index()

# Create output folder if not exists
os.makedirs("output", exist_ok=True)

# Save Excel report
with pd.ExcelWriter("output/automated_sales_report.xlsx", engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    sales_by_region.to_excel(writer, sheet_name="Sales by Region", index=False)
    profit_by_category.to_excel(writer, sheet_name="Profit by Category", index=False)
    monthly_sales.to_excel(writer, sheet_name="Monthly Sales", index=False)

# Create simple chart
plt.figure()
plt.plot(monthly_sales['Order Date'].astype(str), monthly_sales['Sales'])
plt.xticks(rotation=45)
plt.title("Monthly Sales Trend")
plt.tight_layout()
plt.savefig("output/monthly_sales_trend.png")

print("Automated report generated successfully!")
print(f"Total Sales: {total_sales}")
print(f"Total Profit: {total_profit}")
print(f"Profit Margin: {profit_margin:.2%}")
