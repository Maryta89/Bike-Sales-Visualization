import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime


file_path = "Bike_Sales_Visualizations_Lab.xlsx"


dfs = pd.read_excel(file_path, sheet_name=None)


timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# -------------------
# Part 1: Line Chart
# -------------------
df_line = dfs['Revenue and Profit by Year'].iloc[2:7, :3]
df_line.columns = ['Year', 'Annual Profit', 'Annual Revenue']

plt.figure(figsize=(8, 5))
plt.plot(df_line['Year'], df_line['Annual Profit'], marker='o', label='Annual Profit')
plt.plot(df_line['Year'], df_line['Annual Revenue'], marker='o', label='Annual Revenue')
plt.title("Revenue vs. Profits")
plt.xlabel("Year")
plt.ylabel("US Dollars")
plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
plt.grid(True)
plt.tight_layout()
plt.savefig(f"line_chart_{timestamp}.png")  
plt.close()

# -------------------
# Part 2: Column Chart
# -------------------
df_column = dfs['Product Revenue by Country'].iloc[2:10, :5]
df_column.columns = ['Country', 'Mountain Bikes', 'Road Bikes', 'Touring Bikes', 'Accessories']

for col in df_column.columns[1:]:
    df_column[col] = pd.to_numeric(df_column[col], errors='coerce')

df_column.set_index('Country', inplace=True)

df_column.plot(kind='bar', stacked=True, figsize=(8, 5))
plt.title("Product Revenue by Country")
plt.xlabel("Country")
plt.ylabel("US Dollars")
plt.legend(loc='center left', bbox_to_anchor=(1, 0.5))
plt.tight_layout()
plt.savefig(f"column_chart_{timestamp}.png")
plt.close()

# -------------------
# Part 3: Pie Chart
# -------------------
df_pie = dfs['Revenue by Age Group'].iloc[2:7, :2]
df_pie.columns = ['Age Group', 'Revenue']
df_pie['Revenue'] = pd.to_numeric(df_pie['Revenue'], errors='coerce')

plt.figure(figsize=(6, 6))
plt.pie(df_pie['Revenue'], labels=df_pie['Age Group'], autopct='%1.1f%%')
plt.title("Revenue Comparison by Age Group")
plt.tight_layout()
plt.savefig(f"pie_chart_{timestamp}.png")
plt.close()

print("All Plots have Saved!")
