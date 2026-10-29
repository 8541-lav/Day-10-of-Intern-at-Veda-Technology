Day 10 – Simple KPI Tracking Sheet
Project Overview
This project creates a simple KPI tracking sheet using the Global Superstore dataset. The goal is to summarize sales performance in one place using Excel formulas and a PivotTable.
Dataset
- File: GlobalSuperstore.xlsx
- Source sheet: Orders
- Records: 51,290 rows (including the header row in the worksheet count)
- Columns used:
  - Product Name — Column Q
  - Sales — Column R
  - Quantity — Column S
  - Row ID — Column A
KPIs Tracked
KPI	Method
Total Revenue	=SUM(Orders!R:R)
Total Units Sold	=SUM(Orders!S:S)
Average Sales per Data Row (AOV label used in the task)	=B4/(COUNTA(Orders!A:A)-1)
Top Product	PivotTable: Product Name in Rows, Sales in Values (Sum of Sales), sorted largest to smallest


Results
- Total Revenue: ₹12,642,501.91
- Total Units Sold: 178,312
- Average Sales per Data Row: ₹246.50
- Top Product: Enter the first product shown after sorting the PivotTable by Sum of Sales, largest to smallest.
Note about AOV: The formula above calculates average sales per data row, not true average order value. To calculate true AOV, divide total sales by the number of unique Order IDs. The current sheet keeps the formula used during this task.

How to Reproduce
1. Open GlobalSuperstore.xlsx.
2. Add a worksheet named Day10_KPI.
3. Enter the KPI labels in cells A4:A7.
4. Enter the formulas in B4:B6 as described above.
5. Create a PivotTable from the Orders data.
6. Put Product Name in Rows and Sales in Values, confirm the aggregation is Sum of Sales, then sort largest to smallest.
7. Copy the first product name into cell B7.
8. Apply borders, bold headings, and currency formatting to the revenue and average value cells.
Tools Used
- Microsoft Excel or a compatible spreadsheet application
- SUM formula
- COUNTA formula
- PivotTable sorting and aggregation
Conclusion
The KPI sheet provides a quick summary of overall revenue, total quantity sold, average sales per data row, and the product with the highest total sales. The Top Product field should match the first product in the sorted PivotTable.
