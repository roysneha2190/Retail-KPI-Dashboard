# Retail Data Pipeline & Automated Excel Dashboard
This project is a Python-based data pipeline that processes raw retail sales data, stores it in a structured local SQL database, and automatically generates a formatted Excel dashboard for business reporting.

Instead of manually building charts and formatting tables every week, this script completely automates the workflow from raw data to a presentation-ready spreadsheet.

---

## Key Features

* **Automated ETL Pipeline:** Extracts raw transaction records from CSV files, cleans missing data points, and transforms data fields using Pandas.
* **SQL Database Storage:** Programmatically designs a relational SQLite3 schema to store transaction records securely inside a local `.db` file.
* **Direct SQL Metrics Engine:** Runs background SQL queries directly against database tables to calculate dynamic business performance KPIs.
* **Programmatic Excel Formatting:** Leverages `openpyxl` to auto-adjust column widths, format numeric values as currency, and build structured data tables.
* **Executive Visualization Canvas:** Automatically injects clean, color-blocked KPI summary cards alongside native Excel charts without any manual clicking.

---

## Core SQL Logic

```sql
/* 1. Calculate overall business totals for orders, gross sales, and net profits */
SELECT COUNT(*) AS Total_Orders, SUM(Total_Sales) AS Revenue, SUM(Profit) AS Profit FROM sales;

/* 2. Break down sales performance and group the totals by individual product categories */
SELECT Product_Category, SUM(Total_Sales) AS Revenue FROM sales GROUP BY Product_Category;

/* 3. Filter and identify the top 5 highest-performing country markets by sales volume */
SELECT Country, SUM(Total_Sales) AS Sales FROM sales GROUP BY Country ORDER BY Sales DESC LIMIT 5;


**HOW TO SET UP AND RUN**
# Step 1: Install the required Python automation libraries
pip install pandas openpyxl

# Step 2: Execute the master script to trigger the data pipeline
python build_db.py

# Step 3: Open the generated workbook to view the dashboard canvas
open Retail_Dashboard_Data.xlsx
