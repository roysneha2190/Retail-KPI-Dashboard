import sqlite3
import pandas as pd

# Connect directly to the database file you already made
conn = sqlite3.connect("RetailIntelligence.db")

# Write any SQL query you want right here
query = """
SELECT 
    SUM(Total_Sales) AS Total_Revenue,
    AVG(Total_Sales) AS Average_Order_Value
FROM sales;
"""

# Read it and print it out cleanly
df = pd.read_sql_query(query, conn)
print(df)

conn.close()
