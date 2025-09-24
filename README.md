# EXCEL
**_Excel for Data Analytics_**

**Date:** 24 September, 2025 (Wednesday)
**Project Title:** Retail Sales Analytics in Excel
**Datasets Used:** Customers, Products, Sales, Returns
**Skills Demonstrated:** Data Cleaning, Transformation, Lookup Functions, Pivot Tables, Dashboarding, Business Insights

**Step 1: Data Preparation**
Created 4 sheets: Customers, Sales, Products, Returns with headers and sample values.
Cleaned and standardized column names.

**Step 2: Data Transformation**
1. _In Sales:_ 
		Added Profit Margin: =ROUND(H2/F2,2) (in %)
		Added Sales Month: =TEXT(B2,"mmm")
2. _In Customers:_
		Created Age Groups: =IF(Age<30,"Young",IF(Age<50,"Middle-Aged","Senior"))
3. _Used XLOOKUP (or VLOOKUP) to enrich Sales data:_
		Loyalty Status: =XLOOKUP(I2, Customers!A:A, Customers!F:F, "Not Found")
		Cost Price: =XLOOKUP(D2, Products!B:B, Products!E:E, "Not Found")
		Selling Price: =VLOOKUP(D2, Products!B:F, 5, FALSE)

**Step 3: Pivot Tables**
1. _Sales by Region & Category_ → Region (Rows), Category (Columns), Sales (Values)
2. _Top 10 Products by Sales_ → Product (Rows), Sales (Values), sorted by Top 10
3. _Customer Loyalty vs Age Group_ → AgeGroup (Rows), LoyaltyStatus (Columns), Count(CustomerID)
4. _Monthly Sales Trend_ → Group Date by Month, visualize as Line Chart
5. _Profit by Sub-Category_ → Sub-Category (Rows), Profit (Values)

**Step 4: Data Visualization**
1. _Added Slicers:_ Region, Category, Year
2. _Created KPI Cards:_
	Total Sales = =SUM(Sales!F2:F51)
	Total Profit = =SUM(Sales!H2:H51)
	No. of Customers = =COUNTA(Customers!A2:A21)
3. _Applied Conditional Formatting:_
	Profit < 50 → Red highlight
	LoyaltyStatus → Gold = Green, Silver = Blue, Bronze = Gray

**Step 6: Returns Analysis**
1. Linked Returns to Sales using OrderID:
	Return Date = =XLOOKUP(A2, Returns!A:A, Returns!B:B, "Not Returned")
	Return Reason = =XLOOKUP(A2, Returns!A:A, Returns!C:C, "Not Returned")
2. Built:
	Pivot → % of Orders Returned by Category
	Pie Chart → Reasons for Return

**How to Use**
Download the file: _Retail-Sales-Analytics.xlsx_
Open in **Excel** 2016 or later
Go to the Dashboard sheet to explore charts & insights


Course Link:https://www.linkedin.com/learning/excel-managing-and-analyzing-data-25384011?u=107510546
