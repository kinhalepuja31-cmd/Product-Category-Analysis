# Product-Category-Analysis
This project contains a comprehensive global sales dataset formatted in Excel (product_count_analysis.xlsx), detailing unique consumer transactions across multiple geographical regions. The main focus of this analysis is an automated Category Summary Dashboard located at the bottom-left of the worksheet, which parses through the master records to tally inventory counts and dynamically spotlight the highest-volume product category.
Dataset Architecture
The master ledger spans from columns A through W and contains 46 transactional entries (Rows 2–47) across the following attributes:
•	Transactional Metadata: Row ID, Order ID, Order Date, Ship Date, Ship Mode
•	Customer Profiles: Customer ID, Customer Name, Segment
•	Geographic Metrics: City, State, Country, Postal Code, Market, Region
•	Product Parameters: Product ID, Product Name, Sub-Category, Category
•	Financial Metrics: Sales, Quantity, Discount, Profit, Shipping Cost, Order Priority
Dashboard Automation & Logic
The Category Summary Table (Rows 49–59) utilizes dynamic Excel formulas to eliminate manual counting, automatically updating values if the core transactional database scales.
1. Section: Category Summary
Calculates individual inventory distributions based on the main Category column .
•	Office Supplies Count :
excel
=COUNTIF(T$2:T$47, "Office Supplies")
.
  Result: 26
•	Furniture Count :
excel
=COUNTIF(T$2:T$47, "Furniture")
  Result: 11
•	Technology Count :
excel
=COUNTIF(T$2:T$47, "Technology")
	Result: 9
2. Section: Insights & Analytics
Extracts high-level business intelligence from the summary distribution rows.
•	Largest Category Name :
Dynamically cross-references the counts to return the category string with the highest metric.
excel
=INDEX(B50:B52, MATCH(MAX(C50:C52), C50:C52, 0))
  Result: Office Supplies
•	Largest Category Count :
Extracts the peak count value within the calculated criteria array.
excel
=MAX(C50:C52)
  Result: 26
•	Total Products Checked :
Ensures complete dataset integrity by running a full check on total active text lines.
excel
=COUNTA(T2:T47)
	Result: 46


