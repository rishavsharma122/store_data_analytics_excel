# store_data_analytics_excel
Canada Store Data Analysis Dashboard (Excel Project)
Project Overview
This Excel project analyzes sales performance for the Canada Store using an interactive dashboard built in Microsoft Excel. The dashboard provides insights into sales, orders, customer demographics, order status, sales channels, top-performing states, and product categories.
The project is designed for beginners to intermediate Excel users who want to learn:
•	Data cleaning and preparation
•	Pivot Tables
•	Pivot Charts
•	Slicers and Filters
•	Dashboard creation
•	Business analysis using Excel
________________________________________
Dataset Information
Source Sheet
Sheet Name: canada Store
Dataset Columns
Column Name	Description
Order ID	Unique order number
Cust ID	Unique customer ID
Gender	Customer gender
Age	Customer age
Age Group	Customer age category
Date	Order date
Month	Order month
Status	Order delivery status
Channel	Sales channel
SKU	Product SKU
Category	Product category
Size	Product size
Qty	Quantity sold
Currency	Currency type
Amount	Sales amount
Ship City	Customer shipping city
Ship State	Customer shipping state
Ship Postal Code	Postal code
Ship Country	Country name
B2B	Business-to-business indicator
________________________________________
Dashboard Objectives
The dashboard answers the following business questions:
1.	Compare sales and orders using a single chart
2.	Identify the month with the highest sales and orders
3.	Analyze purchasing behavior between men and women
4.	Track different order statuses
5.	Identify top contributing states to sales
6.	Analyze the relationship between age and gender
7.	Determine which sales channel contributes the highest sales
8.	Identify the highest-selling product category
________________________________________
Dashboard Components
1. Sales vs Orders
Purpose
Compares total sales amount with the number of orders month-wise.
Visualization Used
•	Combo Chart
o	Column Chart → Sales Amount
o	Line Chart → Orders Count
Insights
•	Detect seasonal sales trends
•	Compare order volume with revenue growth
________________________________________
2. Men vs Women Analysis
Purpose
Shows gender-wise contribution to total sales.
Visualization Used
•	Pie Chart / Doughnut Chart
Insights
•	Identify the dominant customer segment
•	Understand customer buying patterns
________________________________________
3. Order Status Analysis
Purpose
Displays different order statuses.
Status Types
•	Delivered
•	Cancelled
•	Refunded
•	Returned
Visualization Used
•	Pie Chart
Insights
•	Track operational efficiency
•	Monitor return and cancellation percentages
________________________________________
4. Top States by Sales
Purpose
Identifies the top-performing states contributing to revenue.
Visualization Used
•	Bar Chart
Insights
•	Determine high-performing geographical regions
•	Improve regional marketing strategies
________________________________________
5. Age and Gender Analysis
Purpose
Analyzes purchasing patterns based on age groups and gender.
Visualization Used
•	Clustered Column Chart
Insights
•	Identify target customer groups
•	Understand demographic behavior
________________________________________
6. Sales Channel Analysis
Purpose
Compares sales contribution from different channels.
Example Channels
•	Amazon
•	Myntra
•	Ajio
•	Flipkart
•	Meesho
Visualization Used
•	Pie Chart
Insights
•	Identify the best-performing platform
•	Optimize marketplace strategies
________________________________________
Excel Features Used
Feature	Purpose
Pivot Tables	Data summarization
Pivot Charts	Interactive visualizations
Slicers	Dashboard filtering
Conditional Formatting	Highlight key metrics
Data Cleaning	Remove inconsistencies
Sorting & Filtering	Organize dataset
Charts	Data visualization
________________________________________
Formulas Used in the Project
1. MONTH Formula
Used to extract month values from dates.
=MONTH(Date)
Purpose
Converts order dates into month numbers for monthly analysis.
________________________________________
2. TEXT Formula
Used to display month names.
=TEXT(Date,"MMMM")
Purpose
Converts date values into readable month names.
________________________________________
3. IF Formula
Used to categorize age groups.
=IF(Age<30,"Young",IF(Age<50,"Adult","Senior"))
Purpose
Creates customer age groups.
________________________________________
4. SUM Formula
Calculates total sales.
=SUM(AmountRange)
Purpose
Finds overall revenue.
________________________________________
5. COUNT Formula
Counts total orders.
=COUNT(OrderIDRange)
Purpose
Calculates the total number of orders.
________________________________________
6. SUMIFS Formula
Calculates sales based on conditions.
=SUMIFS(AmountRange,GenderRange,"Women")
Purpose
Calculates gender-wise or state-wise sales.
________________________________________
7. COUNTIFS Formula
Counts records based on multiple conditions.
=COUNTIFS(StatusRange,"Delivered")
Purpose
Counts delivered or cancelled orders.
________________________________________
8. VLOOKUP Formula
Used for data lookup.
=VLOOKUP(LookupValue,TableArray,ColumnNumber,FALSE)
Purpose
Fetches related information from another table.
________________________________________
9. Percentage Formula
Used to calculate contribution percentages.
=Part/Total
Purpose
Calculates sales contribution percentages.
________________________________________
Pivot Tables Used
Pivot Table	Purpose
Sales vs Orders	Monthly sales and order comparison
Men vs Women	Gender-based sales analysis
Order Status	Order tracking analysis
Top Sales States	State contribution analysis
Age & Gender	Customer demographic analysis
Channels	Channel-wise sales comparison
________________________________________
Slicers Added
The dashboard includes slicers for interactive filtering.
Common Slicers
•	Month
•	Channel
•	Category
•	Gender
Benefits
•	Interactive dashboard experience
•	Real-time data filtering
•	Better business insights
________________________________________
Dashboard Design Explanation
Layout Structure
Top Section
•	Project title
•	KPI summary cards
Middle Section
•	Sales trends
•	Gender analysis
•	Order status
Bottom Section
•	State analysis
•	Channel analysis
•	Category analysis
________________________________________
Key Insights from Dashboard
•	Identify peak sales months
•	Understand customer demographics
•	Track operational performance
•	Compare sales channels
•	Analyze regional performance
•	Improve marketing decisions
________________________________________
Learning Outcomes
By completing this project, you will learn:
•	How to clean raw Excel data
•	How to build Pivot Tables
•	How to create interactive charts
•	How to use slicers in dashboards
•	How to analyze business data
•	How to design professional dashboards
________________________________________
Tools Used
•	Microsoft Excel
•	Pivot Tables
•	Pivot Charts
•	Slicers
•	Formulas & Functions
________________________________________
File Structure
Sheet Name	Description
canada Store	Raw dataset
Sales vs orders	Monthly sales analysis
Men vs Women	Gender analysis
Order Status	Status analysis
Top 5 Sales	Top states analysis
Age and Gender	Demographic analysis
Channels	Sales channel analysis
canada store report 2026	Final interactive dashboard
________________________________________
Conclusion
This Excel dashboard project demonstrates how Excel can be used as a powerful business intelligence tool for analyzing sales data and generating actionable insights. The project combines data cleaning, formulas, Pivot Tables, charts, and dashboard design to create a professional reporting solution.
________________________________________
Author
Excel Dashboard Project – Canada Store Data Analysis
