# Excel Mastery Challenge: Sales Data Analysis

## README

---

## 1. Project Overview

This project is an Excel-based sales data analysis challenge designed to demonstrate skills in data cleaning, transformation, analysis, and visualization. The goal is to extract meaningful business insights from raw retail sales data to support decision-making.

The analysis focuses on identifying sales trends, top-performing products, regional performance, customer behavior, and profitability insights using Excel functions, PivotTables, and dashboards.

---

## 2. Objectives

* Clean and standardize raw sales data
* Perform descriptive and advanced statistical analysis
* Build dynamic formulas for business insights
* Create PivotTables and interactive dashboards
* Visualize key trends using charts

---

## 3. Dataset Description

The dataset simulates retail transactions and includes fields such as:

* Order ID
* Order Date
* Product ID
* Product Category
* Region
* Customer ID
* Quantity
* Sales Amount
* Customer Segment (if applicable)

---

## 4. Workbook Structure

The Excel file is organized into the following sheets:

### 1. Raw Data

Original unprocessed dataset imported from CSV.

### 2. Cleaned Data

Contains cleaned and transformed data with:

* Standardized text values (TRIM, proper casing)
* Correct data types (dates, currency, integers)
* Removed duplicates
* Added helper columns (Year, Month, Day of Week)

### 3. Summary Sheet

Displays key metrics and KPIs such as:

* Total Sales
* Average Order Value
* Number of Orders
* Median Quantity Sold
* Total Profit

### 4. Input Sheet

Used for dynamic analysis:

* Region dropdown (Data Validation)
* Target sales input with restrictions
* Discount scenario calculations

### 5. Dashboard

Interactive visualization layer containing:

* PivotTables
* Slicers
* Charts (bar, line, column)

---

## 5. Data Cleaning Process

The following steps were applied:

* Removed duplicate records based on Order ID and Product ID
* Standardized text fields using TRIM()
* Ensured correct formatting for dates and numbers
* Created helper columns:

  * Year using YEAR(Order Date)
  * Month Name using TEXT(MONTH())
  * Day of Week using TEXT(WEEKDAY())

---

## 6. Key Calculations & Formulas Used

* Total Sales: `SUM(Sales Amount)`
* Average Order Value: `AVERAGE(Sales Amount)`
* Profit: `Sales Amount - Product Cost`
* Conditional Logic:

  * IF(Sales Amount > 500, "Large Order", "Standard Order")
* Criteria-Based Functions:

  * SUMIF, SUMIFS
  * AVERAGEIF
  * COUNTIF
* Lookup Functions:

  * XLOOKUP / VLOOKUP for Product Cost retrieval

---

## 7. Analysis Performed

* Sales performance by product category
* Regional sales comparison
* Monthly and yearly sales trends
* Top 5 best-selling products
* Profitability by category and region
* Order segmentation based on value and quantity

---

## 8. Visualization & Dashboard

The dashboard includes:

* Column charts for category-wise sales
* Line charts for monthly sales trends
* Bar charts for top products
* Interactive slicers for filtering by region, category, and year
* PivotTables summarizing key insights

---

## 9. Excel Features Used

* PivotTables & PivotCharts
* Conditional Formatting
* Data Validation (Dropdown lists)
* Tables (Ctrl + T)
* Freeze Panes
* XLOOKUP / VLOOKUP
* Logical functions (IF, AND, OR)
* Statistical functions (SUM, AVERAGE, MEDIAN)

---

## 10. Key Insights (Expected Output)- Identification of top-performing products and regions

* Understanding seasonal sales trends
* Profit contribution by category
* Detection of high-value customers and orders

---

## 11. Best Practices Applied

* Dynamic formulas (no hardcoded values)
* Structured workbook organization
* Clean and readable formatting
* Interactive dashboard design

---

## 12. File Naming Convention

Final file should be saved as:

**Sales_Analysis_YourName.xlsx**

---

## 13. Notes

This project demonstrates practical Excel skills applicable in business analytics, reporting, and decision-making roles. All analysis is fully dynamic and designed for scalability.

---

## End of README
