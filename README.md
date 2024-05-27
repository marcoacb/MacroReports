# MacroReports
Daily reports with Excel Macros

## Description
This project is a VBA Excel Macro designed for reviewing daily SLAs. It allows you to get data from SQL, do calculations, build charts with daily information, and send this report through email.


## Table of Contents
1. [Background](#background)
2. [Technologies Used](#technologies-used)
3. [Preview](#preview)
4. [Features](#features)
5. [Sample Data](#sample-data)
6. [Usage](#usage)
   - [Sheet 1: Executive Summary](#sheet-1-executive-summary)
   - [Sheet 2: Sales Analysis](#sheet-2-sales-analysis)
   - [Sheet 3: Customer Analysis](#sheet-3-customer-analysis)
   - [Sheet 4: Product Analysis](#sheet-4-product-analysis)
7. [Contacts and Support](#contacts-and-support)

## Background
This project addressed a specific need for the service we provided to our client, a financial institution. The task involved retrieving information from the ticketing tool, calculating resolution times according to the established SLAs, and then sending these results via email.

## Technologies Used
The following technologies were used to develop this dashboard:
- **SQL**: For data querying and manipulation.
- **ODBC**: For connecting to DB
- **VBA Excel**: For data transformation, programming code, and charts.
- **Task Scheduler**: For automating duties.

## Preview
![Example of email sent](image/MacroReport1.jpg)

## Features
- Interactive visualizations: line charts, bar charts, pie charts, and more.
- Dynamic filters: filter data by date, category, and other dimensions.
- Customizable panels: adjust layout and content according to your needs.
- Data export: export visualizations and data in multiple formats.

## Sample Data
The dashboard uses sample data from the `example_dataset.csv` file. This dataset includes information on sales, customers, and products.

## Usage
### Sheet 1: Executive Summary
The Executive Summary sheet provides an overview of the most important data, including:
- **Key Performance Indicators (KPIs)**: Displays the most relevant KPIs such as total sales, number of customers, etc.
- **Summary Charts**: Bar and line charts summarizing overall trends.

### Sheet 2: Sales Analysis
This sheet allows you to analyze sales in detail:
- **Date Filters**: Select specific date ranges to filter the data.
- **Monthly Sales Chart**: Visualizes sales month by month.
- **Sales Distribution by Region**: Heatmap showing sales by region.

### Sheet 3: Customer Analysis
This sheet focuses on customer data:
- **Customer Segmentation**: Pie chart showing customer segmentation by categories such as age or income.
- **Purchase History**: Detailed table with customer purchase history.

### Sheet 4: Product Analysis
This sheet lets you see product performance:
- **Sales by Product Category**: Bar chart showing sales by category.
- **Top Selling Products**: Table with the top-selling products and their key metrics.

## Contacts and Support
For any questions or support, contact [Marco Chang](mailto:marcochangbegazo@gmail.com).
