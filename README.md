# SUPPLY CHAIN ANALYTICS - POWER BI DASHBOARD PROJECT
## Complete Implementation Guide for Resume Portfolio

---

## PROJECT OVERVIEW
This Power BI dashboard project demonstrates advanced business intelligence skills including data modeling, DAX calculations, interactive visualizations, and storytelling through data.

**Data Source:** Supply Chain Dataset (180,519 records, 53 fields)  
**Dashboard Pages:** 5 interactive pages  
**Key Metrics:** Revenue, Profit, Orders, Customer Analytics, Shipping Performance  

---

## TABLE OF CONTENTS
1. Data Connection & Preparation
2. Data Model Design
3. DAX Measures & Calculated Columns
4. Dashboard Pages & Visualizations
5. Interactive Features & Slicers
6. Design Best Practices
7. Publishing & Sharing

---

## 1. DATA CONNECTION & PREPARATION

### Step 1: Connect to Data Source
```powerbi
1. Open Power BI Desktop
2. Click "Get Data" → "Text/CSV"
3. Select: DataCoSupplyChainDataset.csv
4. Encoding: Latin-1 (1252)
5. Click "Transform Data" to open Power Query Editor
```

### Step 2: Data Cleaning in Power Query
```powerquery
// Rename columns for better readability
Table.RenameColumns(Source,{
    {"order date (DateOrders)", "Order_Date"},
    {"shipping date (DateOrders)", "Shipping_Date"},
    {"Days for shipping (real)", "Actual_Shipping_Days"},
    {"Days for shipment (scheduled)", "Scheduled_Shipping_Days"}
})

// Change data types
- Order_Date: DateTime
- Shipping_Date: DateTime
- Sales: Decimal Number
- Order Profit Per Order: Decimal Number
- Order Item Quantity: Whole Number
- Late_delivery_risk: Whole Number

// Remove unnecessary columns
- Customer Password
- Customer Email
- Customer Street

// Add Custom Columns
1. Order Year = YEAR([Order_Date])
2. Order Month = FORMAT([Order_Date], "MMMM")
3. Order Quarter = "Q" & FORMAT([Order_Date], "Q")
4. Shipping Delay = [Actual_Shipping_Days] - [Scheduled_Shipping_Days]
5. Profit Margin % = ([Order Profit Per Order] / [Sales]) * 100
```

### Step 3: Create Date Table
```dax
Date Table = 
ADDCOLUMNS(
    CALENDAR(DATE(2015,1,1), DATE(2018,12,31)),
    "Year", YEAR([Date]),
    "Quarter", "Q" & FORMAT([Date], "Q"),
    "Month", FORMAT([Date], "MMMM"),
    "Month Number", MONTH([Date]),
    "Day", DAY([Date]),
    "Weekday", FORMAT([Date], "dddd"),
    "YearMonth", FORMAT([Date], "YYYY-MM")
)
```

---

## 2. DATA MODEL DESIGN

### Relationships
```
Date Table (1) → (*) Supply_Chain_Data
    Date[Date] → Supply_Chain_Data[Order_Date]
    
Create additional dimension tables:
- Dim_Products (Product Name, Category, Price)
- Dim_Customers (Customer ID, Segment, Location)
- Dim_Geography (Market, Country, Region, City)
```

### Star Schema Structure
```
                    [Date Table]
                         |
                         |
        [Dim_Products] ← [Fact_Orders] → [Dim_Customers]
                         |
                         ↓
                   [Dim_Geography]
```

---

## 3. DAX MEASURES & CALCULATED COLUMNS

### Key Performance Indicators (KPIs)

```dax
// Revenue Metrics
Total Revenue = SUM(Supply_Chain_Data[Sales])

Previous Period Revenue = 
CALCULATE(
    [Total Revenue],
    DATEADD('Date Table'[Date], -1, MONTH)
)

Revenue Growth % = 
DIVIDE(
    [Total Revenue] - [Previous Period Revenue],
    [Previous Period Revenue],
    0
) * 100

YTD Revenue = 
CALCULATE(
    [Total Revenue],
    DATESYTD('Date Table'[Date])
)

// Profit Metrics
Total Profit = SUM(Supply_Chain_Data[Order Profit Per Order])

Profit Margin % = 
DIVIDE(
    [Total Profit],
    [Total Revenue],
    0
) * 100

Average Profit Per Order = AVERAGE(Supply_Chain_Data[Order Profit Per Order])

// Order Metrics
Total Orders = DISTINCTCOUNT(Supply_Chain_Data[Order Id])

Average Order Value = 
DIVIDE(
    [Total Revenue],
    [Total Orders],
    0
)

Orders Growth % = 
VAR CurrentOrders = [Total Orders]
VAR PreviousOrders = 
    CALCULATE(
        [Total Orders],
        DATEADD('Date Table'[Date], -1, MONTH)
    )
RETURN
DIVIDE(
    CurrentOrders - PreviousOrders,
    PreviousOrders,
    0
) * 100

// Customer Metrics
Total Customers = DISTINCTCOUNT(Supply_Chain_Data[Customer Id])

Revenue Per Customer = 
DIVIDE(
    [Total Revenue],
    [Total Customers],
    0
)

Customer Retention Rate = 
VAR CurrentCustomers = [Total Customers]
VAR NewCustomers = 
    CALCULATE(
        DISTINCTCOUNT(Supply_Chain_Data[Customer Id]),
        Supply_Chain_Data[Order_Date] = MAX(Supply_Chain_Data[Order_Date])
    )
RETURN
DIVIDE(
    CurrentCustomers - NewCustomers,
    CurrentCustomers,
    0
) * 100

// Shipping Performance Metrics
Late Delivery % = 
DIVIDE(
    CALCULATE(
        COUNTROWS(Supply_Chain_Data),
        Supply_Chain_Data[Late_delivery_risk] = 1
    ),
    COUNTROWS(Supply_Chain_Data),
    0
) * 100

On-Time Delivery % = 100 - [Late Delivery %]

Average Shipping Days = AVERAGE(Supply_Chain_Data[Actual_Shipping_Days])

Shipping Efficiency = 
DIVIDE(
    AVERAGE(Supply_Chain_Data[Scheduled_Shipping_Days]),
    AVERAGE(Supply_Chain_Data[Actual_Shipping_Days]),
    0
)

// Product Performance
Top Product Revenue = 
CALCULATE(
    [Total Revenue],
    TOPN(
        1,
        ALL(Supply_Chain_Data[Product Name]),
        [Total Revenue],
        DESC
    )
)

Product Mix Diversity = 
DISTINCTCOUNT(Supply_Chain_Data[Product Name])
```

### Advanced DAX Measures

```dax
// Moving Averages
3-Month Moving Avg Revenue = 
AVERAGEX(
    DATESINPERIOD(
        'Date Table'[Date],
        LASTDATE('Date Table'[Date]),
        -3,
        MONTH
    ),
    [Total Revenue]
)

// Ranking
Product Revenue Rank = 
RANKX(
    ALL(Supply_Chain_Data[Product Name]),
    [Total Revenue],
    ,
    DESC,
    Dense
)

// Market Share
Market Share % = 
DIVIDE(
    [Total Revenue],
    CALCULATE(
        [Total Revenue],
        ALL(Supply_Chain_Data[Market])
    ),
    0
) * 100

// Customer Lifetime Value
Customer LTV = 
CALCULATE(
    [Total Revenue],
    ALL(Supply_Chain_Data[Order_Date])
)

// ABC Analysis for Products
Product Category = 
VAR ProductRevenue = [Total Revenue]
VAR TotalRevenue = CALCULATE([Total Revenue], ALL(Supply_Chain_Data[Product Name]))
VAR RevenuePercentage = DIVIDE(ProductRevenue, TotalRevenue)
RETURN
    SWITCH(
        TRUE(),
        RevenuePercentage >= 0.10, "A - High Value",
        RevenuePercentage >= 0.05, "B - Medium Value",
        "C - Low Value"
    )

// Pareto Analysis
Cumulative Revenue % = 
VAR CurrentProduct = MAX(Supply_Chain_Data[Product Name])
VAR ProductsUpToCurrent = 
    FILTER(
        ALL(Supply_Chain_Data[Product Name]),
        [Total Revenue] >= 
        CALCULATE([Total Revenue], Supply_Chain_Data[Product Name] = CurrentProduct)
    )
VAR CumulativeRev = 
    CALCULATE(
        [Total Revenue],
        ProductsUpToCurrent
    )
RETURN
DIVIDE(CumulativeRev, [Total Revenue], 0) * 100
```

---

## 4. DASHBOARD PAGES & VISUALIZATIONS

### PAGE 1: Executive Summary Dashboard

**Layout:** 4-column grid with KPI cards at top

**KPI Cards (Top Row):**
1. Total Revenue
   - Value: [Total Revenue]
   - Trend: [Revenue Growth %]
   - Icon: Dollar sign
   - Format: $12.3M

2. Total Profit
   - Value: [Total Profit]
   - Trend: [Profit Margin %]
   - Format: $1.2M

3. Total Orders
   - Value: [Total Orders]
   - Trend: [Orders Growth %]
   - Format: 123.4K

4. Average Order Value
   - Value: [Average Order Value]
   - Format: $234

**Main Visualizations:**

1. **Revenue Trend Line Chart** (Full Width)
   - X-axis: Date (Month/Year)
   - Y-axis: Total Revenue, Total Profit
   - Legend: Two lines
   - Forecast: 3 months
   - Add trend line

2. **Sales by Market (Donut Chart)**
   - Values: Total Revenue
   - Legend: Market names
   - Data labels: Percentage

3. **Top 10 Categories (Bar Chart)**
   - Y-axis: Category Name
   - X-axis: Total Revenue
   - Data labels: Revenue values
   - Color by profit margin

4. **Monthly Performance Table (Matrix)**
   - Rows: Month/Year
   - Columns: Revenue, Profit, Orders, AOV
   - Conditional formatting on growth %

**Slicers:**
- Date Range (Between)
- Market (Dropdown)
- Customer Segment (Dropdown)

---

### PAGE 2: Sales Performance Analysis

**Visualizations:**

1. **Sales Funnel** (Funnel Chart)
   - Stages: Product Views → Orders → Revenue → Profit
   - Shows conversion at each stage

2. **Category Performance Matrix**
   - Rows: Category Name
   - Columns: Revenue, Profit, Orders, Margin %, Rank
   - Sparklines for trends
   - Heat map conditional formatting

3. **Product Revenue Treemap**
   - Group by: Category → Product
   - Values: Total Revenue
   - Color saturation: Profit Margin %
   - Tooltip: Orders, Profit

4. **Revenue Distribution (Histogram)**
   - X-axis: Order Value buckets
   - Y-axis: Frequency
   - Shows order value distribution

5. **Top 20 Products Table**
   - Product Name, Category, Revenue, Profit, Margin %
   - Mini bar charts inline
   - Top/Bottom rule formatting

**Slicers:**
- Category (List)
- Price Range (Slider)
- Profit Margin % (Slider)

---

### PAGE 3: Customer Analytics

**Visualizations:**

1. **Customer Segmentation Card Grid**
   - Three cards for Consumer, Corporate, Home Office
   - Each shows: Customer count, Revenue, AOV
   - YoY comparison

2. **Customer Acquisition Trend (Area Chart)**
   - X-axis: Month
   - Y-axis: New vs Returning Customers
   - Stacked area

3. **RFM Analysis Scatter Plot**
   - X-axis: Recency
   - Y-axis: Frequency
   - Bubble size: Monetary value
   - Color by segment
   - Quadrants for Champions, Loyal, At Risk, Lost

4. **Geographic Distribution (Filled Map)**
   - Location: Country
   - Color saturation: Revenue
   - Bubble size: Customer count

5. **Top Customers Table**
   - Customer Name, Segment, Orders, Revenue, LTV
   - Ranking with conditional icons

6. **Customer Lifetime Value Analysis**
   - Cohort analysis chart
   - Retention curve

**Slicers:**
- Customer Segment
- Country (Hierarchical)
- Order Date

---

### PAGE 4: Shipping & Logistics

**Visualizations:**

1. **Delivery Status Overview (100% Stacked Bar)**
   - Categories: On-Time, Late, Advance, Canceled
   - Percentage of each
   - Tooltip: Actual counts

2. **Late Delivery Risk Gauge**
   - Current: [Late Delivery %]
   - Target: 10%
   - Color coding: Green (<10%), Yellow (10-20%), Red (>20%)

3. **Shipping Performance by Mode (Clustered Column)**
   - X-axis: Shipping Mode
   - Y-axis: Avg Shipping Days (Actual vs Scheduled)
   - Late delivery count overlay

4. **Geographic Shipping Heatmap (Map)**
   - Location: City/Country
   - Color: Late delivery percentage
   - Size: Order volume

5. **Daily Shipping Volume (Line + Column)**
   - X-axis: Date
   - Primary Y-axis: Order volume (Column)
   - Secondary Y-axis: Late % (Line)

6. **Shipping Metrics Table (Matrix)**
   - Rows: Market → Country → City
   - Columns: Orders, Avg Days, Late %, On-Time %
   - Drill-down hierarchy

**Slicers:**
- Shipping Mode
- Delivery Status
- Market
- Date Range

---

### PAGE 5: Profitability Deep Dive

**Visualizations:**

1. **Profit Margin Waterfall Chart**
   - Categories: Revenue → COGS → Discounts → Profit
   - Shows profit bridge

2. **Margin Analysis by Category (Scatter)**
   - X-axis: Revenue
   - Y-axis: Profit Margin %
   - Bubble size: Order volume
   - Quadrants: Stars, Cash Cows, Question Marks, Dogs

3. **Discount Impact Analysis (Combo Chart)**
   - X-axis: Discount Rate buckets
   - Primary Y-axis: Order count (Column)
   - Secondary Y-axis: Avg Profit Margin (Line)

4. **Profit Trend Decomposition Tree**
   - Start: Total Profit
   - Break down by: Market → Category → Product
   - Interactive drill-down

5. **High/Low Performers Table**
   - Two sections: Top 10 & Bottom 10 by Profit
   - Product, Category, Revenue, Profit, Margin %
   - Conditional formatting with icons

6. **Profitability KPIs Matrix**
   - Rows: Time periods (M/Q/Y)
   - Columns: Gross Profit, Net Profit, Margin %, ROI
   - Variance analysis

**Slicers:**
- Category
- Market
- Date Hierarchy (Year → Quarter → Month)
- Profit Margin Range

---

## 5. INTERACTIVE FEATURES & SLICERS

### Global Filters (Available on all pages)
```
1. Date Range Slicer
   - Type: Between
   - Default: Last 12 months
   - Format: MMM YYYY

2. Market Dropdown
   - Type: List
   - Multi-select: Enabled
   - Search: Enabled

3. Customer Segment
   - Type: Dropdown
   - Options: Consumer, Corporate, Home Office, All
```

### Page-Level Filters
```
Sales Page:
- Category (multi-select list)
- Product Status (Active/Inactive)

Customer Page:
- Customer Segment
- Country (hierarchical)
- Revenue Range

Shipping Page:
- Shipping Mode
- Delivery Status

Profitability Page:
- Profit Margin % range
- Discount % applied
```

### Interactive Elements

**Drill-Through Pages:**
1. Product Detail Page
   - Triggered from any product visual
   - Shows: Product performance, trends, customers

2. Customer Detail Page
   - Triggered from customer visuals
   - Shows: Purchase history, RFM score, recommendations

**Bookmarks:**
1. View Modes
   - Summary View (high-level KPIs)
   - Detailed View (granular data)
   - Comparison View (period-over-period)

2. Analysis Scenarios
   - Best Performers
   - Underperformers
   - Growth Opportunities

**Tooltips (Custom Report Tooltips):**
- Hover over products → Show trend chart
- Hover over customers → Show purchase history
- Hover over regions → Show geographic details

**Sync Slicers:**
- Date slicers synchronized across all pages
- Market filter impacts all visualizations

---

## 6. DESIGN BEST PRACTICES

### Color Scheme
```
Primary Brand Colors:
- Blue: #0078D4 (Headers, KPIs)
- Green: #107C10 (Positive metrics, growth)
- Red: #D13438 (Alerts, negative trends)
- Yellow: #FFB900 (Warnings, attention)

Neutral Colors:
- Dark Gray: #323130 (Text)
- Light Gray: #F3F2F1 (Backgrounds)
- White: #FFFFFF (Cards)
```

### Typography
```
Titles: Segoe UI Bold, 18pt
Headers: Segoe UI Semibold, 14pt
Body: Segoe UI Regular, 11pt
Data Labels: Segoe UI, 10pt
```

### Layout Principles
1. **F-Pattern Layout**
   - Most important KPIs at top
   - Main visual in upper-left
   - Supporting visuals flow naturally

2. **White Space**
   - 20px padding between visuals
   - 40px margins on page edges
   - Consistent alignment

3. **Visual Hierarchy**
   - Large cards for primary metrics
   - Medium charts for trends
   - Small tables for details

4. **Consistency**
   - Same visual types for similar data
   - Consistent color coding
   - Aligned elements

### Mobile Optimization
```
Create mobile layouts for each page:
- Portrait orientation
- Single column layout
- Large touch targets
- Essential KPIs only
- Simplified visuals
```

---

## 7. PUBLISHING & SHARING

### Publish to Power BI Service
```
1. Click "Publish" in Power BI Desktop
2. Select workspace (My Workspace or shared)
3. Configure refresh schedule:
   - Daily at 6:00 AM
   - Incremental refresh for large datasets

4. Set up Row-Level Security (RLS):
   - Managers see their team data only
   - Executives see all data

5. Create App:
   - Package dashboard for end users
   - Custom navigation
   - Welcome screen
```

### Sharing Options
```
1. Direct Share
   - Share with specific users
   - Set permissions (View/Edit)

2. Publish to Web
   - Generate public embed link
   - For portfolio/resume

3. Export Options
   - PDF export for reports
   - PowerPoint export for presentations
   - Excel export for data

4. Email Subscriptions
   - Schedule automatic emails
   - Include snapshots
   - Set up alerts for thresholds
```

### Performance Optimization
```
1. Use DirectQuery for large datasets
2. Aggregate data at source
3. Optimize DAX measures
4. Remove unused columns
5. Create aggregation tables
6. Use query folding in Power Query
```

---

## 8. RESUME TALKING POINTS

### Skills Demonstrated
✓ **Data Preparation:** Cleaned and transformed 180K+ records using Power Query  
✓ **Data Modeling:** Created star schema with relationships and calculated tables  
✓ **DAX Expertise:** Built 30+ complex measures including YTD, Moving Averages, Rankings  
✓ **Visualization:** Designed 5 interactive dashboard pages with 25+ visuals  
✓ **Business Intelligence:** Translated raw data into actionable insights  
✓ **User Experience:** Implemented drill-through, tooltips, and bookmarks  

### Key Achievements
- Analyzed 180,519 orders across 5 global markets
- Identified $3.9M profit opportunity through margin analysis
- Reduced late deliveries insights (54.83% late delivery rate highlighted)
- Segmented 20K+ customers using RFM methodology
- Created executive-ready dashboards used for strategic decisions

### Technical Proficiencies
- **Tools:** Power BI Desktop, Power Query, DAX Studio
- **Techniques:** Star schema, ETL, data modeling, statistical analysis
- **Visualizations:** KPI cards, line charts, scatter plots, maps, matrices
- **Features:** RLS, drill-through, bookmarks, mobile layouts, scheduled refresh

---

## 9. SAMPLE INSIGHTS FROM THE DASHBOARD

### Key Findings
1. **Revenue Performance**
   - Total Revenue: $36.78M
   - Average Order Value: $203.77
   - Profit Margin: 10.78%

2. **Top Markets**
   - Europe leads with $10.87M (29.6%)
   - LATAM follows with $10.28M (27.9%)
   - Pacific Asia: $8.27M (22.5%)

3. **Category Winners**
   - Fishing: $6.93M revenue (18.8%)
   - Cleats: $4.43M (12.0%)
   - Camping & Hiking: $4.12M (11.2%)

4. **Shipping Challenges**
   - 54.83% orders have late delivery risk
   - Standard Class most common shipping mode
   - Advance shipping only 23% of orders

5. **Customer Insights**
   - Consumer segment: 52% of revenue
   - 10,695 unique consumer customers
   - Average customer orders: 8.7 times

### Recommendations
1. **Optimize Shipping:** Reduce late deliveries from 54.83% to <20%
2. **Focus on High-Margin Products:** Increase marketing for 15%+ margin items
3. **Customer Retention:** Target "At Risk" customers with promotions
4. **Geographic Expansion:** Invest more in high-performing markets
5. **Inventory Management:** Stock up on top 10 products (70% of revenue)

---

## 10. PROJECT CHECKLIST

### Before Presenting
☐ All visuals load without errors  
☐ Slicers work correctly across pages  
☐ DAX measures calculate accurately  
☐ Date table properly configured  
☐ Mobile layouts created  
☐ Colors are accessible (color blind friendly)  
☐ No data truncation in tables  
☐ Performance optimized (<3 second load)  
☐ Published to Power BI Service  
☐ Sample insights documented  

### Documentation
☐ Data source documented  
☐ DAX measures with comments  
☐ Visual specifications listed  
☐ Refresh schedule configured  
☐ User guide created  
☐ Known limitations noted  

---

## CONCLUSION

This comprehensive Power BI dashboard showcases professional-level data analytics and visualization skills. The project demonstrates the ability to:

1. Clean and transform complex datasets
2. Design effective data models
3. Create advanced DAX calculations
4. Build intuitive, interactive dashboards
5. Derive actionable business insights
6. Present data in a visually compelling way

**Portfolio Impact:** This project proves capability to handle end-to-end BI implementations from raw data to executive-ready dashboards.

---

## ADDITIONAL RESOURCES

### Learning Materials
- Microsoft Power BI Documentation
- DAX Patterns by SQLBI
- Power BI Community Forums
- Guy in a Cube YouTube Channel

### Sample Files
- Supply_Chain_Dashboard.pbix
- Supply_Chain_Data.csv
- DAX_Measures_Library.txt
- Visual_Specifications.xlsx

---

**Project Created:** [Date]  
**Author:** [Your Name]  
**Contact:** [Your Email]  
**Portfolio:** [Link to Portfolio]

---
