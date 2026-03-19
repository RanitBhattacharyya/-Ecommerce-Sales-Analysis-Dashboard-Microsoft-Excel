📌 Project Overview
An end-to-end data analysis project built entirely in Microsoft Excel. Starting from raw ecommerce transaction data (2011–2014), the project covers data cleaning, transformation, and the construction of a fully interactive sales dashboard with dynamic slicers, KPI cards, and visualisations.

🗂️ Dataset
FieldDetailsSourceEcommerce transaction recordsPeriod2011 – 2014Records~2,100+ ordersGeographyUnited States (by state and region)SegmentsConsumer, Corporate, Home Office
Key fields in raw data: Order ID, Order Date, Ship Date, Customer Segment, Product Category, Sub-Category, Sales, Profit, Quantity, Region, State

🔧 Step 1 — Data Cleaning
The raw dataset required the following cleaning steps before any analysis:

Removed duplicate rows using Excel's Remove Duplicates tool
Standardised date formats — converted all Order Date and Ship Date columns to consistent DD/MM/YYYY format
Fixed inconsistent category labels — e.g. trailing spaces and mixed capitalisation in Category and Sub-Category columns
Handled missing/blank values — checked all key numeric columns (Sales, Profit, Quantity) using COUNTBLANK() and filters
Validated data types — ensured Sales and Profit columns were stored as Number, not Text
Trimmed whitespace — applied TRIM() across text columns to prevent lookup errors in pivot tables


⚙️ Step 2 — Data Transformation
Once clean, new calculated columns and structures were built:

Profit Margin % — =Profit/Sales formatted as percentage
Year column — extracted from Order Date using =YEAR()
Month column — extracted using =MONTH() for monthly trend analysis
YoY Growth — calculated by comparing current year totals vs prior year using pivot table values and % difference fields
Region groupings — validated and standardised: Central, East, South, West
Named ranges — applied to key data tables for cleaner formula references
Pivot Tables — built as the data engine behind every chart and KPI metric on the dashboard


📈 Step 3 — Dashboard Construction
The dashboard was built using the following Excel features:
FeatureUsed ForPivot TablesAggregating sales, profit, quantity by category/time/regionPivot ChartsBar charts, line sparklines, donut chartSlicersInteractive filters — Year, Region, SegmentConditional FormattingKPI colour coding (green = positive YoY, red = negative)Sparkline ChartsMini trend lines inside KPI cardsMap ChartSales by US State choroplethData ValidationControlled inputs for filter areas

📊 Key Results & Insights
KPI Summary
MetricValueYoY GrowthTotal Sales$470,532.51▲ +20.62%Total Profit$61,618.60▲ +14.41%Total Quantity7,979 units▲ +27.45%No. of Orders2,102▲ +28.64%Profit Margin13.10%▼ -3.09%
Category-Level Insights
CategoryProfitTechnology$16.99K ✅ Most profitableOffice Supplies$12.79KFurniture-$1.32K ⚠️ Loss-making
Top 5 Sub-Categories by Sales

Chairs — $71.73K
Phones — $68.31K
Storage — $45.05K
Accessories — $40.52K
Tables — $39.15K

Sales Share by Category

Office Supplies: ~36.24%
Furniture: ~34.66%
Technology: ~29.17%


🔮 Business Implications & Future Use

Profit margin compression — Revenue grew 20.6% but margin fell 3.09%. This signals either rising costs, heavier discounting, or a shift in product mix toward lower-margin items. Warrants a deeper cost analysis.
Furniture is a problem category — Despite holding 34.6% of sales share, Furniture generates a net loss. A pricing review or supplier renegotiation is recommended.
Technology is the profit engine — Highest profit despite lowest sales share. Investment in Technology product lines would improve overall margin.
Q4 seasonal spike — October–December shows the strongest sales concentration. Inventory planning and marketing spend should be front-loaded into Q3 to capitalise on this.
Order volume outpacing profit — Orders grew 28.6% vs profit growth of 14.4%. Operational efficiency (fulfilment, returns, discounts) should be examined at scale.
Regional and segment filters — The dashboard slicers allow any stakeholder to drill into Central/East/South/West performance or Consumer vs Corporate vs Home Office behaviour independently.


🛠️ Tools Used

Microsoft Excel — sole tool for cleaning, transformation, analysis, and visualisation

Pivot Tables & Pivot Charts
Slicers & Timeline
Map Charts
Sparklines
Conditional Formatting
Named Ranges & Data Validation
