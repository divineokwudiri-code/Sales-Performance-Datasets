# Sales-Performance-Datasets

## Project Overview
This project presents key insights and discusses potential solutions for a product-sales business by analysing 500 customer orders placed across four regions and seven product categories, and by offering actionable recommendations aimed at improving order value, balancing regional performance and strengthening the product mix.

Key questions explored encompass metrics such as total revenue and profitability, regional and item category performance, order-value trends over time, customer concentration and product contribution.

The analysis is based on sales data for the calendar year 2026 (1 January – 31 December 2026).

Microsoft Excel has been employed to execute this analysis, leveraging the Power Pivot Data Model, PivotTables, PivotCharts, slicers and GETPIVOTDATA-linked KPI cards to build an interactive dashboard.

The interactive Excel dashboard can be downloaded [here](https://github.com/divineokwudiri-code/Sales-Performance-Datasets/blob/main/Sales%20Performance.xlsx). 

The original dataset can be found in the datasets sheet of the workbook, or downloaded [here](https://github.com/divineokwudiri-code/Sales-Performance-Datasets/blob/main/Sale%20Analytics%20Datasets.xlsx).

## Table of Contents
- Dataset
- Methodology
- Executive Summary
- Recommendations
- Limitations
- Tools Used

### Dataset
The dataset contains one row per customer order and is stored as an Excel table on the DataSets sheet. It was loaded into the workbook's Data Model so that every PivotTable, chart and slicer draws from a single source.

Attribute	Value
- Records	500 orders
- Period was between	1 Jan 2026 – 31 Dec 2026
- Customers	206 (Customer ID)
- Products	50 (Product ID)
- Item categories	7 (Chair, Notebook, USB Drive, Backpack, Charger, Pen Set, Desk Lamp)
- Regions	4 (East, West, North, South)
- Columns	17 (12 input, 5 calculated)

The data model consists of a single fact table, so no table relationships were required:

### Methodology
The methodology employed for this project comprised the following steps: defining the requirements, data collection, data validation and preparation, data modelling, calculation of measures, pivot analysis, dashboard design, and ultimately reconciliation and sharing of insights.

1. Requirements
The dashboard was designed to answer five business questions:
- How much revenue and profit did the business generate, and at what margin?
- Which regions and item categories contribute most?
- How did sales and order values change over the year?
- How concentrated is revenue across customers and products?
- Where are the opportunities to grow average order value?

2. Data validation
After importing the data, I validated it before building any analysis:
- Duplicates: No duplicate Order IDs (500 unique of 500).
- Missing values: No blank cells in any of the 12 input columns.
- Dates: Every order falls inside 2026, and the Year column agrees with Order Date on all rows.
- Calculated columns: Sales amount and profit were re-computed independently for every row and matched the workbook formulas exactly.
- Profit sanity check: No order has negative profit.
- Reconciliation: Totals on every PivotTable (revenue, cost, profit, units, orders) were reconciled back to the raw table.

3. Calculated columns

Five calculated columns were added to the table:
- Sales Amount = [Quantity Sold] * [Unit Price]
- Profit       = [Sales Amount] - [Cost Amount]
- Quarter      = "Q" & ROUNDUP(MONTH([Order Date]) / 3, 0)
- Month        = TEXT([Order Date], "MMM")
- Month No     = MONTH([Order Date])

Month No is used to sort the month labels chronologically (Jan → Dec) instead of alphabetically.

4. KPIs
Five headline KPIs are pulled from a PivotTable into the dashboard cards using **GETPIVOTDATA**, so they update automatically with the slicers:

KPI	Definition
- Total Revenue	Sum of Sales Amount
- Total Quantity Sold	Sum of Quantity Sold
- Total Orders	Count of Order ID
- Average Order Value	Average of Sales Amount
- Total Customers	Distinct count of Customer ID

  ![monthly sales trend](https://github.com/divineokwudiri-code/Sales-Performance-Datasets/blob/main/Screenshot%202026-10-05%20153503.png)

  ![profit by region](https://github.com/divineokwudiri-code/Sales-Performance-Datasets/blob/main/Screenshot%202026-10-05%20153836.png)

5. Dashboard design
The dashboard (Dashboard sheet) follows a simple layout: KPI cards across the top, charts in a consistent grid below, and four slicers (Region, Item Category, Month, Customer) connected to all nine PivotTables so a single click filters every visual at once. Fonts, colours and sizes are kept consistent across all charts.

![DASHBOARD](https://github.com/divineokwudiri-code/Sales-Performance-Datasets/blob/main/Screenshot%202026-10-05%20152938.png)

### Executive Summary
Below are the major insights that emerged from the analysis:
- Overall performance: In 2026 the business generated 1,443,358 in sales from 500 orders and 5,314 units, at a total cost of 967,465 a profit of 475,893 and an overall margin of 33.0%. The average order was worth 2,887 and contained about 10.6 units.
- Regions: East is the largest region (406,021, 28.1% of sales), followed by West (367,091), North (335,657) and South (334,589). The differences in sales between regions are not statistically significant. North has the lowest margin (31.1% versus 33.4–33.9% elsewhere): it produces 23.3% of sales but only 22.0% of profit.
- Item categories: Chair is the largest category (345,730, 24.0% of sales, 130 orders), followed by Notebook (19.0%), USB Drive (17.2%), Backpack (16.0%), Charger (10.2%), Pen Set (8.3%) and Desk Lamp (5.5%). Margins are virtually identical across categories (31.9% – 34.6%). What does differ is order size: Backpack (3,438) and Notebook (3,304) orders are roughly 40% larger than USB Drive orders (2,315), which is statistically significant.
- Average order value is falling: The number of orders is steady at 122–127 per quarter, but the average order value rose to 3,100 in Q2 and then fell to 2,658 in Q4 (−14%). As a result, H2 sales (684,485) were 9.8% below H1 (758,873) and H2 profit was 5.9% lower, even though order volumes did not drop. Margin percentage actually improved slightly (31.9% in Q1 to 33.9% in Q4), so the decline is about order size, not cost.
- Monthly pattern: Sales peaked in May (165,750) and January (162,912) and were lowest in June (88,776) and February (91,837). Month-to-month swings are large (up to ±50%), but with only 32–50 orders per month they are not statistically distinguishable from normal variation, so no firm seasonality can be claimed.
- Regional shifts through the year. East fell from 123.7k in Q1 to 82.3k in Q4 and West fell from 111.3k in Q2 to 62.4k in Q4, while North (58.6k → 94.9k from Q2 to Q4) and South (54.6k → 92.7k from Q1 to Q4) grew and partly offset the decline.
- Customers: The top customer by ID (C038) contributed 48,582 over 11 orders. Customers placed 5 orders on average (range 1–11). There is no dependency on a few large accounts.
- Products: The best seller is P027 (54,343) and the smallest is P005 (12,044). Low-volume products do not have weak margins, so sales volume alone should not decide which products to drop.
- Order activity: Orders were placed on 260 of 365 days, averaging 1.9 orders per active day. Saturday is the strongest day by sales (259,130) and Tuesday the weakest (160,890), but the weekday difference is not statistically significant. The largest single order was 9,534 (ORD00364); the top 10% of orders make up 26.1% of revenue.

### Recommendations
Based on the analysis, the following recommendations are proposed:
- Protect average order value: Order volume is stable, but average order value fell 14% from Q2 to Q4. Introduce measures that lift basket size: bundles across categories, volume-based discounts that still preserve margin, and minimum-order incentives. Since margin % does not fall with larger orders, pushing order size is a low-risk lever. Track average order value monthly as a headline KPI.
- Cross-sell high-value categories: Backpack and Notebook orders are about 40% larger than USB Drive orders, while Chair and USB Drive generate the most orders. Pair high-frequency products (Chair, USB Drive) with higher-value items (Backpack, Notebook) in bundles or at checkout.
- Grow the highest-margin, under-sold category: Charger has the best margin (34.6%) but only 10.2% of sales. Increasing its share (for example through promotions or bundling) would lift overall profit per order. Desk Lamp (5.5% of sales) and Pen Set (8.3%) are the smallest categories and are candidates for review, but they carry normal margins, so the case for removing them is weak; consider repositioning or bundling first.
- Investigate North's margin gap: Review pricing, discounting and cost of supply in the region. This difference is not statistically significant on one year of data, so monitor it before making structural changes.
- Diagnose the East and West slowdown: Identify what changed in these regions (customer activity, pricing, stock availability, sales coverage) and consider transferring what worked in North and South.
- Review the product range by contribution and margin, not volume alone: 18 of 50 products generate the last 20% of sales. Some low-selling products have above-average margins, so evaluate each product on profit per order and growth potential before discontinuing anything.
- Smooth the weekly rhythm: Saturday is the strongest day and Tuesday the weakest (about 38% lower). The difference is not statistically significant, so treat it as a hypothesis to test: run a mid-week promotion for a few weeks and compare results.

### Limitations
While the analysis provides valuable insights, there are a few limitations to consider that might impact the effectiveness of the recommendations:
- One year of data and small monthly samples: With 500 orders, each month has only 32–50 orders and each region-quarter roughly 30. Many apparent patterns (monthly swings, regional shifts, weekday effects) are within normal random variation and cannot be confirmed statistically. There is no prior year for year-on-year comparison or to confirm seasonality.
- Profit is simplified: Profit is sales minus a single cost figure. There are no discounts, returns, taxes, shipping, overhead or marketing costs, and no order has a negative profit, so true profitability may be lower.
- No customer, channel or staff detail: The data does not include customer demographics, satisfaction or reviews, order channel (online or in-store), sales representative, or repeat-versus-new flags, so segmentation and retention recommendations are based on order counts only.
  

### Tools Used
Microsoft Excel – Power Pivot Data Model, PivotTables, PivotCharts, slicers, GETPIVOTDATA, calculated columns
Design – consistent colour palette, KPI cards and slicer-driven interactivity

Contact
Questions or feedback are welcome: DIVINE · [divineokwudiri219@gmail.com] · [LinkedIn](https://www.linkedin.com/in/divine-okwudiri-3ab106300/)
