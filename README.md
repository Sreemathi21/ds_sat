# PROJECT: RETAIL STOCK INSIGHT — COMPLETE DS SPRINT ASSESSMENT

Act as a senior Data Scientist, Retail Business Analyst, Python Developer, Excel Analytics Expert, and Business Presentation Designer.

Complete my Data Science Skill Based Assessment 2 project end-to-end, following the exact requirements given by Sri Eshwar College of Engineering, Coimbatore.

**Course:** U23AD491 – Data Science  
**Assessment:** DS Sprint – Skill Based Assessment 2  
**Organization:** VantageGro Retail Pvt. Ltd.  
**Use Case:** Retail Stock Insight  
**Business objective:** Understand demand, improve stock planning, reduce stock-outs and excess inventory, and support better management decisions.

IMPORTANT: Do NOT perform Round 1 (Quiz). Complete Round 2, Round 3, and Round 4 fully.

I am a beginner. Make the project professional, reproducible, well-documented, and easy for me to explain during assessment.

## 1. DATASET ACQUISITION — DO THIS YOURSELF

Do not ask me to upload a dataset or manually find files.

First, search for a suitable public retail sales and inventory dataset that can support this use case. Prefer real data from a trustworthy source such as Kaggle, UCI, government open-data portals, or another credible repository.

The project requires these five logical datasets:

1. `transactions.csv`
   - date
   - store
   - product
   - quantity
   - selling_price
   - discount
   - promotion

2. `products.csv`
   - product
   - category
   - brand
   - price
   - shelf_life
   - supplier

3. `stores.csv`
   - store
   - city
   - store_type
   - customer_count
   - region

4. `inventory.csv`
   - date
   - store
   - product
   - opening
   - received
   - sold
   - closing
   - reorder_level
   - lead_time

5. `external_factors.csv`
   - date
   - city or store
   - weather
   - holiday
   - festival
   - weekend
   - local_event

### Dataset selection rules

- Prefer real data with sufficient transaction history, multiple products, and multiple stores.
- Check the license and source before using any downloaded data.
- If a single public dataset does not contain all required fields, combine compatible public datasets only when their relationships can be justified.
- Do not pretend that missing real-world fields were collected from the source.
- If complete data cannot be obtained, create a realistic synthetic retail dataset that matches the five schemas above.
- Clearly label every synthetic field and dataset. Never present simulated findings as actual findings from a real company.
- Use a fixed random seed for reproducibility.
- Use a sufficient amount of data for meaningful analysis, preferably at least 10,000 transactions, 20 stores, 100 products, and 180 days of history.
- Ensure product IDs, store IDs, and dates are consistent across files.
- Save all source files, record the source URLs and licenses, and document any assumptions.

Deliver the original dataset files and a README explaining the source, schema, assumptions, and how to reproduce the analysis.

## 2. ROUND 2 — EDA AND FEATURE ENGINEERING (30 MARKS)

Use Python in Jupyter Notebook or Google Colab.

Required libraries:
- pandas
- NumPy
- Matplotlib
- Seaborn
- scikit-learn where appropriate
- openpyxl for Excel preparation if required

Create a complete notebook named `Retail_Stock_Insight_EDA.ipynb`.

### Step A: Load and understand the data

- Load all five CSV files.
- Display the first five rows, shape, column names, data types, and descriptive statistics.
- Explain what each dataset represents.
- Identify primary keys, candidate keys, and relationships between the datasets.
- Check whether the data is suitable for the business questions.

### Step B: Data quality assessment

Detect and report:
- Missing weather values and missing product fields.
- Duplicate transactions.
- Inconsistent category names, including differences in capitalization and spelling.
- Invalid or zero/negative quantities where inappropriate.
- Invalid selling prices, product prices, and discounts.
- Impossible dates and inconsistent data types.
- Negative inventory balances or other suspicious stock values.
- Inventory arithmetic mismatches.
- Referential integrity issues between transactions, products, stores, and external factors.

Create a data quality report showing the issue, number of affected records, percentage affected, correction method, and remaining issues.

### Step C: Clean the data

- Remove true duplicate transactions without deleting legitimate repeated purchases.
- Normalize category names using a documented mapping.
- Convert dates and numerical columns to appropriate types.
- Handle missing product information using defensible methods and explicit unknown categories where necessary.
- Handle missing weather values without inventing certainty.
- Flag invalid records before correcting, excluding, or retaining them.
- Validate discounts and selling prices.
- Reconcile inventory only when sufficient evidence exists. Otherwise, flag the mismatch rather than silently changing the values.
- Avoid data leakage and avoid using future information to calculate past inventory decisions.
- Preserve an audit trail of records removed, modified, or flagged.

Save the cleaned files in a folder named `cleaned_data`.

Expected outputs:
- `cleaned_transactions.csv`
- `cleaned_products.csv`
- `cleaned_stores.csv`
- `cleaned_inventory.csv`
- `cleaned_external_factors.csv`
- `data_quality_report.csv`

### Step D: Combine datasets correctly

Create a transaction-level analytical dataset by joining transactions to products, stores, and date-appropriate external factors.

Avoid many-to-many joins and duplicated revenue caused by incorrect joins.

Create a separate store-product-date inventory dataset. Do not join inventory to transaction rows in a way that multiplies inventory quantities.

Document:
- Join keys.
- Join types.
- Row counts before and after joining.
- Unmatched records.
- Grain of each analytical table.

Save the appropriate analysis-ready datasets.

### Step E: Descriptive analysis

Calculate:
- Total units sold.
- Gross sales before discounts, where supported by the data.
- Net revenue after discounts.
- Average selling price.
- Average order or transaction value only if a valid order identifier exists.
- Revenue by product category.
- Revenue and units by store.
- Daily, weekly, and monthly sales trends.
- Average demand by weekday and weekend.
- Promotion versus non-promotion sales.
- Store, product, and category contribution to total revenue.
- Inventory available, inventory coverage, and reorder status where the required inputs exist.

Clearly distinguish sales revenue from profit. Do not calculate profit unless cost data is available.

### Step F: Create meaningful visualizations

Use readable Matplotlib and Seaborn charts with titles, axis labels, legends, and appropriate formatting.

Create at least these visualizations:

1. Revenue by category.
2. Top 10 products by revenue.
3. Bottom-performing products, with sufficient data safeguards.
4. Sales by store.
5. Daily or weekly sales trend.
6. Promotion versus non-promotion units sold.
7. Weekday versus weekend demand.
8. Demand by weather condition.
9. Revenue by region or store type.
10. Product demand variability.
11. Inventory position versus reorder level.
12. Products and stores with stock-out risk.
13. Correlation heatmap for appropriate numerical features.

Save all figures in `outputs/figures`.

Every chart must include a short interpretation and its business relevance. Do not invent findings before calculating them.

### Step G: Answer all six business questions

Provide evidence-based answers to:

1. Which categories generate the most revenue?
2. Which stores show increasing or decreasing sales?
3. How do promotions affect units sold?
4. How does weekend demand differ from weekday demand?
5. Which products show volatile demand?
6. Where are stock-outs frequent?

For each question provide:
- The calculation or analytical method.
- The actual result from the dataset.
- A supporting table or chart.
- Business interpretation.
- A practical recommendation.

Use comparable time windows when comparing stores. Distinguish correlation from causation when discussing promotions or external factors.

### Step H: Feature engineering

Create useful features such as:

- `year`, `month`, `day`, `day_of_week`, `week_of_year`.
- `is_weekend`, `is_holiday`, `is_festival`.
- `discount_percentage`.
- `net_revenue`.
- `average_selling_price`.
- `promotion_flag`.
- Lagged sales features such as 1-day, 7-day, and 28-day lags when supported by the time series.
- Rolling 7-day and 28-day average demand using only past observations.
- Demand standard deviation and coefficient of variation.
- Sales trend indicators.
- Inventory coverage in days.
- Stock-out flag.
- Reorder risk flag.
- Estimated lead-time demand and safety stock, when sufficient history exists.
- ABC product classification based on cumulative revenue contribution.
- XYZ classification based on demand variability, where data sufficiency permits.

Use defensible definitions. Explain each feature and why it matters to retail stock planning.

### Step I: Feature encoding and selection

- Apply one-hot encoding or another appropriate encoding to categorical variables.
- Scale numerical features only when required by the selected algorithm.
- Exclude identifiers from inappropriate numerical interpretations.
- Avoid target leakage.
- Use feature importance, correlation analysis, or another justified selection method.
- Apply PCA only if useful; do not include dimensionality reduction merely to satisfy a checklist.
- If a predictive model is used, use chronological train/test separation rather than random splitting for time-dependent forecasting.

If forecasting is appropriate, compare a simple baseline with a suitable model. Report MAE, RMSE, and any other justified metric. Do not claim a forecasting model is necessary if the available data cannot support reliable evaluation.

### Step J: Round 2 deliverables

Save:
- `Retail_Stock_Insight_EDA.ipynb`
- All cleaned CSV files.
- Feature-engineered analytical datasets.
- Data quality report.
- Summary tables.
- Visualizations.
- Model evaluation results if a model is used.
- A concise findings report in Markdown or PDF.

The notebook must run from beginning to end with clear Markdown headings, comments, explanations, and reproducible results.

## 3. ROUND 3 — DATA ANALYSIS USING MICROSOFT EXCEL (30 MARKS)

Create a professional, fully usable Excel workbook named `Retail_Stock_Insight_Analysis.xlsx`.

Use `openpyxl` or another appropriate method to generate the workbook. Ensure that the file opens correctly in Microsoft Excel and contains actual data, working formulas, formatted tables, charts, and PivotTable-compatible analysis data.

### Required worksheets

1. `Executive Summary`
2. `Revenue by Category`
3. `Sales by Store`
4. `Promotion Analysis`
5. `Weekday vs Weekend`
6. `Inventory Analysis`
7. `Products Needing Attention`
8. `Monthly Trends`
9. `Cleaned Transactions`
10. `Cleaned Products`
11. `Cleaned Stores`
12. `Cleaned Inventory`
13. `External Factors`
14. `Data Quality Report`
15. `Data Dictionary`

### Required Excel analysis

A. Revenue by category:
- Category-wise units and revenue.
- Revenue contribution percentage.
- Ranking of categories.
- Bar chart showing category revenue.

B. Sales by store:
- Total revenue and units sold per store.
- Average daily sales, where appropriate.
- Comparison of store performance.
- Store sales chart.

C. Promotion comparison:
- Units sold and revenue during promotion versus non-promotion periods.
- Percentage difference.
- Compare similar products and periods where possible.
- Clearly state that observational comparisons do not prove causation.

D. Weekday and weekend demand:
- Average daily units sold.
- Total units sold.
- Revenue comparison.
- Difference in percentage.
- Chart comparing the two groups.

E. Inventory position:
- Opening stock.
- Stock received.
- Units sold.
- Closing stock.
- Reorder level.
- Lead time.
- Inventory arithmetic validation.
- Stock-out and reorder-risk flags.
- Inventory coverage where valid.

F. Products needing attention:
- Product and store.
- Recent demand.
- Current inventory.
- Reorder level.
- Lead time.
- Demand variability.
- Risk classification.
- Recommended action.
- Reason for recommendation.

G. Monthly trends:
- Month-wise revenue and units.
- Growth or decline compared with the previous month.
- Line charts.

### Excel formulas and statistical functions

Use formulas where suitable, such as:
- `SUMIFS`
- `COUNTIFS`
- `AVERAGEIFS`
- `IF`
- `IFERROR`
- `XLOOKUP` or `INDEX`/`MATCH`
- `SUM`
- `AVERAGE`
- `MEDIAN`
- `STDEV.S`
- `RANK.EQ`
- `COUNTIF`

Formulas must reference the relevant data and calculate meaningful results. Do not replace every formula with hardcoded output values.

Use Excel Tables with clear headers, filters, appropriate number formats, frozen header rows, and conditional formatting.

### PivotTables and PivotCharts

Include PivotTable-ready source tables and instructions for creating native Excel PivotTables and PivotCharts.

If your available tooling cannot create genuine native Excel PivotTables, do not falsely label ordinary summary tables as PivotTables. Instead:
- Include correctly structured source tables.
- Provide equivalent formula-driven summaries and charts.
- Add a worksheet named `PivotTable Guide` describing the exact fields to drag into Rows, Columns, Values, and Filters.
- Explain how to refresh the analysis in Excel.

If native PivotTables can be created and verified, include them.

### Executive Summary worksheet

Present:
- Total revenue.
- Total units sold.
- Number of stores.
- Number of products.
- Best-performing category.
- Best-performing store.
- Number of products at risk.
- Inventory validation issues.
- Weekday/weekend demand difference.
- Promotion comparison.

Include suitable charts, concise findings, and recommended actions.

Make the workbook visually professional with consistent formatting, readable fonts, sensible column widths, currency formatting, percentages, filters, and conditional formatting.

Verify that formulas reference correct columns and that the workbook is not corrupted.

## 4. ROUND 4 — PITCHING BUSINESS INSIGHTS (20 MARKS)

Create a concise, professional PowerPoint presentation named `Retail_Stock_Insight_Business_Pitch.pptx`.

Target audience: retail management, store managers, inventory planners, and business decision-makers.

Duration: approximately 5–7 minutes.

Create approximately 8–10 slides:

1. Title: Retail Stock Insight — VantageGro Retail Pvt. Ltd.
2. Business problem and project objective.
3. Dataset overview and analytical methodology.
4. Revenue insights: leading categories and products.
5. Store performance: stores increasing or decreasing in sales.
6. Promotion and weekday/weekend demand insights.
7. Inventory risks: stock-outs, low stock, excess inventory, and volatile demand.
8. Key products and stores requiring immediate attention.
9. Data-driven stock-planning recommendations.
10. Expected business impact, limitations, and closing message.

### Presentation rules

- Use actual calculated results and charts from the analysis.
- Do not invent revenue figures, stock-out counts, percentages, or business improvements.
- Include evidence for every major claim.
- Explain what each finding means for the business.
- Distinguish measured results from projected benefits.
- Do not promise specific reductions in stock-outs without a validated evaluation.
- Include the most urgent products and stores in a clear action table.
- Use short, readable bullet points rather than paragraphs.
- Use consistent professional styling and a clean business presentation layout.
- Include speaker notes or a separate `Presentation_Speaking_Script.docx` with simple, beginner-friendly explanations for every slide.

### Recommendations must be actionable

Consider recommendations such as:
- Reorder products that are below appropriate reorder thresholds.
- Prioritize products with high demand and low inventory coverage.
- Review safety stock for products with volatile demand.
- Account for supplier lead times.
- Prepare inventory before holidays, festivals, and weekends when the data supports a demand pattern.
- Evaluate promotions that increase demand without creating unnecessary inventory pressure.
- Review persistently slow-moving products to reduce excess stock.
- Compare stores to identify opportunities to transfer excess inventory.
- Use ABC/XYZ segmentation to prioritize planning effort.

Only recommend a specific action for a product or store when the available data supports it. Clearly identify which actions are immediate, short-term, and long-term.

## 5. FINAL PROJECT FOLDER

Organize every deliverable in this structure:

`Retail_Stock_Insight_Project/`

- `README.md`
- `data/raw/`
- `data/cleaned/`
- `notebooks/Retail_Stock_Insight_EDA.ipynb`
- `outputs/figures/`
- `outputs/tables/`
- `reports/Data_Quality_Report.pdf`
- `reports/Business_Insights_Report.pdf`
- `excel/Retail_Stock_Insight_Analysis.xlsx`
- `presentation/Retail_Stock_Insight_Business_Pitch.pptx`
- `presentation/Presentation_Speaking_Script.docx`

Add other files where genuinely necessary.

Include a reproducibility section with environment requirements, installation commands, data source information, execution instructions, and assumptions.

## 6. QUALITY ASSURANCE — MANDATORY

Before declaring the project complete:

1. Verify all input files exist and can be loaded.
2. Verify key columns and data types.
3. Check for duplicate records and join multiplication.
4. Validate revenue calculations against the source data.
5. Reconcile inventory equations where appropriate.
6. Confirm engineered features use no future information.
7. Run the notebook from beginning to end.
8. Verify the Excel workbook opens and contains the expected sheets, formulas, data, and charts.
9. Check that PowerPoint slides contain consistent, evidence-based findings.
10. Check that the speaking script matches the presentation.
11. Ensure the same figures and conclusions appear consistently across Python, Excel, and PowerPoint.
12. Report any failed checks honestly and fix them where possible.

## 7. HOW TO WORK

- Do not stop after explaining the approach.
- Do not give me only Python code or instructions to create the files myself.
- Obtain or construct the dataset, run the analysis, generate the actual files, and verify them.
- Do not ask me to select a dataset unless there is a genuine blocker.
- Prefer a suitable real dataset; use clearly labelled synthetic data when necessary.
- If you cannot create a particular deliverable in your environment, explain the limitation and provide the closest working alternative.
- Use beginner-friendly explanations in the notebook and speaking script, but keep the final deliverables professional.
- Base all findings on calculated evidence.
- Keep a clear record of the work performed.

## FINAL RESPONSE

When finished, provide:
1. A brief summary of the project completed.
2. The dataset source and whether any data is synthetic.
3. The most important business findings.
4. Links to download the notebook, Excel workbook, PowerPoint, reports, and complete project package.
5. A checklist showing which assessment requirements were completed.
6. Any limitations, assumptions, or manual Excel steps still required.

The project is complete only when the actual deliverable files have been created and verified—not when a plan or code snippet has merely been provided.