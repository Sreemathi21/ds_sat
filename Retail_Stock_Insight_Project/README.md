# Retail Stock Insight — Project Package

This package contains a realistic synthetic retail stock analytics workflow for the DS Sprint assessment.

## Dataset note
The project uses a synthetic retail dataset designed to match the required assessment schema and business use case, while remaining reproducible and explainable.

## Files included
- data/raw/ — source files
- data/cleaned/ — cleaned CSV files and quality report
- notebooks/Retail_Stock_Insight_EDA.ipynb — notebook for exploratory analysis and feature engineering
- outputs/figures/ — graphs and visuals
- outputs/tables/ — summary tables
- reports/Data_Quality_Report.pdf — data quality summary
- reports/Business_Insights_Report.pdf — business insights summary
- excel/Retail_Stock_Insight_Analysis.xlsx — Excel workbook
- presentation/Retail_Stock_Insight_Business_Pitch.pptx — presentation deck
- presentation/Presentation_Speaking_Script.docx — presenter notes

## Key findings
- Total revenue: ₹33,955,176.95
- Best category: Personal Care
- Best store: ST016
- Inventory issues reviewed: 3
- Weekend vs weekday demand difference: -51.80%

## Reproduction
1. Install requirements:
   pandas numpy matplotlib seaborn scikit-learn openpyxl python-pptx python-docx reportlab nbformat nbconvert ipykernel jupyter
2. Run this script from the root.
3. Open the generated notebook and project files for review.
