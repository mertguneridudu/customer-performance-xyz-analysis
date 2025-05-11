
# 📊 Customer Performance Analysis and Segmentation (XYZ Analysis)

### Description  
This project is an Excel-based analytical study designed to help you gain a deep understanding of your customer portfolio, monitor performance over time, and implement strategic segmentation. All data used in this project is fully anonymized and does not include real commercial information; fictional values are used for demonstration purposes.

---

## 🎯 Project Objective

- **Segmentation:** Classify customers into X, Y, Z categories based on value (Sales Amount, Order Count, Quantity) and consistency (temporal fluctuations).
- **Tracking & Analysis:** Monitor segment transitions quarterly and annually to identify loyal customers and detect early signs of performance decline.
- **Strategy Development:** Create customized marketing and sales strategies tailored to different customer segments.
- **Resource Optimization:** Maximize resource utilization through data-driven decision making.

---

## ⚙️ Features

- **XYZ Analysis:** Classifies customer metrics (Sales Amount, Order Count, Quantity) into X, Y, Z and their subcategories.
- **Color-Coded Dashboard:** Visual, intuitive reporting for quarterly and yearly performance comparisons.
- **Dynamic Formula Base:** Percentile ranges and weighting values can be easily customized on the "Formula" sheet.
- **Long-Term Evaluation:** Provides an aggregated performance analysis across the entire time span.

---

## 📁 File Structure & Sheets

| Sheet Name               | Purpose                                                                 |
|--------------------------|-------------------------------------------------------------------------|
| Main Comparison Quarter  | Visualizes customer performance changes and segment transitions quarterly. |
| Formula                  | Defines XYZ analysis criteria, percentiles, and assigned scores for subcategories. |
| 2017, 2018, …, 2023      | Lists detailed calculations, subcategory assignments, and scoring per quarter for each year. |
| Main Comparison Data     | Evaluates cumulative performance across all years to determine long-term segmentation. |

---

## 📊 Analysis Methodology (XYZ Classification & Weighting)

- **Data Segmentation:** Each metric (Sales Amount, Order Count, Quantity) is classified into X (top %), Y (middle %), and Z (bottom %) categories.
- **Subcategories:** X, Y, and Z are further divided into X1–X3, Y1–Y3, Z1–Z3 subsegments.
- **Weighting:** Numerical scores are assigned to each subcategory through tables defined on the "Formula" sheet.
- **Score Calculation:** Total or weighted average score is calculated based on the customer’s assigned subcategories.
- **Final XYZ Rating:** The final score is mapped to an overall X, Y, or Z rating based on thresholds defined on the "Formula" sheet.

> **Note:** You can freely customize the percentile thresholds and scoring logic in the "Formula" sheet to align with your business strategies.

---

## 🚀 Getting Started

1. Clone the repository:  
   ```bash  
   git clone https://github.com/mertguneridudu/customer-performance-xyz-analysis.git  
   ```

2. Open the Excel file: `Customer_Performance_XYZ_Analysis.xlsx`

3. Load your data into the **Data** sheet.(Ensure your data covers the 2017-2023 analysis period.)

4. Update analysis parameters in the **Formula** sheet as needed.

5. Insert the “Refresh All Pivots” macro:

Press Alt + F11 to open the VBA editor.

In the Project Explorer, right‑click VBAProject (Customer_Performance_XYZ_Analysis.xlsm) → Insert → Module.

Paste this code into the new module:

```bash  
Sub RefreshAllPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
```
Save the workbook.

6. Run the macro before viewing results:

In Excel: Developer tab → Macros → select RefreshAllPivots → Run.

7. Review the results on **Main Comparison Quarter** and **Main Comparison Data** sheets.

---

> **Important Note:** Developed as a personal project during my tenure at a company, this workbook contains only **placeholder and anonymized** data. All customer information, sales figures, and other metrics have been completely altered; no real company data is present. This project is shared for the purpose of illustrating the methodology and functionality it employs.
