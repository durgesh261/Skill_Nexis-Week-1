# Week 1: Excel & Data Fundamentals — Practice Assignment & Sales Dashboard

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-239120?style=for-the-badge&logo=python&logoColor=white)
![License](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

Comprehensive completion of **Week 1: Excel & Data Fundamentals** from the Data Analyst Course, covering end-to-end data cleaning, descriptive statistics with native Excel formulas, pivot table summarization, conditional formatting, dynamic charting, and a corporate **Executive Sales Performance Dashboard**.

---

## 📁 Repository Contents

| File / Folder | Description |
| :--- | :--- |
| **`Week1_Practice_Assignment_Solutions.xlsx`** | Complete workbook solving all 5 practice questions with live Excel formulas (`AVERAGE`, `MEDIAN`, `MODE.SNGL`, `VAR.S`), Pivot summaries, Conditional Formatting (Top 10), and Category Bar Chart. |
| **`Week1_Sales_Performance_Dashboard.xlsx`** | Executive-ready interactive dashboard featuring 6 KPI summary cards, Yearly Sales Trend line chart (2012–2015), Category revenue column chart, Segment/Department analysis, and regional rankings. |
| **`global_superstore_2016.xlsx`** | Raw Global Superstore 2016 dataset (51,290 records across 24 attributes, 3 sheets: `Orders`, `Returns`, `People`). |
| **`build_assignment.py`** | Automated Python script to generate `Week1_Practice_Assignment_Solutions.xlsx` using `pandas` and `openpyxl`. |
| **`build_dashboard.py`** | Automated Python script to generate `Week1_Sales_Performance_Dashboard.xlsx` with custom corporate styling, metrics, and charts. |

---

## 📝 Part 1: Week 1 Practice Assignment Solutions

### Question 1: Remove Duplicates & Handle Missing Values in a Dataset
- **Duplicate Check**: An exhaustive duplicate analysis across primary key (`Row ID`) and compound key combinations (`Order ID` + `Product ID`) confirmed **0 duplicate rows** across all 51,290 records.
- **Missing Value Audit**:
  - `Postal Code` contained **41,296 null entries (80.5%)**.
  - All other 23 columns (e.g., `Sales`, `Profit`, `Customer ID`, `Order Date`, `Category`, `Region`) had **0 missing values (100% complete)**.
- **Cleaning Strategy**:
  - Rather than deleting 80.5% of international transactions, missing values in `Postal Code` were systematically imputed with `"N/A - Non-US"`.
  - This preserves order volume, financial accuracy, and global geographical representation while eliminating null cells.

---

### Question 2: Create a Pivot Table Showing Region-Wise Sales
Aggregated sales performance across all 23 regions globally:

| Rank | Region | Order Count | Total Sales Revenue ($) | % of Total Sales | Average Order Value ($) |
| :---: | :--- | :---: | :---: | :---: | :---: |
| 1 | **Western Europe** | 5,883 | $1,731,929.74 | 13.70% | $294.40 |
| 2 | **Central America** | 5,616 | $1,223,101.10 | 9.67% | $217.79 |
| 3 | **Oceania** | 3,487 | $1,100,184.73 | 8.70% | $315.51 |
| 4 | **Southeastern Asia** | 3,129 | $884,423.17 | 7.00% | $282.65 |
| 5 | **Southern Asia** | 2,655 | $866,572.70 | 6.85% | $326.39 |
| 6 | **Eastern Asia** | 2,374 | $855,059.39 | 6.76% | $360.18 |
| 7 | **Western US** | 3,203 | $725,457.82 | 5.74% | $226.49 |
| 8 | **Eastern US** | 2,848 | $678,781.24 | 5.37% | $238.34 |
| 9 | **Northern Europe** | 2,204 | $636,779.25 | 5.04% | $288.92 |
| 10 | **South America** | 2,988 | $617,223.70 | 4.88% | $206.57 |
| ... | *Remaining 13 Regions* | 16,903 | $2,322,989.00 | 18.37% | $137.43 |
| **Total** | **Global Superstore** | **51,290** | **$12,642,501.91** | **100.00%** | **$246.49** |

---

### Question 3: Calculate Mean, Median, and Mode using Excel Formulas
All formulas in sheet `Q3_Descriptive_Statistics` dynamically reference `Cleaned_Data`:

| Statistical Metric | Applied Excel Formula | Sales ($) | Profit ($) | Quantity | Discount Rate (%) |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **Mean (Average)** | `=AVERAGE(Cleaned_Data!Range)` | **$246.49** | **$28.61** | **3.48** | **14.29%** |
| **Median (Middle Value)** | `=MEDIAN(Cleaned_Data!Range)` | **$85.05** | **$9.24** | **3.00** | **0.00%** |
| **Mode (Most Frequent)** | `=MODE.SNGL(Cleaned_Data!Range)`| **$12.96** | **$0.00** | **2** | **0.00%** |
| **Sample Variance** | `=VAR.S(Cleaned_Data!Range)` | **237,719.98** | **30,394.77** | **5.19** | **0.0451** |
| **Sample Std. Deviation** | `=STDEV.S(Cleaned_Data!Range)` | **$487.57** | **$174.34** | **2.28** | **21.23%** |
| **Minimum Value** | `=MIN(Cleaned_Data!Range)` | $0.44 | -$6,599.98 | 1 | 0.00% |
| **Maximum Value** | `=MAX(Cleaned_Data!Range)` | $22,638.48 | $8,399.98 | 14 | 85.00% |
| **Total Sum** | `=SUM(Cleaned_Data!Range)` | $12,642,501.91 | $1,467,457.29 | 178,312 | - |
| **Total Count** | `=COUNT(Cleaned_Data!Range)` | 51,290 | 51,290 | 51,290 | 51,290 |

#### Analytical Insight:
- **Skewness of Sales**: The Mean ($246.49) is nearly **3x the Median ($85.05)**, demonstrating strong positive (right) skewness due to large corporate technology and furniture orders.
- **Discount Impact on Profit**: The mode for profit is $0.00, and standard deviation is high ($174.34), showing heavy margin erosion when discounts exceed 20%.

---

### Question 4: Conditional Formatting to Highlight Top 10 Sales
- **Excel Formatting Rule**: Native Top 10 rule applied (`Home > Conditional Formatting > Top/Bottom Rules > Top 10 Items...`).
- **Style Applied**: Soft green fill (`#C6EFCE`) with dark green bold text (`#006100`).
- **Top 10 High-Value Orders Globally**:

| Rank | Order ID | Order Date | Customer Name | Country | Category | Product Name | Sales ($) | Profit ($) |
| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :---: | :---: |
| 1 | `CA-2012-SM20320140-40985` | 2012-03-17 | Sean Miller | United States | Technology | Cisco TelePresence System EX90 | **$22,638.48** | -$1,811.08 |
| 2 | `CA-2014-TC20980140-41915` | 2014-10-03 | Tamara Chand | United States | Technology | Canon imageCLASS 2200 Advanced Copier | **$17,499.95** | +$8,399.98 |
| 3 | `CA-2015-RB19360140-42087` | 2015-03-24 | Raymond Buch | United States | Technology | Canon imageCLASS 2200 Advanced Copier | **$13,999.96** | +$6,719.98 |
| 4 | `CA-2015-TA21385140-42300` | 2015-10-23 | Tom Ashbrook | United States | Technology | Canon imageCLASS 2200 Advanced Copier | **$11,199.97** | +$3,919.99 |
| 5 | `CA-2015-HL15040140-42326` | 2015-11-18 | Hunter Lopez | United States | Technology | Canon imageCLASS 2200 Advanced Copier | **$10,499.97** | +$5,039.99 |
| 6 | `CA-2014-AB10105140-41991` | 2014-12-18 | Adrian Barton | United States | Technology | Cisco TelePresence System EX90 | **$9,892.74** | +$4,946.37 |
| 7 | `CA-2012-SC20095140-41174` | 2012-09-22 | Sanjit Chand | United States | Technology | GBC DocuBind P400 Electric Binding System | **$9,449.95** | +$4,630.48 |
| 8 | `US-2014-BS11365140-41746` | 2014-04-17 | Bill Shonely | United States | Furniture | Rivers Edge Round Conference Table | **$9,099.93** | +$2,365.98 |
| 9 | `CA-2014-SE20110140-41672` | 2014-02-02 | Seth Vernon | United States | Technology | Global Troy Executive Leather High-Back | **$8,749.95** | +$2,799.98 |
| 10 | `CA-2014-CC12370140-41783` | 2014-05-24 | Christopher Conant | United States | Technology | Canon imageCLASS 2200 Advanced Copier | **$8,399.98** | +$1,120.00 |

---

### Question 5: Bar Chart Showing Category-Wise Revenue
Embedded in sheet `Q5_Category_Revenue_Chart`:

| Category | Order Items | Total Revenue ($) | Total Profit ($) | Profit Margin % | % Revenue Share |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Technology** | 10,141 | **$4,744,557.50** | $663,778.74 | **13.99%** | **37.53%** |
| **Furniture** | 9,860 | **$4,110,874.19** | $285,412.87 | **6.94%** | **32.52%** |
| **Office Supplies** | 31,289 | **$3,787,070.23** | $518,265.68 | **13.69%** | **29.95%** |
| **Total** | **51,290** | **$12,642,501.91** | **$1,467,457.29** | **11.61%** | **100.00%** |

---

## 📊 Part 2: Week 1 Project — Sales Performance Dashboard

### Dashboard Architecture & Metrics
The workbook `Week1_Sales_Performance_Dashboard.xlsx` is structured as follows:

1. **Executive KPI Cards**:
   - 💰 **Total Revenue**: `$12,642,502` *(+26.3% YoY growth in 2015)*
   - 📈 **Total Profit**: `$1,467,457` *(11.61% overall net margin)*
   - 📦 **Total Orders**: `25,728` unique orders *(51,290 items shipped)*
   - 🏷️ **Average Order Value**: `$491.39` per order *(+$246.49 per item)*
   - 🎯 **Profit Margin**: `11.61%`
   - 🏆 **Top Category**: `Technology` *(37.5% revenue share)*

2. **Yearly Sales Performance & Growth**:
   - **2012**: $2,259,451.55 | $248,940.81 Profit (11.02% Margin)
   - **2013**: $2,677,438.73 (+18.50% YoY) | $307,415.28 Profit (11.48% Margin)
   - **2014**: $3,405,708.34 (+27.20% YoY) | $406,935.23 Profit (11.95% Margin)
   - **2015**: $4,299,903.29 (+26.26% YoY) | $504,165.97 Profit (11.73% Margin)
   - **Cumulative 4-Year Revenue Growth**: **+90.31%**

3. **Department / Customer Segment Breakdown**:
   - **Consumer**: $6,507,949.41 (51.48% share, $749,239.78 profit, 11.51% margin)
   - **Corporate**: $3,824,697.52 (30.25% share, $441,208.33 profit, 11.54% margin)
   - **Home Office**: $2,309,854.98 (18.27% share, $277,009.18 profit, 11.99% margin)

4. **Strategic Business Recommendations**:
   - **Fix Furniture Margins**: Despite contributing $4.11M (32.5% of revenue), Furniture generates only 6.94% margin due to heavy discounts on Tables (-$62.4K loss) and Bookcases. Introduce maximum discount limits (< 15%).
   - **Scale High-Performing Tech**: Technology delivers 37.5% of revenue and 45.2% of total profit with a strong 14.0% margin. Increase inventory of Copiers, Phones, and Accessories.
   - **Regional Expansion**: Western Europe ($1.73M), Central America ($1.22M), and Oceania ($1.10M) are the primary revenue drivers. Replicate customer acquisition strategies in emerging regions such as Canada and Central Asia.

---

## 🛠️ How to Rebuild Workbooks Locally

### Prerequisites
- Python 3.10+
- `pandas`
- `openpyxl`

### Installation
```bash
pip install pandas openpyxl
```

### Execution
Run the scripts in PowerShell or Bash:
```bash
# 1. Generate the Week 1 Assignment Solution Workbook
python build_assignment.py

# 2. Generate the Week 1 Executive Sales Performance Dashboard
python build_dashboard.py
```

---

## 👨‍💻 Author & Course Information
- **Course**: Data Analyst Course — Week 1: Excel & Data Fundamentals
- **Topic**: Data Cleaning, Pivot Tables, Descriptive Statistics, Dashboard Design
- **Dataset**: Global Superstore 2016
