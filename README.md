# Task 2: Customer Feedback Analysis
**RaushByte Technologies – Data Analytics Internship**

---

## 📌 Objective
Analyze customer satisfaction and feedback trends from 8,000 customer records to identify sentiment patterns, satisfaction levels, and premium customer behavior.

---

## ✅ Requirements Completed

| # | Requirement | Status |
|---|-------------|--------|
| 1 | Count positive and negative feedback | ✅ Done |
| 2 | Calculate average satisfaction score | ✅ Done |
| 3 | Analyze premium customers | ✅ Done |
| 4 | Create pie charts and bar graphs | ✅ Done |

---

## 📊 Dataset

- **File:** `customer_analytics_dataset.xlsx`
- **Rows:** 8,000 customers
- **Columns:** 10
  - Customer_ID, Customer_Name, City, Age
  - Customer_Type (Premium / Returning / New)
  - Purchase_Count, Total_Spending
  - Feedback (Excellent / Good / Average / Poor)
  - Satisfaction_Score (1–10)
  - Membership_Years

---

## 🔍 Data Quality

| Check | Result |
|-------|--------|
| Missing Values | 0 |
| Duplicates | 0 |
| Total Records | 8,000 |
| Data Integrity | 100% ✅ |

---

## 📈 Key Findings

### 1. Feedback Distribution
| Feedback | Count | % |
|----------|-------|---|
| Good | 2,064 | 25.8% |
| Poor | 1,998 | 25.0% |
| Excellent | 1,983 | 24.8% |
| Average | 1,955 | 24.4% |

**Sentiment Summary:**
- ✅ **Positive** (Excellent + Good): 4,047 customers — **50.6%**
- ❌ **Negative** (Poor): 1,998 customers — **25.0%**
- ⚪ **Neutral** (Average): 1,955 customers — **24.4%**

### 2. Average Satisfaction Score
- **Overall Average: 5.51 / 10**
- Median: 6.0 / 10
- Range: 1 – 10
- Std Deviation: 2.88

### 3. Premium Customer Analysis
| Metric | Value |
|--------|-------|
| Total Premium Customers | 2,711 (33.9%) |
| Avg Total Spending | ₹2,53,878 |
| Avg Satisfaction | 5.51 / 10 |
| Avg Purchase Count | 25.1 |
| Avg Membership | 5.6 years |

### 4. Satisfaction by Customer Type
| Type | Avg Score |
|------|-----------|
| New | 5.55 / 10 |
| Premium | 5.51 / 10 |
| Returning | 5.48 / 10 |

---

## 📁 Project Structure

```
Task2-Customer-Feedback-Analysis/
│
├── data/
│   └── customer_analytics_dataset.xlsx   ← Raw dataset
│
├── charts/
│   ├── 1_feedback_distribution.png       ← Pie chart: Feedback types
│   ├── 2_satisfaction_by_type.png        ← Bar chart: Satisfaction by type
│   ├── 3_feedback_sentiment.png          ← Pie chart: Positive/Negative/Neutral
│   └── 4_premium_comparison.png         ← Bar chart: Premium vs Non-Premium
│
├── output/
│   └── Task2_Customer_Analysis.xlsx      ← Excel workbook with formulas & charts
│
├── customer_analysis.py                  ← Python analysis script
├── requirements.txt                      ← Python dependencies
└── README.md                             ← This file
```

---

## 🛠️ Tools Used

- **Python 3** — Data analysis and chart generation
- **pandas** — Data loading and aggregation
- **matplotlib** — Visualizations (pie charts, bar graphs)
- **openpyxl** — Excel workbook creation with formulas and embedded charts
- **Excel** — Final workbook with KPIs, COUNTIF/AVERAGEIF formulas, charts

---

## 🚀 How to Run

1. Install dependencies:
   ```bash
   pip install pandas openpyxl matplotlib
   ```

2. Run the analysis script:
   ```bash
   python customer_analysis.py
   ```

3. Open the Excel workbook:
   - Go to `output/Task2_Customer_Analysis.xlsx`
   - Open in Excel and click **Enable Editing** if prompted
   - Review the **Summary** sheet for KPIs and charts

---

## 📋 Excel Workbook

**Sheet 1 – Data:** All 8,000 customer records with formatted table

**Sheet 2 – Summary:**
- KPI cards with live formulas (`COUNTA`, `AVERAGE`, `SUM`, `COUNTIF`)
- Feedback distribution table (`COUNTIF` per type)
- Sentiment summary (Positive / Negative / Neutral counts)
- Satisfaction by customer type (`AVERAGEIF`)
- Premium customer metrics (`AVERAGEIF`, `COUNTIF`)
- 2 embedded Excel charts (Pie + Bar)

---

*Task 2 of 2 | RaushByte Technologies Data Analytics Internship*
