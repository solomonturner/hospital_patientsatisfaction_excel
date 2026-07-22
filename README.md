# Hospital Patient Satisfaction & Operational Performance Analysis (Excel)
**Tools Used:** Microsoft Excel (Power Query, Power Pivot, XLOOKUP, IF/IFS, PivotTables, PivotCharts, Conditional Formatting)

## Executive Summary

Hospitals participating in Medicare and Medicaid reimbursement programs have a financial incentive to maintain high patient satisfaction. Under these reimbursement models, hospitals that consistently perform below established patient satisfaction benchmarks may have up to **2%** of their reimbursements withheld. In this scenario, hospital leadership established an internal patient satisfaction target of **3.8 out of 5** and sought to identify where operational improvements could have the greatest impact without significantly increasing operating costs.

This analysis addressed one key business question:

- Where should the hospital focus resources to improve patient satisfaction without significantly increasing costs?

Using operational data from the hospital's electronic health record (EHR), staffing logs, and departmental reference tables, five departments were evaluated across patient satisfaction, wait times, treatment duration, staffing levels, patient volume, treatment costs, and readmission rates. Excel was used to clean and integrate the data, build a relational data model, and create an interactive dashboard for exploratory analysis.

The analysis found that **no department achieved the hospital's patient satisfaction target of 3.8**. General Surgery demonstrated the greatest opportunity for operational improvement, recording the lowest patient satisfaction score alongside the longest average wait time, longest treatment duration, and highest treatment costs. Emergency also emerged as a priority due to its high patient volume and below-target satisfaction score. Pediatrics came closest to meeting the satisfaction target, while staffing analysis revealed that Evening shifts experienced the highest patient demand despite generally having fewer staff than Day shifts. Conversely, Night shifts accounted for a relatively small share of patient visits while maintaining comparatively high staffing levels and strong patient satisfaction.

These findings suggest that better aligning staffing resources with patient demand may improve patient satisfaction without increasing overall staffing costs. However, because the analysis is based solely on operational data, it identifies relationships rather than causation. Additional information, including patient acuity, case complexity, physician communication, staff experience, and bed availability, would be required before recommending specific operational changes.

---

## Project Context & Objectives

Patient satisfaction has become an increasingly important performance metric for hospitals. Under Medicare and Medicaid reimbursement programs, hospitals that consistently receive poor patient satisfaction scores risk having up to 2% of their reimbursements retained. As a result, hospital administrators must continuously balance improving patient experience while controlling operational costs.

Hospital leadership established an internal patient satisfaction target of **3.8 out of 5** and wanted to determine where resources should be allocated to improve patient experience without significantly increasing costs.

Specifically, they wanted to answer:

- **Where should the hospital focus resources to improve patient satisfaction without significantly increasing costs?**

To answer this question, operational performance was evaluated across five hospital departments by analyzing:

- Patient satisfaction
- Patient wait times
- Treatment duration
- Staffing levels by shift
- Patient volume by shift
- Treatment costs compared with departmental baseline costs
- Readmission rates

The results of this analysis help identify departments where operational improvements may provide the greatest return while maintaining financial efficiency.

---

## Data Sourcing & Preparation

Data was sourced from the hospital's electronic health record (EHR), staffing log, and departmental reference tables. The datasets required integration and the creation of several calculated fields before analysis:

- Cleaned and standardized data via **PowerQuery**.
- Combined patient encounter, staffing, and departmental reference data using **XLOOKUP**.
- Created derived fields using **IF** and **IFS** statements to classify patient visits by shift (Day, Evening, Night) and calculate operational metrics.
- Built a two-way relationship between  datasets using **Power Pivot**, enabling analysis across multiple related tables.
- Developed **PivotTables** to summarize patient satisfaction, wait times, treatment duration, staffing levels, treatment costs, readmission rates, and patient visit volume by department and shift.
- Created **PivotCharts**, including bar charts, scatter plots, and combination charts, to identify operational trends and relationships.
- Applied **Conditional Formatting** to highlight departments performing below the hospital's patient satisfaction target and identify cost variances from departmental baseline costs.
<img width="1741" height="801" alt="image" src="https://github.com/user-attachments/assets/006b2116-6fc8-432f-948c-b8c1b86c1e68" />

---

## Methodology & Exploratory Data Analysis (EDA)

Excel was used throughout this project, from data preparation and transformation to building the dashboard and visualizations.

The dashboard compares each department across several operational performance indicators, including:

- Average patient satisfaction
- Average patient wait time
- Average treatment time
- Average treatment charge to patients
- Departmental baseline cost
- Readmission rate
- Average staffing levels by shift
- Patient visit volume by shift

Several patterns emerged during the exploratory analysis:

- No department achieved the hospital's target patient satisfaction score of **3.8**.
- General Surgery consistently underperformed across multiple operational metrics, recording the lowest patient satisfaction alongside the longest wait times, treatment times, and highest treatment cost.
- Patient volume peaked during the **Evening** shift, yet four of the five departments operated with fewer staff on average than during the Day shift.
- The **Night** shift accounted for approximately **8.5% of total patient visits** while averaging roughly **20% of total staffing**, suggesting staffing resources may not align with patient demand.
- Four of the five departments recorded average or below-average patient wait times during the **Night** shift while also achieving above-average patient satisfaction, indicating that higher staffing levels and shorter wait times may contribute to improved patient experience.

---

## Key Findings & Visualizations

<img width="1740" height="800" alt="image" src="https://github.com/user-attachments/assets/59b3c119-24b3-4147-b254-407d3927a9dd" />


### Headline Metrics

| Metric | Value |
|:---|---:|
| Overall Average Patient Satisfaction | **3.48 / 5.00** |
| Hospital Satisfaction Target | **3.80 / 5.00** |
| Overall Average Wait Time | **53.91 minutes** |
| Overall Readmission Rate | **10.08%** |

---

### Department Performance

No department achieved the hospital's internal patient satisfaction target of **3.8**.

| Department | Average Satisfaction |
|:---|---:|
| Pediatrics | **3.77** |
| Cardiology | **3.62** |
| Emergency | **3.47** |
| Orthopedics | **3.31** |
| General Surgery | **3.21** |

General Surgery consistently performed worst across multiple operational metrics. It recorded the lowest patient satisfaction while also reporting the longest wait time, treatment duration and highest treatment costs, suggesting opportunities to improve operational efficiency before considering additional spending.

Pediatrics came closest to achieving the hospital's satisfaction target, however, the average treatment charge to patients came out below its departmental base cost.

---

### Staffing & Patient Volume

The relationship between staffing levels and patient demand revealed several noteworthy patterns.

The **Evening** shift experienced the highest patient volume across the hospital, yet four of the five departments operated with fewer staff on average than during the Day shift. This mismatch suggests an opportunity to better align staffing resources with patient demand.

Conversely, the **Night** shift accounted for only **approximately 8.5% of total patient visits** while representing roughly **20% of total staffing**. Four of the five departments also recorded average or below-average patient wait times and above-average satisfaction during Night shifts.

Although additional analysis would be required before recommending staffing reductions, these findings suggest Night shifts may be relatively overstaffed compared with patient demand, while Evening shifts may benefit from additional staffing resources.

---

### Wait Times & Patient Satisfaction

Average patient wait times varied across departments, ranging between **38 and 70 minutes**.

There was a correlation between high average wait times and average satisfaction in each department. 

For example:

- The Emergency, Orthopedics, and General Surgery departments recorded average wait times of **55, 64, and 69 minutes**, respectively, while reporting average patient satisfaction scores of **3.47, 3.31, and 3.21**, respectively.

These findings suggest that patient wait time plays a significant role in patient satisfaction.

---

## Limitations & Next Steps

This analysis was conducted using operational hospital data only and therefore required several assumptions based on general healthcare operations.

Patient satisfaction is influenced by many factors beyond those included in this dataset, including:

- Patient acuity
- Case complexity
- Physician communication
- Staff experience
- Bed availability
- Emergency department demand

Without these variables, the analysis identifies relationships rather than proving causation.

### Next Steps

- Incorporate patient acuity and diagnosis data to better account for treatment complexity across departments.
- Compare staffing levels with hourly patient arrivals to determine whether staffing shortages contribute to longer wait times.
- Analyze physician-level performance to identify additional variation within departments.
- Expand the analysis across multiple reporting periods to determine whether the observed trends remain consistent over time.
- Recreate the dashboard in Power BI to enable interactive filtering, drill-through analysis, and automated reporting.

---

## Recommendation

Based on the available evidence, Emergency and General Surgery should be prioritized for operational review, as both departments reported patient satisfaction scores below the hospital target of 3.8 while accounting for a significant proportion of total patient visits. Emergency had the highest patient volume (314 visits) with an average satisfaction score of 3.47, while General Surgery recorded the lowest satisfaction score (3.21), the longest average wait time (69.57 minutes), and the highest average treatment time (89.40 minutes). Improving performance in these departments would likely have the greatest impact on overall patient satisfaction due to their combination of high patient demand and below-target satisfaction scores.

More broadly, the staffing analysis suggests an opportunity to better align resources with patient demand. Evening shifts represent the highest volume period across most departments, yet staffing levels are generally lower compared to Day shifts. Conversely, Night shifts account for a smaller proportion of total patient visits while maintaining relatively strong satisfaction scores and, in some departments, comparable staffing levels.

Rather than increasing overall staffing levels, hospital leadership should evaluate whether staffing schedules can be better aligned with patient demand patterns, particularly during higher-volume Evening periods. A targeted review of department-specific workflows, hourly patient arrivals, patient acuity, and staffing requirements would help determine whether resource adjustments could reduce wait times and improve patient satisfaction while maintaining cost efficiency.
