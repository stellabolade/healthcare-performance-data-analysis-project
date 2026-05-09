# healthcare-performance-data-analysis-project

**Author: Stella Omobolade Obase**
March 2026

---

## 📌 Project Overview

This project involved a comprehensive data cleaning exercise on a healthcare performance dataset containing **5,000 patient encounter records**. The goal was to identify and resolve data quality issues; ranging from formatting inconsistencies to deep logical contradictions, and prepare the dataset for reliable analysis and dashboard development.

**Tool Used:** Microsoft Excel
**Deliverable:** Cleaned dataset + Power BI Dashboard


## Dataset Structure

| Property | Detail |
|---|---|
| File Type | CSV (cleaned output: Excel) |
| Total Records | 5,000 encounter records |
| Unique Patients | 989 |
| Variables | 17 key columns |
| Departments | Emergency, Pediatrics, Orthopedics, General Surgery, Cardiology, Internal Medicine |
| ICD-10 Codes | B34, E11, I10, K35, M54, N39, J18, A09 |
| CPT Procedures | 93000, 70450, 80053, 99284, 81002, 71020 |

---

## 🔍 Key Data Quality Issues Identified

| # | Issue | Description |
|---|---|---|
| 1 | **Inconsistent Data Types** | 17 columns stored as General — reformatted to Text, Numeric, Date/Time, and Accounting |
| 2 | **Inconsistent Capitalization** | Categorical columns (e.g., Admission Type) had mixed casing e.g., "Emergency" vs "emergency" |
| 3 | **Invalid Gender Entries** | 150 records had "None" as gender value |
| 4 | **Department–Age Mismatch** | Pediatrics records contained adult patients (age > 18) |
| 5 | **Multiple Ages per Patient** | Same patient ID recorded with different ages across encounters |
| 6 | **Missing Cost Values** | Blank cells in the Cost of Encounter column |
| 7 | **Admission Type Contradiction** | Emergency department patients listed with Elective admission type |
| 8 | **Outcome–Disposition Conflict** | Records showing "Expired" outcome but "Home" disposition, and vice versa |
| 9 | **Readmission Flag Errors** | Re-admitted patients incorrectly flagged as 0 |
| 10 | **ED Arrival Time for Elective Patients** | Elective admissions incorrectly had ED arrival timestamps |

---

## 🛠️ Cleaning Steps Applied

### 1. Standardization & Formatting
- Applied `PROPER()` and `TRIM()` functions to fix casing and remove excess spaces across all categorical columns.
- Reformatted all 17 columns from General to their correct data types.

### 2. Gender Conflict Resolution
A PivotTable was created to count Male/Female entries per Patient ID. The most frequently occurring gender per patient was assigned using:
```excel
=IF(GETPIVOTDATA("gender", Sheet1!$A$3, "patient_id", A2, "gender", "Male") >
   GETPIVOTDATA("gender", Sheet1!$A$3, "patient_id", A2, "gender", "Female"),
   "Male", "Female")
```

### 3. Pediatrics Age Mismatch
Patients over 18 in the Pediatrics department were flagged as "Misclassified" rather than reassigned, preserving data integrity:
```excel
=IF(AND(K2>18, M2="Pediatrics"), "Misclassified", M2)
```

### 4. Multiple Ages per Patient
The **median age** per Patient ID was calculated and applied consistently across all records for that patient.

### 5. Missing Cost of Encounter Values
Blank cells were identified using Go To Special (Ctrl+G) and filled with the **median cost value** — chosen over the mean due to the skewed distribution of cost data.

### 6. Admission Type vs. Department Fix
```excel
=IF(O2="Emergency", "Emergency", P2)
```
Ensured all Emergency department records had "Emergency" as their admission type.

### 7. Readmission Flag Recalculation (30-Day Rule)
The original readmission flag was unreliable. A new flag was recalculated from scratch using actual admission and discharge timestamps:

```excel
-- Previous Discharge Date
=IFERROR(MAXIFS(F$2:F2, A$2:A2, A2, C$2:C2, "<"&C2, F$2:F2, "<"&C2), " ")

-- Days Since Last Discharge
=IF(D2=" ", " ", C2-D2)

-- Corrected Readmission Flag
=IF(E2<0, "Invalid", IF(AND(E2>0, E2<=30), 1, 0))
```
- `1` = Readmitted within 30 days
- `0` = Admitted after 30 days
- `Invalid` = Negative interval (data error)

### 8. Outcome & Disposition Conflict Resolution
Outcome was first updated based on the corrected readmission flag, then cross-checked against discharge disposition to ensure "Expired" records were consistent across both columns.

### 9. ED Arrival Time for Elective Patients
```excel
=IF(Q2="Elective", " ", F2)
```
ED arrival time cells were blanked out for all elective admissions, as these patients do not arrive through the Emergency Department.

### 10. Code Readability — Lookup Tables
ICD-10 and CPT codes were mapped to human-readable descriptions to improve interpretability for non-clinical stakeholders:

| ICD Code | Description | CPT Code | Description |
|---|---|---|---|
| B34 | Viral Infection | 93000 | ECG |
| E11 | Type II Diabetes | 70450 | CT Brain |
| I10 | Hypertension | 80053 | Lab Test |
| K35 | Appendicitis | 99284 | ED Visit |
| M54 | Dorsalgia | 81002 | Urinalysis |
| N39 | UTI | 71020 | Chest X-Ray |
| J18 | Pneumonia | | |
| A09 | Gastroenteritis | | |

---

## 💡 Key Decisions & Rationale

- **Median over Mean for cost imputation** — The cost distribution was skewed, making the median a more robust central estimate.
- **"Misclassified" flag for Pediatrics** — Rather than arbitrarily reassigning adult patients to other departments, a placeholder was used to preserve original data and flag the issue clearly.
- **Full readmission recalculation** — The 30-day readmission rate is a standard hospital performance metric. Rebuilding it from timestamps ensured accuracy and transparency.

---

## 📊 Dashboard

The cleaned dataset was used to build a performance dashboard in **Power BI** (see `/dashboard/` folder). The dashboard visualizes key metrics including readmission rates, department performance, encounter outcomes, and cost distribution.

---

## 🎯 Skills Demonstrated

- Data cleaning and quality assurance in Microsoft Excel
- Logical data validation and conflict resolution
- Medical data literacy (ICD-10, CPT codes, clinical classifications)
- Use of advanced Excel functions: `IF`, `AND`, `MAXIFS`, `IFERROR`, `GETPIVOTDATA`, `TRIM`, `PROPER`
- Data storytelling and dashboard development in Power BI
- Documentation and reporting

---

## 👩‍💻 About

**Stella Omobolade Obase**
Data Analyst Intern — Dataverse Africa
📅 March 2026

> *"A dataset that looks clean at first glance may still contain hidden issues that only become visible through careful investigation and thoughtful data preparation."*
