# SAP Order Management & Data Integrity Project 
**Period:** January - March 2025  
**Tools:** SAP SD (VA01 / VA02 / VA03 / VK13) | Python (pandas, openpyxl) | Excel | Power Query  

---

## Project Overview

A structured end-to-end SAP order management project simulating a B2B enterprise environment. 
The project covers the full SAP SD order lifecycle - from order creation and pricing validation 
through to discrepancy resolution and KPI reporting.

---

## What This Project Demonstrates

| Skill | Detail |
|---|---|
| Order Entry | 200 customer orders created and maintained (VA01 / VA02 / VA03 equivalent) |
| Pricing Validation | All orders cross-checked against standard condition records (VK13 equivalent) |
| Discrepancy Resolution | 24 errors identified, root-caused, and resolved with full audit trail |
| Reporting | Weekly KPI dashboard - order volume, error rate, resolution performance |
| Process Improvement | Validation SOP developed - 30% estimated reduction in rework time |
| Python Automation | Script automates 4-point validation check across all 200 orders |

---

## Project Files

| File | Description |
|---|---|
| `SAP_Order_Management_Project.xlsx` | Main project workbook - 6 sheets including Order Register, Discrepancy Log, KPI Dashboard, Validation Checklist, and Pricing Master |
| `SAP_Validation_Report.xlsx` | Python-generated validation report - order-level pass/fail results, executive summary, weekly error trend |
| `sap_validation_script.py` | Python script - automates pricing validation, mandatory field checks, line value integrity, and date logic across all orders |

---

## Key Results

- **200** customer orders processed across 10 B2B customers and 12 product lines
- **24** discrepancies identified and fully resolved
- Error rate reduced from **12%** (Week 1) to **0%** (Week 8)
- **4-check validation framework** built and documented as reusable SOP
- Python script flags pricing deviations >5% from VK13 condition records automatically

---

## How the Workflow Maps to SAP

| This Project | SAP Equivalent |
|---|---|
| Order Register (Excel) | VA01 - Create Sales Order |
| Order amendments logged | VA02 - Change Sales Order |
| VA03 review column | VA03 - Display Sales Order |
| Pricing Master sheet | VK13 - Display Condition Records |
| Python validation script | Automated pre-release order check |
| Discrepancy Log | Change document / audit trail |

---

## Skills Demonstrated

`SAP SD` `Order Management` `Data Validation` `Pricing Compliance` 
`Discrepancy Resolution` `Python` `pandas` `openpyxl` `Excel` 
`Power Query` `KPI Reporting` `Process Improvement` `Data Integrity`
