# 🏗️ Payroll Anomaly Detector
### Automated Pre-Processing Validation for U.S. Multi-State Payroll Operations

---

## Overview

In companies with field-based workforces, payroll processing follows a flow that no single software fully covers:

```
Field Supervisors → Timecards (external system) → Export CSV → [GAP] → Import to Payroll Software → Process & Pay
```

**That gap — between the timecard export and the payroll system import — is where most errors happen.**

Missing pay rates, uncaptured overtime, minimum wage breaches, duplicate records, and gross pay anomalies all live in that interval. In a manual workflow, catching them depends entirely on the analyst's attention. This project automates that validation layer.

---

## What This Project Does

`payroll_validator.py` receives a CSV file of weekly timecard data, runs **12 compliance and reconciliation rules**, and generates a formatted Excel report with three sheets:

- **Summary** — KPI cards, severity breakdown, and flags by rule
- **Anomalies** — Full log of every flag with description and potential business impact
- **Dashboard** — Bar charts by rule, by state, and by severity

---

## Real-World Context

This project was designed around the operational reality of **Master Roofing Solutions (MRS)**, a U.S.-based roofing contractor processing payroll for 300–500 field employees weekly across multiple states.

Timecards were submitted by field supervisors and manually entered into **EnterTimeOnline**, then exported as CSV and imported into the payroll system for processing. The validation step between export and import was done manually — this script automates it.

The dataset used for demonstration is simulated but mirrors real payroll structures: `piece_rate`, `hourly`, and `salary` employees across California, Texas, Florida, New York, and Washington, with intentional errors seeded to demonstrate each validation rule.

---

## Validation Rules

### Group 1 — Data Quality
| Rule | Description | Severity |
|------|-------------|----------|
| R02 | Duplicate record: same `employee_id` + `week_ending` | CRITICAL |
| R03 | Missing or zero pay rate (`hourly_rate` or `piece_rate_usd`) | CRITICAL |
| R04 | Invalid `pay_type` value (not hourly, piece_rate, or salary) | HIGH |

### Group 2 — FLSA Overtime (Federal)
| Rule | Description | Severity |
|------|-------------|----------|
| R05 | Employee worked >40h but `ot_hours = 0` — weekly OT not captured | CRITICAL |

### Group 3 — California Overtime (CA Labor Code §510)
| Rule | Description | Severity |
|------|-------------|----------|
| R06 | CA employee: `shift_hours > 8h` but `ot_hours = 0` — daily OT not captured | CRITICAL |
| R07 | CA employee: `shift_hours > 12h` but `double_time_hours = 0` — double time not applied | HIGH |

### Group 4 — Piece-Rate Minimum Wage Floor
| Rule | Description | Severity |
|------|-------------|----------|
| R08 | Piece-rate effective hourly rate < state minimum wage | CRITICAL |

### Group 5 — California Meal Break (CA Labor Code §512)
| Rule | Description | Severity |
|------|-------------|----------|
| R10 | CA employee: shift >5h, meal break not taken — 1st premium owed | HIGH |
| R11 | CA employee: shift >10h, meal break not taken — 2nd premium may apply | MEDIUM |

### Group 6 — Reconciliation (Variance Analysis)
| Rule | Description | Severity |
|------|-------------|----------|
| R22 | Individual gross pay varies >40% vs. personal historical average | HIGH |
| R23 | Company total payroll varies >15% week-over-week | HIGH |
| R24 | Employee hours exceed personal mean + 2 standard deviations | MEDIUM |

---

## Minimum Wage Reference (2025)

| State | Min Wage/hr |
|-------|------------|
| CA    | $16.50     |
| NY    | $16.00     |
| WA    | $16.28     |
| FL    | $13.00     |
| TX    | $7.25      |

---

## Project Scope

**In scope:**
- Pre-processing validation of timecard CSV exports
- Federal compliance (FLSA) and California-specific compliance (Labor Code §510, §512, IWC Wage Order)
- Multi-state minimum wage enforcement
- Variance reconciliation across pay periods

**Out of scope:**
- Direct integration with payroll software (PrismHR, ADP, Paychex, Gusto, etc.)
- Net pay calculation and tax withholding
- Payslip generation
- Employee roster management

**Designed for:**
Companies in construction, roofing, landscaping, staffing, or agriculture with 50–1,000 field employees, multi-state operations, and a timecard export → import workflow.

---

## Project Structure

```
mrs-payroll-anomaly-detector/
│
├── payroll_validator.py       # Main script: validation engine + Excel report generator
├── generate_dataset.py        # Simulated dataset generator with seeded errors
├── mrs_payroll_raw.csv        # Input: simulated timecard data (121 records, 30 employees)
├── MRS_Payroll_Validation_Report.xlsx  # Output: formatted 3-sheet Excel report
├── requirements.txt           # Python dependencies
└── README.md                  # This file
```

---

## How to Run

**1. Clone the repository**
```bash
git clone https://github.com/your-username/mrs-payroll-anomaly-detector.git
cd mrs-payroll-anomaly-detector
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Run the validator**
```bash
python payroll_validator.py
```

The report will be generated as `MRS_Payroll_Validation_Report.xlsx` in the same directory.

**4. To regenerate the simulated dataset**
```bash
python generate_dataset.py
```

---

## Input File Format

The script expects a CSV with the following columns:

| Column | Type | Description |
|--------|------|-------------|
| `employee_id` | string | Unique employee identifier (e.g., MRS001) |
| `employee_name` | string | Full name |
| `state` | string | Work state (CA, TX, FL, NY, WA) |
| `pay_type` | string | `hourly`, `piece_rate`, or `salary` |
| `task_code` | string | Piece-rate task identifier (piece_rate only) |
| `week_ending` | date | Pay period end date (YYYY-MM-DD) |
| `shift_hours` | float | Total shift duration in hours |
| `regular_hours` | float | Regular hours worked |
| `ot_hours` | float | Overtime hours reported |
| `double_time_hours` | float | Double time hours (CA only) |
| `pieces_completed` | int | Units completed (piece_rate only) |
| `piece_rate_usd` | float | Rate per piece in USD |
| `hourly_rate` | float | Hourly rate in USD (hourly/salary only) |
| `meal_break_taken` | bool | Whether meal break was taken (True/False) |

---

## Sample Output

**Summary Sheet — 92 flags detected across 121 records:**

| Severity | Count |
|----------|-------|
| CRITICAL | 43 |
| HIGH | 40 |
| MEDIUM | 9 |

**Top flags by rule:**
- R03 (missing pay rate): 20 flags
- R06 (CA daily OT): 9 flags
- R05 (FLSA weekly OT): 8 flags
- R22 (gross pay variance): 24 flags

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| Python 3.x | Core scripting language |
| pandas | Data loading, transformation, and validation logic |
| numpy | Statistical calculations (std deviation for R24) |
| openpyxl | Excel report generation with formatting and charts |


---

## Author

**Edgar Blanco**
Payroll Analyst | Data & Automation | Remote · Rionegro, Antioquia, Colombia

3+ years of full-cycle U.S. payroll operations across multi-state environments.
Experience with PrismHR, EnterTimeOnline, California Labor Code compliance, and piece-rate workforce management.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://linkedin.com/in/your-profile)

---

## License

MIT License — free to use, adapt, and build upon with attribution.
