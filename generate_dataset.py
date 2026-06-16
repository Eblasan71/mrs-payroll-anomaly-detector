"""
MRS Payroll Dataset Generator — 200 employees x 4 weeks = 800 records
Master Roofing Solutions | Simulated timecard data with seeded errors
"""

import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import date

fake = Faker("en_US")
random.seed(42)
np.random.seed(42)

# ─── CONSTANTS ───────────────────────────────────────────────────────────────

STATES = ["CA", "TX", "FL", "NY", "WA"]
MIN_WAGE = {"CA": 16.50, "TX": 7.25, "FL": 13.00, "NY": 16.00, "WA": 16.28}
PAY_TYPES_VALID = ["hourly", "piece_rate", "salary"]

PIECE_RATES = {
    "shingle_install":   2.80,
    "felt_underlayment": 1.20,
    "ridge_cap":         3.50,
    "flashing":          4.00,
    "tear_off":          1.80,
    "decking_repair":    5.00,
}

WEEK_ENDINGS = [
    date(2025, 5, 16),
    date(2025, 5, 23),
    date(2025, 5, 30),
    date(2025, 6,  6),
]

N_EMPLOYEES = 200

# ─── EMPLOYEE ROSTER ─────────────────────────────────────────────────────────

employees = []
for i in range(1, N_EMPLOYEES + 1):
    state    = random.choices(STATES, weights=[35, 25, 20, 10, 10])[0]
    pay_type = random.choices(PAY_TYPES_VALID, weights=[20, 70, 10])[0]
    task     = random.choice(list(PIECE_RATES.keys()))
    employees.append({
        "employee_id":    f"MRS{i:03d}",
        "employee_name":  fake.name(),
        "state":          state,
        "pay_type":       pay_type,
        "task_code":      task if pay_type == "piece_rate" else None,
        "piece_rate_usd": PIECE_RATES[task] if pay_type == "piece_rate" else None,
        "hourly_rate":    round(random.uniform(17, 28), 2) if pay_type != "piece_rate" else None,
    })

# ─── GENERATE BASE RECORDS ───────────────────────────────────────────────────

records = []
for emp in employees:
    for week in WEEK_ENDINGS:
        shift_hours = round(random.uniform(6, 11), 1)
        actual_total = round(shift_hours * 5, 1)
        reg_hours   = min(actual_total, 40.0)
        ot_hours    = round(max(actual_total - 40, 0), 1)
        pieces      = random.randint(60, 180) if emp["pay_type"] == "piece_rate" else None

        records.append({
            "employee_id":       emp["employee_id"],
            "employee_name":     emp["employee_name"],
            "state":             emp["state"],
            "pay_type":          emp["pay_type"],
            "task_code":         emp["task_code"],
            "week_ending":       week,
            "shift_hours":       shift_hours,
            "regular_hours":     reg_hours,
            "ot_hours":          ot_hours,
            "double_time_hours": 0.0,
            "pieces_completed":  pieces,
            "piece_rate_usd":    emp["piece_rate_usd"],
            "hourly_rate":       emp["hourly_rate"],
            "meal_break_taken":  True,
        })

df = pd.DataFrame(records)

# ─── SEED ERRORS ─────────────────────────────────────────────────────────────

hourly_salary_ids = [e["employee_id"] for e in employees if e["pay_type"] != "piece_rate"]
piece_rate_ids    = [e["employee_id"] for e in employees if e["pay_type"] == "piece_rate"]
ca_ids            = [e["employee_id"] for e in employees if e["state"] == "CA"]

# E01 — R03: hourly_rate = 0 (8 employees)
e01_ids = random.sample(hourly_salary_ids, 8)
df.loc[df["employee_id"].isin(e01_ids), "hourly_rate"] = 0.0

# E02 — R03: piece_rate_usd = NaN (6 employees)
e02_ids = random.sample([x for x in piece_rate_ids if x not in e01_ids], 6)
df.loc[df["employee_id"].isin(e02_ids), "piece_rate_usd"] = np.nan

# E03 — R05: regular_hours > 40 but ot_hours = 0 (12 employees, weeks 2 & 3)
used = set(e01_ids + e02_ids)
e03_ids = random.sample([x for x in df["employee_id"].unique() if x not in used], 12)
mask_e03 = df["employee_id"].isin(e03_ids) & df["week_ending"].isin([WEEK_ENDINGS[1], WEEK_ENDINGS[2]])
df.loc[mask_e03, "regular_hours"] = round(random.uniform(43, 52), 1)
df.loc[mask_e03, "ot_hours"] = 0.0
used.update(e03_ids)

# E04 — R06: CA daily OT ignored (10 CA employees, week 1)
ca_pool = [x for x in ca_ids if x not in used]
e04_ids = random.sample(ca_pool, min(10, len(ca_pool)))
mask_e04 = df["employee_id"].isin(e04_ids) & (df["week_ending"] == WEEK_ENDINGS[0])
df.loc[mask_e04, "shift_hours"] = round(random.uniform(8.5, 10), 1)
df.loc[mask_e04, "ot_hours"] = 0.0
used.update(e04_ids)

# E05 — R07: CA double time missing (6 CA employees, week 4)
ca_pool2 = [x for x in ca_pool if x not in used]
e05_ids = random.sample(ca_pool2, min(6, len(ca_pool2)))
mask_e05 = df["employee_id"].isin(e05_ids) & (df["week_ending"] == WEEK_ENDINGS[3])
df.loc[mask_e05, "shift_hours"] = round(random.uniform(12.5, 14), 1)
df.loc[mask_e05, "double_time_hours"] = 0.0
used.update(e05_ids)

# E06 — R04: pay_type = "contract" (8 employees)
remaining = [x for x in df["employee_id"].unique() if x not in used]
e06_ids = random.sample(remaining, 8)
df.loc[df["employee_id"].isin(e06_ids), "pay_type"] = "contract"
used.update(e06_ids)

# E07 — R02: duplicate records (4 employees)
remaining2 = [x for x in df["employee_id"].unique() if x not in used]
e07_ids = random.sample(remaining2, 4)
for eid in e07_ids:
    dup = df[(df["employee_id"] == eid) & (df["week_ending"] == WEEK_ENDINGS[2])].copy()
    df = pd.concat([df, dup], ignore_index=True)

# E08 — R10: CA meal break not taken (10 CA employees, weeks 1 & 3)
ca_pool3 = [x for x in ca_ids if x not in used]
e08_ids = random.sample(ca_pool3, min(10, len(ca_pool3)))
mask_e08 = df["employee_id"].isin(e08_ids) & df["week_ending"].isin([WEEK_ENDINGS[0], WEEK_ENDINGS[2]])
df.loc[mask_e08, "meal_break_taken"] = False

# E09 — R08: piece-rate < min wage (8 employees, weeks 3 & 4)
e09_ids = random.sample([x for x in piece_rate_ids if x not in e02_ids], 8)
mask_e09 = df["employee_id"].isin(e09_ids) & df["week_ending"].isin([WEEK_ENDINGS[2], WEEK_ENDINGS[3]])
df.loc[mask_e09, "piece_rate_usd"] = 0.15
df.loc[mask_e09, "pieces_completed"] = random.randint(10, 25)

# ─── SORT & EXPORT ───────────────────────────────────────────────────────────

df = df.sort_values(["employee_id", "week_ending"]).reset_index(drop=True)
df.to_csv("/home/claude/mrs_payroll_raw.csv", index=False)

print(f"Total records    : {len(df)}")
print(f"Unique employees : {df['employee_id'].nunique()}")
print(f"Week endings     : {sorted(df['week_ending'].unique())}")
print(f"\nPay type distribution:\n{df['pay_type'].value_counts()}")
print(f"\nState distribution:\n{df['state'].value_counts()}")
print(f"\nSeeded errors:")
print(f"  E01 (rate=0)       : {len(e01_ids)} employees")
print(f"  E02 (rate=null)    : {len(e02_ids)} employees")
print(f"  E03 (OT hidden)    : {len(e03_ids)} employees")
print(f"  E04 (CA daily OT)  : {len(e04_ids)} employees")
print(f"  E05 (CA DT miss)   : {len(e05_ids)} employees")
print(f"  E06 (bad pay_type) : {len(e06_ids)} employees")
print(f"  E07 (duplicates)   : {len(e07_ids)} employees")
print(f"  E08 (meal break)   : {len(e08_ids)} employees")
print(f"  E09 (MW breach)    : {len(e09_ids)} employees")
