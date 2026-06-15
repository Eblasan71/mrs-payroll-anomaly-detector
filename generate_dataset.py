import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import date, timedelta

fake = Faker("en_US")
random.seed(42)
np.random.seed(42)

# ─── CONSTANTS ───────────────────────────────────────────────────────────────

STATES = ["CA", "TX", "FL", "NY", "WA"]
MIN_WAGE = {"CA": 16.50, "TX": 7.25, "FL": 13.00, "NY": 16.00, "WA": 16.28}
PAY_TYPES_VALID = ["hourly", "piece_rate", "salary"]

# MRS task codes (piece-rate dominant, roofing context)
PIECE_RATES = {
    "shingle_install":   2.80,
    "felt_underlayment": 1.20,
    "ridge_cap":         3.50,
    "flashing":          4.00,
    "tear_off":          1.80,
    "decking_repair":    5.00,
}

# 4 week-ending dates (Fridays)
WEEK_ENDINGS = [
    date(2025, 5, 16),
    date(2025, 5, 23),
    date(2025, 5, 30),
    date(2025, 6,  6),
]

# ─── EMPLOYEE ROSTER ─────────────────────────────────────────────────────────

employees = []
for i in range(1, 31):
    state = random.choices(STATES, weights=[35, 25, 20, 10, 10])[0]
    pay_type = random.choices(PAY_TYPES_VALID, weights=[20, 70, 10])[0]
    task = random.choice(list(PIECE_RATES.keys()))
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
        reg_hours   = min(shift_hours * 5, 40.0)          # 5-day week cap at 40 baseline
        actual_total = round(shift_hours * 5, 1)
        ot_hours    = round(max(actual_total - 40, 0), 1)
        pieces      = random.randint(60, 180) if emp["pay_type"] == "piece_rate" else None

        records.append({
            "employee_id":      emp["employee_id"],
            "employee_name":    emp["employee_name"],
            "state":            emp["state"],
            "pay_type":         emp["pay_type"],
            "task_code":        emp["task_code"],
            "week_ending":      week,
            "shift_hours":      shift_hours,
            "regular_hours":    reg_hours,
            "ot_hours":         ot_hours,
            "double_time_hours":0.0,
            "pieces_completed": pieces,
            "piece_rate_usd":   emp["piece_rate_usd"],
            "hourly_rate":      emp["hourly_rate"],
            "meal_break_taken": True,
        })

df = pd.DataFrame(records)

# ─── SEED ERRORS ─────────────────────────────────────────────────────────────

def flag_rows(emp_ids, week=None):
    mask = df["employee_id"].isin(emp_ids)
    if week:
        mask &= df["week_ending"] == week
    return mask

# Pick specific employee pools by pay type / state for realistic errors
hourly_salary_ids = [e["employee_id"] for e in employees if e["pay_type"] != "piece_rate"]
piece_rate_ids    = [e["employee_id"] for e in employees if e["pay_type"] == "piece_rate"]
ca_ids            = [e["employee_id"] for e in employees if e["state"] == "CA"]

# E01 — R03: hourly_rate = 0 (3 hourly/salary employees, all weeks)
e01_ids = random.sample(hourly_salary_ids, 3)
df.loc[df["employee_id"].isin(e01_ids), "hourly_rate"] = 0.0

# E02 — R03: piece_rate_usd = NaN (2 piece-rate employees, all weeks)
e02_ids = random.sample([x for x in piece_rate_ids if x not in e01_ids], 2)
df.loc[df["employee_id"].isin(e02_ids), "piece_rate_usd"] = np.nan

# E03 — R05: regular_hours > 40 but ot_hours = 0 (4 employees, weeks 2 & 3)
e03_ids = random.sample([x for x in df["employee_id"].unique().tolist() if x not in e01_ids + e02_ids], 4)
mask_e03 = df["employee_id"].isin(e03_ids) & df["week_ending"].isin([WEEK_ENDINGS[1], WEEK_ENDINGS[2]])
df.loc[mask_e03, "regular_hours"] = round(random.uniform(43, 52), 1)
df.loc[mask_e03, "ot_hours"] = 0.0   # ← error: OT not captured

# E04 — R06: CA employees shift_hours > 8 but ot_hours = 0 (3 CA employees, week 1)
ca_pool = [x for x in ca_ids if x not in e01_ids + e02_ids + e03_ids]
e04_ids = random.sample(ca_pool, min(3, len(ca_pool)))
mask_e04 = df["employee_id"].isin(e04_ids) & (df["week_ending"] == WEEK_ENDINGS[0])
df.loc[mask_e04, "shift_hours"] = round(random.uniform(8.5, 10), 1)
df.loc[mask_e04, "ot_hours"] = 0.0   # ← error: daily OT not captured

# E05 — R07: CA employees shift_hours > 12, no double_time (2 CA employees, week 4)
ca_pool2 = [x for x in ca_pool if x not in e04_ids]
e05_ids = random.sample(ca_pool2, min(2, len(ca_pool2)))
mask_e05 = df["employee_id"].isin(e05_ids) & (df["week_ending"] == WEEK_ENDINGS[3])
df.loc[mask_e05, "shift_hours"] = round(random.uniform(12.5, 14), 1)
df.loc[mask_e05, "double_time_hours"] = 0.0   # ← error: should have DT

# E06 — R04: pay_type = "contract" (2 employees, all weeks)
remaining = [x for x in df["employee_id"].unique() if x not in
             e01_ids + e02_ids + e03_ids + e04_ids + e05_ids]
e06_ids = random.sample(remaining, 2)
df.loc[df["employee_id"].isin(e06_ids), "pay_type"] = "contract"

# E07 — R02: duplicate record (same employee_id + week_ending appears twice)
e07_id = random.choice([x for x in remaining if x not in e06_ids])
dup_row = df[(df["employee_id"] == e07_id) & (df["week_ending"] == WEEK_ENDINGS[2])].copy()
df = pd.concat([df, dup_row], ignore_index=True)

# E08 — R10: CA, shift_hours > 5, meal_break_taken = False (3 CA employees, weeks 1 & 3)
ca_pool3 = [x for x in ca_ids if x not in e01_ids + e02_ids]
e08_ids = random.sample(ca_pool3, min(3, len(ca_pool3)))
mask_e08 = df["employee_id"].isin(e08_ids) & df["week_ending"].isin([WEEK_ENDINGS[0], WEEK_ENDINGS[2]])
df.loc[mask_e08, "meal_break_taken"] = False

# E09 — R08: piece_rate earnings/hr < state min wage (2 piece-rate employees, weeks 3 & 4)
e09_ids = random.sample([x for x in piece_rate_ids if x not in e02_ids], 2)
mask_e09 = df["employee_id"].isin(e09_ids) & df["week_ending"].isin([WEEK_ENDINGS[2], WEEK_ENDINGS[3]])
df.loc[mask_e09, "piece_rate_usd"] = 0.15   # artificially low → breach
df.loc[mask_e09, "pieces_completed"] = random.randint(10, 25)

# ─── SORT & EXPORT ───────────────────────────────────────────────────────────

df = df.sort_values(["employee_id", "week_ending"]).reset_index(drop=True)
df.to_csv("/home/claude/mrs_payroll_raw.csv", index=False)

print(f"Total records : {len(df)}")
print(f"Unique employees : {df['employee_id'].nunique()}")
print(f"Week endings : {sorted(df['week_ending'].unique())}")
print(f"\nPay type distribution:\n{df['pay_type'].value_counts()}")
print(f"\nState distribution:\n{df['state'].value_counts()}")
print(f"\nSample of seeded errors:")
print(f"  E01 (rate=0)    : {e01_ids}")
print(f"  E02 (rate=null) : {e02_ids}")
print(f"  E03 (OT hidden) : {e03_ids}")
print(f"  E04 (CA daily OT): {e04_ids}")
print(f"  E05 (CA DT miss): {e05_ids}")
print(f"  E06 (bad type)  : {e06_ids}")
print(f"  E07 (duplicate) : {e07_id}")
print(f"  E08 (meal break): {e08_ids}")
print(f"  E09 (MW breach) : {e09_ids}")
