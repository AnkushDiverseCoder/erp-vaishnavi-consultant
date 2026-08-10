# Monthly Salary & Weekly Rest — Full Guide

This explains how a **monthly-salaried** employee's pay is calculated, the new
**"Pay Full Monthly Salary"** option, and worked examples.

## Where to find it
**Payroll Config** → **Working Days** section → switch **"Pay Full Monthly Salary
when present all working days"** ON, then pick the **Absent-day deduction divisor**.

---

## The two modes

### Mode A — Prorate by paid days  (default, unchanged)
```
Earned = (Monthly Gross ÷ Working-Days-Basis) × Paid Days
Paid Days = Present + Paid Weekly-Off + Paid Holidays
```
Full salary happens only when Paid Days equals the divisor — which is reliable
**only when Working Days Basis = Calendar Days**. With Fixed 26 / Fixed 30 it can
over- or under-pay, because the weekly-rest days are added on top of a fixed
divisor. (This is the trap that made the full salary hard to hit.)

### Mode B — Pay Full Monthly Salary  (new, recommended for full-salary clients)
```
Present on ALL working days  → FULL monthly gross (always)
Each genuine ABSENT day      → deduct (Gross ÷ Divisor)
Earned = Gross − (Absent Days × Gross ÷ Divisor)
```
- **Immune to the weekly-rest count.** Whether the month has 4 or 5 Sundays, an
  employee who misses no working day gets the **full** salary.
- The **Divisor** (Calendar / 26 / 30) only decides how much **one absent day**
  costs — it does *not* affect the full-attendance amount.

---

## Worked examples — Gross ₹19,800, August (31 days)

Say August has **5 Sundays** ⇒ 26 working days.

| Attendance | Absent days | Mode B (Calendar ÷31) | Mode B (÷26) |
|---|---|---|---|
| Present all 26 working days | 0 | **₹19,800** (full) | **₹19,800** (full) |
| 1 working day absent | 1 | 19,800 − 19,800/31 = **₹19,161** | 19,800 − 19,800/26 = **₹19,038** |
| 2 working days absent | 2 | **₹18,523** | **₹18,277** |

If instead August had **4 Sundays** ⇒ 27 working days, an employee present on all
27 still gets the **full ₹19,800** in Mode B — no manual adjustment needed.

> "Absent days" = month days − present − weekly-off − holidays. Weekly-off and
> holidays are never counted as absent, so they never reduce the salary.

---

## How to set up BAJAJ EARTH (monthly salary, weekly rest paid in full)
1. **Payroll Config → Working Days**
   - Turn **ON** "Pay Full Monthly Salary when present all working days".
   - Absent-day deduction = **Gross ÷ Calendar Days** (recommended).
2. **Rest Day / Weekly Off** → **Fixed Sunday** (so real Sundays are counted).
3. **Weekly Off Policy** → **Paid**.
4. Make sure the employee's salary type is **Monthly** (not Daily Wages).
5. Save, then process the month — a fully-present employee reads **₹19,800**.

---

## Notes
- **Default is OFF**, so every existing establishment keeps its current
  calculation exactly. Only establishments you switch ON use the new mode.
- Applies to **monthly salary** only. Daily-wage employees are unchanged
  (they are paid per day worked, with the weekly-off policy you already set).
- Gross stays a whole number (rounded) everywhere — reports, ECR, 3A, uploads.
- Takes effect when a month is **processed / re-processed** (Save & Continue),
  since stored values are recomputed then.
