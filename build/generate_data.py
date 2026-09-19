from __future__ import annotations

import json
import math
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
OUTPUT = ROOT / "deliverables"
DATA.mkdir(exist_ok=True)
OUTPUT.mkdir(exist_ok=True)

rng = np.random.default_rng(240918)
months = pd.date_range("2023-01-01", "2026-12-01", freq="MS")
tiers = pd.DataFrame({
    "plan": ["Starter", "Growth", "Enterprise"],
    "monthly_price": [99.0, 349.0, 1499.0],
    "starting_mix": [0.61, 0.31, 0.08],
})

scenario = pd.DataFrame([
    ["Base", 0.045, 0.018, 1.00, 1.00, 1.00, 1.00],
    ["Upside", 0.060, 0.014, 1.04, 1.10, 1.12, 0.95],
    ["Downside", 0.025, 0.026, 0.97, 0.75, 0.82, 1.12],
], columns=["scenario", "monthly_customer_growth", "monthly_churn", "price_multiplier", "hiring_multiplier", "marketing_multiplier", "cloud_unit_cost_multiplier"])

# Monthly customer and subscription history by plan.
customer_rows = []
end_by_plan = {"Starter": 320, "Growth": 145, "Enterprise": 28}
for i, month in enumerate(months):
    forecast = month >= pd.Timestamp("2026-07-01")
    for _, p in tiers.iterrows():
        plan = p["plan"]
        start = end_by_plan[plan]
        base_growth = {"Starter": .040, "Growth": .048, "Enterprise": .030}[plan]
        base_churn = {"Starter": .024, "Growth": .015, "Enterprise": .008}[plan]
        seasonality = 1 + 0.12 * math.sin((month.month - 1) / 12 * 2 * math.pi)
        if forecast:
            new = round(start * base_growth * seasonality)
            churned = round(start * base_churn)
        else:
            new = max(1, round(start * base_growth * seasonality * rng.normal(1.0, .10)))
            churned = max(0, round(start * base_churn * rng.normal(1.0, .12)))
        ending = start + new - churned
        expansion = {"Starter": 0.0, "Growth": 0.012, "Enterprise": 0.022}[plan]
        realized_price = p["monthly_price"] * (1 + 0.035) ** max(0, month.year - 2023)
        mrr = ending * realized_price * (1 + expansion)
        customer_rows.append([month, plan, start, new, churned, ending, realized_price, mrr, forecast])
        end_by_plan[plan] = ending

subscriptions = pd.DataFrame(customer_rows, columns=["month", "plan", "starting_customers", "new_customers", "churned_customers", "ending_customers", "price_per_month", "mrr", "is_forecast"])

# Employee roster with planned and actual dates.
departments = ["Engineering", "Sales", "Marketing", "Customer Success", "G&A"]
base_salaries = {"Engineering": 145000, "Sales": 112000, "Marketing": 105000, "Customer Success": 92000, "G&A": 125000}
employees = []
emp_id = 1001
for dept, count in zip(departments, [19, 12, 7, 8, 6]):
    for j in range(count):
        planned = pd.Timestamp("2022-09-01") + pd.DateOffset(months=int(rng.integers(0, 45)))
        actual = planned + pd.Timedelta(days=int(rng.integers(-20, 46)))
        if actual > pd.Timestamp("2026-06-30"):
            actual = pd.NaT
        salary = round(base_salaries[dept] * rng.uniform(.78, 1.28), -2)
        employees.append([f"E{emp_id}", dept, planned.date(), actual.date() if pd.notna(actual) else None, salary, 0.22])
        emp_id += 1
employees = pd.DataFrame(employees, columns=["employee_id", "department", "planned_hire_date", "actual_hire_date", "annual_salary", "benefits_rate"])

# Monthly headcount and personnel cost by department.
hc_rows = []
for month in months:
    month_end = month + pd.offsets.MonthEnd(0)
    for dept in departments:
        sub = employees[employees.department.eq(dept)]
        planned = int(pd.to_datetime(sub.planned_hire_date).le(month_end).sum())
        actual = int(pd.to_datetime(sub.actual_hire_date).le(month_end).sum())
        use_hc = actual if month < pd.Timestamp("2026-07-01") else planned
        monthly_salary = sub.loc[pd.to_datetime(sub.actual_hire_date if month < pd.Timestamp("2026-07-01") else sub.planned_hire_date).le(month_end), "annual_salary"].sum() / 12
        benefits = monthly_salary * .22
        hc_rows.append([month, dept, planned, actual, use_hc, monthly_salary, benefits, monthly_salary + benefits])
headcount = pd.DataFrame(hc_rows, columns=["month", "department", "planned_headcount", "actual_headcount", "reported_headcount", "salary_cost", "benefits_cost", "personnel_cost"])

# Operating expenses and budget.
opex_rows = []
for i, month in enumerate(months):
    total_mrr = subscriptions.loc[subscriptions.month.eq(month), "mrr"].sum()
    scale = 1 + i * .012
    budget_map = {
        "Marketing": 90000 * scale,
        "Cloud & API": total_mrr * .115,
        "Software": 26000 * scale,
        "Office & Admin": 31000 * scale,
        "Professional Services": 18000 * scale,
    }
    for category, budget in budget_map.items():
        fixed = category not in ("Marketing", "Cloud & API")
        if month < pd.Timestamp("2026-07-01"):
            bias = {"Marketing": 1.07, "Cloud & API": 1.09, "Software": 1.03, "Office & Admin": .97, "Professional Services": 1.05}[category]
            actual = budget * bias * rng.normal(1.0, .045)
        else:
            actual = np.nan
        opex_rows.append([month, category, round(budget, 2), round(actual, 2) if not np.isnan(actual) else None, "Fixed" if fixed else "Variable"])
opex = pd.DataFrame(opex_rows, columns=["month", "category", "budget", "actual", "cost_type"])

# Consolidated P&L actual/budget and base forecast.
financial_rows = []
starting_cash = 18_000_000.0
cash = starting_cash
for month in months:
    revenue = subscriptions.loc[subscriptions.month.eq(month), "mrr"].sum()
    budget_revenue = revenue / (1.025 if month < pd.Timestamp("2026-07-01") else 1.0)
    cogs = revenue * (.19 + .008 * math.sin(month.month / 12 * 2 * math.pi))
    payroll = headcount.loc[headcount.month.eq(month), "personnel_cost"].sum()
    other_opex = opex.loc[opex.month.eq(month), "actual"].sum(min_count=1)
    budget_other = opex.loc[opex.month.eq(month), "budget"].sum()
    actual_revenue = revenue if month < pd.Timestamp("2026-07-01") else None
    actual_cogs = cogs * rng.normal(1.0, .018) if month < pd.Timestamp("2026-07-01") else None
    actual_payroll = payroll if month < pd.Timestamp("2026-07-01") else None
    actual_opex = other_opex if month < pd.Timestamp("2026-07-01") else None
    forecast_revenue = revenue
    forecast_cogs = cogs
    forecast_payroll = payroll
    forecast_opex = budget_other
    operating_cash = forecast_revenue - forecast_cogs - forecast_payroll - forecast_opex
    capex = 45000 if month.month in (1, 7) else 15000
    financing = 4_000_000 if month == pd.Timestamp("2024-09-01") else 0
    ending_cash = cash + operating_cash - capex + financing
    financial_rows.append([month, budget_revenue, budget_revenue * .19, payroll * 1.02, budget_other, actual_revenue, actual_cogs, actual_payroll, actual_opex, forecast_revenue, forecast_cogs, forecast_payroll, forecast_opex, cash, capex, financing, ending_cash])
    cash = ending_cash
financials = pd.DataFrame(financial_rows, columns=["month", "budget_revenue", "budget_cogs", "budget_payroll", "budget_opex", "actual_revenue", "actual_cogs", "actual_payroll", "actual_opex", "forecast_revenue", "forecast_cogs", "forecast_payroll", "forecast_opex", "starting_cash", "capex", "financing", "ending_cash"])

for frame, name in [(tiers, "pricing_plans"), (scenario, "scenario_assumptions"), (subscriptions, "subscriptions"), (employees, "employees"), (headcount, "headcount_monthly"), (opex, "operating_expenses"), (financials, "financials_monthly")]:
    frame.to_csv(DATA / f"{name}.csv", index=False, date_format="%Y-%m-%d")

db_path = DATA / "saas_fpa.db"
with sqlite3.connect(db_path) as conn:
    for frame, name in [(tiers, "pricing_plans"), (scenario, "scenario_assumptions"), (subscriptions, "subscriptions"), (employees, "employees"), (headcount, "headcount_monthly"), (opex, "operating_expenses"), (financials, "financials_monthly")]:
        frame.to_sql(name, conn, if_exists="replace", index=False)

actual_2026 = financials[financials.month.between("2026-01-01", "2026-06-01")]
summary = {
    "as_of": "2026-06-30",
    "latest_mrr": float(subscriptions.loc[subscriptions.month.eq(pd.Timestamp("2026-06-01")), "mrr"].sum()),
    "latest_arr": float(subscriptions.loc[subscriptions.month.eq(pd.Timestamp("2026-06-01")), "mrr"].sum() * 12),
    "ytd_revenue": float(actual_2026.actual_revenue.sum()),
    "ytd_budget_revenue": float(actual_2026.budget_revenue.sum()),
    "ytd_operating_loss": float((actual_2026.actual_revenue - actual_2026.actual_cogs - actual_2026.actual_payroll - actual_2026.actual_opex).sum()),
    "ending_cash_jun_2026": float(financials.loc[financials.month.eq(pd.Timestamp("2026-06-01")), "ending_cash"].iloc[0]),
}
(DATA / "management_summary.json").write_text(json.dumps(summary, indent=2))

# Advanced SaaS analytics layer: cohorts, ARR bridge, accrual/cash timing, back-test, and liquidity simulation.
metric_rows = []
prev_mrr = None
for month in months:
    sm = subscriptions[subscriptions.month.eq(month)]
    start_mrr = 0.0 if prev_mrr is None else prev_mrr
    new_mrr = float((sm.new_customers * sm.price_per_month).sum())
    churn_mrr = float((sm.churned_customers * sm.price_per_month).sum())
    expansion_mrr = float(sm.mrr.sum() - start_mrr - new_mrr + churn_mrr) if prev_mrr is not None else 0.0
    ending_mrr = float(sm.mrr.sum())
    grr = max(0.0, 1 - churn_mrr / start_mrr) if start_mrr else 1.0
    nrr = (start_mrr - churn_mrr + expansion_mrr) / start_mrr if start_mrr else 1.0
    marketing = float(opex.loc[(opex.month.eq(month)) & opex.category.eq("Marketing"), "actual" if month < pd.Timestamp("2026-07-01") else "budget"].sum())
    new_customers = int(sm.new_customers.sum())
    cac = marketing / new_customers if new_customers else np.nan
    gross_margin = 1 - float(financials.loc[financials.month.eq(month), "forecast_cogs"].iloc[0]) / ending_mrr
    arpc = ending_mrr / sm.ending_customers.sum()
    monthly_churn = sm.churned_customers.sum() / sm.starting_customers.sum()
    ltv = arpc * gross_margin / monthly_churn if monthly_churn else np.nan
    metric_rows.append([month, start_mrr, new_mrr, expansion_mrr, churn_mrr, ending_mrr, grr, nrr, cac, ltv, ltv / cac if cac else np.nan, gross_margin])
    prev_mrr = ending_mrr
saas_metrics = pd.DataFrame(metric_rows, columns=["month", "starting_mrr", "new_mrr", "expansion_mrr", "churned_mrr", "ending_mrr", "grr", "nrr", "cac", "ltv", "ltv_cac_ratio", "gross_margin"])

cohort_rows = []
cohort_months = pd.date_range("2023-01-01", "2026-06-01", freq="MS")
for cohort_month in cohort_months:
    acquired = int(rng.integers(24, 58))
    base_arpc = float(rng.uniform(280, 430))
    churn_rate = float(rng.uniform(.014, .027))
    for observation_month in months[months >= cohort_month]:
        age = (observation_month.year - cohort_month.year) * 12 + observation_month.month - cohort_month.month
        retained = round(acquired * ((1 - churn_rate) ** age))
        cohort_mrr = retained * base_arpc * ((1.006) ** age)
        cohort_rows.append([cohort_month, observation_month, age, acquired, retained, retained / acquired, cohort_mrr])
cohort_analysis = pd.DataFrame(cohort_rows, columns=["cohort_month", "observation_month", "age_months", "customers_acquired", "customers_retained", "logo_retention", "cohort_mrr"])

schedule_rows = []
deferred = 0.0
for _, r in financials.iterrows():
    revenue = r.forecast_revenue
    annual_prepay_share = .34
    new_annual_billings = revenue * annual_prepay_share * 12 / 12
    monthly_billings = revenue * (1 - annual_prepay_share)
    billings = new_annual_billings + monthly_billings + revenue * .025
    deferred_change = billings - revenue
    deferred_end = deferred + deferred_change
    collections = billings * .985
    ar_change = billings - collections
    schedule_rows.append([r.month, revenue, billings, collections, deferred, deferred_change, deferred_end, ar_change])
    deferred = deferred_end
revenue_schedule = pd.DataFrame(schedule_rows, columns=["month", "recognized_revenue", "billings", "cash_collections", "deferred_revenue_start", "change_in_deferred_revenue", "deferred_revenue_end", "change_in_ar"])

statement_rows = []
for _, r in financials.iterrows():
    sched = revenue_schedule[revenue_schedule.month.eq(r.month)].iloc[0]
    revenue = r.actual_revenue if pd.notna(r.actual_revenue) else r.forecast_revenue
    cogs = r.actual_cogs if pd.notna(r.actual_cogs) else r.forecast_cogs
    payroll = r.actual_payroll if pd.notna(r.actual_payroll) else r.forecast_payroll
    other = r.actual_opex if pd.notna(r.actual_opex) else r.forecast_opex
    ebitda = revenue - cogs - payroll - other
    depreciation = 12000 + 0.15 * r.capex
    operating_income = ebitda - depreciation
    cfo = ebitda + sched.change_in_deferred_revenue - sched.change_in_ar
    cfi = -r.capex
    cff = r.financing
    net_cash = cfo + cfi + cff
    statement_rows.append([r.month, revenue, cogs, revenue-cogs, payroll, other, ebitda, depreciation, operating_income, cfo, cfi, cff, net_cash, r.ending_cash, sched.deferred_revenue_end])
statements = pd.DataFrame(statement_rows, columns=["month", "revenue", "cogs", "gross_profit", "payroll", "other_opex", "ebitda", "depreciation", "operating_income", "cash_from_operations", "cash_from_investing", "cash_from_financing", "net_change_in_cash", "ending_cash", "deferred_revenue"])

backtest_rows = []
actual_months = financials[financials.actual_revenue.notna()].reset_index(drop=True)
for i in range(12, len(actual_months)):
    hist = actual_months.iloc[i-3:i]
    avg_growth = hist.actual_revenue.pct_change().dropna().mean()
    predicted = hist.actual_revenue.iloc[-1] * (1 + avg_growth)
    actual = actual_months.actual_revenue.iloc[i]
    backtest_rows.append([actual_months.month.iloc[i], predicted, actual, predicted-actual, abs(predicted-actual)/actual])
forecast_backtest = pd.DataFrame(backtest_rows, columns=["month", "predicted_revenue", "actual_revenue", "forecast_error", "absolute_percentage_error"])

simulations = []
base_cash = summary["ending_cash_jun_2026"]
base_forecast = financials[financials.month.between("2026-07-01", "2026-12-01")]
for simulation_id in range(1, 2001):
    cash_path = base_cash
    growth_shock = rng.normal(0, .012)
    churn_shock = rng.normal(0, .006)
    cost_shock = rng.normal(1, .045)
    for _, r in base_forecast.iterrows():
        revenue_sim = r.forecast_revenue * (1 + growth_shock - churn_shock)
        outflow_sim = (r.forecast_cogs + r.forecast_payroll + r.forecast_opex) * cost_shock + r.capex
        cash_path += revenue_sim - outflow_sim + r.financing
    simulations.append([simulation_id, growth_shock, churn_shock, cost_shock, cash_path, cash_path < 1_000_000, cash_path < 0])
monte_carlo = pd.DataFrame(simulations, columns=["simulation_id", "growth_shock", "churn_shock", "cost_multiplier", "dec_ending_cash", "below_1m", "below_zero"])
mc_summary = pd.DataFrame([
    ["P05 ending cash", monte_carlo.dec_ending_cash.quantile(.05)],
    ["P50 ending cash", monte_carlo.dec_ending_cash.quantile(.50)],
    ["P95 ending cash", monte_carlo.dec_ending_cash.quantile(.95)],
    ["Probability below $1M", monte_carlo.below_1m.mean()],
    ["Probability below zero", monte_carlo.below_zero.mean()],
    ["Revenue forecast MAPE", forecast_backtest.absolute_percentage_error.mean()],
], columns=["metric", "value"])

for frame, name in [(saas_metrics, "saas_metrics"), (cohort_analysis, "cohort_analysis"), (revenue_schedule, "revenue_schedule"), (statements, "integrated_statements"), (forecast_backtest, "forecast_backtest"), (monte_carlo, "monte_carlo_results"), (mc_summary, "model_risk_summary")]:
    frame.to_csv(DATA / f"{name}.csv", index=False, date_format="%Y-%m-%d")
with sqlite3.connect(db_path) as conn:
    for frame, name in [(saas_metrics, "saas_metrics"), (cohort_analysis, "cohort_analysis"), (revenue_schedule, "revenue_schedule"), (statements, "integrated_statements"), (forecast_backtest, "forecast_backtest"), (monte_carlo, "monte_carlo_results"), (mc_summary, "model_risk_summary")]:
        frame.to_sql(name, conn, if_exists="replace", index=False)

print(json.dumps({"rows": {"subscriptions": len(subscriptions), "employees": len(employees), "headcount": len(headcount), "opex": len(opex), "financials": len(financials)}, "summary": summary}, indent=2))
