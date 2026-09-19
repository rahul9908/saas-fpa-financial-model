# Power BI FP&A Dashboard Build Guide

## Import and model

Import the CSV files from `../data/`. Create a calendar table from 2023-01-01 through 2026-12-31 and mark it as the date table.

Relationships:

- `Calendar[Date]` 1:* `financials_monthly[month]`
- `Calendar[Date]` 1:* `subscriptions[month]`
- `Calendar[Date]` 1:* `headcount_monthly[month]`
- `Calendar[Date]` 1:* `operating_expenses[month]`
- `pricing_plans[plan]` 1:* `subscriptions[plan]`
- `Calendar[Date]` 1:* `saas_metrics[month]`
- `Calendar[Date]` 1:* `revenue_schedule[month]`
- `Calendar[Date]` 1:* `integrated_statements[month]`
- Department and expense-category dimensions may be created from distinct values for cleaner filtering.

Use single-direction filtering from dimensions to facts. Hide surrogate technical fields and raw additive headcount fields where they could be misused across departments or months.

## Dashboard pages

1. **Executive Financial Summary**: KPI cards for MRR, ARR, revenue variance, gross margin, operating loss, ending cash, and runway; monthly revenue and cash trend; actual/forecast flag.
2. **Revenue and Customer Metrics**: MRR by plan, customer bridge, new versus churned customers, gross logo churn, ARPC, and CAC proxy.
3. **Budget versus Actuals**: monthly revenue and expense variance waterfall, category matrix, favorable/unfavorable conditional formatting, and quarterly drill-down.
4. **Expense and Headcount Analysis**: expense mix, fixed versus variable cost, department headcount variance, personnel cost trend, and planned versus actual hires.
5. **Cash Burn and Runway**: starting and ending cash, net cash flow, monthly burn, rolling burn, financing events, and runway.
6. **Forecast and Scenarios**: active case, six-month outlook, key driver table, scenario comparison, and growth/churn sensitivity heat map.

7. **SaaS Unit Economics**: GRR, NRR, CAC, LTV, LTV/CAC, and ARR bridge.
8. **Cohorts and Forecast QA**: cohort-retention curve, forecast MAPE/bias, and model-risk notes.
9. **Liquidity Simulation**: P05/P50/P95 cash, probability below $1M, and probability below zero.

## Visual conventions

- Actuals: navy; forecast: teal; budget: gray.
- Favorable variance: green; unfavorable variance: red.
- Display currency in USD thousands or millions and state the unit in each title.
- Add an as-of date and synthetic-data disclosure to every page footer.

Native `.pbix` generation is outside this package; the included fact tables, schema, measures, and page specification are the build-ready Power BI deliverable.
