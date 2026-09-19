-- SaaS FP&A portfolio model: reusable management-reporting queries

-- Monthly revenue and customer metrics by plan
SELECT month, plan,
       SUM(ending_customers) AS ending_customers,
       SUM(new_customers) AS new_customers,
       SUM(churned_customers) AS churned_customers,
       SUM(mrr) AS mrr,
       SUM(mrr) * 12 AS arr,
       CASE WHEN SUM(starting_customers) = 0 THEN NULL
            ELSE 1.0 * SUM(churned_customers) / SUM(starting_customers) END AS gross_logo_churn
FROM subscriptions
GROUP BY month, plan
ORDER BY month, plan;

-- Budget versus actual with favorable/unfavorable convention
SELECT month,
       budget_revenue, actual_revenue,
       actual_revenue - budget_revenue AS revenue_variance,
       budget_payroll + budget_opex + budget_cogs AS budget_expense,
       actual_payroll + actual_opex + actual_cogs AS actual_expense,
       (budget_payroll + budget_opex + budget_cogs)
       - (actual_payroll + actual_opex + actual_cogs) AS expense_variance
FROM financials_monthly
WHERE actual_revenue IS NOT NULL
ORDER BY month;

-- Headcount variance by department
SELECT month, department, planned_headcount, actual_headcount,
       actual_headcount - planned_headcount AS headcount_variance,
       personnel_cost
FROM headcount_monthly
ORDER BY month, department;

-- Rolling three-month burn rate and runway
WITH cashflow AS (
  SELECT month, ending_cash,
         forecast_revenue - forecast_cogs - forecast_payroll - forecast_opex - capex AS net_cash_flow
  FROM financials_monthly
), burn AS (
  SELECT month, ending_cash, net_cash_flow,
         AVG(CASE WHEN net_cash_flow < 0 THEN -net_cash_flow ELSE 0 END)
         OVER (ORDER BY month ROWS BETWEEN 2 PRECEDING AND CURRENT ROW) AS avg_3m_burn
  FROM cashflow
)
SELECT month, ending_cash, net_cash_flow, avg_3m_burn,
       CASE WHEN avg_3m_burn = 0 THEN NULL ELSE ending_cash / avg_3m_burn END AS runway_months
FROM burn
ORDER BY month;
