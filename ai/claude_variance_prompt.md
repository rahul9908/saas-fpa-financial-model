# Claude-ready variance commentary prompt

You are drafting management commentary for a fictional SaaS company. Use only the structured values supplied below. Do not infer causes that the data does not support. Separate facts from hypotheses, label favorable and unfavorable variances, quantify every material statement, and state that management must validate operational explanations.

Required sections:

1. Revenue and customer performance
2. Gross margin and cloud/API cost
3. Payroll, headcount, and operating expenses
4. Cash burn, runway, and funding risk
5. Forecast and scenario implications

For every explanation, include a validation tag: `Validated`, `Needs owner confirmation`, or `Data does not support a cause`.

Input data should come from the workbook's BvA, Revenue Model, Expense Model, Headcount, Cash Flow, and Scenarios sheets. Return a concise executive summary followed by a validation table listing claim, calculation, source cells, status, and reviewer note.

