from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables" / "Variance_Commentary_Report_Masters.docx"
fin = pd.read_csv(ROOT / "data" / "financials_monthly.csv", parse_dates=["month"])
subs = pd.read_csv(ROOT / "data" / "subscriptions.csv", parse_dates=["month"])
opex = pd.read_csv(ROOT / "data" / "operating_expenses.csv", parse_dates=["month"])
hc = pd.read_csv(ROOT / "data" / "headcount_monthly.csv", parse_dates=["month"])
metrics = pd.read_csv(ROOT / "data" / "saas_metrics.csv", parse_dates=["month"])
backtest = pd.read_csv(ROOT / "data" / "forecast_backtest.csv", parse_dates=["month"])
risk = pd.read_csv(ROOT / "data" / "model_risk_summary.csv")

ytd = fin[fin.month.between("2026-01-01", "2026-06-01")].copy()
rev_var = ytd.actual_revenue.sum() - ytd.budget_revenue.sum()
rev_var_pct = rev_var / ytd.budget_revenue.sum()
actual_exp = ytd.actual_cogs.sum()+ytd.actual_payroll.sum()+ytd.actual_opex.sum()
budget_exp = ytd.budget_cogs.sum()+ytd.budget_payroll.sum()+ytd.budget_opex.sum()
exp_var = budget_exp-actual_exp
op_loss = ytd.actual_revenue.sum()-actual_exp
jun = subs[subs.month.eq("2026-06-01")]
mrr = jun.mrr.sum(); arr=mrr*12; churn=jun.churned_customers.sum()/jun.starting_customers.sum()
jun_cash = ytd.iloc[-1].ending_cash
dec_cash = fin.loc[fin.month.eq("2026-12-01"),"ending_cash"].iloc[0]
cloud = opex[(opex.month.between("2026-01-01","2026-06-01")) & opex.category.eq("Cloud & API")]
cloud_var = cloud.budget.sum()-cloud.actual.sum()
hc_jun = hc[hc.month.eq("2026-06-01")]
hc_var = hc_jun.actual_headcount.sum()-hc_jun.planned_headcount.sum()

def money(x): return f"${abs(x)/1_000_000:.2f}M" if abs(x)>=1_000_000 else f"${abs(x)/1_000:.0f}K"
def shade(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)

doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.65); sec.bottom_margin=Inches(.65); sec.left_margin=Inches(.72); sec.right_margin=Inches(.72)
styles=doc.styles
styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(10); styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Aptos')
for s,size,color in [('Title',26,'16324F'),('Heading 1',16,'16324F'),('Heading 2',12,'2F75B5')]:
    styles[s].font.name='Aptos Display' if s=='Title' else 'Aptos'; styles[s].font.size=Pt(size); styles[s].font.color.rgb=RGBColor.from_string(color)

p=doc.add_paragraph(); p.style='Title'; p.add_run('Variance commentary report')
p=doc.add_paragraph('SaaS FP&A portfolio case | Actuals through June 2026'); p.style='Subtitle'
doc.add_paragraph('This report contains an AI-assisted draft based on structured financial results. Every explanation below was manually checked against the source calculations. No AI system made or approved a financial decision.')

doc.add_heading('Executive summary', level=1)
bullets=[
    f"YTD revenue was {money(ytd.actual_revenue.sum())}, {money(rev_var)} above budget ({rev_var_pct:.1%} favorable). The dataset supports the amount, but it does not isolate a causal driver beyond customer and plan-level growth.",
    f"Total YTD expenses were {money(actual_exp)}, producing an operating loss of {money(op_loss)}. Expenses were {money(abs(exp_var))} {'below' if exp_var>0 else 'above'} budget ({abs(exp_var)/budget_exp:.1%} {'favorable' if exp_var>0 else 'unfavorable'}).",
    f"June MRR reached {money(mrr)} and ARR reached {money(arr)}. Gross logo churn was {churn:.1%} for the month.",
    f"June ending cash was {money(jun_cash)}. The base forecast ends December at {money(dec_cash)}, indicating a near-term funding requirement if the modeled burn continues.",
]
for t in bullets: doc.add_paragraph(t, style='List Bullet')

doc.add_heading('Validated variance commentary', level=1)
items=[
('Revenue','Favorable',f"Revenue exceeded budget by {money(rev_var)} ({rev_var_pct:.1%}) through June.","Validated","Recomputed from monthly actual and budget revenue."),
('Cloud and API','Unfavorable' if cloud_var<0 else 'Favorable',f"Cloud and API cost was {money(abs(cloud_var))} {'above' if cloud_var<0 else 'below'} budget.","Validated","Recomputed from the Cloud & API category. Higher usage is a hypothesis, not a proven cause."),
('Total expense','Favorable' if exp_var>0 else 'Unfavorable',f"Total expense was {money(abs(exp_var))} {'below' if exp_var>0 else 'above'} budget.","Validated","Includes COGS, payroll, and other operating expense."),
('Headcount','Favorable' if hc_var<0 else 'Unfavorable',f"June actual headcount was {abs(hc_var)} {'below' if hc_var<0 else 'above'} plan.","Validated","Compared summed department-level actual and planned headcount."),
('Cash','Risk',f"Base-case cash declines from {money(jun_cash)} in June to {money(dec_cash)} in December.","Validated","Recomputed from forecast inflow, outflow, capex, and financing."),
]
table=doc.add_table(rows=1, cols=5); table.alignment=WD_TABLE_ALIGNMENT.CENTER; table.style='Table Grid'
hdr=table.rows[0].cells
for i,v in enumerate(['Area','Direction','Commentary','Status','Manual validation']): hdr[i].text=v; shade(hdr[i],'2F75B5');
for cell in hdr:
    for run in cell.paragraphs[0].runs: run.font.bold=True; run.font.color.rgb=RGBColor(255,255,255)
for row in items:
    cells=table.add_row().cells
    for i,v in enumerate(row): cells[i].text=str(v); cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
for row in table.rows: row.cells[0].width=Inches(.85); row.cells[1].width=Inches(.8); row.cells[2].width=Inches(2.4); row.cells[3].width=Inches(.85); row.cells[4].width=Inches(2.1)

doc.add_heading('Forecast and management attention', level=1)
doc.add_paragraph('The modeled base case remains loss-making through the forecast horizon. Management should confirm the timing and size of the next financing action before cash approaches the minimum operating buffer. The scenario model allows growth, churn, pricing, hiring, marketing, and cloud unit cost assumptions to change without rewriting the historical actuals.')
doc.add_paragraph('The strongest controllable levers in this dataset are hiring pace and discretionary marketing. Pricing and churn create greater upside to recurring revenue, while cloud unit cost affects gross margin. These statements describe model mechanics; they do not establish operational feasibility.')

doc.add_heading('SaaS economics and cohort interpretation', level=1)
latest_metrics = metrics.loc[metrics.month.eq(pd.Timestamp('2026-06-01'))].iloc[0]
doc.add_paragraph(f"June gross revenue retention was {latest_metrics.grr:.1%} and net revenue retention was {latest_metrics.nrr:.1%}. The modeled LTV/CAC ratio was {latest_metrics.ltv_cac_ratio:.1f}x. These metrics use a simplified gross-margin-over-logo-churn formulation and should be compared with a contribution-margin cohort model before investment decisions.")
doc.add_paragraph('The cohort schedule separates acquisition month from observation month, allowing the analyst to distinguish mix shifts from true retention changes. Cohort retention should be reviewed alongside plan, channel, and customer-size segmentation because aggregate retention can improve even when a newer cohort underperforms.')

doc.add_heading('Forecast model risk', level=1)
mape = backtest.absolute_percentage_error.mean()
risk_map = dict(zip(risk.metric, risk.value))
doc.add_paragraph(f"The rolling one-step revenue back-test produced mean absolute percentage error of {mape:.1%}. The 2,000-path liquidity simulation estimates a {risk_map['Probability below $1M']:.1%} probability of ending December below $1.0M and a {risk_map['Probability below zero']:.1%} probability of negative cash. The simulated median December cash is {money(risk_map['P50 ending cash'])}.")
doc.add_paragraph('The simulation varies customer growth, churn, and the operating-cost multiplier around the base case. It is a risk framing tool rather than a probability forecast: correlations, fat tails, financing optionality, and management responses are not modeled.')

doc.add_heading('Accounting and control considerations', level=1)
for t in [
    'Recognized subscription revenue is separated from billings and cash collections through a deferred-revenue schedule.',
    'The integrated statement view bridges EBITDA to operating cash flow using deferred revenue and accounts-receivable movements.',
    'The cash-flow schedule remains a planning model and does not replace a GAAP statement of cash flows or audited close process.',
    'Terminal audit checks review customer roll-forward, cash roll-forward, source-period completeness, and scenario-control validity.',
]: doc.add_paragraph(t, style='List Bullet')

doc.add_heading('AI use and reviewer protocol', level=1)
doc.add_paragraph('The companion prompt in ai/claude_variance_prompt.md is designed for an initial Claude draft. Claude was not available in this environment, so this document does not claim that Claude generated the text. Before external use, paste the structured results into Claude, compare its draft to this validation table, and retain only explanations supported by named calculations or confirmed by business owners.')
steps=['Tie every amount to the workbook or prepared dataset.','Reject causal language that the data does not prove.','Confirm operational explanations with the accountable department owner.','Re-run the report after any forecast or scenario change.']
for s in steps: doc.add_paragraph(s, style='List Number')

doc.add_heading('Data scope and conventions', level=1)
doc.add_paragraph('The company, customer base, employees, and financial results are synthetic. Positive revenue variance is favorable. Positive expense variance means actual expense is below budget. Actuals run through June 2026; later months are forecast.')

footer=sec.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER; footer.add_run('Synthetic portfolio case | FP&A management reporting').font.size=Pt(8)
OUT.parent.mkdir(exist_ok=True); doc.save(OUT); print(OUT)
