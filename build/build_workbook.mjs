import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const root = path.resolve(".");
const dataDir = path.join(root, "data");
const outDir = path.join(root, "deliverables");
await fs.mkdir(outDir, { recursive: true });

function parseCsv(text) {
  const lines = text.trim().split(/\r?\n/);
  const parseLine = (line) => {
    const out=[]; let cur="", q=false;
    for(let i=0;i<line.length;i++){const c=line[i]; if(c==='"'){if(q&&line[i+1]==='"'){cur+='"';i++;}else q=!q;} else if(c===','&&!q){out.push(cur);cur="";}else cur+=c;} out.push(cur); return out;
  };
  const headers=parseLine(lines[0]);
  return [headers, ...lines.slice(1).map(line=>parseLine(line).map((v,i)=>{
    if(v==="") return null;
    if(headers[i].includes("date") || headers[i]==="month") return new Date(v+(/T/.test(v)?"":"T00:00:00"));
    if(["plan","department","category","cost_type","employee_id","scenario"].includes(headers[i])) return v;
    if(v==="True"||v==="true") return true; if(v==="False"||v==="false") return false;
    const n=Number(v); return Number.isFinite(n)?n:v;
  }))];
}
const csv = async name => parseCsv(await fs.readFile(path.join(dataDir, name+".csv"), "utf8"));
const subs = await csv("subscriptions");
const employees = await csv("employees");
const hc = await csv("headcount_monthly");
const opex = await csv("operating_expenses");
const fin = await csv("financials_monthly");
const scenarios = await csv("scenario_assumptions");
const metrics = await csv("saas_metrics");
const cohorts = await csv("cohort_analysis");
const statements = await csv("integrated_statements");
const backtest = await csv("forecast_backtest");
const monteCarlo = await csv("monte_carlo_results");
const riskSummary = await csv("model_risk_summary");

const wb=Workbook.create();
const names=["Executive Summary","Assumptions","SaaS Metrics","ARR Bridge","Cohort Analytics","Revenue Model","Expense Model","Headcount","BvA","Cash Flow","Statements","Scenarios","Sensitivity","Forecast QA","Monte Carlo","Actuals Data","Subscriptions","Employees","Opex Data","Audit","ReadMe"];
const ws={}; for(const n of names) ws[n]=wb.worksheets.add(n);

const navy="#16324F", teal="#00A6A6", blue="#2F75B5", pale="#EAF3F8", green="#D9EAD3", red="#F4CCCC", amber="#FFF2CC", gray="#667085", white="#FFFFFF";
function title(sh,text,sub="") { sh.showGridLines=false; sh.getRange("A1:L1").merge(); sh.getRange("A1").values=[[text]]; sh.getRange("A1:L1").format={fill:navy,font:{bold:true,color:white,size:20},rowHeight:34,verticalAlignment:"center"}; if(sub){sh.getRange("A2:L2").merge();sh.getRange("A2").values=[[sub]];sh.getRange("A2:L2").format={font:{italic:true,color:gray,size:10},rowHeight:22};} }
function header(r){r.format={fill:blue,font:{bold:true,color:white},borders:{preset:"all",style:"thin",color:white},verticalAlignment:"center"};}
function input(r){r.format={fill:amber,font:{color:"#7F6000",bold:true},borders:{preset:"all",style:"thin",color:"#D6B656"}};}
function money(r){r.setNumberFormat('$#,##0;[Red]($#,##0);-');}
function pct(r){r.setNumberFormat('0.0%;[Red](0.0%);-');}
function writeRaw(sh, rows, tableName){sh.getRangeByIndexes(0,0,rows.length,rows[0].length).values=rows;header(sh.getRangeByIndexes(0,0,1,rows[0].length));sh.freezePanes.freezeRows(1);sh.getUsedRange().format.autofitColumns();sh.getRange("A:A").setNumberFormat("yyyy-mm-dd");const t=sh.tables.add(sh.getUsedRange().address,true,tableName);t.style="TableStyleMedium2";}
writeRaw(ws["Actuals Data"],fin,"FinancialsTable");
writeRaw(ws["Subscriptions"],subs,"SubscriptionsTable");
writeRaw(ws["Employees"],employees,"EmployeesTable");
writeRaw(ws["Opex Data"],opex,"OpexTable");

// Assumptions and active case.
title(ws.Assumptions,"Assumptions and scenario controls","Yellow cells are editable. Historical actuals do not change with the selector.");
ws.Assumptions.getRange("A4:B4").values=[["Active scenario","Base"]]; header(ws.Assumptions.getRange("A4")); input(ws.Assumptions.getRange("B4"));
ws.Assumptions.getRange("B4").dataValidation={rule:{type:"list",values:["Base","Upside","Downside"]}};
ws.Assumptions.getRange("A7:G11").values=[["Scenario","Monthly growth","Monthly churn","Price multiplier","Hiring multiplier","Marketing multiplier","Cloud cost multiplier"],...scenarios.slice(1)];
header(ws.Assumptions.getRange("A7:G7"));input(ws.Assumptions.getRange("B8:G10"));pct(ws.Assumptions.getRange("B8:C10"));ws.Assumptions.getRange("D8:G10").setNumberFormat("0.00x");
ws.Assumptions.getRange("A13:B19").values=[["Active driver","Value"],["Monthly growth",null],["Monthly churn",null],["Price multiplier",null],["Hiring multiplier",null],["Marketing multiplier",null],["Cloud cost multiplier",null]];
header(ws.Assumptions.getRange("A13:B13"));
for(let r=14;r<=19;r++) ws.Assumptions.getRange(`B${r}`).formulas=[[`=_xlfn.XLOOKUP($B$4,$A$8:$A$10,${String.fromCharCode(66+r-14)}$8:${String.fromCharCode(66+r-14)}$10)`]];
pct(ws.Assumptions.getRange("B14:B15")); ws.Assumptions.getRange("B16:B19").setNumberFormat("0.00x");ws.Assumptions.getRange("A:G").format.columnWidth=20;

// Revenue model, formula-driven from raw subscription table.
title(ws["Revenue Model"],"Revenue model","Monthly customer bridge, MRR, ARR, churn, growth, and ARPC");
const revHeaders=[["Month","Starting customers","New customers","Churned customers","Ending customers","MRR","ARR","Revenue growth","Gross churn","ARPC","Period"]];
ws["Revenue Model"].getRange("A4:K4").values=revHeaders;header(ws["Revenue Model"].getRange("A4:K4"));
const monthList=fin.slice(1).map(r=>[r[0]]);ws["Revenue Model"].getRange(`A5:A${4+monthList.length}`).values=monthList;ws["Revenue Model"].getRange(`A5:A${4+monthList.length}`).setNumberFormat("mmm-yy");
for(let r=5;r<=52;r++){
  ws["Revenue Model"].getRange(`B${r}`).formulas=[[`=SUMIFS(Subscriptions!$C$2:$C$145,Subscriptions!$A$2:$A$145,$A${r})`]];
  ws["Revenue Model"].getRange(`C${r}`).formulas=[[`=SUMIFS(Subscriptions!$D$2:$D$145,Subscriptions!$A$2:$A$145,$A${r})`]];
  ws["Revenue Model"].getRange(`D${r}`).formulas=[[`=SUMIFS(Subscriptions!$E$2:$E$145,Subscriptions!$A$2:$A$145,$A${r})`]];
  ws["Revenue Model"].getRange(`E${r}`).formulas=[[`=SUMIFS(Subscriptions!$F$2:$F$145,Subscriptions!$A$2:$A$145,$A${r})`]];
  ws["Revenue Model"].getRange(`F${r}`).formulas=[[`=SUMIFS(Subscriptions!$H$2:$H$145,Subscriptions!$A$2:$A$145,$A${r})`]];
  ws["Revenue Model"].getRange(`G${r}`).formulas=[[`=F${r}*12`]];
  ws["Revenue Model"].getRange(`H${r}`).formulas=[[r===5?"=0":`=IFERROR(F${r}/F${r-1}-1,0)`]];
  ws["Revenue Model"].getRange(`I${r}`).formulas=[[`=IFERROR(D${r}/B${r},0)`]];
  ws["Revenue Model"].getRange(`J${r}`).formulas=[[`=IFERROR(F${r}/E${r},0)`]];
  ws["Revenue Model"].getRange(`K${r}`).formulas=[[`=IF(A${r}<DATE(2026,7,1),"Actual","Forecast")`]];
}
money(ws["Revenue Model"].getRange("F5:G52"));money(ws["Revenue Model"].getRange("J5:J52"));pct(ws["Revenue Model"].getRange("H5:I52"));ws["Revenue Model"].freezePanes.freezeRows(4);ws["Revenue Model"].getRange("A:K").format.columnWidth=16;

// Graduate-level SaaS operating metrics and ARR waterfall.
title(ws["SaaS Metrics"],"SaaS unit economics","GRR, NRR, CAC, LTV, LTV/CAC, and gross margin by month");
ws["SaaS Metrics"].getRange("A4:L4").values=[metrics[0]];header(ws["SaaS Metrics"].getRange("A4:L4"));ws["SaaS Metrics"].getRangeByIndexes(4,0,metrics.length-1,metrics[0].length).values=metrics.slice(1);
ws["SaaS Metrics"].getRange("A5:A52").setNumberFormat("mmm-yy");money(ws["SaaS Metrics"].getRange("B5:F52"));pct(ws["SaaS Metrics"].getRange("G5:H52"));money(ws["SaaS Metrics"].getRange("I5:J52"));ws["SaaS Metrics"].getRange("K5:K52").setNumberFormat("0.0x");pct(ws["SaaS Metrics"].getRange("L5:L52"));ws["SaaS Metrics"].freezePanes.freezeRows(4);ws["SaaS Metrics"].getRange("A:L").format.columnWidth=16;

title(ws["ARR Bridge"],"ARR movement bridge","Starting ARR plus new and expansion ARR less churned ARR equals ending ARR");
ws["ARR Bridge"].getRange("A4:G4").values=[["Month","Starting ARR","New ARR","Expansion ARR","Churned ARR","Ending ARR","Bridge check"]];header(ws["ARR Bridge"].getRange("A4:G4"));
for(let i=0;i<48;i++){const r=5+i;ws["ARR Bridge"].getRange(`A${r}`).formulas=[[`='SaaS Metrics'!A${r}`]];ws["ARR Bridge"].getRange(`B${r}`).formulas=[[`='SaaS Metrics'!B${r}*12`]];ws["ARR Bridge"].getRange(`C${r}`).formulas=[[`='SaaS Metrics'!C${r}*12`]];ws["ARR Bridge"].getRange(`D${r}`).formulas=[[`='SaaS Metrics'!D${r}*12`]];ws["ARR Bridge"].getRange(`E${r}`).formulas=[[`='SaaS Metrics'!E${r}*12`]];ws["ARR Bridge"].getRange(`F${r}`).formulas=[[`='SaaS Metrics'!F${r}*12`]];ws["ARR Bridge"].getRange(`G${r}`).formulas=[[`=B${r}+C${r}+D${r}-E${r}-F${r}`]];}
ws["ARR Bridge"].getRange("A5:A52").setNumberFormat("mmm-yy");money(ws["ARR Bridge"].getRange("B5:G52"));ws["ARR Bridge"].getRange("G5:G52").conditionalFormats.add("cellIs",{operator:"notEqual",formula:0,format:{fill:red}});ws["ARR Bridge"].getRange("A:G").format.columnWidth=18;

title(ws["Cohort Analytics"],"Customer cohort retention","Average logo retention by months since acquisition");
ws["Cohort Analytics"].getRange("A4:C4").values=[["Age (months)","Average logo retention","Cohorts observed"]];header(ws["Cohort Analytics"].getRange("A4:C4"));
for(let age=0;age<=24;age++){const rows=cohorts.slice(1).filter(x=>x[2]===age);const r=5+age;ws["Cohort Analytics"].getRange(`A${r}:C${r}`).values=[[age,rows.reduce((s,x)=>s+x[5],0)/rows.length,rows.length]];}
pct(ws["Cohort Analytics"].getRange("B5:B29"));ws["Cohort Analytics"].getRange("A:C").format.columnWidth=24;

// Expense model by category with budget/actual/variance.
title(ws["Expense Model"],"Expense model","Operating expenses by category and cost behavior");
ws["Expense Model"].getRange("A4:F4").values=[["Month","Category","Cost type","Budget","Actual / Forecast","Variance (F)" ]];header(ws["Expense Model"].getRange("A4:F4"));
const ox=opex.slice(1).map(r=>[r[0],r[1],r[4],r[2],r[3]??r[2],null]);ws["Expense Model"].getRangeByIndexes(4,0,ox.length,6).values=ox;
for(let r=5;r<5+ox.length;r++) ws["Expense Model"].getRange(`F${r}`).formulas=[[`=D${r}-E${r}`]];
ws["Expense Model"].getRange(`A5:A${4+ox.length}`).setNumberFormat("mmm-yy");money(ws["Expense Model"].getRange(`D5:F${4+ox.length}`));ws["Expense Model"].getRange(`F5:F${4+ox.length}`).conditionalFormats.add("cellIs",{operator:"lessThan",formula:0,format:{fill:red,font:{color:"#9C0006"}}});ws["Expense Model"].freezePanes.freezeRows(4);ws["Expense Model"].getRange("A:F").format.columnWidth=21;

// Headcount model.
title(ws.Headcount,"Headcount model","Planned versus actual staffing and monthly personnel cost");
ws.Headcount.getRange("A4:H4").values=[["Month","Department","Planned HC","Actual HC","Reported HC","Salary","Benefits","Personnel cost"]];header(ws.Headcount.getRange("A4:H4"));
ws.Headcount.getRangeByIndexes(4,0,hc.length-1,8).values=hc.slice(1);ws.Headcount.getRange(`A5:A${3+hc.length}`).setNumberFormat("mmm-yy");money(ws.Headcount.getRange(`F5:H${3+hc.length}`));ws.Headcount.freezePanes.freezeRows(4);ws.Headcount.getRange("A:H").format.columnWidth=19;

// BvA monthly P&L.
title(ws.BvA,"Budget versus actuals","Favorable variance is positive for revenue and when actual expense is below budget");
ws.BvA.getRange("A4:N4").values=[["Month","Budget revenue","Actual revenue","Revenue var","Budget COGS","Actual COGS","COGS var","Budget payroll","Actual payroll","Payroll var","Budget other opex","Actual other opex","Opex var","Operating variance"]];header(ws.BvA.getRange("A4:N4"));
for(let i=0;i<48;i++){const r=5+i, src=2+i;ws.BvA.getRange(`A${r}`).formulas=[[`='Actuals Data'!A${src}`]];ws.BvA.getRange(`B${r}`).formulas=[[`='Actuals Data'!B${src}`]];ws.BvA.getRange(`C${r}`).formulas=[[`='Actuals Data'!F${src}`]];ws.BvA.getRange(`D${r}`).formulas=[[`=C${r}-B${r}`]];ws.BvA.getRange(`E${r}`).formulas=[[`='Actuals Data'!C${src}`]];ws.BvA.getRange(`F${r}`).formulas=[[`='Actuals Data'!G${src}`]];ws.BvA.getRange(`G${r}`).formulas=[[`=E${r}-F${r}`]];ws.BvA.getRange(`H${r}`).formulas=[[`='Actuals Data'!D${src}`]];ws.BvA.getRange(`I${r}`).formulas=[[`='Actuals Data'!H${src}`]];ws.BvA.getRange(`J${r}`).formulas=[[`=H${r}-I${r}`]];ws.BvA.getRange(`K${r}`).formulas=[[`='Actuals Data'!E${src}`]];ws.BvA.getRange(`L${r}`).formulas=[[`='Actuals Data'!I${src}`]];ws.BvA.getRange(`M${r}`).formulas=[[`=K${r}-L${r}`]];ws.BvA.getRange(`N${r}`).formulas=[[`=D${r}+G${r}+J${r}+M${r}`]];}
ws.BvA.getRange("A5:A52").setNumberFormat("mmm-yy");money(ws.BvA.getRange("B5:N52"));ws.BvA.getRange("D5:N52").conditionalFormats.add("cellIs",{operator:"lessThan",formula:0,format:{fill:red,font:{color:"#9C0006"}}});ws.BvA.freezePanes.freezeRows(4);ws.BvA.getRange("A:N").format.columnWidth=16;

// Cash flow.
title(ws["Cash Flow"],"Cash flow forecast","Indirect operating cash proxy, capex, financing, burn, and runway");
ws["Cash Flow"].getRange("A4:J4").values=[["Month","Starting cash","Cash inflow","Operating outflow","Capex","Financing","Net cash flow","Ending cash","Burn rate","Runway months"]];header(ws["Cash Flow"].getRange("A4:J4"));
for(let i=0;i<48;i++){const r=5+i,s=2+i;ws["Cash Flow"].getRange(`A${r}`).formulas=[[`='Actuals Data'!A${s}`]];ws["Cash Flow"].getRange(`B${r}`).formulas=[[`='Actuals Data'!N${s}`]];ws["Cash Flow"].getRange(`C${r}`).formulas=[[`='Actuals Data'!J${s}`]];ws["Cash Flow"].getRange(`D${r}`).formulas=[[`=SUM('Actuals Data'!K${s}:M${s})`]];ws["Cash Flow"].getRange(`E${r}`).formulas=[[`='Actuals Data'!O${s}`]];ws["Cash Flow"].getRange(`F${r}`).formulas=[[`='Actuals Data'!P${s}`]];ws["Cash Flow"].getRange(`G${r}`).formulas=[[`=C${r}-D${r}-E${r}+F${r}`]];ws["Cash Flow"].getRange(`H${r}`).formulas=[[`=B${r}+G${r}`]];ws["Cash Flow"].getRange(`I${r}`).formulas=[[`=MAX(0,-G${r})`]];ws["Cash Flow"].getRange(`J${r}`).formulas=[[`=IFERROR(H${r}/AVERAGE(I${Math.max(5,r-2)}:I${r}),0)`]];}
ws["Cash Flow"].getRange("A5:A52").setNumberFormat("mmm-yy");money(ws["Cash Flow"].getRange("B5:I52"));ws["Cash Flow"].getRange("J5:J52").setNumberFormat("0.0");ws["Cash Flow"].freezePanes.freezeRows(4);ws["Cash Flow"].getRange("A:J").format.columnWidth=17;

title(ws.Statements,"Integrated statements","Accrual P&L and cash-flow bridge with deferred revenue");
ws.Statements.getRangeByIndexes(3,0,statements.length,statements[0].length).values=statements;header(ws.Statements.getRangeByIndexes(3,0,1,statements[0].length));ws.Statements.getRange(`A5:A${3+statements.length}`).setNumberFormat("mmm-yy");money(ws.Statements.getRange(`B5:O${3+statements.length}`));ws.Statements.freezePanes.freezeRows(4);ws.Statements.getRange("A:O").format.columnWidth=17;

// Scenario outputs for six-month forecast, calculated from drivers.
title(ws.Scenarios,"Scenario comparison","Six-month outlook from July to December 2026");
ws.Scenarios.getRange("A4:H4").values=[["Scenario","Growth","Churn","Price","Hiring","Marketing","Cloud cost","Dec-26 ending cash"]];header(ws.Scenarios.getRange("A4:H4"));
ws.Scenarios.getRange("A5:G7").values=scenarios.slice(1);
for(let r=5;r<=7;r++) ws.Scenarios.getRange(`H${r}`).formulas=[[`='Cash Flow'!H46 + SUM('Cash Flow'!G47:G52) + ('Revenue Model'!F46*6*(B${r}-C${r})*D${r}) - (SUM('Headcount'!H215:H240)*(E${r}-1)) - (SUMIFS('Expense Model'!E:E,'Expense Model'!B:B,"Marketing")*(F${r}-1)) - (SUMIFS('Expense Model'!E:E,'Expense Model'!B:B,"Cloud & API")*(G${r}-1))`]];
pct(ws.Scenarios.getRange("B5:C7"));ws.Scenarios.getRange("D5:G7").setNumberFormat("0.00x");money(ws.Scenarios.getRange("H5:H7"));ws.Scenarios.getRange("A:H").format.columnWidth=19;

// Sensitivity grid: growth versus churn impact on projected MRR.
title(ws.Sensitivity,"MRR sensitivity","December 2026 MRR response to monthly growth and churn assumptions");
ws.Sensitivity.getRange("A5:F5").values=[["Growth / churn",0.010,0.015,0.020,0.025,0.030]];header(ws.Sensitivity.getRange("A5:F5"));ws.Sensitivity.getRange("A6:A10").values=[[.02],[.03],[.04],[.05],[.06]];header(ws.Sensitivity.getRange("A6:A10"));
for(let r=6;r<=10;r++)for(let c=2;c<=6;c++){const col=String.fromCharCode(64+c);ws.Sensitivity.getRange(`${col}${r}`).formulas=[[`='Revenue Model'!F46*(1+$A${r}-${col}$5)^6`]];}
pct(ws.Sensitivity.getRange("A6:A10"));pct(ws.Sensitivity.getRange("B5:F5"));money(ws.Sensitivity.getRange("B6:F10"));ws.Sensitivity.getRange("B6:F10").conditionalFormats.add("colorScale",{colors:["#F8696B","#FFEB84","#63BE7B"],thresholds:["min",{type:"percentile",value:50},"max"]});ws.Sensitivity.getRange("A:F").format.columnWidth=18;

title(ws["Forecast QA"],"Forecast validation","Rolling one-step revenue back-test and error diagnostics");
ws["Forecast QA"].getRangeByIndexes(3,0,backtest.length,backtest[0].length).values=backtest;header(ws["Forecast QA"].getRangeByIndexes(3,0,1,backtest[0].length));ws["Forecast QA"].getRange(`A5:A${3+backtest.length}`).setNumberFormat("mmm-yy");money(ws["Forecast QA"].getRange(`B5:D${3+backtest.length}`));pct(ws["Forecast QA"].getRange(`E5:E${3+backtest.length}`));ws["Forecast QA"].getRange("G4:H6").values=[["Diagnostic","Result"],["Mean absolute percentage error",null],["Bias",null]];header(ws["Forecast QA"].getRange("G4:H4"));ws["Forecast QA"].getRange("H5").formulas=[[`=AVERAGE(E5:E${3+backtest.length})`]];ws["Forecast QA"].getRange("H6").formulas=[[`=AVERAGE(D5:D${3+backtest.length})`]];pct(ws["Forecast QA"].getRange("H5"));money(ws["Forecast QA"].getRange("H6"));ws["Forecast QA"].getRange("A:H").format.columnWidth=22;

title(ws["Monte Carlo"],"Liquidity risk simulation","2,000 stochastic paths for growth, churn, and operating cost through December 2026");
ws["Monte Carlo"].getRange("A4:B10").values=riskSummary;header(ws["Monte Carlo"].getRange("A4:B4"));money(ws["Monte Carlo"].getRange("B5:B7"));pct(ws["Monte Carlo"].getRange("B8:B10"));ws["Monte Carlo"].getRange("D4:J4").values=[monteCarlo[0]];header(ws["Monte Carlo"].getRange("D4:J4"));ws["Monte Carlo"].getRangeByIndexes(4,3,monteCarlo.length-1,monteCarlo[0].length).values=monteCarlo.slice(1);pct(ws["Monte Carlo"].getRange(`E5:F${3+monteCarlo.length}`));ws["Monte Carlo"].getRange(`G5:G${3+monteCarlo.length}`).setNumberFormat("0.00x");money(ws["Monte Carlo"].getRange(`H5:H${3+monteCarlo.length}`));ws["Monte Carlo"].freezePanes.freezeRows(4);ws["Monte Carlo"].getRange("A:J").format.columnWidth=18;

// Audit.
title(ws.Audit,"Model audit","Terminal checks only; no model output depends on this sheet");
ws.Audit.getRange("A4:C9").values=[["Check","Result","Status"],["Customer roll-forward",null,null],["Cash roll-forward",null,null],["BvA actual periods",null,null],["Scenario selector",null,null],["No negative customers",null,null]];header(ws.Audit.getRange("A4:C4"));
ws.Audit.getRange("B5").formulas=[[`=ABS(SUM('Revenue Model'!B6:B52)-SUM('Revenue Model'!E5:E51))`]];ws.Audit.getRange("B6").formulas=[[`=ABS(SUM('Cash Flow'!B6:B52)-SUM('Cash Flow'!H5:H51))`]];ws.Audit.getRange("B7").formulas=[[`=COUNT(BvA!C5:C52)`]];ws.Audit.getRange("B8").formulas=[[`=COUNTIF(Assumptions!A8:A10,Assumptions!B4)`]];ws.Audit.getRange("B9").formulas=[[`=MIN('Revenue Model'!E5:E52)`]];
for(let r=5;r<=9;r++) ws.Audit.getRange(`C${r}`).formulas=[[r===7?`=IF(B${r}=42,"PASS","REVIEW")`:r===9?`=IF(B${r}>=0,"PASS","REVIEW")`:`=IF(B${r}<0.01,"PASS","REVIEW")`]];
ws.Audit.getRange("C5:C9").conditionalFormats.add("containsText",{text:"PASS",format:{fill:green,font:{color:"#274E13",bold:true}}});ws.Audit.getRange("A:C").format.columnWidth=24;

// Executive summary with linked KPIs and chart source.
title(ws["Executive Summary"],"SaaS FP&A executive summary","Actuals through June 2026; forecast through December 2026 | Selected case linked from Assumptions");
ws["Executive Summary"].getRange("A4:B10").values=[["KPI","Value"],["Selected scenario",null],["Jun-26 MRR",null],["Jun-26 ARR",null],["YTD revenue variance",null],["YTD operating loss",null],["Jun-26 ending cash",null]];header(ws["Executive Summary"].getRange("A4:B4"));
ws["Executive Summary"].getRange("B5").formulas=[["=Assumptions!B4"]];ws["Executive Summary"].getRange("B6").formulas=[["='Revenue Model'!F46"]];ws["Executive Summary"].getRange("B7").formulas=[["='Revenue Model'!G46"]];ws["Executive Summary"].getRange("B8").formulas=[["=SUM(BvA!D41:D46)"]];ws["Executive Summary"].getRange("B9").formulas=[["=SUM(BvA!C41:C46)-SUM(BvA!F41:F46)-SUM(BvA!I41:I46)-SUM(BvA!L41:L46)"]];ws["Executive Summary"].getRange("B10").formulas=[["='Cash Flow'!H46"]];money(ws["Executive Summary"].getRange("B6:B10"));
ws["Executive Summary"].getRange("A13:C25").values=[["Month","MRR","Ending cash"],...Array.from({length:12},(_,i)=>[fin[37+i][0],null,null])];header(ws["Executive Summary"].getRange("A13:C13"));for(let r=14;r<=25;r++){const source=41+(r-14);ws["Executive Summary"].getRange(`B${r}`).formulas=[[`='Revenue Model'!F${source}`]];ws["Executive Summary"].getRange(`C${r}`).formulas=[[`='Cash Flow'!H${source}`]];}ws["Executive Summary"].getRange("A14:A25").setNumberFormat("mmm-yy");money(ws["Executive Summary"].getRange("B14:C25"));
const chart=ws["Executive Summary"].charts.add("line",ws["Executive Summary"].getRange("A13:C25"));chart.titleText="MRR and ending cash trend";chart.hasLegend=true;chart.legend.position="bottom";chart.setPosition("E4","L19");ws["Executive Summary"].getRange("A:C").format.columnWidth=22;

title(ws.ReadMe,"Model guide","Synthetic SaaS portfolio case study");ws.ReadMe.getRange("A4:B18").values=[["Topic","Details"],["Period","Jan-2023 to Dec-2026; actuals through Jun-2026"],["Scenario control","Change Assumptions!B4; active drivers update automatically"],["Revenue","Customer bridge by plan, MRR, ARR, growth, churn, and ARPC"],["Unit economics","GRR, NRR, CAC, LTV, LTV/CAC, cohort retention, and ARR movement"],["Accrual model","Recognized revenue, billings, collections, accounts receivable, and deferred revenue"],["Statements","Integrated P&L and cash-flow bridge with ending cash and deferred revenue"],["Forecast QA","Rolling back-test, MAPE, and forecast bias"],["Risk","2,000-path Monte Carlo simulation of December liquidity"],["Expenses","Payroll, benefits, marketing, cloud/API, software, office/admin, professional services"],["Cash flow","Operating cash proxy, capex, financing, burn, runway"],["Variance convention","Positive is favorable for revenue and expense"],["Power BI","Load CSV files in /data and use /powerbi model and measures"],["AI commentary","Claude-ready prompt plus manual validation; no claim that Claude ran here"],["Disclosure","All entities and amounts are synthetic"]];header(ws.ReadMe.getRange("A4:B4"));ws.ReadMe.getRange("A:B").format.columnWidth=34;ws.ReadMe.getRange("B5:B18").format.wrapText=true;

for(const n of names){const u=ws[n].getUsedRange();if(u){u.format.font={name:"Aptos",size:10};u.format.verticalAlignment="center";}}
await wb.recalculate();
const check=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},maxChars:5000});
console.log(check.ndjson || String(check));
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(path.join(outDir,"SaaS_FP&A_Model_Masters.xlsx"));
const preview=await wb.render({sheetName:"Executive Summary",range:"A1:L25",scale:1,format:"png"});await fs.writeFile(path.join(root,"build","workbook_preview.png"),new Uint8Array(await preview.arrayBuffer()));
console.log("Wrote workbook");
