import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const root=path.resolve(".");
const skill="C:/Users/rahul/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations";
const tmp=path.join(root,"build","deck_masters"); const out=path.join(root,"deliverables","SaaS_FP&A_Management_Deck_Masters.pptx");
await fs.mkdir(tmp,{recursive:true}); await fs.mkdir(path.dirname(out),{recursive:true});
const {resolvePresentationFont,applyPresentationChartFont,finalizePresentation}=await import(pathToFileURL(path.join(skill,"container_tools/artifact_tool_utils.mjs")).href);
const font=resolvePresentationFont();

function parseCsv(text){const lines=text.trim().split(/\r?\n/);const p=l=>{const a=[];let c='',q=false;for(let i=0;i<l.length;i++){const x=l[i];if(x==='"'){if(q&&l[i+1]==='"'){c+='"';i++;}else q=!q;}else if(x===','&&!q){a.push(c);c='';}else c+=x;}a.push(c);return a};const h=p(lines[0]);return lines.slice(1).map(l=>Object.fromEntries(p(l).map((v,i)=>[h[i],v===''?null:(Number.isFinite(Number(v))?Number(v):v)])));}
const fin=parseCsv(await fs.readFile(path.join(root,"data","financials_monthly.csv"),"utf8"));
const subs=parseCsv(await fs.readFile(path.join(root,"data","subscriptions.csv"),"utf8"));
const hc=parseCsv(await fs.readFile(path.join(root,"data","headcount_monthly.csv"),"utf8"));
const opex=parseCsv(await fs.readFile(path.join(root,"data","operating_expenses.csv"),"utf8"));
const scenarios=parseCsv(await fs.readFile(path.join(root,"data","scenario_assumptions.csv"),"utf8"));
const metrics=parseCsv(await fs.readFile(path.join(root,"data","saas_metrics.csv"),"utf8"));
const cohorts=parseCsv(await fs.readFile(path.join(root,"data","cohort_analysis.csv"),"utf8"));
const riskSummary=parseCsv(await fs.readFile(path.join(root,"data","model_risk_summary.csv"),"utf8"));
const ytd=fin.filter(r=>r.month>='2026-01-01'&&r.month<='2026-06-01');
const sum=(a,k)=>a.reduce((s,r)=>s+(r[k]||0),0); const fmt=x=>Math.abs(x)>=1e6?`$${(x/1e6).toFixed(1)}M`:`$${Math.round(x/1e3)}K`;
const revA=sum(ytd,'actual_revenue'), revB=sum(ytd,'budget_revenue'), expA=sum(ytd,'actual_cogs')+sum(ytd,'actual_payroll')+sum(ytd,'actual_opex'), opLoss=revA-expA;
const junSubs=subs.filter(r=>r.month.startsWith('2026-06')); const mrr=sum(junSubs,'mrr'), arr=mrr*12, churn=sum(junSubs,'churned_customers')/sum(junSubs,'starting_customers');
const junCash=fin.find(r=>r.month.startsWith('2026-06')).ending_cash, decCash=fin.find(r=>r.month.startsWith('2026-12')).ending_cash;
const prs=Presentation.create({slideSize:{width:1280,height:720}}); const navy="#16324F", teal="#00A6A6", blue="#2F75B5", gray="#667085", pale="#EAF3F8", red="#C00000", green="#548235", white="#FFFFFF";
function textbox(slide,text,left,top,width,height,size=22,color=navy,bold=false){const s=slide.shapes.add({geometry:"textbox",position:{left,top,width,height},fill:"none",line:{fill:"none",width:0}});s.text=text;s.text.style={typeface:font,fontSize:size,color,bold,autoFit:"shrinkText",verticalAlignment:"middle"};return s;}
function base(title,section){const s=prs.slides.add();s.background.fill=white;textbox(s,section.toUpperCase(),72,32,260,22,13,teal,true);textbox(s,title,72,58,1136,62,34,navy,true);textbox(s,"Synthetic SaaS portfolio case | Actuals through June 2026",72,675,850,18,10,gray,false);return s;}
function chartFont(c){applyPresentationChartFont(c,{fontFamily:font});}

// Cover
{const s=prs.slides.add();s.background.fill=navy;textbox(s,"SaaS FP&A budgeting and forecasting model",84,150,1030,130,48,white,true);textbox(s,"Management review | Actuals through June 2026 | Forecast through December 2026",88,300,980,50,22,"#B9E8E8",false);textbox(s,"Revenue, expenses, headcount, cash, scenarios, and funding outlook",88,390,900,42,20,white,false);textbox(s,"All entities and values are synthetic",88,620,600,28,13,"#CBD5E1",false);}

// Executive summary
{const s=base("Revenue is ahead of budget, while cash declines toward a funding trigger","Executive summary");
 const labels=["Jun MRR", "Jun ARR", "YTD revenue variance", "YTD operating loss", "Jun cash", "Dec cash forecast"]; const vals=[fmt(mrr),fmt(arr),fmt(revA-revB),fmt(opLoss),fmt(junCash),fmt(decCash)];
 for(let i=0;i<6;i++){const x=72+(i%3)*380,y=155+Math.floor(i/3)*150;textbox(s,labels[i],x,y,320,28,15,gray,false);textbox(s,vals[i],x,y+35,320,60,34,(i===3||i===5)?red:(i===2?green:navy),true);}
 textbox(s,"Base forecast cash falls below $1.0M by December. Management should validate financing timing and the minimum operating buffer.",72,500,1120,90,24,navy,false);
 s.speakerNotes.textFrame.setText("Source: synthetic model outputs. Revenue variance = actual less budget. Expense and loss figures use the modeled accrual P&L.");}

// Revenue
{const s=base("MRR reaches $430K in June as the customer base expands","Revenue and customers");
 const ms=fin.filter(r=>r.month>='2025-07-01'&&r.month<='2026-06-01').map(r=>r.month.slice(0,7)); const mrrs=ms.map(m=>Math.round(sum(subs.filter(r=>r.month.startsWith(m)),'mrr')/100)); const customers=ms.map(m=>sum(subs.filter(r=>r.month.startsWith(m)),'ending_customers'));
 const c=s.charts.add("line",{position:{left:72,top:150,width:760,height:440},categories:ms.map(m=>m.slice(5)+"/"+m.slice(2,4)),series:[{name:"MRR ($K)",values:mrrs.map(v=>v/10),fill:blue}],hasLegend:false,dataLabels:{showValue:false}});chartFont(c);
 textbox(s,`June ARR\n${fmt(arr)}`,880,170,280,100,30,navy,true);textbox(s,`Gross logo churn\n${(churn*100).toFixed(1)}%`,880,300,280,100,30,teal,true);textbox(s,`Ending customers\n${customers.at(-1).toLocaleString()}`,880,430,280,100,30,navy,true);
 s.speakerNotes.textFrame.setText("Source: subscriptions.csv. MRR is ending customers multiplied by realized plan price, including modeled expansion for Growth and Enterprise plans.");}

// BvA
{const s=base("Revenue outperformance does not offset the current expense base","Budget versus actuals");const cats=["Revenue","COGS","Payroll","Other opex"];const actual=[revA,sum(ytd,'actual_cogs'),sum(ytd,'actual_payroll'),sum(ytd,'actual_opex')].map(v=>Math.round(v/10000)/100);const budget=[revB,sum(ytd,'budget_cogs'),sum(ytd,'budget_payroll'),sum(ytd,'budget_opex')].map(v=>Math.round(v/10000)/100);
 const c=s.charts.add("bar",{position:{left:72,top:150,width:800,height:430},categories:cats,series:[{name:"Budget ($M)",values:budget,fill:"#AAB7C4"},{name:"Actual ($M)",values:actual,fill:blue}],barOptions:{direction:"column",grouping:"clustered"},hasLegend:true,dataLabels:{showValue:true,position:"outEnd"}});chartFont(c);
 textbox(s,`Revenue\n${fmt(revA-revB)} favorable`,910,185,270,95,25,green,true);textbox(s,`Total expense\n${fmt(expA-(sum(ytd,'budget_cogs')+sum(ytd,'budget_payroll')+sum(ytd,'budget_opex')))} over budget`,910,320,270,110,24,red,true);textbox(s,`Operating loss\n${fmt(opLoss)}`,910,470,270,80,25,red,true);
 s.speakerNotes.textFrame.setText("Source: financials_monthly.csv, January-June 2026. Expense variance wording is actual less budget on this slide.");}

// Expense/headcount
{const s=base("Personnel remains the largest cost; headcount trails plan","Expense and headcount");const cats=["Payroll","Marketing","Cloud & API","Software","Office & Admin","Professional services"];const values=[sum(ytd,'actual_payroll'),...cats.slice(1).map(k=>sum(opex.filter(r=>r.month>='2026-01-01'&&r.month<='2026-06-01'&&r.category.toLowerCase()===k.toLowerCase()),'actual'))].map(v=>Math.round(v/10000)/100);
 const c=s.charts.add("bar",{position:{left:72,top:150,width:790,height:430},categories:cats,series:[{name:"YTD actual ($M)",values,fill:blue}],barOptions:{direction:"bar",grouping:"clustered"},hasLegend:false,dataLabels:{showValue:true,position:"outEnd"}});chartFont(c);
 const j=hc.filter(r=>r.month.startsWith('2026-06'));const plan=sum(j,'planned_headcount'),act=sum(j,'actual_headcount');textbox(s,`June actual HC\n${act}`,910,185,260,90,28,navy,true);textbox(s,`June plan HC\n${plan}`,910,315,260,90,28,gray,true);textbox(s,`Variance\n${act-plan}`,910,445,260,90,28,act-plan<=0?green:red,true);
 s.speakerNotes.textFrame.setText("Source: headcount_monthly.csv and operating_expenses.csv. Headcount variance is actual less planned at June month-end.");}

// Cash
{const s=base("Base-case cash declines to $0.9M by December","Cash burn and runway");const f=fin.filter(r=>r.month>='2026-01-01');const cats=f.map(r=>r.month.slice(5,7)+"/"+r.month.slice(2,4));const cash=f.map(r=>Math.round(r.ending_cash/10000)/100);const c=s.charts.add("line",{position:{left:72,top:150,width:860,height:430},categories:cats,series:[{name:"Ending cash ($M)",values:cash,fill:red}],hasLegend:false,dataLabels:{showValue:false}});chartFont(c);
 textbox(s,"Forecast",950,165,220,30,15,teal,true);textbox(s,`June cash\n${fmt(junCash)}`,950,225,220,90,27,navy,true);textbox(s,`December cash\n${fmt(decCash)}`,950,355,220,90,27,red,true);textbox(s,"Financing work should begin before the modeled cash buffer becomes operationally restrictive.",950,485,220,110,17,navy,false);
 s.speakerNotes.textFrame.setText("Source: financials_monthly.csv. Cash flow is a planning proxy based on modeled revenue, costs, capex, and financing; it is not a GAAP statement of cash flows.");}

// Scenarios
{const s=base("Cash outcomes reflect the tradeoff between growth and hiring pace","Forecast scenarios");const names=scenarios.map(r=>r.scenario);const ending=names.map((n,i)=>{const sc=scenarios[i];return Math.round((decCash + mrr*6*(sc.monthly_customer_growth-sc.monthly_churn)*sc.price_multiplier - sum(hc.filter(r=>r.month>='2026-07-01'),'personnel_cost')*(sc.hiring_multiplier-1) - sum(opex.filter(r=>r.month>='2026-07-01'&&r.category==='Marketing'),'budget')*(sc.marketing_multiplier-1) - sum(opex.filter(r=>r.month>='2026-07-01'&&r.category==='Cloud & API'),'budget')*(sc.cloud_unit_cost_multiplier-1))/10000)/100;});
 const c=s.charts.add("bar",{position:{left:72,top:160,width:760,height:410},categories:names,series:[{name:"Dec ending cash ($M)",values:ending,fill:teal}],barOptions:{direction:"column",grouping:"clustered"},hasLegend:true,dataLabels:{showValue:false}});chartFont(c);
 textbox(s,"Upside",900,170,250,28,16,teal,true);textbox(s,`December cash: $${ending[1].toFixed(2)}M\nGrowth improves, but faster hiring consumes cash in this case.`,900,205,280,110,18,navy,false);textbox(s,"Downside",900,350,250,28,16,red,true);textbox(s,`December cash: $${ending[2].toFixed(2)}M\nReduced hiring preserves cash despite weaker revenue.`,900,385,280,110,18,navy,false);
 s.speakerNotes.textFrame.setText("Scenario outputs are modeled estimates. Driver definitions are in scenario_assumptions.csv and the Excel Scenarios sheet.");}

// Advanced analytics
{const s=base("Retention and acquisition economics sharpen the growth view","SaaS unit economics");const r=metrics.slice(-12);const cats=r.map(x=>x.month.slice(5,7)+"/"+x.month.slice(2,4));const c=s.charts.add("line",{position:{left:72,top:155,width:760,height:420},categories:cats,series:[{name:"GRR (%)",values:r.map(x=>Math.round(x.grr*1000)/10),fill:blue},{name:"NRR (%)",values:r.map(x=>Math.round(x.nrr*1000)/10),fill:teal}],hasLegend:true,dataLabels:{showValue:false}});chartFont(c);const x=r.at(-1);textbox(s,`LTV / CAC\n${x.ltv_cac_ratio.toFixed(1)}x`,900,185,250,90,28,navy,true);textbox(s,`CAC\n${fmt(x.cac)}`,900,320,250,90,28,navy,true);textbox(s,`LTV\n${fmt(x.ltv)}`,900,455,250,90,28,teal,true);}
{const s=base("Cohort retention separates mix effects from true customer durability","Cohort analytics");const ages=Array.from({length:25},(_,i)=>i);const vals=ages.map(a=>{const q=cohorts.filter(x=>x.age_months===a);return Math.round(q.reduce((z,x)=>z+x.logo_retention,0)/q.length*1000)/10});const c=s.charts.add("line",{position:{left:72,top:155,width:850,height:420},categories:ages.map(String),series:[{name:"Average logo retention (%)",values:vals,fill:teal}],hasLegend:false,dataLabels:{showValue:false}});chartFont(c);textbox(s,"Interpretation",965,175,210,30,16,teal,true);textbox(s,"Cohort retention should be segmented by plan, channel, and customer size before operational decisions.",965,220,220,180,19,navy,false);}
{const s=base("Liquidity risk remains material across stochastic outcomes","Forecast model risk");const rm=Object.fromEntries(riskSummary.map(x=>[x.metric,x.value]));const c=s.charts.add("bar",{position:{left:72,top:170,width:740,height:390},categories:["P05","P50","P95"],series:[{name:"Dec ending cash ($M)",values:[rm['P05 ending cash'],rm['P50 ending cash'],rm['P95 ending cash']].map(v=>Math.round(v/10000)/100),fill:blue}],barOptions:{direction:"column",grouping:"clustered"},hasLegend:true,dataLabels:{showValue:false}});chartFont(c);textbox(s,`Below $1M\n${(rm['Probability below $1M']*100).toFixed(1)}%`,880,190,280,100,29,red,true);textbox(s,`Below zero\n${(rm['Probability below zero']*100).toFixed(1)}%`,880,335,280,100,29,red,true);textbox(s,`Back-test MAPE\n${(rm['Revenue forecast MAPE']*100).toFixed(1)}%`,880,480,280,80,25,navy,true);}

// Close
{const s=base("Management priorities for the next forecast cycle","Actions and controls");const items=[["Funding","Confirm the minimum cash buffer and start financing planning before Q4."],["Retention","Validate churn drivers by plan and owner before assigning corrective actions."],["Hiring","Tie each forecast hire to capacity, quota, or control requirements."],["Cloud cost","Separate usage growth from unit-price and architecture effects."],["Reporting","Refresh actuals monthly, reconcile source totals, and review AI commentary manually."]];let y=155;for(const [h,b] of items){textbox(s,h,90,y,210,46,22,teal,true);textbox(s,b,310,y,850,46,21,navy,false);y+=92;}s.speakerNotes.textFrame.setText("Actions are portfolio-case recommendations based on the model mechanics and should be validated against real operating constraints before use.");}

for(let i=0;i<prs.slides.items.length;i++){const slide=prs.slides.items[i];const img=await prs.export({slide,format:"png",scale:1});await fs.writeFile(path.join(tmp,`slide-${i+1}.png`),new Uint8Array(await img.arrayBuffer()));}
const candidate=path.join(tmp,"candidate.pptx");await (await PresentationFile.exportPptx(prs)).save(candidate);
await finalizePresentation({explicitTotalSlideCount:11,requiredNativeTableOwnerSlides:[],requiredNativeChartOwnerSlides:[3,4,5,6,7,8,9,10],materializeLiteralChartWorkbooks:true,workspaceDir:root,candidatePath:candidate,finalPath:out,pythonExecutable:"C:/Users/rahul/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe",integrityValidatorPath:path.join(skill,"container_tools/inspect_presentation_package_integrity.py"),layoutValidatorPath:path.join(skill,"container_tools/inspect_presentation_layout_geometry.py"),layoutArgs:["--expected-slide-size-emu","12192000,6858000","--validate-heading-fit"],fontPolicy:{basis:"design",families:[font]},verifyArtifactToolImport:true,receiptPath:path.join(tmp,"validation-masters.json")});
console.log(out);
