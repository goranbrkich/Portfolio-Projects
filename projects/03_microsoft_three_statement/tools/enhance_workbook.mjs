// Add charts, explanatory Notes and an auditable OLS/HC3 worksheet.
// Workbook authoring requires the Codex primary runtime and @oai/artifact-tool.
import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const [project,input,output,metadataPath,qa] = process.argv.slice(2);
const metadata=JSON.parse(await fs.readFile(metadataPath,'utf8'));
const layout=JSON.parse(await fs.readFile(`${project}/evidence/workbook_layout.json`,'utf8'));
const regression=JSON.parse(await fs.readFile(`${project}/evidence/regression/regression_results.json`,'utf8'));
const code=await fs.readFile(`${project}/src/regression_analysis.py`,'utf8');
const criticalValues=JSON.parse(await fs.readFile(`${project}/evidence/regression/student_t_critical_values.json`,'utf8'));
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(input));
await fs.mkdir(qa,{recursive:true});
const blue='#17365A', light='#5B9BD5', orange='#B8833F', gray='#7F8EA3';
const author=w.comments.setSelf({displayName:'Goran Brkich'}).id;
const noteCounts={},panels=[],views=[],chartChecks=[];
const formulasBefore=Object.fromEntries(metadata.map(s=>[s.name,Object.fromEntries(Object.entries(s.cells).filter(([,c])=>c.formula).map(([a,c])=>[a,c.formula]))]));
const quote=name=>`'${name.replaceAll("'","''")}'`;
const cols=Array.from({length:15},(_,i)=>String.fromCharCode(69+i));
const common={
 'Accounts receivable, net':'Amounts owed by customers after allowances. Forecast receivables equal annual revenue times receivable days divided by 365. Higher balances consume operating cash.',
 'Receivables':'Amounts owed by customers after allowances. Forecast receivables equal annual revenue times receivable days divided by 365. Higher balances consume operating cash.',
 'Inventories':'Products and materials held for sale or use. Forecast inventory equals cost of revenue times inventory days divided by 365.',
 'Accounts payable':'Amounts owed to suppliers, including disclosed capital expenditure payables. The forecast separates operating payables from capital payables before deriving CFO.',
 'Total assets':'Current assets plus PPE, lease assets, investments, goodwill, intangibles and other long-term assets. A year-end balance, not an annual cash-flow amount.',
 'Total liabilities':'Current and long-term obligations, including debt, tax, unearned revenue and lease liabilities. Finance leases are included in the relevant other-liability categories.',
 'Total liabilities and stockholders equity':'Total liabilities plus stockholders equity. It must equal total assets under the independent balance-sheet check.',
 'Total current assets':'Cash, short-term investments, receivables, inventory and other current assets. Amounts expected to be available or used within the operating cycle.',
 'Total current liabilities':'Obligations classified as current, including payables, current debt, compensation, taxes, unearned revenue and other current liabilities.',
 'Accrued compensation':'Employee compensation owed but not yet paid. The forecast uses a revenue-based planning ratio; increases can provide an operating cash timing benefit.',
 'Operating lease right-of-use assets':'Assets representing rights to use property under operating leases. Forecast carrying amounts follow legacy amortization and new lease additions.',
 'Intangible assets, net':'Identifiable intangible assets after accumulated amortization. The forecast follows the disclosed amortization schedule without unannounced acquisitions.',
 'Other current assets':'Other short-term operating assets as reported. Forecast balances use a revenue-based planning ratio and their increases consume operating cash.',
 'Other long-term assets':'Other noncurrent assets as reported. Forecast balances use a revenue-based planning ratio and their changes affect the operating cash reconciliation.',
 'Short-term investments':'Investment reserve separate from cash. The forecast follows planned growth and the minimum reserve, liquidating eligible balances before liquidity borrowing.',
 'Equity and other investments':'Noncurrent equity and other investments. The forecast increases the opening balance by the assumed cash purchases without unrealized valuation gains.',
 'Current portion of long-term debt':'Debt carrying amount classified as due within the next year. The schedule applies contractual principal timing and the relevant book adjustments.',
 'Long-term debt':'Debt carrying amount classified beyond the next year. It excludes separately modeled lease liabilities.',
 'Short-term debt':'Reported short-term borrowing. The forecast funding schedule determines incremental liquidity borrowing and its classification.',
 'Impairment, integration and restructuring':'Historical reported other operating expenses. No unannounced impairment, integration or restructuring events are forecast.',
 'Revenue':'Recognized sales from products and services. Consolidated forecast revenue is the sum of the three segment forecasts. Revenue is different from cash collected.',
 'Operating income':'Gross profit less research and development, sales and marketing, general and administrative costs, and other operating expenses. It excludes nonoperating items and income tax.',
 'Net income':'Income before taxes less income tax expense. It includes noncash items, so it differs from operating cash flow.',
 'Gross profit':'Revenue less cost of revenue. It measures profit before operating expenses.',
 'Cost of revenue':'Expenses assigned to delivering products and services. Forecast amounts include allocated depreciation, amortization and operating lease expense.',
 'Free cash flow':'Operating cash flow less cash capital expenditure. Cash capex is negative in the cash-flow statement, so the worksheet adds that signed amount to CFO. This measure excludes finance lease principal.',
 'FCF after finance lease principal':'Free cash flow less cash finance lease principal payments. It makes the financing cash burden of leased assets visible.',
 'Cash capex':'Cash spent on capital assets, including modeled capitalized cash interest. It excludes newly recognized noncash finance lease assets.',
 'Cash capital expenditure':'Signed cash outflow for capital investment. Its absolute amount includes capitalized cash interest. New finance lease assets are noncash and excluded.',
 'Operating cash flow':'Cash generated by operations after noncash adjustments and changes in operating assets and liabilities. It excludes cash capital expenditure.',
 'Net cash from operations':'Net income plus noncash adjustments and signed operating balance changes. Depreciation and stock-based compensation are added back once.',
 'Operating margin':'Operating income divided by revenue for the same fiscal year. A declining margin can reflect infrastructure depreciation and operating cost assumptions.',
 'Gross margin':'Gross profit divided by revenue for the same fiscal year.',
 'Net margin':'Net income divided by revenue for the same fiscal year.',
 'Revenue growth':'Current-year revenue divided by prior-year revenue less one. Growth across a segment definition change is unavailable.',
 'Diluted earnings per share':'Net income divided by diluted weighted average shares. Both numerator and shares are in millions, producing dollars per share.',
 'Diluted EPS':'Net income divided by diluted weighted average shares. The forecast share count has its own driver and does not mechanically follow buyback cash.',
 'Basic earnings per share':'Net income divided by basic weighted average shares, before dilution from potentially dilutive securities.',
 'Cash and short-term investments':'Cash and cash equivalents plus short-term investments. This is a liquidity measure, not total corporate cash resources.',
 'Debt carrying amount':'Debt face principal adjusted for discounts, issuance costs, hedge accounting and exchange premiums. Finance lease liabilities are reported separately.',
 'Finance lease liabilities':'Outstanding finance lease obligations, including current and noncurrent portions. Additions recognize an asset and liability without immediate cash capex.',
 'Property and equipment, net':'Gross PPE less accumulated depreciation. Forecast net PPE follows additions and annual depreciation cohorts.',
 'Total stockholders equity':'Common stock and paid-in capital plus retained earnings and accumulated other comprehensive income.',
 'Stock-based compensation':'Noncash employee equity compensation. It remains embedded in operating expenses, is added back once in CFO, and increases paid-in capital.',
 'Accumulated other comprehensive income':'Equity items recognized outside net income. The forecast holds the FY2026 balance constant.',
 'Income tax expense':'Historical reported tax expense. Forecast income before taxes multiplied by the effective tax-rate assumption.',
 'Effective tax rate':'Income tax expense divided by income before tax for the same period. The forecast is a planning assumption, not statutory tax advice.',
 'Receivable days':'Ending receivables divided by annual revenue multiplied by 365. Forecast receivables equal revenue times days divided by 365.',
 'Receivable days on revenue':'Ending receivables divided by annual revenue multiplied by 365. Increasing this driver ties up more operating cash.',
 'Inventory days':'Ending inventory divided by annual cost of revenue multiplied by 365. Forecast inventory equals cost of revenue times days divided by 365.',
 'Operating payable days':'Operating payables, excluding capital expenditure payables, divided by annual cost of revenue multiplied by 365.',
 'Operating accounts payable':'Trade payables related to operations. Capital expenditure payables are excluded before calculating the operating cash-flow effect.',
 'Capital expenditure payables':'Unpaid capital asset purchases at year-end. Their increase raises capital additions but is excluded from the operating payable cash adjustment.',
 'Operating income plus disclosed D&A':'Operating income plus separately disclosed PPE depreciation and intangible amortization. This is an analytical earnings proxy, not Microsoft-reported EBITDA.',
 'Net debt excluding leases':'Debt carrying amount less cash and short-term investments. Negative net debt indicates net liquidity under this definition.',
 'Net debt including finance leases':'Debt carrying amount plus finance lease liabilities less cash and short-term investments. Operating leases remain separately disclosed.',
 'Cash dividends and repurchases':'Cash shareholder distributions. This does not include noncash stock-based compensation or employee share issuance proceeds.',
 'Cash conversion after finance lease principal':'FCF after finance lease principal divided by the operating income plus disclosed D&A proxy. The numerator reflects reinvestment and lease payments.',
 'Operating current working capital proxy':'Receivables plus inventory and other current assets, less modeled operating payables, compensation, current unearned revenue and other operating current liabilities. Cash, debt and capital payables are excluded.',
 'Change in net PPE / revenue growth amount':'Annual change in net PPE divided by the annual revenue increase. This measures incremental asset intensity and is unavailable when its denominator is zero.',
 'Basic weighted average shares':'Average shares outstanding during the fiscal year, in millions. This is not the year-end share count.',
 'Diluted weighted average shares':'Weighted average shares including potential dilution, in millions. The forecast uses a separate net-change assumption.',
 'Goodwill':'Acquisition-related asset recognized when purchase consideration exceeds identifiable net assets. Held constant in the forecast because no unannounced acquisitions or impairments are assumed.',
 'Land held constant':'Land is excluded from depreciable PPE and remains at its FY2026 amount in the forecast.',
 'Closing cash and cash equivalents':'Opening cash plus operating, investing and financing cash flows and the FX cash effect. This closing balance links to the balance sheet.',
 'Liquidity borrowing requirement':'Funding needed after available short-term investments above the reserve are liquidated. Borrowing occurs at year-end and starts accruing interest next year. This is not a committed credit facility.',
 'Other income and expense, net':'Net nonoperating items. The forecast uses investment income less recognized interest expense and does not extrapolate FY2026 investment gains.',
 'Research and development':'Research and product development expense. Forecast base costs are driven by revenue and then receive allocated D&A and lease expense.',
 'Sales and marketing':'Selling and marketing expense driven by its revenue-based forecast assumption.',
 'General and administrative':'Corporate administration expense driven by its revenue-based forecast assumption.',
 'Retained earnings':'Accumulated earnings after dividends and the retained-earnings share of repurchases. Forecast closing balance is opening retained earnings plus net income less these distributions.',
 'Common stock and paid-in capital':'Equity contributed through shares and compensation. Forecast opening balance plus SBC and cash issuance less the allocated paid-in-capital portion of repurchases.',
};
const extras={
 Overview:'Revenue can expand while the operating margin falls if infrastructure depreciation grows faster than revenue. Review the cash capex and free cash flow chart together. A lower capex-to-revenue ratio releases cash, but does not remove finance lease principal payments. The separate lease balance shows the obligations attached to noncash capacity additions. Ending liquidity also depends on dividends, repurchases and investment-reserve assumptions. Use the linked observations below to compare FY2026 with the active FY2031 case, then investigate the underlying schedules before interpreting the increase as distributable cash.',
 'Equity Research':'The difference between net income and operating cash flow identifies the role of noncash charges and operating balance changes. Compare FCF with FCF after finance lease principal to avoid overlooking leased-capacity funding. A strong EPS path can coexist with a lower operating margin because the diluted share count changes independently. FY2026 investment gains should not be treated as a recurring operating driver. Assess the capex decline, receivable days and lease payment schedule together before translating these forecasts into a valuation. The charts and observations update with the selected case.',
 'Investment Banking':'Compare both debt definitions before assessing financing capacity. A company can show net cash excluding leases and still carry substantial finance lease commitments. An increase in the earnings proxy partly reflects noncash D&A and does not itself provide cash for debt service. Current-ratio improvement measures short-term coverage but does not establish a committed borrowing facility. Shareholder distributions compete with capex and lease payments for operating cash. Review the debt face-to-book bridge and contractual maturities before using these metrics in a financing or transaction analysis.',
 'Private Equity':'Separate cash capex from noncash finance lease additions when estimating total infrastructure commitments. SBC increases the difference between accounting earnings and cash, while also affecting shareholder dilution. FCF after finance lease principal is more informative for cash availability than the earnings proxy alone. The working-capital measure excludes cash, debt and capital payables so it can support operating diligence. The incremental PPE ratio is sensitive to the size of the revenue increase and is not an investment return. These measures support operating diligence without assuming an acquisition or an LBO.',
 'FP&A':'Explain the operating income change by revenue contributions and changes in each cost category. The bridge is additive and should end at the same FY2027 operating income as the income statement. Forecast segment demand also increases capex, lease additions and subsequent depreciation, so revenue growth does not flow directly to profit. The expense chart separates R&D, marketing and administration to show where spending assumptions change. Use the active-case observations to identify the largest planning trade-offs, then change the owning driver on Assumptions and review cash as well as profit.',
 Assumptions:'Compare the three revenue growth drivers with investment intensity before editing a case. Faster demand may increase depreciation and finance lease obligations even when cash capex is a lower share of sales. Receivable and payable days change cash conversion rather than revenue recognition. The useful-life inputs affect expense timing and net PPE, not the initial cash capex budget. Allocation percentages distribute already calculated expenses and must total 100%. Tail assumptions affect current versus long-term classification in the final forecast year. Change one driver at a time and review its statement and schedule consequences.',
 'Income Statement':'Revenue growth and the operating margin should be interpreted together. Segment cost rates exclude the D&A and lease costs that the expense schedules add separately. A rising depreciation burden can reduce margin even if the underlying operating cost ratios improve. Net income also depends on opening-balance investment income, recognized interest and the effective tax rate. EPS reflects the independent diluted-share driver. The forecast does not repeat FY2026 investment gains. Use the profit and margin charts to distinguish revenue expansion from changes in operating profitability and nonoperating income.',
 'Balance Sheet':'Growing PPE and lease assets show the capital committed to infrastructure. Cash is an output of the cash-flow calculation and funding policy. It is not entered to force the balance sheet to reconcile. Receivables, deferred revenue and operating liabilities are tied to planning drivers, while goodwill, land and AOCI remain constant. Finance lease current liabilities are included in other current liabilities and their noncurrent portion in other long-term liabilities. Compare asset growth with equity and liability funding, then check the independent reconciliations on Audit.',
 'Cash Flow':'Read the profit-to-cash adjustments before evaluating FCF. Operating asset increases consume cash and operating liability increases release cash. Historical cash-flow changes need not equal simple balance-sheet differences because reported cash-flow statements can include acquisitions, FX and reclassifications. Forecast cash capex already includes capitalized cash interest, and finance lease principal appears in financing cash flow. The reserve liquidation and borrowing lines show the funding sequence. Compare the two FCF measures in the charts to see the cash burden of finance leases and assess the payout assumptions alongside reinvestment.',
 Segments:'The three revenue paths identify which businesses generate the forecast expansion. Segment margins also include explicitly allocated D&A, lease expense and operating costs. They therefore reflect allocation assumptions as well as segment growth and base cost rates. FY2023 to FY2026 have the current segment definition, making those years the appropriate common historical comparison for the new charts. Earlier definitions remain in the original data and should not be joined into a continuous growth series. Consolidated operating income deducts unallocated corporate expense after summing segment operating income.',
 'Expense Plan':'The base R&D rate and separately allocated D&A and leases together produce total R&D. SBC remains embedded in the operating cost assumptions and is added back once in CFO. Adding it again as a separate expense would double count it. Depreciation increases as successive PPE cohorts enter service, so expense can keep rising after capex intensity declines. Compare the category chart with the infrastructure expense chart to distinguish revenue-driven spending from accumulated asset costs. The segment allocation changes presentation by business without changing the consolidated expense total.',
 'Working Capital':'Receivable days determine how much recognized revenue remains uncollected at year-end. Inventory and payable days use cost of revenue rather than sales. Capital payables are separated from operating payables to prevent unpaid capex from being counted as operating cash generation. Changes in unearned revenue can release operating cash before revenue is recognized. The cash-effect rows reverse the sign of operating asset increases and retain the sign of operating liability increases. Review the day-based assumptions and cash effects together rather than treating a growing balance as automatically favorable.',
 'PPE and Intangibles':'Cash capex, the change in capital payables and new finance lease assets jointly determine total PPE additions. These are different funding routes for capital assets and cannot be added to cash outflow a second time. Legacy assets and each new cohort have separate depreciation calculations. The half-year convention recognizes partial commissioning-year expense. Increasing useful lives lowers near-term depreciation and raises net carrying value, so it changes profit timing rather than investment cost. Intangibles follow the disclosed amortization schedule. Compare additions and depreciation to explain the forecast net PPE path.',
 'Debt and Leases':'Contractual debt principal, accounting adjustments and lease principal have distinct schedules. The face-to-book bridge matters when comparing balance-sheet debt with contractual repayment amounts. Legacy lease cash payments contain interest, while new leases use modeled terms and rates. Finance lease principal is a financing outflow. Operating lease expense and liability reductions need a cash-flow reconciliation because their timing differs. Liquidating investments above the reserve precedes additional borrowing. Evaluate financing obligations alongside liquid assets and the shareholder payout plan before describing excess cash as available for distribution.',
 'Tax and Equity':'The tax-rate assumption affects earnings and equity, while tax balance changes also affect cash. SBC and employee share issuance increase paid-in capital, and repurchases are divided between paid-in capital and retained earnings. Dividends reduce retained earnings and financing cash flow under the same-year payment assumption. Diluted shares remain an independent forecast input, so buyback cash does not guarantee a matching EPS increase. Review distributions against FCF after lease principal before interpreting rising equity or earnings as spare funding capacity.',
 'Historical Inputs':'The source report year identifies the filing supplying the latest comparative presentation by the fixed cutoff. It can differ from the fiscal year of the observation. Restatements and reclassifications are retained rather than replaced with older original values. Reported cash-flow adjustments preserve their signs. Blank or unavailable note observations must not be treated as zero. The new chart compares revenue, CFO and cash capex using the ten actual years only. These historical cells feed both the operating model and the separate regression worksheet. Refresh the underlying sources before extending the fiscal-year horizon.',
 'Filing Notes':'Face debt describes principal owed, whereas carrying debt reflects accounting adjustments. Lease additions recognize assets and liabilities without a matching immediate cash capex payment. Capital payables also bridge asset additions and cash paid. The disclosed maturity and amortization schedules provide the starting point for forecast timing. Amounts rounded in notes can differ slightly from statement aggregates. FY2032 tail assumptions are identified separately on Assumptions. The charts use recent disclosed years and preserve unavailable observations rather than inserting zeros.',
 Audit:'Read each difference against its stated accounting relationship and units. A small floating-point residual is different from a missing input or a genuine imbalance. Forecast balance-sheet, cash, PPE, debt and equity checks should remain within the existing USD 0.01 million tolerance. Allocation checks use percentages rather than USD amounts. The audit sheet reads the model and does not drive business calculations. A case change requires recalculation before judging the differences. The independent Python verification compares saved statement values, while the added regression verification compares the OLS and HC3 outputs separately.',
 'Data >>':'Historical Inputs contains the audited statements, Filing Notes contains supporting disclosures, and Audit checks the resulting model. The primary source and fiscal-year cutoff remain unchanged in this revision. Source rows preserve reported meanings and signs. The regression worksheet reads revenue and CFO from historical cells and never feeds back into those inputs. Notes explain the key source and calculation cells. Use the source report year and URLs to trace an observation before updating it. Charts display the existing data and formula-linked forecast rather than an additional dataset.',
 'Model Guide':'Use cell Notes for definitions and exact formula references, and the larger analytical panels for interpretation. All added operating charts read the same active forecast case. Financial Modeling uses actual FY2017 to FY2026 observations, so switching forecast cases does not change its historical regression. Its Excel formulas calculate OLS and HC3 without an embedded Python service. Copy the displayed code into VSCode to reproduce the historical results with NumPy and SciPy. Regression estimates describe historical association and do not replace the linked operating forecast.',
};

function note(sheet,address,body){
 w.notes.add({id:`${sheet.name}:${address}`,target:{cell:{sheetName:sheet.name,sheetId:sheet.sheetId,address}},authorId:author,createdAt:'',body:{plainText:body}});
 noteCounts[sheet.name]=(noteCounts[sheet.name]||0)+1;
}
function description(label,sheet,row){
 if(sheet==='Historical Inputs'){
   for(const [group,map] of Object.entries(layout.raw)){
     const key=Object.entries(map).find(([,r])=>r===row)?.[0];if(!key)continue;
     if(group==='cashflow'&&key==='da_other')return 'Reported depreciation, amortization and other noncash adjustments. This aggregate can include impairments and need not equal the sum of separate note disclosures.';
     if(group==='cashflow'&&key==='other_financing')return 'Reported aggregate other financing cash flow. Historical finance lease principal is separated in the presented Cash Flow sheet where note data permit.';
     const owner=group==='income'?'Income Statement':group==='balance'?'Balance Sheet':'Cash Flow';
     const section=group==='income'?'IS':group==='balance'?'BS':'CF';
     const mapped=metadata.find(s=>s.name===owner).labels.find(([a])=>Number(a.slice(1))===layout[section][key]);
     if(mapped&&common[mapped[1]])return common[mapped[1]]+' Primary filing value. The original reporting sign and fiscal year are preserved.';
   }
 }
 if(common[label])return common[label];
 const l=label.toLowerCase();
 if(l.includes('cash effect')||l.startsWith('change in '))return 'Signed operating cash adjustment. An increase in an operating asset consumes cash; an increase in an operating liability provides cash. Forecasts use the balance changes shown in the referenced schedule. Historical reported changes can include acquisition, FX and classification effects.';
 if(l.includes('cash capex')&&l.includes('/'))return 'Cash capital expenditure divided by the named revenue or earnings denominator in the same year. The capex numerator excludes noncash finance lease additions.';
 if(l.includes('allocated'))return 'Expense distributed using the explicit allocation inputs on Assumptions. These allocations are modeling assumptions rather than a Microsoft disclosure of these costs by segment. The formula identifies the owning expense schedule and percentage.';
 if(l.includes('margin'))return 'The named profit measure divided by revenue for the same fiscal year. Review numerator and denominator references in the formula. Percentages are stored as decimals.';
 if(l.includes('/'))return 'Ratio of the two measures named in the row label for the same fiscal year. Read the formula below for the exact numerator, denominator and sign treatment. Percentage or multiple formatting does not change the underlying value.';
 if(l.includes('opening'))return 'Beginning-of-year balance. Forecast periods carry forward the prior closing balance, with FY2026 providing the first forecast opening balance.';
 if(l.includes('closing'))return 'Year-end balance from the relevant roll-forward. Follow the formula to opening balances, additions, expenses, payments or other changes.';
 if(l.includes('depreciation'))return 'Noncash expense for PPE. Forecast legacy assets and new annual cohorts use the useful-life inputs and half-year convention. Expense reduces asset carrying value and is added back once in operating cash flow.';
 if(l.includes('amortization')&&l.includes('intangible'))return 'Noncash intangible expense. Forecast amortization follows the expected schedule disclosed in the FY2026 filing, subject to the remaining asset balance.';
 if(l.includes('principal'))return 'Reduction of debt or lease principal, separate from interest expense. Read the formula for the specific contractual or modeled payment schedule and sign.';
 if(l.includes('interest'))return 'Interest or investment income calculated using the named rate and balance. Capitalized interest is part of the capex budget; the remaining recognized interest affects net income.';
 if(l.includes('unearned')||l.includes('deferred rev'))return 'Customer payments billed or received before revenue recognition. Forecast balances use revenue-based planning ratios. Increases in the operating liability can provide operating cash.';
 if(l.includes('cash')||l.includes('cff')||l.includes('cfi'))return 'Cash balance or signed cash flow as named in the row. Positive cash-flow amounts are inflows and negative amounts are outflows. Follow the formula or source reference below for the classification.';
 if(l.includes('cohort'))return 'Remaining carrying amount of the named annual PPE cohort. Depreciation starts with a half-year charge and subsequent years use that cohort\'s useful life.';
 if(l.includes('less')||sheet==='Audit')return 'Independent reconciliation of the named balances or flows. USD differences use the existing 0.01 million tolerance. Allocation differences measure percentage points. n.a. means the check does not apply.';
 if(l.includes('growth'))return 'Current period divided by the prior matching period less one, or an editable forecast growth assumption. Preserve unavailable growth across definition changes.';
 return `${label}. ${sheet==='Historical Inputs'||sheet==='Filing Notes'?'Primary filing observation or disclosed schedule. Keep its original fiscal year, units and reported sign.':'This line belongs to the '+sheet+' calculation. Read its formula to trace the referenced inputs and intermediate totals.'}`;
}

// Explain each original metric label and every populated numeric/formula cell in its row.
for(const m of metadata){
 const s=w.worksheets.getItem(m.name), labels=Object.fromEntries(m.labels.map(([a,t])=>[Number(a.slice(1)),t]));
 note(s,'C2',`${m.name}. Amounts are USD millions unless the row states a ratio, days, years or per-share value. Hover over annotated labels or values for definitions and formula references.`);
 if(m.cells.E4?.formula)note(s,'E4','Linked active case name. Change only Assumptions D4: 1 Base, 2 Upside, 3 Downside. All forecast charts use the active case.');
 for(const [rowText,label] of Object.entries(labels)){
   const row=Number(rowText);if(row<8)continue;
   let parent=label;
   if(m.name==='Assumptions'&&['Base','Upside','Downside'].includes(label)){
     for(let k=row-1;k>=8;k--)if(labels[k]&&!['Base','Upside','Downside'].includes(labels[k])){parent=labels[k];break;}
   }
   const items=Object.entries(m.cells).filter(([a,c])=>Number(a.replace(/[A-Z]+/,''))===row&&!a.startsWith('C')&&(c.formula||(/^[D-S]\d+$/.test(a)&&c.value!=='')));
   if(!items.length)continue;
   const meaning=description(parent,m.name,row);
   note(s,`C${row}`,meaning+(parent!==label?` This is the editable ${label} case input.`:''));
   for(const [a,c] of items){
     if(!/^[D-S]\d+$/.test(a))continue;
     const period=m.cells[`${a.match(/^[A-Z]+/)[0]}7`]?.value;
     const formula=c.formula?`\nFormula: =${c.formula}\nCross-sheet references identify the owning schedule. Same-sheet references identify intermediate calculations.`:'\nStored input or primary source observation. It is not calculated in this cell.';
     let detail=meaning+(parent!==label?` Editable ${label} scenario input.`:'');
     if(m.name==='Assumptions'&&row>=9&&row<=200&&((row-9)%4===0))detail+=' The active forecast row selects the case input below it and feeds the operating model.';
     if(c.value==='n.a.')detail+=' This period is unavailable or not applicable. Do not interpret it as zero.';
     note(s,a,`${parent}${period?` (${period})`:''}.\n${detail}${formula}`);
   }
 }
 if(m.name==='Assumptions')note(s,'D4','Authoritative editable scenario selector. Enter 1 for Base, 2 for Upside or 3 for Downside. The existing validation and active driver rows remain in place.');
 if(m.name==='Data >>')note(s,'C7',extras['Data >>']);
}

const obs={
 Overview:[`="Selected "&Assumptions!G4&" case: FY2031 revenue is "&TEXT('Income Statement'!S8/1000,"0.0")&"bn and operating margin is "&TEXT('Income Statement'!S20,"0.0%")&"."`,`="Cash capex falls from "&TEXT(-'Cash Flow'!N30/'Income Statement'!N8,"0.0%")&" of revenue in FY2026 to "&TEXT(-'Cash Flow'!S30/'Income Statement'!S8,"0.0%")&" in FY2031."`,`="FY2031 FCF is "&TEXT('Cash Flow'!S56/1000,"0.0")&"bn; after finance lease principal it is "&TEXT('Cash Flow'!S57/1000,"0.0")&"bn."`],
 'Equity Research':[`="FY2031 operating margin: "&TEXT(S12,"0.0%")&". FCF margin: "&TEXT(S22,"0.0%")&"."`,`="FY2031 finance lease principal reduces FCF by "&TEXT((S18-S19)/1000,"0.0")&"bn."`,`="Diluted EPS changes from "&TEXT(N15,"0.00")&" in FY2026 to "&TEXT(S15,"0.00")&" in FY2031, using an independent share-count driver."`],
 'Investment Banking':[`="FY2031 net debt excluding leases: "&TEXT(S16/1000,"0.0")&"bn; including finance leases: "&TEXT(S17/1000,"0.0")&"bn."`,`="FY2031 finance lease liabilities: "&TEXT(S12/1000,"0.0")&"bn. Operating leases: "&TEXT(S13/1000,"0.0")&"bn."`,`="FY2031 lease-inclusive debt / earnings proxy: "&TEXT(S20,"0.00")&"x. The denominator is not reported EBITDA."`],
 'Private Equity':[`="FY2031 cash capex: "&TEXT(S10/1000,"0.0")&"bn; noncash finance lease additions: "&TEXT(S11/1000,"0.0")&"bn."`,`="FY2031 FCF after lease principal: "&TEXT(S15/1000,"0.0")&"bn; cash conversion on the earnings proxy: "&TEXT(S16,"0.0%")&"."`,`="FY2031 SBC: "&TEXT(S13/1000,"0.0")&"bn. Assess the cash add-back alongside the independent dilution assumption."`],
 'FP&A':[`="FY2027 operating income changes by "&TEXT((O19-N19)/1000,"0.0")&"bn under the selected case."`,`="FY2031 Intelligent Cloud revenue share: "&TEXT(S9/SUM(S8:S10),"0.0%")&"."`,`="PPE depreciation rises from "&TEXT(N16/1000,"0.0")&"bn in FY2026 to "&TEXT(S16/1000,"0.0")&"bn in FY2031."`],
 Assumptions:[`="FY2031 segment growth: PBP "&TEXT(S9,"0.0%")&", IC "&TEXT(S13,"0.0%")&", MPC "&TEXT(S17,"0.0%")&"."`,`="FY2031 cash capex / revenue: "&TEXT(S45,"0.0%")&". New finance lease assets / revenue: "&TEXT(S53,"0.0%")&"."`,`="FY2031 collection days: "&TEXT(S69,"0.0")&"; operating payable days: "&TEXT(S77,"0.0")&". These drivers affect cash conversion."`],
 'Income Statement':[`="FY2026 to FY2031 revenue CAGR: "&TEXT((S8/N8)^(1/5)-1,"0.0%")&"."`,`="Operating margin changes from "&TEXT(N20,"0.0%")&" to "&TEXT(S20,"0.0%")&"; net margin changes from "&TEXT(N29,"0.0%")&" to "&TEXT(S29,"0.0%")&"."`,`="FY2031 diluted shares: "&TEXT(S32,"#,##0")&"m. EPS: "&TEXT(S33,"0.00")&" dollars."`],
 'Balance Sheet':[`="FY2031 net PPE represents "&TEXT(S16/S22,"0.0%")&" of assets; cash and short-term investments represent "&TEXT(S10/S22,"0.0%")&"."`,`="FY2031 equity funds "&TEXT(S45/S22,"0.0%")&" of total assets under the selected payout assumptions."`,`="FY2031 current assets / current liabilities: "&TEXT(S14/S32,"0.00")&"x. Cash remains linked to the cash-flow statement."`],
 'Cash Flow':[`="FY2031 CFO: "&TEXT(S28/1000,"0.0")&"bn. Cash capex: "&TEXT(-S30/1000,"0.0")&"bn. FCF: "&TEXT(S56/1000,"0.0")&"bn."`,`="Finance lease principal reduces FY2031 FCF by "&TEXT((S56-S57)/1000,"0.0")&"bn."`,`="FY2031 reserve liquidation: "&TEXT(S60/1000,"0.0")&"bn; liquidity borrowing requirement: "&TEXT(S61/1000,"0.0")&"bn."`],
 Segments:[`="FY2031 revenue shares: PBP "&TEXT(S8/S51,"0.0%")&", IC "&TEXT(S22/S51,"0.0%")&", MPC "&TEXT(S36/S51,"0.0%")&"."`,`="FY2031 modeled operating margins: PBP "&TEXT(S19,"0.0%")&", IC "&TEXT(S33,"0.0%")&", MPC "&TEXT(S47,"0.0%")&"."`,`="FY2031 unallocated corporate operating expense: "&TEXT(S54/1000,"0.0")&"bn. It is deducted after the segment total."`],
 'Expense Plan':[`="FY2031 operating expenses: "&TEXT(S24/1000,"0.0")&"bn, or "&TEXT(S24/S8,"0.0%")&" of revenue."`,`="FY2031 PPE depreciation: "&TEXT(S10/1000,"0.0")&"bn; operating lease expense: "&TEXT(S12/1000,"0.0")&"bn."`,`="FY2031 total R&D includes "&TEXT(SUM(S15:S17)/1000,"0.0")&"bn of allocated D&A and operating lease expense."`],
 'Working Capital':[`="FY2031 receivable days: "&TEXT(S9,"0.0")&"; inventory days: "&TEXT(S11,"0.0")&"; operating payable days: "&TEXT(S16,"0.0")&"."`,`="FY2031 accounts payable: "&TEXT(S15/1000,"0.0")&"bn, including "&TEXT(S14/1000,"0.0")&"bn of capital payables."`,`="FY2031 receivables cash effect: "&TEXT(S31/1000,"0.0")&"bn; operating payables cash effect: "&TEXT(S35/1000,"0.0")&"bn."`],
 'PPE and Intangibles':[`="FY2031 PPE additions: "&TEXT(S15/1000,"0.0")&"bn versus depreciation of "&TEXT(S21/1000,"0.0")&"bn."`,`="FY2031 closing net PPE: "&TEXT(S24/1000,"0.0")&"bn. New finance lease assets contribute "&TEXT(S14/1000,"0.0")&"bn of annual additions."`,`="FY2031 remaining intangible assets: "&TEXT(S43/1000,"0.0")&"bn, after the disclosed amortization schedule."`],
 'Debt and Leases':[`="FY2031 debt face value: "&TEXT(S12/1000,"0.0")&"bn; carrying amount: "&TEXT(S18/1000,"0.0")&"bn."`,`="FY2031 finance lease liabilities: "&TEXT(S42/1000,"0.0")&"bn; operating lease liabilities: "&TEXT(S63/1000,"0.0")&"bn."`,`="FY2031 liquidity borrowing: "&TEXT(S11/1000,"0.0")&"bn. The funding sequence uses eligible investment reserves first."`],
 'Tax and Equity':[`="FY2031 effective tax rate: "&TEXT(S9,"0.0%")&". Income tax expense: "&TEXT(S10/1000,"0.0")&"bn."`,`="FY2031 dividends and repurchases: "&TEXT(SUM(S18:S19)/1000,"0.0")&"bn, compared with net income of "&TEXT(S14/1000,"0.0")&"bn."`,`="FY2031 equity: "&TEXT(S30/1000,"0.0")&"bn. Diluted shares: "&TEXT(S34,"#,##0")&"m."`],
 'Historical Inputs':[`="FY2026 revenue source report: "&TEXT(N106,"0")&". Balance sheet source report: "&TEXT(N107,"0")&". Cash-flow source report: "&TEXT(N108,"0")&"."`,`="FY2026 reported revenue: "&TEXT(N9/1000,"0.0")&"bn; operating cash flow: "&TEXT(N80/1000,"0.0")&"bn."`,`="FY2017 to FY2026 actuals are the regression sample. FY2016 is opening-balance context only."`],
 'Filing Notes':[`="FY2026 debt face value: "&TEXT(N24/1000,"0.0")&"bn. Discount and issuance costs: "&TEXT(N25/1000,"0.0")&"bn."`,`="FY2026 finance lease liabilities: "&TEXT(N12/1000,"0.0")&"bn; operating lease liabilities: "&TEXT(N15/1000,"0.0")&"bn."`,`="FY2026 capital expenditure payables: "&TEXT(N23/1000,"0.0")&"bn. These are separate from operating payables in the forecast."`],
 Audit:[`="Largest absolute forecast balance-sheet difference: "&TEXT(MAX(MAX(O8:S8),-MIN(O8:S8)),"0.000000")&" USD millions."`,`="Largest absolute forecast ending-cash difference: "&TEXT(MAX(MAX(O9:S9),-MIN(O9:S9)),"0.000000")&" USD millions."`,`="Checks are terminal diagnostics. An unavailable check is not a passed check, and no business calculation reads this sheet."`],
};
const roleSheets=['Equity Research','Investment Banking','Private Equity','FP&A'];
for(const name of roleSheets)obs[name]=obs[name].map(f=>f.replace(/\b([NOS])(\d+)\b/g,(_,col,row)=>({N:'I',O:'J',S:'N'})[col]+row));
const chartSpecs={
 'Equity Research':[['Cash conversion (USD billions)',[[17,'CFO'],[18,'FCF'],[19,'FCF after lease principal']],'money'],['Operating and FCF margins',[[12,'Operating margin'],[22,'FCF margin']],'percent']],
 'Investment Banking':[['Debt, leases and liquidity (USD billions)',[[11,'Debt'],[12,'Finance leases'],[15,'Liquid assets']],'money'],['Debt / earnings proxy',[[19,'Debt only'],[20,'Including finance leases']],'multiple']],
 'Private Equity':[['Cash and noncash investment (USD billions)',[[10,'Cash capex'],[11,'Finance lease additions']],'money'],['Cash conversion and capex / earnings proxy',[[16,'Cash conversion'],[22,'Cash capex / earnings proxy']],'percent']],
 'FP&A':[['Segment revenue (USD billions)',[[8,'PBP'],[9,'Intelligent Cloud'],[10,'MPC']],'money',6],['Operating expenses (USD billions)',[[12,'R&D'],[13,'Sales and marketing'],[14,'G&A']],'money']],
 Assumptions:[['Active segment revenue growth',[[9,'PBP'],[13,'Intelligent Cloud'],[17,'MPC']],'percent',10],['Active capital investment ratios',[[45,'Cash capex / revenue'],[53,'Finance lease additions / revenue']],'percent',10]],
 'Income Statement':[['Revenue and earnings (USD billions)',[[8,'Revenue'],[19,'Operating income'],[28,'Net income']],'money'],['Profit margins',[[12,'Gross margin'],[20,'Operating margin'],[29,'Net margin']],'percent']],
 'Balance Sheet':[['Asset composition (USD billions)',[[10,'Liquid assets'],[16,'Net PPE'],[22,'Total assets']],'money'],['Liability and equity funding (USD billions)',[[40,'Liabilities'],[45,'Equity']],'money']],
 'Cash Flow':[['Operating cash and reinvestment (USD billions)',[[28,'CFO'],[30,'Cash capex',-1],[56,'FCF']],'money'],['FCF before and after lease principal (USD billions)',[[56,'FCF'],[57,'After finance lease principal']],'money']],
 Segments:[['Revenue on the current segment definition (USD billions)',[[8,'PBP'],[22,'Intelligent Cloud'],[36,'MPC']],'money',6],['Modeled segment operating margins',[[19,'PBP'],[33,'Intelligent Cloud'],[47,'MPC']],'percent',6]],
 'Expense Plan':[['Operating expense categories (USD billions)',[[18,'R&D'],[20,'Sales and marketing'],[21,'G&A']],'money'],['Infrastructure expense (USD billions)',[[10,'PPE depreciation'],[11,'Intangible amortization'],[12,'Operating leases']],'money']],
 'Working Capital':[['Operating working-capital days',[[9,'Receivables'],[11,'Inventory'],[16,'Operating payables']],'days',9],['Operating balance cash effects (USD billions)',[[31,'Receivables'],[35,'Operating payables'],[36,'Unearned revenue']],'money',10]],
 'PPE and Intangibles':[['PPE additions and depreciation (USD billions)',[[8,'Cash capex'],[15,'Total additions'],[21,'Depreciation']],'money',9],['Net PPE and intangible balances (USD billions)',[[24,'Net PPE'],[43,'Intangibles']],'money',9]],
 'Debt and Leases':[['Debt and lease obligations (USD billions)',[[18,'Debt carrying amount'],[42,'Finance leases'],[63,'Operating leases']],'money',9],['Liquid assets (USD billions)',[[78,'Short-term investments'],[80,'Cash']],'money',9]],
 'Tax and Equity':[['Income and tax expense (USD billions)',[[14,'Net income'],[10,'Income tax']],'money'],['Shareholder cash distributions (USD billions)',[[18,'Repurchases'],[19,'Dividends']],'money']],
 'Historical Inputs':[['Historical revenue and operating cash (USD billions)',[[9,'Revenue'],[80,'CFO'],[89,'Cash capex',-1]],'money',0,10]],
 'Filing Notes':[['Disclosed lease liabilities (USD billions)',[[12,'Finance leases'],[15,'Operating leases']],'money',6,4],['Disclosed depreciation and amortization (USD billions)',[[8,'PPE depreciation'],[9,'Intangible amortization']],'money',6,4]],
};

function styleChart(ch,title,fmt){
 ch.title=title;ch.titleTextStyle.typeface='Arial';ch.titleTextStyle.fontSize=14;
 ch.legend={position:'top',textStyle:{typeface:'Arial',fontSize:11}};
 ch.xAxis={axisType:'textAxis',textStyle:{typeface:'Arial',fontSize:11}};
 ch.yAxis={numberFormatCode:fmt==='money'?'0.0,"bn"':fmt==='percent'?'0%':fmt==='multiple'?'0.0"x"':fmt==='number'?'0.00':'0.0',numberFormatSourceLinked:false,textStyle:{typeface:'Arial',fontSize:11}};
 ch.series.items.forEach((series,i)=>{const color=[blue,light,orange,gray][i%4];series.fill=color;series.line={fill:color,style:i===1?'dashed':'solid',width:2};});
 chartChecks.push({title,series:ch.series.items.map(s=>({values:s.formula,categories:s.categoryFormula}))});
}
function addChart(s,spec,chartRow,helperRow,index){
 const [title,series,fmt,requestedStart=0,requestedCount=15-requestedStart]=spec;
 const role=roleSheets.includes(s.name),start=role?(s.name==='FP&A'&&index===0?1:0):requestedStart,count=role?10-start:requestedCount;
 const rangeEnd=String.fromCharCode(67+series.length);
 const headers=[['Fiscal year',...series.map(v=>v[1])]];
 s.getRange(`C${helperRow}:${rangeEnd}${helperRow}`).values=headers;
 const data=[];
 for(let k=0;k<count;k++){
   const year=(role?2022:2017)+start+k,col=cols[start+k];
   data.push([`="FY${year} ${year<=2026?'A':'E'}"`,...series.map(([r,,sign=1])=>`=IF(OR(${col}${r}="n.a.",ISBLANK(${col}${r})),"",${sign===-1?'-':''}${col}${r})`)]);
 }
 s.getRange(`C${helperRow+1}:${rangeEnd}${helperRow+count}`).formulas=data;
 s.getRange(`C${helperRow}:${rangeEnd}${helperRow+count}`).format.font={name:'Arial',size:10};
 s.getRange(`D${helperRow+1}:${rangeEnd}${helperRow+count}`).setNumberFormat(fmt==='percent'?'0.0%':'#,##0.0;(#,##0.0);"–"');
 s.getRange(`C${helperRow}:${rangeEnd}${helperRow}`).format.fill='#E7EEF7';
 const ch=s.charts.add('line',s.getRange(`C${helperRow}:${rangeEnd}${helperRow+count}`));
 styleChart(ch,title,fmt);ch.setPosition(`${index===0?'C':'K'}${chartRow}`,`${index===0?'J':'T'}${chartRow+17}`);
 if(index===0&&!(chartSpecs[s.name]?.length>1))ch.setPosition(`C${chartRow}`,`S${chartRow+17}`);
 note(s,`C${helperRow}`,`${title}. Formula-linked chart data. A means actual and E means the active-case estimate. Unavailable historical observations remain blank.`);
}

// Preserve the original native explanatory shapes, expanding their text and bounds.
// Text/extent replacement uses the established DrawingML helper after XLSX export.
for(const m of metadata){
 const s=w.worksheets.getItem(m.name);
 const original=m.shapes;
 if(m.name==='Model Guide'){
   const guide=[
    'Read the three statements together. Revenue and profit describe annual performance, the balance sheet records year-end resources and obligations, and cash flow explains funding. An increasing earnings forecast can require substantial asset investment. The original financial-data cutoff remains fixed. The new analytical worksheet uses actuals only and does not extend the forecast horizon.',
    'Charts and live observations follow the same authoritative selector. Historical statements remain fixed. Review a later-year driver as well as the first forecast year when comparing cases, because asset cohorts and lease balances accumulate through time. Changes in one driver can affect several later statements. Excel and Python forecast input stores remain separate.',
    'A useful analysis follows the economic sequence rather than reading each statement in isolation. Higher segment revenue can raise cash capex and lease additions. These additions increase future depreciation and lease expense, which affect earnings. Working-capital balances, distributions and contractual payments then determine cash. Audit observes the result without driving it.',
    'Constant asset or equity balances are explicit planning choices rather than forecasts of every possible event. Useful-life and allocation assumptions affect profit timing and segment presentation. Independent share counts mean repurchase cash does not determine diluted EPS mechanically. Review these assumptions when interpreting cash accumulation, operating margins and earnings growth.',
    'Use the current segment definition for comparisons beginning FY2023. Do not treat the older definition as a continuous growth series. Annual statement comparatives can contain restatements or classification changes, so their source report year matters. The regression uses consolidated revenue and CFO, preserving the selected comparative presentations rather than constructing an alternative history.',
    'Use Historical Inputs for reported statement values and Filing Notes for debt, lease and asset disclosures. Amounts are USD millions, while shares are millions and per-share values are dollars. Ratios use decimals. Notes can be rounded differently from statement aggregates. Trace the source report year and URL before refreshing an observation or extending the sample.',
    'Professional views apply different financial questions to one operating model. They do not create separate forecasts. The regression is another diagnostic view of historical relationships and does not produce an investment recommendation. Small-sample fit does not establish valuation, acquisition feasibility or a stable causal relationship. Preserve the separation between a model output and the interpretation it supports.',
    extras['Model Guide'],
   ];
   original.forEach((p,i)=>panels.push({sheet:m.name,name:p.name,title:p.name,body:p.text.split(' | ').slice(1).join(' ')+ '\n\n'+guide[i],row:6+i*13,width:1750,height:225}));
   views.push({sheet:m.name,range:'C7:S18'});
   continue;
 }
 let first=original.length?Number(original[0].from.row):10;
 original.forEach((p,i)=>panels.push({sheet:m.name,name:p.name,title:p.name,body:p.text.split(' | ').slice(1).join(' ')+'\n\n'+extras[m.name],row:first+i*13,width:1750,height:225}));
 if(!original.length&&m.name==='Data >>'){
   s.shapes.add({geometry:'rect',anchor:{from:{row:11,col:2},extent:{widthPx:1750,heightPx:225}},fill:'#F1F5FA',line:{fill:'#B8C7D9',width:1}}).name='Data interpretation';
   panels.push({sheet:m.name,name:'Data interpretation',title:'Data interpretation',body:extras[m.name],row:11,width:1750,height:225});
   views.push({sheet:m.name,range:'C2:S24'});continue;
 }
 const obsRow=first+Math.max(original.length,1)*13+1;
 if(obs[m.name]){
   s.getRange(`C${obsRow}`).values=[['Selected-case observations']];
   s.getRange(`C${obsRow}:S${obsRow}`).format.font={name:'Arial',size:11,bold:true,color:blue};
   for(let i=0;i<obs[m.name].length;i++){
     const row=obsRow+i+1;s.getRange(`C${row}:S${row}`).merge();
     s.getRange(`C${row}`).formulas=[[obs[m.name][i]]];
     s.getRange(`C${row}:S${row}`).format.font={name:'Arial',size:10,color:'#172B46'};
     s.getRange(`C${row}:S${row}`).format.rowHeight=23;
     note(s,`C${row}`,'Live analytical observation. This formula reads the existing statement or schedule and updates with the selected case. '+obs[m.name][i]);
   }
 }
 const chartRow=obsRow+6;
 for(const [i,spec] of (chartSpecs[m.name]||[]).entries())addChart(s,spec,chartRow,chartRow+21+i*19,i);
 views.push({sheet:m.name,range:`C${first+1}:T${chartRow+(chartSpecs[m.name]?18:0)}`});
}

// Historical regression worksheet. It has no dependency back into the operating model.
const s=w.worksheets.add('Financial Modeling');s.showGridLines=false;s.tabColor=blue;
s.getRange('A1:B380').format.columnWidth=2;
s.getRange('C1:C380').format.columnWidth=11;
s.getRange('D1:E380').format.columnWidth=16;
s.getRange('F1:T380').format.columnWidth=13;
s.getRange('C1:T380').format.font={name:'Arial',size:10,color:'#172B46'};
s.getRange('C1:T380').format.rowHeight=20;
s.getRange('C2:T2').merge();s.getRange('C2').values=[['Financial Modeling']];s.getRange('C2').format.font={name:'Arial',size:16,bold:true,color:blue};
s.getRange('C3:T3').merge();s.getRange('C3').values=[['Historical FY2017–FY2026. OLS of ln(CFO) on ln(revenue), with HC3 robust standard errors. USD millions.']];
s.getRange('C4:D4').merge();s.getRange('C4').values=[['Case Selected:']];s.getRange('E4').formulas=[['=Assumptions!G4']];s.getRange('E4').format.font={name:'Arial',size:10,color:'#008000'};
s.getRange('C6:T6').merge();s.getRange('C6').values=[['Historical regressions do not change with the forecast case and do not feed the three-statement forecasts.']];
s.freezePanes.freezeRows(11);s.freezePanes.freezeColumns(3);
s.getRange('V10:W10').values=[['Degrees of freedom','95% two-sided t critical']];
s.getRange('V10:W10').format={fill:'#E7EEF7',font:{name:'Arial',size:10,bold:true},wrapText:true,rowHeight:36};
s.getRange('V11:W40').values=criticalValues;s.getRange('V11:V40').setNumberFormat('0');s.getRange('W11:W40').setNumberFormat('0.000000');
s.getRange('V1:W40').format.columnWidth=18;
note(s,'W10','Student t reference values at the 97.5th percentile, generated with scipy.stats.t.ppf(0.975,df). Degrees of freedom 1 to 30 cover both supplied samples. The 95% confidence intervals read the critical value for their residual degrees of freedom.');
const metrics=['Observations','Residual degrees of freedom','Mean x','Mean y','Sxx','Slope','Intercept','Residual sum of squares','Total sum of squares','R squared','Adjusted R squared','OLS residual standard error','OLS slope standard error','OLS intercept standard error','HC3 slope standard error','HC3 intercept standard error','HC3 slope t statistic','HC3 slope p value','95% t critical value','HC3 slope CI lower','HC3 slope CI upper','Durbin Watson','Maximum leverage','Leverage reference 4/n','HC3 intercept t statistic','HC3 intercept p value','HC3 intercept CI lower','HC3 intercept CI upper'];
function regressionBlock(first,last,summary,changes=false){
 const at=k=>`$E$${summary+k}`;
 const header=first-1;
 s.getRange(`C${header}:T${header}`).values=[['Fiscal year','Revenue','CFO',changes?'Delta ln revenue':'ln revenue',changes?'Delta ln CFO':'ln CFO','Centered x','Centered x squared','Centered x times y','Fitted y','Residual','Leverage h','HC3 weight',changes?'Fitted CFO growth':'Fitted median CFO','Residual squared','Centered y squared','Centered x times weight','x squared times weight','Residual change squared']];
 s.getRange(`C${header}:T${header}`).format={fill:blue,font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:44,verticalAlignment:'center'};
 const matrix=[];
 for(let row=first;row<=last;row++){
   const i=row-first,col=cols[i+(changes?1:0)];
   matrix.push([`=${changes?2018+i:2017+i}`,`='Historical Inputs'!${col}9`,`='Historical Inputs'!${col}80`,changes?`=F${13+i}-F${12+i}`:`=LN(D${row})`,changes?`=G${13+i}-G${12+i}`:`=LN(E${row})`,`=F${row}-${at(2)}`,`=H${row}^2`,`=H${row}*G${row}`,`=${at(6)}+${at(5)}*F${row}`,`=G${row}-K${row}`,`=1/${at(0)}+I${row}/${at(4)}`,`=L${row}^2/(1-M${row})^2`,changes?`=EXP(K${row})-1`:`=EXP(K${row})`,`=L${row}^2`,`=(G${row}-${at(3)})^2`,`=H${row}*N${row}`,`=I${row}*N${row}`,row===first?'=0':`=(L${row}-L${row-1})^2`]);
 }
 s.getRange(`C${first}:T${last}`).formulas=matrix;
 s.getRange(`C${first}:C${last}`).setNumberFormat('"FY"0');s.getRange(`D${first}:E${last}`).setNumberFormat('#,##0');
 s.getRange(`F${first}:T${last}`).setNumberFormat('0.000000');s.getRange(`O${first}:O${last}`).setNumberFormat(changes?'0.0%':'#,##0');
 s.getRange(`D${first}:E${last}`).format.font={name:'Arial',size:10,color:'#008000'};
 const f=[`=COUNT(F${first}:F${last})`,`=${at(0)}-2`,`=AVERAGE(F${first}:F${last})`,`=AVERAGE(G${first}:G${last})`,`=SUM(I${first}:I${last})`,`=SUM(J${first}:J${last})/${at(4)}`,`=${at(3)}-${at(5)}*${at(2)}`,`=SUM(P${first}:P${last})`,`=SUM(Q${first}:Q${last})`,`=1-${at(7)}/${at(8)}`,`=1-(${at(7)}/${at(1)})/(${at(8)}/(${at(0)}-1))`,`=SQRT(${at(7)}/${at(1)})`,`=${at(11)}/SQRT(${at(4)})`,`=${at(11)}*SQRT(1/${at(0)}+${at(2)}^2/${at(4)})`,`=SQRT(SUM(S${first}:S${last}))/${at(4)}`,`=SQRT(SUM(N${first}:N${last})/${at(0)}^2+${at(2)}^2*${at(14)}^2-2*${at(2)}*SUM(R${first}:R${last})/(${at(0)}*${at(4)}))`,`=${at(5)}/${at(14)}`,`=TDIST(ABS(${at(16)}),${at(1)},2)`,`=VLOOKUP(${at(1)},$V$11:$W$40,2,FALSE)`,`=${at(5)}-${at(18)}*${at(14)}`,`=${at(5)}+${at(18)}*${at(14)}`,`=SUM(T${first}:T${last})/${at(7)}`,`=MAX(M${first}:M${last})`,`=4/${at(0)}`,`=${at(6)}/${at(15)}`,`=TDIST(ABS(${at(24)}),${at(1)},2)`,`=${at(6)}-${at(18)}*${at(15)}`,`=${at(6)}+${at(18)}*${at(15)}`];
 for(let i=0;i<metrics.length;i++){s.getRange(`C${summary+i}:D${summary+i}`).merge();s.getRange(`C${summary+i}`).values=[[metrics[i]]];}
 s.getRange(`E${summary}:E${summary+metrics.length-1}`).formulas=f;s.getRange(`E${summary}:E${summary+27}`).setNumberFormat('0.000000');
 s.getRange(`E${summary+17}`).setNumberFormat('0.000E+00');s.getRange(`E${summary+25}`).setNumberFormat('0.000E+00');
 for(let r=first;r<=last;r++)for(let ci=2;ci<=19;ci++){
   const address=String.fromCharCode(65+ci)+r;
   const def=['Fiscal year. Actual observations only.','Historical revenue in USD millions.','Historical operating cash flow in USD millions.',changes?'Annual change in ln revenue.':'Natural logarithm of positive revenue.',changes?'Annual change in ln CFO.':'Natural logarithm of positive CFO.','x minus sample mean x.','Squared centered regressor. Summed to Sxx.','Centered x multiplied by y. Summed and divided by Sxx for the slope.','OLS fitted value in the model\'s logarithmic units.','Observed y less fitted y.','Leverage: 1/n plus centered x squared / Sxx. High values amplify the HC3 correction.','HC3 weight: residual squared divided by (1 minus leverage) squared.',changes?'EXP(fitted log change)-1: predicted proportional CFO change.':'EXP(fitted ln CFO): fitted conditional median in USD millions. This is not a bias-corrected conditional mean.','Squared OLS residual. Summed to RSS.','Squared y minus mean y. Summed to SST.','Centered x times HC3 weight, used in the intercept variance.','Centered x squared times HC3 weight, used in the slope variance.','Squared change in consecutive residuals. The first row is zero by construction, used in Durbin Watson.'][ci-2];
   note(s,address,def+'\nFormula: '+matrix[r-first][ci-2]);
 }
 metrics.forEach((name,i)=>note(s,`E${summary+i}`,name+'. '+(name.includes('HC3')?'HC3 covariance changes uncertainty estimates while retaining OLS coefficients. Student t inference uses n-2 degrees of freedom and is approximate in this short sample. ':'')+'\nFormula: '+f[i]));
 return {first,last,summary,changes,expected:changes?regression.first_differences:regression.levels};
}
const blocks=[regressionBlock(12,21,28),regressionBlock(84,92,98,true)];
s.getRange('C9:T9').merge();s.getRange('C9').values=[['Log levels regression']];s.getRange('C9').format.font={name:'Arial',size:11,bold:true,color:blue};
s.getRange('C81:T81').merge();s.getRange('C81').values=[['First differences regression']];s.getRange('C81').format.font={name:'Arial',size:11,bold:true,color:blue};
for(const p of [
 {title:'Interpreting the levels regression',row:26,height:215,body:'OLS estimates the historical association between ln(CFO) and ln(revenue). The slope is an elasticity: a 1% revenue difference corresponds approximately to the slope percent difference in fitted CFO. The high in-sample fit partly reflects the common upward time trend. Ten annual observations provide little evidence about future structural stability.\n\nThe fitted CFO series uses EXP(fitted ln CFO), so it is a conditional median rather than a bias-corrected conditional mean. It is a diagnostic comparison and does not replace the integrated forecast.'},
 {title:'What HC3 changes',row:38,height:225,body:'HC3 uses squared residuals divided by (1 minus leverage) squared in the covariance calculation. It changes standard errors, t statistics, p values and confidence intervals, while leaving OLS coefficients and fitted values unchanged. The visible helper columns show each observation\'s contribution.\n\nInference uses Student t with n minus two degrees of freedom. This is an explicit small-sample approximation. HC3 does not correct serial correlation, omitted variables, shared trends or structural breaks. Durbin Watson and leverage are descriptive diagnostics here, not proof that these problems are absent.'},
 {title:'Reading the first differences result',row:97,height:360,body:'The second regression uses annual log changes in CFO and revenue. It removes the level trend and leaves nine observations. The slope describes the association between growth changes, so it is not directly interchangeable with the levels elasticity.\n\nA wider confidence interval reflects the short sample and limited variation. Review the maximum leverage and the individual leverage column. A large leverage value makes one fiscal year influential and increases its HC3 weight. The 4/n reference is a descriptive screening threshold, not a statistical test.\n\nCompare coefficients and residuals in both specifications. Agreement does not establish causality. The original operating forecast retains its segment, expense, asset, lease, working-capital and payout assumptions.'},
 ]){
 s.shapes.add({geometry:'rect',anchor:{from:{row:p.row,col:7},extent:{widthPx:1360,heightPx:p.height}},fill:'#F1F5FA',line:{fill:'#B8C7D9',width:1}}).name=p.title;
 panels.push({sheet:s.name,name:p.title,title:p.title,body:p.body,row:p.row,width:1360,height:p.height});
}
for(const [title,r1,r2,row,ix,fmt] of [['Actual and fitted CFO (USD billions)','E11:E21','O11:O21',59,0,'money'],['Log-level residuals','L11:L21',null,59,1,'number'],['Actual and fitted CFO log changes','G83:G92','K83:K92',131,0,'number'],['First-difference leverage','M83:M92',null,131,1,'number']]){
 const ranges=[s.getRange(row===59?'C11:C21':'C83:C92'),s.getRange(r1)];if(r2)ranges.push(s.getRange(r2));
 const ch=s.charts.add('line',ranges);styleChart(ch,title,fmt);ch.setPosition(`${ix===0?'C':'K'}${row}`,`${ix===0?'J':'T'}${row+17}`);
}
s.getRange('C151:T151').merge();s.getRange('C151').values=[['Python reproduction in VSCode']];s.getRange('C151').format.font={name:'Arial',size:11,bold:true,color:blue};
const codeLines=code.replace(/\n$/,'').split('\n'),codeStart=160,codeEnd=codeStart+codeLines.length-1;
const instructions=[`Select C${codeStart}:C${codeEnd}, copy and paste into regression_analysis.py in VSCode.`,'Install once: python -m pip install numpy scipy matplotlib','Run: python regression_analysis.py --out regression_output','For edited data, export the input columns as year,revenue,cfo CSV and use --csv input.csv.','The supplied code embeds the original actuals. Excel edits do not rewrite this embedded Python data.','Method: statsmodels OLSResults.HC3_se documentation. The Python script implements the same matrix formula directly.'];
instructions.forEach((text,i)=>{s.getRange(`C${152+i}:T${152+i}`).merge();s.getRange(`C${152+i}`).values=[[text]];});
for(let i=0;i<codeLines.length;i++){const r=codeStart+i;s.getRange(`C${r}:T${r}`).merge();s.getRange(`C${r}`).values=[[codeLines[i]]];s.getRange(`C${r}:T${r}`).format.font={name:'Courier New',size:9,color:'#172B46'};s.getRange(`C${r}:T${r}`).format.rowHeight=15;}
note(s,'C2','Standalone historical regression. This worksheet never feeds the three-statement model. All coefficients and HC3 statistics are calculated with editable Excel formulas.');
note(s,'C151','The complete code below is supplied in src/regression_analysis.py. It runs outside Excel with NumPy and SciPy. Excel does not require a Python-in-Excel subscription.');
note(s,'C160','Beginning of the complete Python file. Copy the indicated single-column range so each cell becomes one Python source line.');
views.push({sheet:s.name,range:'C9:T76'},{sheet:s.name,range:'C81:T148'},{sheet:s.name,range:'C151:T185'});

w.recalculate();
const metricsMap={0:'n',1:'df',5:'slope',6:'intercept',7:'rss',8:'sst',9:'r_squared',10:'adjusted_r_squared',11:'residual_standard_error',18:'critical_t95',21:'durbin_watson',22:'max_leverage',23:'leverage_reference'};
const comparisons=[];
for(const b of blocks){
 const values=s.getRange(`E${b.summary}:E${b.summary+27}`).values.flat();
 const meanX=b.expected.observations.reduce((sum,o)=>sum+o.x,0)/b.expected.n;
 const meanY=b.expected.observations.reduce((sum,o)=>sum+o.y,0)/b.expected.n;
 const sxx=b.expected.observations.reduce((sum,o)=>sum+(o.x-meanX)**2,0);
 for(let k=0;k<values.length;k++){
   const expected=metricsMap[k]?b.expected[metricsMap[k]]:({2:meanX,3:meanY,4:sxx,12:b.expected.se_ols[1],13:b.expected.se_ols[0],14:b.expected.se_hc3[1],15:b.expected.se_hc3[0],16:b.expected.t_hc3[1],17:b.expected.p_hc3[1],19:b.expected.ci95_hc3[1][0],20:b.expected.ci95_hc3[1][1],24:b.expected.t_hc3[0],25:b.expected.p_hc3[0],26:b.expected.ci95_hc3[0][0],27:b.expected.ci95_hc3[0][1]})[k];
   const difference=Number(values[k])-expected;
   const tolerance=[17,25].includes(k)?1e-12+Math.abs(expected)*1e-7:1e-8*Math.max(1,Math.abs(expected));
   if(typeof values[k]!=='number'||!Number.isFinite(expected)||!Number.isFinite(difference)||Math.abs(difference)>tolerance)throw Error(`Regression mismatch ${b.changes?'changes':'levels'} ${metrics[k]}: ${values[k]} vs ${expected}`);
   comparisons.push({model:b.changes?'first_differences':'levels',metric:metrics[k],excel:values[k],python:expected,difference});
 }
}
for(const m of metadata){
 const sheet=w.worksheets.getItem(m.name);
 for(const [a,f] of Object.entries(formulasBefore[m.name]))if(sheet.getRange(a).formulas[0][0]!==`=${f}`)throw Error('Original formula changed: '+m.name+'!'+a);
}
const scan=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:60},summary:'enhanced workbook error scan'});
await fs.writeFile(`${qa}/formula_error_scan.ndjson`,scan.ndjson);
await fs.writeFile(`${project}/evidence/workbook_enhancements.json`,JSON.stringify({revision_date:'2026-10-10',note_counts:noteCounts,chart_count:chartChecks.length+2,added_charts:chartChecks,panels,regression_comparisons:comparisons,python_code_range:`Financial Modeling!C${codeStart}:C${codeEnd}`,original_formulas_preserved:true,views},null,2));
await (await SpreadsheetFile.exportXlsx(w)).save(output);
console.log(JSON.stringify({output,notes:Object.values(noteCounts).reduce((a,b)=>a+b,0),charts:chartChecks.length+2,regression_comparisons:comparisons.length,panels:panels.length,code_range:`C${codeStart}:C${codeEnd}`,error_scan:scan.ndjson}));
