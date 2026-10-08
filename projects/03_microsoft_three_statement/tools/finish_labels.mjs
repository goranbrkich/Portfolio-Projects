// Apply business labels while preserving the completed financial model.
// Requires the Codex primary runtime with @oai/artifact-tool.
import fs from 'node:fs/promises';
import {FileBlob, SpreadsheetFile} from '@oai/artifact-tool';

const [project, outputPath, qaPath] = process.argv.slice(2);
const layout = JSON.parse(await fs.readFile(`${project}/evidence/workbook_layout.json`, 'utf8'));
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(`${project}/deliverables/Microsoft_Three_Statement_Model.xlsx`));
const labels = {
  IS: {
    revenue:'Revenue', cogs:'Cost of revenue', gross_profit:'Gross profit', rd:'Research and development', sales_marketing:'Sales and marketing', ga:'General and administrative', other_opex:'Impairment, integration and restructuring', operating_income:'Operating income', other_income:'Other income and expense, net', pretax_income:'Income before income taxes', income_tax:'Income tax expense', net_income:'Net income', shares_basic:'Basic weighted average shares', shares_diluted:'Diluted weighted average shares', eps_diluted:'Diluted earnings per share', eps_basic:'Basic earnings per share'
  },
  BS: {
    cash:'Cash and cash equivalents', st_investments:'Short-term investments', liquid_assets:'Cash and short-term investments', receivables:'Accounts receivable, net', inventory:'Inventories', other_current_assets:'Other current assets', current_assets:'Total current assets', ppe:'Property and equipment, net', op_rou:'Operating lease right-of-use assets', equity_investments:'Equity and other investments', goodwill:'Goodwill', intangibles:'Intangible assets, net', other_lt_assets:'Other long-term assets', total_assets:'Total assets', payables:'Accounts payable', st_debt:'Short-term debt', current_debt:'Current portion of long-term debt', compensation:'Accrued compensation', current_tax:'Short-term income taxes', current_deferred_rev:'Short-term unearned revenue', securities_lending:'Securities lending payable', other_current_liab:'Other current liabilities', current_liabilities:'Total current liabilities', long_debt:'Long-term debt', long_tax:'Long-term income taxes', long_deferred_rev:'Long-term unearned revenue', deferred_tax_liab:'Deferred income tax liabilities', op_lease_long:'Noncurrent operating lease liabilities', other_lt_liab:'Other long-term liabilities', total_liabilities:'Total liabilities', paid_in_capital:'Common stock and paid-in capital', retained_earnings:'Retained earnings', aoci:'Accumulated other comprehensive income', total_equity:'Total stockholders equity', total_le:'Total liabilities and stockholders equity'
  },
  CF: {
    net_income:'Net income', sbc:'Stock-based compensation', investment_adjustment:'Investment and derivative loss or gain adjustment', deferred_taxes:'Deferred income taxes', ar_change:'Change in accounts receivable', inventory_change:'Change in inventories', other_ca_change:'Change in other current assets', other_lta_change:'Change in other long-term assets', ap_change:'Change in operating accounts payable', deferred_rev_change:'Change in unearned revenue', tax_change:'Change in income taxes payable', other_cl_change:'Change in other current operating liabilities', other_ltl_change:'Change in other long-term operating liabilities', other_operating_adjustments:'Other operating adjustments', cfo:'Net cash from operations', capex:'Cash capital expenditure', acquisitions:'Acquisitions and purchases of other assets', investments_purchased:'Purchases of investments', investments_matured:'Maturities of investments', investments_sold:'Sales of investments', other_investing:'Other investing and securities lending', cfi:'Net cash from investing', debt_issued:'New debt and liquidity borrowing', net_st_debt:'Net proceeds or repayments of short-term debt', debt_repaid:'Repayments of debt', stock_issued:'Cash proceeds from share issuance', buybacks:'Cash share repurchases', dividends:'Cash dividends', other_financing:'Other financing excluding finance lease principal', cff:'Net cash from financing', fx_cash:'Effect of exchange rates on cash', cash_change:'Net change in cash', opening_cash:'Opening cash and cash equivalents', closing_cash:'Closing cash and cash equivalents'
  }
};
for (const [name, section] of [['Income Statement','IS'],['Balance Sheet','BS'],['Cash Flow','CF']]) {
  const sheet = wb.worksheets.getItem(name);
  for (const [metric, label] of Object.entries(labels[section])) sheet.getRange(`C${layout[section][metric]}`).values = [[label]];
}
wb.recalculate();
const scan = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:40},summary:'final formula error scan'});
await fs.mkdir(qaPath,{recursive:true});
await fs.writeFile(`${qaPath}/formula_error_scan.ndjson`,scan.ndjson);
for (const name of ['Income Statement','Balance Sheet','Cash Flow']) {
  const image = await wb.render({sheetName:name,range:'C2:S56',scale:1.3,format:'png'});
  await fs.writeFile(`${qaPath}/${name.replaceAll(' ','_')}.png`,new Uint8Array(await image.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(JSON.stringify({output:outputPath,errorScan:scan.ndjson}));
