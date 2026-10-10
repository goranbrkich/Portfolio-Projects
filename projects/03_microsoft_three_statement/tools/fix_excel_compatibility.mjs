import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const [mode, input, outputDir, specPath] = process.argv.slice(2);
if (!['inspect', 'notes', 'fix'].includes(mode) || !input || !outputDir) {
  throw new Error('Usage: node fix_excel_compatibility.mjs inspect|fix input.xlsx output-dir [spec.json]');
}
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
if (mode === 'notes') {
  console.log(workbook.help('workbook.notes', { include: 'index,examples,notes', maxChars: 3600 }).ndjson);
  process.exit(0);
}
if (mode === 'inspect') {
  console.log((await workbook.inspect({ kind: 'sheet', include: 'id,name', maxChars: 3500 })).ndjson);
  console.log(workbook.help('fx.FIXED', { include: 'index,examples,notes', maxChars: 1800 }).ndjson);
  for (const [sheetName, range, name] of [
    ['Income Statement', 'C7:S35', 'before_income_statement'],
    ['Working Capital', 'C24:S39', 'before_working_capital'],
    ['Income Statement', 'C53:S56', 'before_observations'],
    ['Financial Modeling', 'C11:T21', 'before_regression_table'],
  ]) {
    const preview = await workbook.render({ sheetName, range, scale: 1, format: 'png' });
    await fs.writeFile(path.join(outputDir, `${name}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  console.log('Baseline views saved.');
  process.exit(0);
}

// The edit plan is read-only XML evidence from the supplied workbook.
const spec = JSON.parse(await fs.readFile(specPath, 'utf8'));
for (const edit of spec.observationEdits) {
  const sheet = workbook.worksheets.getItem(edit.sheet);
  const oldFormula = sheet.getRange(edit.cell).formulas[0][0];
  if (oldFormula !== edit.oldFormula) throw new Error(`Source formula mismatch: ${edit.sheet}!${edit.cell}`);
  sheet.getRange(edit.cell).formulas = [[edit.newFormula]];
  workbook.notes.add({
    id: `${edit.sheet}:${edit.cell}`,
    target: { cell: { sheetName: edit.sheet, sheetId: sheet.sheetId, address: edit.cell } },
    authorId: '', createdAt: '', body: { plainText: edit.newNote },
  });
}
for (const rowPlan of spec.rowPlans) {
  const sheet = workbook.worksheets.getItem(rowPlan.sheet);
  for (const range of rowPlan.ranges) sheet.getRange(range).format.rowHeight = spec.tableRowHeightPoints;
}
workbook.recalculate();
const cases = [];
for (const [number, name] of [[1, 'Base'], [2, 'Upside'], [3, 'Downside']]) {
  workbook.worksheets.getItem('Assumptions').getRange('D4').values = [[number]];
  workbook.recalculate();
  const observations = spec.observationEdits.map(edit => ({
    sheet: edit.sheet, cell: edit.cell,
    formula: workbook.worksheets.getItem(edit.sheet).getRange(edit.cell).formulas[0][0],
    value: workbook.worksheets.getItem(edit.sheet).getRange(edit.cell).values[0][0],
  }));
  const failures = observations.filter(o => typeof o.value !== 'string' || /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)/.test(o.value));
  if (failures.length) throw new Error(`Observation recalculation failed for ${name}: ${JSON.stringify(failures)}`);
  const checks = ['8', '9', '10', '11', '12', '16', '17', '18', '19', '20', '21'].flatMap(r =>
    ['O', 'P', 'Q', 'R', 'S'].map(c => ({ cell: `${c}${r}`, value: workbook.worksheets.getItem('Audit').getRange(`${c}${r}`).values[0][0] })));
  const checkFailures = checks.filter(c => typeof c.value !== 'number' || Math.abs(c.value) > 1e-6);
  if (checkFailures.length) throw new Error(`Statement checks failed for ${name}: ${JSON.stringify(checkFailures)}`);
  cases.push({ number, name, observations, auditChecks: checks });
}
workbook.worksheets.getItem('Assumptions').getRange('D4').values = [[spec.originalCase]];
workbook.recalculate();
const errorScan = await workbook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 100 }, maxChars: 2000, summary: 'Final recalculation error scan' });
console.log(errorScan.ndjson);
await fs.writeFile(path.join(outputDir, 'artifact_recalculation.json'), JSON.stringify({ cases, errorScan: errorScan.ndjson, rowPlans: spec.rowPlans, tableRowHeightPoints: spec.tableRowHeightPoints, originalCase: spec.originalCase }, null, 2) + '\n');
const result = await SpreadsheetFile.exportXlsx(workbook);
await result.save(path.join(outputDir, 'Microsoft_Three_Statement_Model.xlsx'));
console.log(`Exported ${spec.observationEdits.length} observation fixes and ${spec.rowPlans.length} row plans.`);
for (const [sheetName, range, name] of spec.previewRanges) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: 'png' });
  await fs.writeFile(path.join(outputDir, `${name}.png`), new Uint8Array(await preview.arrayBuffer()));
}
console.log('Changed views saved.');
