const ExcelJS = require('exceljs');
async function inspect() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('THREE.xlsx');
  const ws = wb.worksheets[0];
  
  console.log('Worksheet dimensions:', ws.rowCount, 'rows', ws.columnCount, 'columns');
  
  // Log first 15 rows to understand header structure
  for (let i = 1; i <= Math.min(15, ws.rowCount); i++) {
    const row = ws.getRow(i);
    const values = [];
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      if (colNum <= 30) { // Limit to 30 columns
        values.push(typeof cell.value === 'object' && cell.value !== null 
          ? (cell.value.result !== undefined ? '{Formula}' : (cell.value.richText ? cell.value.richText.map(rt => rt.text).join('') : '{Obj}')) 
          : cell.value);
      }
    });
    console.log(`Row ${i}:`, values);
  }
}
inspect().catch(console.error);
