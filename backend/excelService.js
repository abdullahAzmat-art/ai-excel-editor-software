import ExcelJS from 'exceljs';

/**
 * Helper to analyze the workbook, finding header row, student column index, and subject columns.
 * 
 * @param {ExcelJS.Worksheet} worksheet 
 * @returns {{ headerRowNumber: number, nameColIndex: number, subjects: Array<{name: string, colIndex: number}>, students: Array<{name: string, rowNumber: number}> }}
 */
function parseWorksheetMetadata(worksheet) {
  let headerRowNumber = 1;
  let nameColIndex = 1;
  let foundHeader = false;

  // 1. Scan sheet for cell containing "name", "student", or "roll" to determine header row
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (foundHeader) return;
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (foundHeader) return;
      const val = cell.value ? cell.value.toString().toLowerCase().trim() : '';
      if (val.includes('name') || val.includes('student') || val.includes('roll')) {
        headerRowNumber = rowNumber;
        nameColIndex = colNumber;
        foundHeader = true;
      }
    });
  });

  const headerRow = worksheet.getRow(headerRowNumber);
  const subjects = [];
  const students = [];

  // 2. Identify subjects from header row (any column that is not name, calculations, or metadata)
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    if (colNumber === nameColIndex) return;
    const cellValue = cell.value ? cell.value.toString().trim() : '';
    if (!cellValue) return;

    const lowerValue = cellValue.toLowerCase();
    const isCalculationOrMetadataCol = 
      lowerValue.includes('total') || 
      lowerValue.includes('percent') || 
      lowerValue.includes('position') || 
      lowerValue.includes('rank') || 
      lowerValue.includes('avg') || 
      lowerValue.includes('average') || 
      lowerValue.includes('grade') || 
      lowerValue.includes('remark') ||
      lowerValue.includes('roll') ||
      lowerValue.includes('id');
    
    if (!isCalculationOrMetadataCol) {
      subjects.push({ name: cellValue, colIndex: colNumber });
    }
  });

  // 3. Extract student names from rows below header row
  for (let r = headerRowNumber + 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const cellValue = row.getCell(nameColIndex).value;
    if (cellValue !== null && cellValue !== undefined) {
      const name = cellValue.toString().trim();
      if (name) {
        students.push({ name, rowNumber: r });
      }
    }
  }

  return { headerRowNumber, nameColIndex, subjects, students };
}

/**
 * Reads Excel workbook and returns the list of detected student names and subjects.
 * 
 * @param {Buffer} buffer - Excel file buffer
 * @returns {Promise<{students: string[], subjects: string[]}>}
 */
export async function analyzeExcel(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("Excel sheet is empty or invalid.");
  }

  const { subjects, students } = parseWorksheetMetadata(worksheet);

  return {
    students: students.map(s => s.name),
    subjects: subjects.map(sub => sub.name)
  };
}

/**
 * Updates cells with new marks, calculates Total, Percentage, and Position, preserves styling, and returns the updated buffer.
 * 
 * @param {Buffer} buffer - Excel file buffer
 * @param {Array<{student: string, subject: string, marks: number}>} updates - validated updates array
 * @returns {Promise<Buffer>} updated Excel file buffer
 */
export async function applyUpdates(buffer, updates) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("Excel sheet is empty or invalid.");
  }

  const { headerRowNumber, nameColIndex, subjects, students } = parseWorksheetMetadata(worksheet);

  // 1. Locate/Create Total, Percentage, Position columns
  let totalColIndex = -1;
  let percentageColIndex = -1;
  let positionColIndex = -1;
  let maxColIndex = 0;

  const headerRow = worksheet.getRow(headerRowNumber);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    if (colNumber > maxColIndex) maxColIndex = colNumber;
    const val = cell.value ? cell.value.toString().toLowerCase().trim() : '';
    if (val.includes('total')) totalColIndex = colNumber;
    else if (val.includes('percent')) percentageColIndex = colNumber;
    else if (val.includes('position') || val.includes('rank')) positionColIndex = colNumber;
  });

  // Append new columns if they do not exist
  if (totalColIndex === -1) {
    totalColIndex = ++maxColIndex;
    const cell = headerRow.getCell(totalColIndex);
    cell.value = "Total";
    cell.font = { bold: true, name: 'Calibri', size: 11 };
    cell.alignment = { horizontal: 'center' };
  }
  if (percentageColIndex === -1) {
    percentageColIndex = ++maxColIndex;
    const cell = headerRow.getCell(percentageColIndex);
    cell.value = "Percentage";
    cell.font = { bold: true, name: 'Calibri', size: 11 };
    cell.alignment = { horizontal: 'center' };
  }
  if (positionColIndex === -1) {
    positionColIndex = ++maxColIndex;
    const cell = headerRow.getCell(positionColIndex);
    cell.value = "Position";
    cell.font = { bold: true, name: 'Calibri', size: 11 };
    cell.alignment = { horizontal: 'center' };
  }

  // 2. Put marks into the cells
  for (const update of updates) {
    const matchedStudent = students.find(s => s.name.toLowerCase() === update.student.toLowerCase());
    const matchedSubject = subjects.find(sub => sub.name.toLowerCase() === update.subject.toLowerCase());

    if (matchedStudent && matchedSubject) {
      const row = worksheet.getRow(matchedStudent.rowNumber);
      const cell = row.getCell(matchedSubject.colIndex);
      cell.value = Number(update.marks);
    }
  }

  // 3. Calculate Total and Percentage for each student row
  for (const student of students) {
    const row = worksheet.getRow(student.rowNumber);
    let total = 0;
    
    for (const sub of subjects) {
      const val = row.getCell(sub.colIndex).value;
      if (val !== null && val !== undefined && !isNaN(Number(val))) {
        total += Number(val);
      }
    }

    // Write Total
    const totalCell = row.getCell(totalColIndex);
    totalCell.value = total;

    // Write Percentage (assuming each subject is out of 100)
    const pctCell = row.getCell(percentageColIndex);
    const percentage = subjects.length > 0 ? (total / subjects.length) : 0;
    pctCell.value = Number(percentage.toFixed(2));
  }

  // 4. Calculate Position (competition/dense ranking: equal totals get equal ranks, skip rank numbers appropriately)
  const sortedStudents = students.map(s => {
    const row = worksheet.getRow(s.rowNumber);
    const totalVal = Number(row.getCell(totalColIndex).value) || 0;
    return { ...s, total: totalVal };
  }).sort((a, b) => b.total - a.total);

  let currentRank = 1;
  for (let i = 0; i < sortedStudents.length; i++) {
    if (i > 0 && sortedStudents[i].total < sortedStudents[i - 1].total) {
      currentRank = i + 1; // standard competition rank
    }
    const row = worksheet.getRow(sortedStudents[i].rowNumber);
    const posCell = row.getCell(positionColIndex);
    posCell.value = currentRank;
    posCell.alignment = { horizontal: 'center' };
  }

  // Return the updated workbook binary buffer
  return await workbook.xlsx.writeBuffer();
}
