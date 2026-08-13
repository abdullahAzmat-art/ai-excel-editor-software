import ExcelJS from 'exceljs';

/**
 * Parses a custom side-by-side template (like THREE.xlsx).
 * Maps every student block by finding "Name:" cells.
 */
function parseWorksheetMetadata(worksheet) {
  const blocks = [];

  // 1. Find all "Name:" cells to identify student blocks
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      const val = cell.value ? cell.value.toString().trim() : '';
      if (val === 'Name:') {
        // Student name is one column to the right
        const studentNameCell = row.getCell(colNumber + 1);
        let studentName = studentNameCell.value ? studentNameCell.value.toString().trim() : 'Unknown';

        // Subjects start 5 rows down (Row 11 if Name is Row 6)
        const subjectStartRow = rowNumber + 5; 

        const subjects = [];
        let currentRow = subjectStartRow;
        
        while (currentRow <= worksheet.rowCount) {
          const subCell = worksheet.getRow(currentRow).getCell(colNumber + 1);
          const subVal = subCell.value ? subCell.value.toString().trim() : '';
          
          if (!subVal || subVal.includes('Total Marks:')) {
            break; // End of subject list for this student block
          }
          
          subjects.push({
            name: subVal,
            row: currentRow,
            marksCol: colNumber + 3 // "Obtained Marks" is 3 cols to the right of "Name:" column
          });
          
          currentRow++;
        }

        blocks.push({
          studentName,
          nameRow: rowNumber,
          nameCol: colNumber,
          subjects
        });
      }
    });
  });

  // Consolidate global subject list from the first block to give to the LLM
  const globalSubjects = blocks.length > 0 ? blocks[0].subjects.map(s => s.name) : [];
  const studentsList = blocks.map(b => b.studentName);

  return { blocks, students: studentsList, subjects: globalSubjects };
}

export async function analyzeExcel(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("Excel sheet is empty or invalid.");

  const { students, subjects } = parseWorksheetMetadata(worksheet);

  return { students, subjects };
}

export async function applyMarksOnly(buffer, updates) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("Excel sheet is empty or invalid.");

  const { blocks } = parseWorksheetMetadata(worksheet);

  for (const update of updates) {
    const targetBlock = blocks.find(b => b.studentName.toLowerCase() === update.student.toLowerCase());
    if (targetBlock) {
      const targetSubject = targetBlock.subjects.find(s => s.name.toLowerCase() === update.subject.toLowerCase());
      if (targetSubject) {
        const cell = worksheet.getRow(targetSubject.row).getCell(targetSubject.marksCol);
        cell.value = Number(update.marks);
      }
    }
  }

  return await workbook.xlsx.writeBuffer();
}

export async function calculateResults(buffer) {
  // Stubbed out: The template has native Excel formulas!
  // We simply return the buffer unchanged to prevent overwriting those native formulas.
  return { updatedBuffer: buffer, resultsPreview: [] };
}

/**
 * Converts a number to an ordinal string (1 → "1st", 2 → "2nd", etc.)
 */
function toOrdinal(n) {
  const s = ['th', 'st', 'nd', 'rd'];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

/**
 * Calculates positions for all students by summing their obtained marks,
 * ranking them with dense competition ranking, and writing the ordinal
 * position string into both the "Current Position" and "Sum up Position"
 * cells of each student's block.
 *
 * Returns the updated buffer and a preview array for the UI.
 */
export async function calculatePositions(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("Excel sheet is empty or invalid.");

  const { blocks } = parseWorksheetMetadata(worksheet);

  // 1. Sum obtained marks for every student
  const studentTotals = blocks.map(block => {
    let total = 0;
    for (const subject of block.subjects) {
      const val = worksheet.getRow(subject.row).getCell(subject.marksCol).value;
      if (val !== null && val !== undefined && !isNaN(Number(val))) {
        total += Number(val);
      }
    }
    return { ...block, total };
  });

  // 2. Sort descending by total (dense ranking)
  studentTotals.sort((a, b) => b.total - a.total);

  let rank = 1;
  const preview = [];

  for (let i = 0; i < studentTotals.length; i++) {
    if (i > 0 && studentTotals[i].total < studentTotals[i - 1].total) {
      rank = i + 1;
    }

    const ordinal = toOrdinal(rank);
    const { nameCol, nameRow, studentName, total } = studentTotals[i];

    // Position value column is nameCol + 5
    const posCol = nameCol + 5;

    // Current Position → rows nameRow-1 and nameRow (rows 5 & 6)
    worksheet.getRow(nameRow - 1).getCell(posCol).value = ordinal;
    worksheet.getRow(nameRow).getCell(posCol).value = ordinal;

    // Sum up Position → rows nameRow+1 and nameRow+2 (rows 7 & 8)
    worksheet.getRow(nameRow + 1).getCell(posCol).value = ordinal;
    worksheet.getRow(nameRow + 2).getCell(posCol).value = ordinal;

    preview.push({ student: studentName, total, position: rank, ordinal });
  }

  const updatedBuffer = await workbook.xlsx.writeBuffer();
  return { updatedBuffer, preview };
}

