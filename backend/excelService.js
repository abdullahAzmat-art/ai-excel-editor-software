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
