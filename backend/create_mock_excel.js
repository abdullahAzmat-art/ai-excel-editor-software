import ExcelJS from 'exceljs';
import path from 'path';

async function create() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Results');
  
  // Headers
  worksheet.addRow(['Student Name', 'Math', 'Biology', 'Urdu', 'Total', 'Percentage', 'Position']);
  
  // Apply bold formatting to header row
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, name: 'Calibri', size: 11 };
  
  // Student rows (leave Total, Percentage, and Position empty)
  worksheet.addRow(['Ali', 50, 60, 45, '', '', '']);
  worksheet.addRow(['Ahmed', 70, 65, 80, '', '', '']);
  worksheet.addRow(['Sara', 85, 90, 88, '', '', '']);
  worksheet.addRow(['Zain', 40, 50, 42, '', '', '']);

  const outputPath = path.join(process.cwd(), 'student_results.xlsx');
  await workbook.xlsx.writeFile(outputPath);
  console.log('Mock Excel file created successfully at:', outputPath);
}

create().catch(console.error);
