import fs from 'fs';
import ExcelJS from 'exceljs';
import { analyzeExcel, applyMarksOnly } from './excelService.js';

async function verify() {
  console.log('--- Loading THREE.xlsx ---');
  let buffer = fs.readFileSync('THREE.xlsx');

  const { students, subjects } = await analyzeExcel(buffer);
  console.log(`Found ${students.length} students.`);
  console.log('First 5 students:', students.slice(0, 5));
  console.log(`Found ${subjects.length} subjects:`, subjects);

  if (students.length === 0) throw new Error('No students found! Parser needs fixing.');
  if (subjects.length === 0) throw new Error('No subjects found! Parser needs fixing.');

  // Test writing marks for the first two students
  const student1 = students[0];
  const student2 = students[1];
  const subject1 = subjects[0]; // e.g. "English - Lit + Language"
  const subject2 = subjects[2]; // e.g. "Mathematics"

  console.log(`\n--- Applying marks: ${student1} → ${subject1} = 88, ${student2} → ${subject2} = 72 ---`);

  const updates = [
    { student: student1, subject: subject1, marks: 88 },
    { student: student2, subject: subject2, marks: 72 },
  ];

  buffer = await applyMarksOnly(buffer, updates);
  fs.writeFileSync('THREE_updated.xlsx', buffer);
  console.log('Written to THREE_updated.xlsx');

  // Verify by reading back
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];

  // Student 1 is at col 1, subjects start at row 11, "Obtained Marks" is col 1+3=4
  const val1 = ws.getRow(11).getCell(4).value; // Eshal Yousaf, English Obtained Marks
  console.log(`\n✅ ${student1} - ${subject1} obtained marks cell value:`, val1);

  // Student 2 is at col 8, subjects start at row 11, "Obtained Marks" is col 8+3=11
  const val2 = ws.getRow(13).getCell(11).value; // 2nd student, Mathematics row (row 13)
  console.log(`✅ ${student2} - ${subject2} obtained marks cell value:`, val2);

  if (val1 === 88 && val2 === 72) {
    console.log('\n🎉 VERIFICATION PASSED — Marks written correctly to THREE.xlsx template!');
  } else {
    console.log('\n❌ VERIFICATION FAILED — Values do not match expected marks.');
    console.log('Expected:', { val1: 88, val2: 72 }, 'Got:', { val1, val2 });
  }
}

verify().catch(err => {
  console.error('❌ Error:', err.message);
  process.exit(1);
});
