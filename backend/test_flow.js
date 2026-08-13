import fs from 'fs';
import path from 'path';
import fetch from 'node-fetch';
import FormData from 'form-data';
import ExcelJS from 'exceljs';

async function runTest() {
  console.log("Starting server simulation...");
  
  // 1. Upload
  const formData = new FormData();
  formData.append('file', fs.createReadStream('student_results.xlsx'));
  
  let res = await fetch('http://localhost:5000/api/analyze', {
    method: 'POST',
    body: formData
  });
  
  if (!res.ok) throw new Error("Analyze failed: " + await res.text());
  let data = await res.json();
  const sessionId = data.sessionId;
  console.log("Session created:", sessionId);

  // 2. Apply Update
  const updates = [
    { student: 'Ali', subject: 'Math', marks: 99 }
  ];
  
  res = await fetch('http://localhost:5000/api/apply-update', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sessionId, updates })
  });
  
  if (!res.ok) throw new Error("Apply failed: " + await res.text());
  console.log("Update applied.");

  // 3. Download
  res = await fetch(`http://localhost:5000/api/download/${sessionId}`);
  if (!res.ok) throw new Error("Download failed: " + await res.text());
  
  const arrayBuffer = await res.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);
  
  // 4. Verify contents
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  
  // Ali is on row 2, Math is col 2 (based on student_results.xlsx)
  const val = worksheet.getRow(2).getCell(2).value;
  console.log("Ali Math Marks in Downloaded File:", val);

  if (val === 99) {
    console.log("SUCCESS: Server maintained state perfectly.");
  } else {
    console.log("FAILED: Downloaded file did not contain the update.");
  }
}

runTest().catch(console.error);
