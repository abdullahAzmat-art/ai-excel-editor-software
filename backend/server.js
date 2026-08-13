import express from 'express';
import cors from 'cors';
import multer from 'multer';
import dotenv from 'dotenv';
import crypto from 'crypto';
import { analyzeExcel, applyMarksOnly } from './excelService.js';
import { extractUpdates } from './llmService.js';
import { llmResponseSchema } from './validation.js';

dotenv.config();

const app = express();
const port = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ 
  storage,
  limits: { fileSize: 20 * 1024 * 1024 } // 20MB for the large THREE.xlsx
});

// In-memory session store
const sessions = {};

/**
 * 1. Upload & Analyze Excel template
 * POST /api/analyze
 */
app.post('/api/analyze', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ error: "Please upload an Excel file." });

    const { students, subjects } = await analyzeExcel(file.buffer);
    
    const sessionId = crypto.randomUUID();
    sessions[sessionId] = {
      buffer: file.buffer,
      originalName: file.originalname,
      students,
      subjects,
      updatedStudents: new Set()
    };

    console.log(`Session ${sessionId}: Found ${students.length} students, ${subjects.length} subjects`);
    res.json({ success: true, sessionId, students, subjects });
  } catch (error) {
    console.error("Analysis Error:", error);
    res.status(500).json({ error: error.message || "Failed to analyze Excel file." });
  }
});

/**
 * 2. Process User Prompt using Groq LLM
 * POST /api/process-prompt
 */
app.post('/api/process-prompt', async (req, res) => {
  try {
    const { sessionId, prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: "Prompt is required." });
    
    const session = sessions[sessionId];
    if (!session) return res.status(404).json({ error: "Session not found or expired." });

    const llmJson = await extractUpdates(prompt, session.students, session.subjects);

    const validationResult = llmResponseSchema.safeParse(llmJson);
    if (!validationResult.success) {
      return res.status(422).json({ 
        error: "Extracted updates did not match expected structure.",
        details: validationResult.error.format()
      });
    }

    res.json({ success: true, updates: validationResult.data.updates });
  } catch (error) {
    console.error("Process Prompt Error:", error);
    res.status(500).json({ error: error.message || "Failed to process prompt." });
  }
});

/**
 * 3. Apply updates to the session's workbook in memory
 * POST /api/apply-update
 */
app.post('/api/apply-update', async (req, res) => {
  try {
    const { sessionId, updates } = req.body;
    const session = sessions[sessionId];
    if (!session) return res.status(404).json({ error: "Session not found or expired." });
    if (!updates || !Array.isArray(updates)) return res.status(400).json({ error: "Valid updates array is required." });

    // Apply marks into the correct cells in the template
    const updatedBuffer = await applyMarksOnly(session.buffer, updates);
    session.buffer = updatedBuffer; // Persist updated workbook in memory

    updates.forEach(u => {
      if (u.student) session.updatedStudents.add(u.student);
    });

    console.log(`Session ${sessionId}: Applied ${updates.length} updates. ${session.updatedStudents.size}/${session.students.length} students updated.`);

    res.json({ 
      success: true, 
      updatedCount: session.updatedStudents.size,
      totalCount: session.students.length,
      updatedStudents: [...session.updatedStudents]
    });
  } catch (error) {
    console.error("Apply Update Error:", error);
    res.status(500).json({ error: error.message || "Failed to apply updates." });
  }
});

/**
 * 4. Download the updated Excel file
 * GET /api/download/:sessionId
 */
app.get('/api/download/:sessionId', (req, res) => {
  const session = sessions[req.params.sessionId];
  if (!session) return res.status(404).send("Session not found or expired.");

  const nameBase = session.originalName.substring(0, session.originalName.lastIndexOf('.'));
  
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${nameBase}_updated.xlsx"`);
  res.send(Buffer.from(session.buffer));
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
