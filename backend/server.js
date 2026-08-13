import express from 'express';
import cors from 'cors';
import multer from 'multer';
import dotenv from 'dotenv';
import { analyzeExcel, applyUpdates } from './excelService.js';
import { extractUpdates } from './llmService.js';
import { llmResponseSchema } from './validation.js';

dotenv.config();

const app = express();
const port = process.env.PORT || 5000;

// Enable CORS and body parsing
app.use(cors());
app.use(express.json());

// Set up file uploading using multer (store in memory buffer)
const storage = multer.memoryStorage();
const upload = multer({ 
  storage,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

/**
 * 1. Upload & Analyze Excel metadata
 * POST /api/analyze
 */
app.post('/api/analyze', upload.single('file'), async (resOrReq, res) => {
  try {
    const file = resOrReq.file;
    if (!file) {
      return res.status(400).json({ error: "Please upload an Excel file." });
    }

    const { students, subjects } = await analyzeExcel(file.buffer);
    res.json({ success: true, students, subjects });
  } catch (error) {
    console.error("Analysis Endpoint Error:", error);
    res.status(500).json({ error: error.message || "Failed to analyze Excel file." });
  }
});

/**
 * 2. Process User Prompt using Groq LLM
 * POST /api/process-prompt
 */
app.post('/api/process-prompt', async (req, res) => {
  try {
    const { prompt, students, subjects } = req.body;
    if (!prompt) {
      return res.status(400).json({ error: "Prompt is required." });
    }
    if (!students || !subjects) {
      return res.status(400).json({ error: "Student list and subject list are required." });
    }

    // Call llmService
    const llmJson = await extractUpdates(prompt, students, subjects);

    // Validate LLM output with Zod
    const validationResult = llmResponseSchema.safeParse(llmJson);
    if (!validationResult.success) {
      console.warn("LLM Output failed Zod validation, raw output:", llmJson);
      return res.status(422).json({ 
        error: "Extracted updates did not match expected structure.",
        details: validationResult.error.format()
      });
    }

    res.json({ success: true, updates: validationResult.data.updates });
  } catch (error) {
    console.error("Process Prompt Endpoint Error:", error);
    res.status(500).json({ error: error.message || "Failed to process prompt." });
  }
});

/**
 * 3. Confirm updates and download updated Excel file
 * POST /api/confirm
 */
app.post('/api/confirm', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    const updatesStr = req.body.updates;

    if (!file) {
      return res.status(400).json({ error: "Please upload the original Excel file." });
    }
    if (!updatesStr) {
      return res.status(400).json({ error: "Updates list is required." });
    }

    let updates;
    try {
      updates = JSON.parse(updatesStr);
    } catch (e) {
      return res.status(400).json({ error: "Updates must be a valid JSON array string." });
    }

    // Apply updates and calculate totals/percentages/positions
    const updatedBuffer = await applyUpdates(file.buffer, updates);

    // Stream updated Excel file back to client
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=updated_result.xlsx`);
    res.send(updatedBuffer);
  } catch (error) {
    console.error("Confirm/Download Endpoint Error:", error);
    res.status(500).json({ error: error.message || "Failed to apply updates and generate file." });
  }
});

// Start Express Server
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
