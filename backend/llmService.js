import Groq from 'groq-sdk';
import dotenv from 'dotenv';

dotenv.config();

// Standard models: llama-3.3-70b-versatile, llama3-8b-8192, mixtral-8x7b-32768
const MODEL_NAME = 'llama-3.3-70b-versatile';

/**
 * Sends a natural language prompt along with detected students and subjects to Groq LLM.
 * Returns the parsed and validated JSON object containing the updates array.
 * 
 * @param {string} promptText - The user prompt describing updates (e.g., "Ali got 78 in Math")
 * @param {string[]} students - List of students detected in the Excel sheet
 * @param {string[]} subjects - List of subjects detected in the Excel sheet
 * @returns {Promise<{updates: Array<{student: string, subject: string, marks: number}>}>}
 */
export async function extractUpdates(promptText, students, subjects) {
  const apiKey = process.env.GROQ_API_KEY;
  if (!apiKey || apiKey === 'your_groq_api_key_here') {
    throw new Error("GROQ_API_KEY is not set. Please add it to your backend/.env file.");
  }

  const groq = new Groq({ apiKey });

  const systemPrompt = `You are a student results extraction assistant. Your task is to extract student names, subject names, and marks from a user prompt.

Available students in the sheet: ${JSON.stringify(students)}
Available subjects in the sheet: ${JSON.stringify(subjects)}

Match the students and subjects in the user's description to the exact spelling/casing in the provided list.
For example, if the list contains "Mathematics" and the user says "Math", output "Mathematics".
If the list contains "Ali Ahmed" and the user says "Ali", output "Ali Ahmed".

Output ONLY valid JSON matching this schema:
{
  "updates": [
    {
      "student": "exact_student_name_from_list",
      "subject": "exact_subject_name_from_list",
      "marks": integer_marks_value
    }
  ]
}

Do not include any chat commentary or markdown formatting. Output ONLY the JSON object.`;

  try {
    const chatCompletion = await groq.chat.completions.create({
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: `Please parse this prompt and map to the student/subject list:\n\n"${promptText}"` }
      ],
      model: MODEL_NAME,
      response_format: { type: 'json_object' }
    });

    const responseContent = chatCompletion.choices[0]?.message?.content;
    if (!responseContent) {
      throw new Error("Empty response received from LLM");
    }

    return JSON.parse(responseContent);
  } catch (error) {
    console.error("LLM Extraction Error:", error);
    throw error;
  }
}
