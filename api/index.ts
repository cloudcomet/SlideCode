import express from "express";
import { convert } from "pptx-to-pdf";
import multer from "multer";
import { GoogleGenAI } from "@google/genai";

// Vercel Serverless Function Configuration
export const config = {
  api: {
    bodyParser: false, // Must be false so multer can parse the multipart/form-data stream
  },
};

const app = express();
app.use(express.json());

// Use multer to handle file uploads in memory
const upload = multer({ storage: multer.memoryStorage() });

// API to convert PPTX buffer to PDF
app.post("/api/convert-pdf", upload.single("pptx"), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: "No pptx file uploaded" });
      return;
    }
    
    const pptxBuffer = req.file.buffer;
    const pdfBuffer = await convert(pptxBuffer);
    
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "attachment; filename=presentation.pdf");
    res.send(pdfBuffer);
  } catch (error: any) {
    console.error("PDF conversion error:", error);
    res.status(500).json({ error: error.message || "Failed to convert PPTX to PDF" });
  }
});

// API to explain errors using Gemini
app.post("/api/explain-error", async (req, res) => {
  try {
    const { code, error } = req.body;
    
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) { 
       return res.status(500).json({ explanation: "Error: GEMINI_API_KEY is not configured." });
    }

    const ai = new GoogleGenAI({ apiKey });
    const prompt = `The user encountered an error while trying to execute the following PPTX generation code.\n\nCode:\n\`\`\`javascript\n${code}\n\`\`\`\n\nError:\n${error}\n\nPlease explain why this error occurred and how to fix it. Provide a brief, helpful explanation. If there is a simple fix to the code, provide it.`;
    
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash-lite",
      contents: prompt,
    });

    res.json({ explanation: response.text });
  } catch (error: any) {
    console.error("Gemini API error:", error);
    res.status(500).json({ error: "Failed to generate explanation" });
  }
});

export default app;
