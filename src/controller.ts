import { Request, Response } from "express";
import { GenerateRequestSchema } from "./schema";
import { callAi } from "./aiService";
import { generateDocx } from "./docxService";

export async function generatePocHandler(req: Request, res: Response): Promise<void> {
  try {
    // 1. Validate request
    const parseResult = GenerateRequestSchema.safeParse(req.body);
    if (!parseResult.success) {
      res.status(400).json({
        error: "Invalid request body",
        details: parseResult.error.flatten(),
      });
      return;
    }

    const input = parseResult.data;
    console.log(`Generating POC for: ${input.clientName}, workloads: ${input.workloads.join(", ")}`);

    // 2. Call AI to generate structured JSON
    const aiOutput = await callAi(input);
    console.log("AI output validated successfully");

    // 3. Generate DOCX from AI output
    const docxBuffer = await generateDocx(aiOutput, input);
    console.log("DOCX generated successfully");

    // 4. Return file
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const timestamp = `${pad(now.getDate())}-${pad(now.getMonth() + 1)}-${now.getFullYear()}-${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
    const filename = `${input.clientName.replace(/[^a-zA-Z0-9]/g, "_")}_POC_${timestamp}.docx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(docxBuffer);
  } catch (err) {
    console.error("Error generating POC:", err);
    res.status(500).json({
      error: "Failed to generate POC document",
      message: err instanceof Error ? err.message : "Unknown error",
    });
  }
}
