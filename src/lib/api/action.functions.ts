import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const InputSchema = z.object({
  action: z.enum(["email", "notes", "instructions"]),
  columns: z.array(z.string()).min(1),
  sampleRows: z.array(z.record(z.string(), z.any())).max(20),
  totalRows: z.number().int().nonnegative(),
});

const PROMPTS: Record<string, string> = {
  email:
    "Based on this data, write a professional email. If there's a Date column, mention date range. If there's Amount/Sales, include totals. If it's a list like farms/students, summarize the list. Make it 4-5 sentences with a Subject line.",
  notes:
    "Create meeting notes/bullet points from this data. 5 bullets max. Focus on key facts, totals, categories, or patterns found in the data.",
  instructions:
    "Write clear instructions for a team based on this data. Example: if it's sales data, tell team targets. If it's farm data, give tasks per farm. If no clear action, create general instructions about the dataset. 4 steps max.",
};

export const generateAction = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => InputSchema.parse(d))
  .handler(async ({ data }) => {
    const apiKey = process.env.LOVABLE_API_KEY;
    if (!apiKey) throw new Error("Missing LOVABLE_API_KEY");

    const prompt = `Columns: ${JSON.stringify(data.columns)}.
Sample data (first ${data.sampleRows.length} of ${data.totalRows} rows): ${JSON.stringify(data.sampleRows)}.

${PROMPTS[data.action]}
Do not assume the data is about sales unless the columns indicate that. Adapt to whatever data is provided.`;

    const res = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Lovable-API-Key": apiKey,
      },
      body: JSON.stringify({
        model: "google/gemini-3-flash-preview",
        messages: [
          { role: "system", content: "You are a helpful office assistant who writes clear, concise business content." },
          { role: "user", content: prompt },
        ],
      }),
    });

    if (res.status === 429) throw new Error("Rate limit reached. Please try again shortly.");
    if (res.status === 402) throw new Error("AI credits exhausted. Please add credits in workspace billing.");
    if (!res.ok) throw new Error(`AI request failed: ${res.status}`);

    const json = await res.json();
    const text: string = json?.choices?.[0]?.message?.content ?? "";
    return { text: text.trim() };
  });
