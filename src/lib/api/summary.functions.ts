import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const InputSchema = z.object({
  columns: z.array(z.string()).min(1),
  sampleRows: z.array(z.record(z.string(), z.any())).max(50),
  totalRows: z.number().int().nonnegative(),
  rangeLabel: z.string().optional(),
});

export const generateSummary = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => InputSchema.parse(d))
  .handler(async ({ data }) => {
    const apiKey = process.env.LOVABLE_API_KEY;
    if (!apiKey) throw new Error("Missing LOVABLE_API_KEY");

    const prompt = `You are a data analyst. Here are the columns: ${JSON.stringify(data.columns)}.
Here is sample data (first ${data.sampleRows.length} rows): ${JSON.stringify(data.sampleRows)}.
Total rows in dataset${data.rangeLabel ? ` (filtered to ${data.rangeLabel})` : ""}: ${data.totalRows}.

Write a 3-sentence summary of what this data is about. Include: 1) What type of data this is, 2) Total rows, 3) 1-2 key insights like highest value, most common category, or patterns.
Do not assume columns mean sales. Adapt to whatever data is given.`;

    const res = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Lovable-API-Key": apiKey,
      },
      body: JSON.stringify({
        model: "google/gemini-3-flash-preview",
        messages: [
          { role: "system", content: "You are a concise data analyst." },
          { role: "user", content: prompt },
        ],
      }),
    });

    if (res.status === 429) throw new Error("Rate limit reached. Please try again shortly.");
    if (res.status === 402) throw new Error("AI credits exhausted. Please add credits in workspace billing.");
    if (!res.ok) throw new Error(`AI request failed: ${res.status}`);

    const json = await res.json();
    const text: string = json?.choices?.[0]?.message?.content ?? "";
    return { summary: text.trim() };
  });
