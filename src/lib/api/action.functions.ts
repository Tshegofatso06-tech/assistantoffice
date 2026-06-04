import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";

const ACTIONS = [
  "email",
  "notes",
  "instructions",
  "trends",
  "scientific",
  "sales",
] as const;

const InputSchema = z.object({
  action: z.enum(ACTIONS),
  columns: z.array(z.string()).min(1),
  sampleRows: z.array(z.record(z.string(), z.any())).max(50),
  totalRows: z.number().int().nonnegative(),
  isChartLike: z.boolean().optional(),
  emptyCells: z.number().int().nonnegative().optional(),
});

const PROMPTS: Record<(typeof ACTIONS)[number], string> = {
  email:
    "Based on this data, write a professional email. If there's a Date column, mention date range. If there's Amount/Sales, include totals. If it's a list like farms/students, summarize the list. Make it 4-5 sentences with a Subject line.",
  notes:
    "Create meeting notes/bullet points from this data. 5 bullets max. Focus on key facts, totals, categories, or patterns found in the data.",
  instructions:
    "Write clear instructions for a team based on this data. If sales data, give targets. If farm data, give tasks per farm. If no clear action, give general instructions about the dataset. 4 steps max.",
  trends:
    "Find 3 trends or patterns in this data. If a Date column exists, check growth/decline over time and any seasonal pattern. If numeric columns exist, identify highest/lowest values, possible correlations, and outliers. Return as 3 short numbered bullets.",
  scientific:
    "This looks like scientific or experimental data. Provide a concise summary covering: 1) What appears to be measured (variables/units if visible), 2) Sample size / number of rows, 3) Key findings (means, ranges, dominant categories), 4) Any anomalies or outliers. Use short labelled sections.",
  sales:
    "Act as a senior sales analyst. Give 3 actionable insights from this sales data: best-performing product/region, slow period or weakest segment, and one concrete recommendation. Keep each insight to 1-2 sentences.",
};

export const generateAction = createServerFn({ method: "POST" })
  .inputValidator((d: unknown) => InputSchema.parse(d))
  .handler(async ({ data }) => {
    const apiKey = process.env.LOVABLE_API_KEY;
    if (!apiKey) throw new Error("Missing LOVABLE_API_KEY");

    const notes: string[] = [];
    if (data.isChartLike) {
      notes.push("This appears to be chart/table data with X (label) and Y (numeric) values.");
    }
    if (data.emptyCells && data.emptyCells > 0) {
      notes.push(`Ignore empty cells (${data.emptyCells} empty cells were detected and should not be counted).`);
    }

    const prompt = `Columns: ${JSON.stringify(data.columns)}.
Sample data (first ${data.sampleRows.length} of ${data.totalRows} rows): ${JSON.stringify(data.sampleRows)}.
${notes.join(" ")}

${PROMPTS[data.action]}
Do not assume the data is about sales unless the columns clearly indicate that. Adapt to whatever data is provided (lab results, surveys, sales, farms, inventory, etc.).`;

    const res = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Lovable-API-Key": apiKey,
      },
      body: JSON.stringify({
        model: "google/gemini-3-flash-preview",
        messages: [
          { role: "system", content: "You are a careful data analyst who writes clear, concise business and scientific content." },
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
