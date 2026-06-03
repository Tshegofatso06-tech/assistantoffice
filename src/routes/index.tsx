import { createFileRoute } from "@tanstack/react-router";
import { useRef, useState } from "react";
import { Upload, FileSpreadsheet, ListChecks, ClipboardList, Mail, Clock, Sparkles } from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "OfficeAssist — Work smarter, not longer" },
      { name: "description", content: "Upload spreadsheets and generate instructions, to-do lists and emails in minutes." },
    ],
  }),
  component: Index,
});

const fakeResults = [
  { label: "Top product", value: "Bread", detail: "R5,000 in sales" },
  { label: "Best day", value: "Friday", detail: "32% of weekly revenue" },
  { label: "Slowest item", value: "Muffins", detail: "Only 12 units sold" },
];

const fakeAnswers: Record<string, string> = {
  instructions:
    "1. Open the shared drive.\n2. Locate the 'Weekly Reports' folder.\n3. Save your file using the format YYYY-MM-DD_Name.xlsx.\n4. Notify your manager via email once uploaded.\n5. Archive last week's file into the 'Archive' subfolder.",
  meeting:
    "• Follow up with Sarah re: Q3 budget by Friday\n• Send updated client proposal to Mark\n• Book meeting room for next Tuesday's review\n• Share marketing assets with the design team\n• Prepare slides for Monday's all-hands",
  email:
    "Subject: Weekly Update — Steady Progress Across the Board\n\nHi team,\n\nThis week we hit 92% of our sales target and onboarded two new clients. Bread continues to lead the category at R5,000. Next week we'll focus on improving slow-moving items and finalising the Q4 plan.\n\nThanks for your hard work,\n— The Team",
};

type TimeRange = "today" | "week" | "month" | "year";

const timeRanges: { key: TimeRange; label: string }[] = [
  { key: "today", label: "Today" },
  { key: "week", label: "This Week" },
  { key: "month", label: "This Month" },
  { key: "year", label: "This Year" },
];

const smartSummaries: Record<TimeRange, string[]> = {
  today: [
    "Trend: Sales are up 4% compared to yesterday.",
    "Best product: Coffee brought in the most money today.",
    "Busiest hour: 10am had the highest number of orders.",
  ],
  week: [
    "Trend: Sales went up 15% compared to last week.",
    "Best product: Bread made the most money this week.",
    "Busiest day: Friday had the highest sales.",
  ],
  month: [
    "Trend: Revenue grew 8% compared to last month.",
    "Best product: Pastries were the top earner this month.",
    "Busiest week: Week 3 brought in the most sales.",
  ],
  year: [
    "Trend: Annual revenue is up 22% compared to last year.",
    "Best product: Bread is the top seller for the year.",
    "Busiest month: December had the highest sales overall.",
  ],
};

function Index() {
  const [uploaded, setUploaded] = useState<string | null>(null);
  const [activeAction, setActiveAction] = useState<keyof typeof fakeAnswers | null>(null);
  const [activeRange, setActiveRange] = useState<TimeRange | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const handleFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) setUploaded(f.name);
  };

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b border-border bg-card">
        <div className="mx-auto max-w-6xl px-6 py-5 flex items-center gap-3">
          <div className="h-9 w-9 rounded-lg bg-primary flex items-center justify-center">
            <FileSpreadsheet className="h-5 w-5 text-primary-foreground" />
          </div>
          <div>
            <h1 className="text-lg font-semibold text-foreground">OfficeAssist</h1>
            <p className="text-xs text-muted-foreground">Less busywork. More real work.</p>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-6 py-10">
        <div className="grid gap-8 lg:grid-cols-[1fr_280px]">
          <section className="space-y-8">
            {/* Upload */}
            <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
              <h2 className="text-base font-semibold text-foreground mb-1">Upload your Excel file</h2>
              <p className="text-sm text-muted-foreground mb-4">We'll pull out the highlights so you don't have to.</p>
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFile}
                className="hidden"
              />
              <button
                onClick={() => fileRef.current?.click()}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2.5 text-sm font-medium text-primary-foreground transition-colors hover:bg-primary/90"
              >
                <Upload className="h-4 w-4" />
                {uploaded ? "Upload another file" : "Upload Excel file"}
              </button>
              {uploaded && (
                <p className="mt-3 text-xs text-muted-foreground">Loaded: <span className="text-foreground font-medium">{uploaded}</span></p>
              )}
            </div>

            {/* Results */}
            {uploaded && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <h2 className="text-base font-semibold text-foreground mb-4">Results</h2>
                <div className="grid gap-4 sm:grid-cols-3">
                  {fakeResults.map((r) => (
                    <div key={r.label} className="rounded-lg bg-secondary p-4">
                      <p className="text-xs uppercase tracking-wide text-muted-foreground">{r.label}</p>
                      <p className="mt-1 text-xl font-semibold text-foreground">{r.value}</p>
                      <p className="text-sm text-muted-foreground">{r.detail}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}


            {activeAction && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <h2 className="text-base font-semibold text-foreground mb-3">Generated for you</h2>
                <pre className="whitespace-pre-wrap rounded-lg bg-secondary p-4 text-sm text-foreground font-sans leading-relaxed">
                  {fakeAnswers[activeAction]}
                </pre>
              </div>
            )}
          </section>

          {/* Sidebar actions */}
          <aside className="space-y-3">
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Quick actions</p>
            <ActionButton icon={ListChecks} label="Make Instructions" onClick={() => setActiveAction("instructions")} active={activeAction === "instructions"} />
            <ActionButton icon={ClipboardList} label="Meeting Notes to To-Do List" onClick={() => setActiveAction("meeting")} active={activeAction === "meeting"} />
            <ActionButton icon={Mail} label="Write Weekly Email" onClick={() => setActiveAction("email")} active={activeAction === "email"} />

            <div className="pt-4 space-y-3">
              <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Smart Summary</p>
              <div className="grid grid-cols-2 gap-2">
                {timeRanges.map((r) => {
                  const active = activeRange === r.key;
                  return (
                    <button
                      key={r.key}
                      onClick={() => setActiveRange(r.key)}
                      className={`rounded-lg border px-3 py-2 text-xs font-medium transition-colors ${
                        active
                          ? "border-primary bg-primary text-primary-foreground"
                          : "border-border bg-card text-foreground hover:bg-secondary"
                      }`}
                    >
                      {r.label}
                    </button>
                  );
                })}
              </div>
              <div className="rounded-lg border border-primary/20 bg-primary/10 p-4">
                <div className="flex items-center gap-2 mb-2">
                  <Sparkles className="h-4 w-4 text-primary" />
                  <h3 className="text-sm font-semibold text-foreground">Smart Summary</h3>
                </div>
                {activeRange ? (
                  <ul className="space-y-2 text-sm text-foreground leading-relaxed">
                    {smartSummaries[activeRange].map((s, i) => (
                      <li key={i} className="flex gap-2">
                        <span className="text-primary font-semibold">{i + 1}.</span>
                        <span>{s}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="text-sm text-muted-foreground">Pick a time range to see a quick summary.</p>
                )}
              </div>
            </div>
          </aside>
        </div>

        {/* Time saved */}
        <div className="mt-12 rounded-xl border border-border bg-primary p-6 text-primary-foreground flex items-center gap-4">
          <Clock className="h-8 w-8 shrink-0 opacity-80" />
          <div>
            <p className="text-sm opacity-80">Before: 3 hours of work.</p>
            <p className="text-lg font-semibold">After: 20 minutes with this website.</p>
          </div>
        </div>
      </main>
    </div>
  );
}

function ActionButton({
  icon: Icon,
  label,
  onClick,
  active,
}: {
  icon: React.ComponentType<{ className?: string }>;
  label: string;
  onClick: () => void;
  active: boolean;
}) {
  return (
    <button
      onClick={onClick}
      className={`w-full flex items-center gap-3 rounded-lg border px-4 py-3 text-left text-sm font-medium transition-colors ${
        active
          ? "border-primary bg-primary text-primary-foreground"
          : "border-border bg-card text-foreground hover:bg-secondary"
      }`}
    >
      <Icon className="h-4 w-4 shrink-0" />
      <span>{label}</span>
    </button>
  );
}
