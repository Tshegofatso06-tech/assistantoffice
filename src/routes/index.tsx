import { createFileRoute } from "@tanstack/react-router";
import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, ListChecks, ClipboardList, Mail, Clock, Sparkles, AlertCircle } from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "OfficeAssist — Work smarter, not longer" },
      { name: "description", content: "Upload spreadsheets and generate instructions, to-do lists and emails in minutes." },
    ],
  }),
  component: Index,
});

const fakeAnswers: Record<string, string> = {
  instructions:
    "1. Open the shared drive.\n2. Locate the 'Weekly Reports' folder.\n3. Save your file using the format YYYY-MM-DD_Name.xlsx.\n4. Notify your manager via email once uploaded.\n5. Archive last week's file into the 'Archive' subfolder.",
  meeting:
    "• Follow up with Sarah re: Q3 budget by Friday\n• Send updated client proposal to Mark\n• Book meeting room for next Tuesday's review\n• Share marketing assets with the design team\n• Prepare slides for Monday's all-hands",
  email:
    "Subject: Weekly Update — Steady Progress Across the Board\n\nHi team,\n\nThis week we hit 92% of our sales target and onboarded two new clients. Next week we'll focus on improving slow-moving items and finalising the Q4 plan.\n\nThanks for your hard work,\n— The Team",
};

type TimeRange = "thisWeek" | "lastWeek" | "thisMonth" | "lastMonth";

const timeRanges: { key: TimeRange; label: string }[] = [
  { key: "thisWeek", label: "This Week" },
  { key: "lastWeek", label: "Last Week" },
  { key: "thisMonth", label: "This Month" },
  { key: "lastMonth", label: "Last Month" },
];

type Row = { date: Date; amount: number; sales: number };

function parseDate(v: unknown): Date | null {
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    // Excel serial date
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0);
  }
  if (typeof v === "string") {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function toNum(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/[^0-9.-]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function startOfDay(d: Date) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getRange(key: TimeRange): { from: Date; to: Date } {
  const now = new Date();
  const today = startOfDay(now);
  const dow = today.getDay(); // 0=Sun
  const monOffset = (dow + 6) % 7; // days since Monday
  const thisMonStart = new Date(today); thisMonStart.setDate(today.getDate() - monOffset);
  const lastMonStart = new Date(thisMonStart); lastMonStart.setDate(thisMonStart.getDate() - 7);
  const thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);

  switch (key) {
    case "thisWeek":
      return { from: thisMonStart, to: new Date(thisMonStart.getTime() + 7 * 86400000) };
    case "lastWeek":
      return { from: lastMonStart, to: thisMonStart };
    case "thisMonth":
      return { from: thisMonthStart, to: new Date(today.getFullYear(), today.getMonth() + 1, 1) };
    case "lastMonth":
      return { from: lastMonthStart, to: thisMonthStart };
  }
}

function summarize(rows: Row[], key: TimeRange): string[] {
  const { from, to } = getRange(key);
  const filtered = rows.filter((r) => r.date >= from && r.date < to);
  if (filtered.length === 0) {
    return [`No records found for ${timeRanges.find((t) => t.key === key)?.label}.`];
  }
  const totalSales = filtered.reduce((s, r) => s + r.sales, 0);
  const totalAmount = filtered.reduce((s, r) => s + r.amount, 0);
  // Best day by sales
  const byDay = new Map<string, number>();
  filtered.forEach((r) => {
    const k = startOfDay(r.date).toDateString();
    byDay.set(k, (byDay.get(k) || 0) + r.sales);
  });
  let bestDay = ""; let bestVal = -Infinity;
  byDay.forEach((v, k) => { if (v > bestVal) { bestVal = v; bestDay = k; } });

  return [
    `Total sales: ${totalSales.toLocaleString()} across ${filtered.length} record${filtered.length === 1 ? "" : "s"}.`,
    `Total amount: ${totalAmount.toLocaleString()}.`,
    `Best day: ${bestDay} with ${bestVal.toLocaleString()} in sales.`,
  ];
}

function Index() {
  const [uploaded, setUploaded] = useState<string | null>(null);
  const [rows, setRows] = useState<Row[] | null>(null);
  const [parseError, setParseError] = useState<string | null>(null);
  const [activeAction, setActiveAction] = useState<keyof typeof fakeAnswers | null>(null);
  const [activeRange, setActiveRange] = useState<TimeRange | null>(null);
  const [summaryError, setSummaryError] = useState<string | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setParseError(null);
    setActiveRange(null);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: null });
      const parsed: Row[] = [];
      for (const r of json) {
        const dateRaw = r.Date ?? r.date;
        const amountRaw = r.Amount ?? r.amount;
        const salesRaw = r.Sales ?? r.sales;
        const date = parseDate(dateRaw);
        if (!date) continue;
        parsed.push({ date, amount: toNum(amountRaw), sales: toNum(salesRaw) });
      }
      if (parsed.length === 0) {
        setParseError("No valid rows found. Make sure your file has Date, Amount and Sales columns.");
        setRows(null);
        setUploaded(null);
        return;
      }
      setRows(parsed);
      setUploaded(f.name);
      setSummaryError(null);
    } catch (err) {
      setParseError("Could not read that file. Please upload a valid .xlsx file.");
      setRows(null);
      setUploaded(null);
    }
  };

  const handleRange = (key: TimeRange) => {
    if (!rows) {
      setSummaryError("Please upload Excel file first");
      setActiveRange(null);
      return;
    }
    setSummaryError(null);
    setActiveRange(key);
  };

  const summary = activeRange && rows ? summarize(rows, activeRange) : null;

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
            <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
              <h2 className="text-base font-semibold text-foreground mb-1">Upload your Excel file</h2>
              <p className="text-sm text-muted-foreground mb-4">
                We'll use the Date, Amount and Sales columns to build your summary.
              </p>
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls"
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
                <p className="mt-3 text-xs text-muted-foreground">
                  Loaded: <span className="text-foreground font-medium">{uploaded}</span> ({rows?.length} rows)
                </p>
              )}
              {parseError && (
                <p className="mt-3 text-xs text-destructive flex items-center gap-1.5">
                  <AlertCircle className="h-3.5 w-3.5" /> {parseError}
                </p>
              )}
            </div>

            {activeAction && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <h2 className="text-base font-semibold text-foreground mb-3">Generated for you</h2>
                <pre className="whitespace-pre-wrap rounded-lg bg-secondary p-4 text-sm text-foreground font-sans leading-relaxed">
                  {fakeAnswers[activeAction]}
                </pre>
              </div>
            )}
          </section>

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
                      onClick={() => handleRange(r.key)}
                      className={`rounded-lg border px-2 py-2 text-xs font-medium transition-colors ${
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
              <div className="rounded-lg border border-primary/20 bg-sky-100 p-4">
                <div className="flex items-center gap-2 mb-2">
                  <Sparkles className="h-4 w-4 text-primary" />
                  <h3 className="text-sm font-semibold text-foreground">Smart Summary</h3>
                </div>
                {summaryError ? (
                  <p className="text-sm text-destructive flex items-center gap-1.5">
                    <AlertCircle className="h-3.5 w-3.5" /> {summaryError}
                  </p>
                ) : summary ? (
                  <ul className="space-y-2 text-sm text-foreground leading-relaxed">
                    {summary.map((s, i) => (
                      <li key={i} className="flex gap-2">
                        <span className="text-primary font-semibold">{i + 1}.</span>
                        <span>{s}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="text-sm text-muted-foreground">
                    {uploaded ? "Pick a time range to see a quick summary." : "Upload a file, then pick a time range."}
                  </p>
                )}
              </div>
            </div>
          </aside>
        </div>

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
