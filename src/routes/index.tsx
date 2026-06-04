import { createFileRoute } from "@tanstack/react-router";
import { useRef, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import * as XLSX from "xlsx";
import { Upload, FileSpreadsheet, ListChecks, ClipboardList, Mail, Clock, Sparkles, AlertCircle, Loader2 } from "lucide-react";
import { generateSummary } from "@/lib/api/summary.functions";
import { generateAction } from "@/lib/api/action.functions";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "OfficeAssist — Work smarter, not longer" },
      { name: "description", content: "Upload spreadsheets and generate instructions, to-do lists and emails in minutes." },
    ],
  }),
  component: Index,
});

type ActionKey = "instructions" | "meeting" | "email";
const actionToServer: Record<ActionKey, "instructions" | "notes" | "email"> = {
  instructions: "instructions",
  meeting: "notes",
  email: "email",
};
const actionLabels: Record<ActionKey, string> = {
  instructions: "Send Instructions",
  meeting: "Write Notes",
  email: "Draft Email",
};

type TimeRange = "thisWeek" | "lastWeek" | "thisMonth" | "lastMonth";

const timeRanges: { key: TimeRange; label: string }[] = [
  { key: "thisWeek", label: "This Week" },
  { key: "lastWeek", label: "Last Week" },
  { key: "thisMonth", label: "This Month" },
  { key: "lastMonth", label: "Last Month" },
];

type Row = Record<string, unknown>;

function findDateColumn(columns: string[]): string | null {
  const lower = columns.map((c) => c.toLowerCase().trim());
  const exact = lower.findIndex((c) => c === "date");
  if (exact !== -1) return columns[exact];
  const partial = lower.findIndex((c) => c.includes("date"));
  return partial !== -1 ? columns[partial] : null;
}

function parseDate(v: unknown): Date | null {
  if (v instanceof Date) return v;
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (typeof v === "string") {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function startOfDay(d: Date) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function getRange(key: TimeRange): { from: Date; to: Date } {
  const today = startOfDay(new Date());
  const dow = today.getDay();
  const monOffset = (dow + 6) % 7;
  const thisMon = new Date(today); thisMon.setDate(today.getDate() - monOffset);
  const lastMon = new Date(thisMon); lastMon.setDate(thisMon.getDate() - 7);
  const thisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  switch (key) {
    case "thisWeek": return { from: thisMon, to: new Date(thisMon.getTime() + 7 * 86400000) };
    case "lastWeek": return { from: lastMon, to: thisMon };
    case "thisMonth": return { from: thisMonth, to: new Date(today.getFullYear(), today.getMonth() + 1, 1) };
    case "lastMonth": return { from: lastMonth, to: thisMonth };
  }
}

function Index() {
  const [uploaded, setUploaded] = useState<string | null>(null);
  const [rows, setRows] = useState<Row[] | null>(null);
  const [columns, setColumns] = useState<string[]>([]);
  const [dateColumn, setDateColumn] = useState<string | null>(null);
  const [parseError, setParseError] = useState<string | null>(null);
  const [activeAction, setActiveAction] = useState<keyof typeof fakeAnswers | null>(null);
  const [activeRange, setActiveRange] = useState<TimeRange | null>(null);
  const [summaryError, setSummaryError] = useState<string | null>(null);
  const [summaryText, setSummaryText] = useState<string | null>(null);
  const [loadingSummary, setLoadingSummary] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const runSummary = useServerFn(generateSummary);

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setParseError(null);
    setActiveRange(null);
    setSummaryText(null);
    setSummaryError(null);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Row>(ws, { defval: null });
      if (json.length === 0) {
        setParseError("File uploaded but no data found");
        setRows(null); setColumns([]); setDateColumn(null); setUploaded(f.name);
        return;
      }
      const colSet = new Set<string>();
      json.forEach((r) => Object.keys(r).forEach((k) => colSet.add(k)));
      const cols = Array.from(colSet);
      setRows(json);
      setColumns(cols);
      setDateColumn(findDateColumn(cols));
      setUploaded(f.name);
    } catch {
      setParseError("Could not read that file. Please upload a valid .xlsx file.");
      setRows(null); setColumns([]); setDateColumn(null); setUploaded(null);
    }
  };

  const handleRange = async (key: TimeRange) => {
    if (!rows) {
      setSummaryError("Please upload Excel file first");
      setActiveRange(null);
      return;
    }
    setSummaryError(null);
    setActiveRange(key);
    setSummaryText(null);
    setLoadingSummary(true);

    let filtered = rows;
    let rangeLabel = timeRanges.find((t) => t.key === key)?.label ?? "";
    let noDateNote = "";
    if (dateColumn) {
      const { from, to } = getRange(key);
      filtered = rows.filter((r) => {
        const d = parseDate(r[dateColumn]);
        return d !== null && d >= from && d < to;
      });
    } else {
      noDateNote = "No date column found, showing full data. ";
      rangeLabel = "all rows";
    }

    try {
      // Serialize Dates to strings for JSON transport
      const sample = filtered.slice(0, 50).map((r) => {
        const out: Record<string, unknown> = {};
        for (const k of columns) {
          const v = r[k];
          out[k] = v instanceof Date ? v.toISOString().slice(0, 10) : v;
        }
        return out;
      });
      if (sample.length === 0) {
        setSummaryText(`${noDateNote}No rows match ${rangeLabel}.`);
        setLoadingSummary(false);
        return;
      }
      const res = await runSummary({
        data: {
          columns,
          sampleRows: sample,
          totalRows: filtered.length,
          rangeLabel,
        },
      });
      setSummaryText(noDateNote + res.summary);
    } catch (err) {
      setSummaryError(err instanceof Error ? err.message : "Failed to generate summary");
    } finally {
      setLoadingSummary(false);
    }
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
            <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
              <h2 className="text-base font-semibold text-foreground mb-1">Upload your Excel file</h2>
              <p className="text-sm text-muted-foreground mb-4">
                Any spreadsheet works. We'll detect columns automatically.
              </p>
              <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} className="hidden" />
              <button
                onClick={() => fileRef.current?.click()}
                className="inline-flex items-center gap-2 rounded-lg bg-primary px-4 py-2.5 text-sm font-medium text-primary-foreground transition-colors hover:bg-primary/90"
              >
                <Upload className="h-4 w-4" />
                {uploaded ? "Upload another file" : "Upload Excel file"}
              </button>
              {uploaded && (
                <div className="mt-3 space-y-1 text-xs text-muted-foreground">
                  <p>Loaded: <span className="text-foreground font-medium">{uploaded}</span>{rows ? ` (${rows.length} rows)` : ""}</p>
                  {columns.length > 0 && (
                    <p>Detected columns: <span className="text-foreground">{columns.join(", ")}</span></p>
                  )}
                  {columns.length > 0 && (
                    <p>{dateColumn ? <>Date column: <span className="text-foreground">{dateColumn}</span></> : "No date column detected — time filters will summarize all rows."}</p>
                  )}
                </div>
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
                      disabled={loadingSummary}
                      className={`rounded-lg border px-2 py-2 text-xs font-medium transition-colors disabled:opacity-60 ${
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
                {loadingSummary ? (
                  <p className="text-sm text-foreground flex items-center gap-2"><Loader2 className="h-3.5 w-3.5 animate-spin" /> Analyzing…</p>
                ) : summaryError ? (
                  <p className="text-sm text-destructive flex items-center gap-1.5">
                    <AlertCircle className="h-3.5 w-3.5" /> {summaryError}
                  </p>
                ) : summaryText ? (
                  <p className="text-sm text-foreground leading-relaxed whitespace-pre-wrap">{summaryText}</p>
                ) : (
                  <p className="text-sm text-muted-foreground">
                    {uploaded ? "Pick a time range to see an AI summary." : "Upload a file, then pick a time range."}
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
