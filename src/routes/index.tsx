import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useRef, useState } from "react";
import { useServerFn } from "@tanstack/react-start";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  ListChecks,
  ClipboardList,
  Mail,
  Clock,
  Sparkles,
  AlertCircle,
  Loader2,
  TrendingUp,
  BarChart3,
  GitCompare,
  FlaskConical,
  DollarSign,
} from "lucide-react";
import { generateSummary } from "@/lib/api/summary.functions";
import { generateAction } from "@/lib/api/action.functions";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "OfficeAssist — Work smarter, not longer" },
      { name: "description", content: "Upload spreadsheets and generate insights, trends, and reports in minutes." },
    ],
  }),
  component: Index,
});

type ActionKey =
  | "instructions"
  | "meeting"
  | "email"
  | "trends"
  | "scientific"
  | "sales";

const actionToServer: Record<ActionKey, "instructions" | "notes" | "email" | "trends" | "scientific" | "sales"> = {
  instructions: "instructions",
  meeting: "notes",
  email: "email",
  trends: "trends",
  scientific: "scientific",
  sales: "sales",
};
const actionLabels: Record<ActionKey, string> = {
  instructions: "Send Instructions",
  meeting: "Write Notes",
  email: "Draft Email",
  trends: "Trends & Patterns",
  scientific: "Scientific Summary",
  sales: "Sales Insights",
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

function findSalesColumn(columns: string[]): string | null {
  const targets = ["sales", "amount", "revenue", "total", "price"];
  for (const t of targets) {
    const idx = columns.findIndex((c) => c.toLowerCase().trim() === t);
    if (idx !== -1) return columns[idx];
  }
  for (const t of targets) {
    const idx = columns.findIndex((c) => c.toLowerCase().includes(t));
    if (idx !== -1) return columns[idx];
  }
  return null;
}

function isEmpty(v: unknown) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function toNumber(v: unknown): number | null {
  if (typeof v === "number" && !isNaN(v)) return v;
  if (typeof v === "string") {
    const n = parseFloat(v.replace(/[,$]/g, ""));
    if (!isNaN(n)) return n;
  }
  return null;
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

function classifyColumns(rows: Row[], columns: string[]) {
  const numericCols: string[] = [];
  const textCols: string[] = [];
  for (const c of columns) {
    let nums = 0, vals = 0;
    for (const r of rows) {
      const v = r[c];
      if (isEmpty(v)) continue;
      vals++;
      if (toNumber(v) !== null) nums++;
    }
    if (vals === 0) continue;
    if (nums / vals > 0.7) numericCols.push(c);
    else textCols.push(c);
  }
  return { numericCols, textCols };
}

function Index() {
  const [uploaded, setUploaded] = useState<string | null>(null);
  const [rows, setRows] = useState<Row[] | null>(null);
  const [columns, setColumns] = useState<string[]>([]);
  const [dateColumn, setDateColumn] = useState<string | null>(null);
  const [salesColumn, setSalesColumn] = useState<string | null>(null);
  const [emptyCells, setEmptyCells] = useState(0);
  const [isChartLike, setIsChartLike] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [activeAction, setActiveAction] = useState<ActionKey | null>(null);
  const [activeRange, setActiveRange] = useState<TimeRange | null>(null);
  const [summaryError, setSummaryError] = useState<string | null>(null);
  const [summaryText, setSummaryText] = useState<string | null>(null);
  const [loadingSummary, setLoadingSummary] = useState(false);
  const [actionText, setActionText] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [loadingAction, setLoadingAction] = useState(false);
  const [copied, setCopied] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const runSummary = useServerFn(generateSummary);
  const runAction = useServerFn(generateAction);

  const { numericCols, textCols } = useMemo(
    () => (rows ? classifyColumns(rows, columns) : { numericCols: [], textCols: [] }),
    [rows, columns],
  );

  const metrics = useMemo(() => {
    if (!rows) return null;
    const numStats = numericCols.map((c) => {
      const nums: number[] = [];
      for (const r of rows) {
        const n = toNumber(r[c]);
        if (n !== null) nums.push(n);
      }
      if (nums.length === 0) return null;
      const total = nums.reduce((a, b) => a + b, 0);
      return {
        column: c,
        total,
        avg: total / nums.length,
        min: Math.min(...nums),
        max: Math.max(...nums),
        count: nums.length,
      };
    }).filter(Boolean) as Array<{ column: string; total: number; avg: number; min: number; max: number; count: number }>;

    const textStats = textCols.map((c) => {
      const freq = new Map<string, number>();
      for (const r of rows) {
        const v = r[c];
        if (isEmpty(v)) continue;
        const k = String(v);
        freq.set(k, (freq.get(k) ?? 0) + 1);
      }
      let top: { value: string; count: number } | null = null;
      for (const [v, n] of freq) {
        if (!top || n > top.count) top = { value: v, count: n };
      }
      return top ? { column: c, mostCommon: top.value, count: top.count } : null;
    }).filter(Boolean) as Array<{ column: string; mostCommon: string; count: number }>;

    return { numStats, textStats };
  }, [rows, numericCols, textCols]);

  const comparison = useMemo(() => {
    if (!rows || numericCols.length === 0 || textCols.length === 0) return null;
    const categoryCols = ["region", "product", "farm", "category", "department", "store", "name"];
    const catCol =
      textCols.find((c) => categoryCols.includes(c.toLowerCase().trim())) ?? textCols[0];
    const valCol = salesColumn && numericCols.includes(salesColumn) ? salesColumn : numericCols[0];
    const totals = new Map<string, number>();
    for (const r of rows) {
      const k = r[catCol];
      const n = toNumber(r[valCol]);
      if (isEmpty(k) || n === null) continue;
      const key = String(k);
      totals.set(key, (totals.get(key) ?? 0) + n);
    }
    const sorted = [...totals.entries()].sort((a, b) => b[1] - a[1]);
    if (sorted.length === 0) return null;
    return {
      categoryCol: catCol,
      valueCol: valCol,
      top: sorted.slice(0, 3),
      bottom: sorted.slice(-3).reverse(),
    };
  }, [rows, numericCols, textCols, salesColumn]);

  const salesAuto = useMemo(() => {
    if (!rows || !salesColumn) return null;
    let total = 0, count = 0;
    const byDay = new Map<string, number>();
    const byCategory = new Map<string, number>();
    const catCol = textCols.find((c) => ["product", "region", "farm", "category"].includes(c.toLowerCase().trim())) ?? textCols[0];
    for (const r of rows) {
      const n = toNumber(r[salesColumn]);
      if (n === null) continue;
      total += n; count++;
      if (dateColumn) {
        const d = parseDate(r[dateColumn]);
        if (d) {
          const k = d.toISOString().slice(0, 10);
          byDay.set(k, (byDay.get(k) ?? 0) + n);
        }
      }
      if (catCol) {
        const k = r[catCol];
        if (!isEmpty(k)) byCategory.set(String(k), (byCategory.get(String(k)) ?? 0) + n);
      }
    }
    if (count === 0) return null;
    const bestDay = [...byDay.entries()].sort((a, b) => b[1] - a[1])[0] ?? null;
    const bestCategory = [...byCategory.entries()].sort((a, b) => b[1] - a[1])[0] ?? null;
    return {
      total,
      avg: total / count,
      bestDay,
      bestCategory,
      categoryLabel: catCol ?? null,
    };
  }, [rows, salesColumn, dateColumn, textCols]);

  const handleAction = async (key: ActionKey) => {
    setActiveAction(key);
    setActionText(null);
    setActionError(null);
    setCopied(false);
    if (!rows || rows.length === 0) {
      setActionError("Upload Excel file first before generating text");
      return;
    }
    setLoadingAction(true);
    try {
      const sample = rows.slice(0, 20).map((r) => {
        const out: Record<string, unknown> = {};
        for (const k of columns) {
          const v = r[k];
          out[k] = v instanceof Date ? v.toISOString().slice(0, 10) : v;
        }
        return out;
      });
      const res = await runAction({
        data: {
          action: actionToServer[key],
          columns,
          sampleRows: sample,
          totalRows: rows.length,
          isChartLike,
          emptyCells,
        },
      });
      setActionText(res.text);
    } catch (err) {
      setActionError(err instanceof Error ? err.message : "Failed to generate");
    } finally {
      setLoadingAction(false);
    }
  };

  const copyAction = async () => {
    if (!actionText) return;
    try {
      await navigator.clipboard.writeText(actionText);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch {}
  };

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setParseError(null);
    setActiveRange(null);
    setSummaryText(null);
    setSummaryError(null);
    setActionText(null);
    setActiveAction(null);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json<Row>(ws, { defval: null });
      if (json.length === 0) {
        setParseError("File uploaded but no data found");
        setRows(null); setColumns([]); setDateColumn(null); setSalesColumn(null);
        setEmptyCells(0); setIsChartLike(false); setUploaded(f.name);
        return;
      }
      const colSet = new Set<string>();
      let empty = 0;
      json.forEach((r) => {
        Object.keys(r).forEach((k) => colSet.add(k));
      });
      const cols = Array.from(colSet);
      json.forEach((r) => {
        for (const c of cols) if (isEmpty(r[c])) empty++;
      });
      const { numericCols: nc, textCols: tc } = classifyColumns(json, cols);
      const chartLike = nc.length >= 1 && tc.length >= 1 && json.length >= 2;
      setRows(json);
      setColumns(cols);
      setDateColumn(findDateColumn(cols));
      setSalesColumn(findSalesColumn(cols));
      setEmptyCells(empty);
      setIsChartLike(chartLike);
      setUploaded(f.name);
    } catch {
      setParseError("Could not read that file. Please upload a valid .xlsx file.");
      setRows(null); setColumns([]); setDateColumn(null); setSalesColumn(null);
      setEmptyCells(0); setIsChartLike(false); setUploaded(null);
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

  const fmt = (n: number) =>
    n >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 2 }) : n.toFixed(2);

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
                Any spreadsheet works — sales, lab results, surveys, inventory.
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
                  {salesColumn && (
                    <p>Sales/amount column: <span className="text-foreground">{salesColumn}</span></p>
                  )}
                  {isChartLike && (
                    <p className="text-foreground">Looks like chart/table data (numeric + label columns).</p>
                  )}
                  {emptyCells > 0 && (
                    <p className="text-amber-600 dark:text-amber-400">⚠ {emptyCells} empty cells skipped</p>
                  )}
                </div>
              )}
              {parseError && (
                <p className="mt-3 text-xs text-destructive flex items-center gap-1.5">
                  <AlertCircle className="h-3.5 w-3.5" /> {parseError}
                </p>
              )}
            </div>

            {rows && rows.length > 0 && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <h2 className="text-base font-semibold text-foreground mb-3">Data preview (first 10 rows)</h2>
                <div className="overflow-auto max-h-80 rounded-lg border border-border">
                  <table className="w-full text-xs">
                    <thead className="bg-secondary sticky top-0">
                      <tr>
                        {columns.map((c) => (
                          <th key={c} className="text-left font-medium px-3 py-2 whitespace-nowrap">{c}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.slice(0, 10).map((r, i) => (
                        <tr key={i} className="border-t border-border">
                          {columns.map((c) => {
                            const v = r[c];
                            const display =
                              v instanceof Date ? v.toISOString().slice(0, 10) : isEmpty(v) ? "—" : String(v);
                            return (
                              <td key={c} className="px-3 py-2 whitespace-nowrap text-foreground">{display}</td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {metrics && (metrics.numStats.length > 0 || metrics.textStats.length > 0) && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <div className="flex items-center gap-2 mb-3">
                  <BarChart3 className="h-4 w-4 text-primary" />
                  <h2 className="text-base font-semibold text-foreground">Key Metrics</h2>
                </div>
                <div className="grid gap-3 sm:grid-cols-2">
                  {metrics.numStats.map((s) => (
                    <div key={s.column} className="rounded-lg border border-border bg-secondary p-3">
                      <p className="text-xs font-medium text-muted-foreground mb-1">{s.column}</p>
                      <div className="grid grid-cols-2 gap-1 text-xs">
                        <span className="text-muted-foreground">Total</span><span className="text-foreground text-right">{fmt(s.total)}</span>
                        <span className="text-muted-foreground">Average</span><span className="text-foreground text-right">{fmt(s.avg)}</span>
                        <span className="text-muted-foreground">Min</span><span className="text-foreground text-right">{fmt(s.min)}</span>
                        <span className="text-muted-foreground">Max</span><span className="text-foreground text-right">{fmt(s.max)}</span>
                        <span className="text-muted-foreground">Count</span><span className="text-foreground text-right">{s.count}</span>
                      </div>
                    </div>
                  ))}
                  {metrics.textStats.map((s) => (
                    <div key={s.column} className="rounded-lg border border-border bg-secondary p-3">
                      <p className="text-xs font-medium text-muted-foreground mb-1">{s.column}</p>
                      <p className="text-xs"><span className="text-muted-foreground">Most common: </span><span className="text-foreground font-medium">{s.mostCommon}</span> <span className="text-muted-foreground">({s.count}×)</span></p>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {salesAuto && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <div className="flex items-center gap-2 mb-3">
                  <DollarSign className="h-4 w-4 text-primary" />
                  <h2 className="text-base font-semibold text-foreground">Sales Overview</h2>
                </div>
                <div className="grid gap-3 sm:grid-cols-2 text-xs">
                  <div className="rounded-lg bg-secondary p-3"><p className="text-muted-foreground">Total Sales</p><p className="text-foreground font-semibold text-sm">{fmt(salesAuto.total)}</p></div>
                  <div className="rounded-lg bg-secondary p-3"><p className="text-muted-foreground">Average Sale</p><p className="text-foreground font-semibold text-sm">{fmt(salesAuto.avg)}</p></div>
                  {salesAuto.bestDay && (
                    <div className="rounded-lg bg-secondary p-3"><p className="text-muted-foreground">Best Day</p><p className="text-foreground font-semibold text-sm">{salesAuto.bestDay[0]} ({fmt(salesAuto.bestDay[1])})</p></div>
                  )}
                  {salesAuto.bestCategory && (
                    <div className="rounded-lg bg-secondary p-3"><p className="text-muted-foreground">Best {salesAuto.categoryLabel}</p><p className="text-foreground font-semibold text-sm">{salesAuto.bestCategory[0]} ({fmt(salesAuto.bestCategory[1])})</p></div>
                  )}
                </div>
              </div>
            )}

            {comparison && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <div className="flex items-center gap-2 mb-3">
                  <GitCompare className="h-4 w-4 text-primary" />
                  <h2 className="text-base font-semibold text-foreground">Comparison</h2>
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  By <span className="text-foreground">{comparison.categoryCol}</span> — values from <span className="text-foreground">{comparison.valueCol}</span>
                </p>
                <div className="grid gap-3 sm:grid-cols-2 text-xs">
                  <div>
                    <p className="font-medium text-foreground mb-1">Top 3</p>
                    <ul className="space-y-1">
                      {comparison.top.map(([k, v]) => (
                        <li key={k} className="flex justify-between rounded bg-secondary px-2 py-1"><span>{k}</span><span className="text-foreground">{fmt(v)}</span></li>
                      ))}
                    </ul>
                  </div>
                  <div>
                    <p className="font-medium text-foreground mb-1">Bottom 3</p>
                    <ul className="space-y-1">
                      {comparison.bottom.map(([k, v]) => (
                        <li key={k} className="flex justify-between rounded bg-secondary px-2 py-1"><span>{k}</span><span className="text-foreground">{fmt(v)}</span></li>
                      ))}
                    </ul>
                  </div>
                </div>
              </div>
            )}

            {activeAction && (
              <div className="rounded-xl border border-border bg-card p-6 shadow-sm">
                <div className="flex items-center justify-between mb-3">
                  <h2 className="text-base font-semibold text-foreground">{actionLabels[activeAction]}</h2>
                  {actionText && (
                    <button
                      onClick={copyAction}
                      className="text-xs font-medium rounded-md border border-border bg-card px-2.5 py-1 hover:bg-secondary"
                    >
                      {copied ? "Copied!" : "Copy"}
                    </button>
                  )}
                </div>
                {loadingAction ? (
                  <p className="text-sm text-foreground flex items-center gap-2"><Loader2 className="h-3.5 w-3.5 animate-spin" /> Generating…</p>
                ) : actionError ? (
                  <p className="text-sm text-destructive flex items-center gap-1.5">
                    <AlertCircle className="h-3.5 w-3.5" /> {actionError}
                  </p>
                ) : actionText ? (
                  <textarea
                    readOnly
                    value={actionText}
                    className="w-full min-h-[200px] rounded-lg bg-secondary p-4 text-sm text-foreground font-sans leading-relaxed resize-y border border-border focus:outline-none focus:ring-1 focus:ring-ring"
                  />
                ) : null}
              </div>
            )}
          </section>

          <aside className="space-y-3">
            <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Quick actions</p>
            <ActionButton icon={ListChecks} label="Send Instructions" onClick={() => handleAction("instructions")} active={activeAction === "instructions"} />
            <ActionButton icon={ClipboardList} label="Write Notes" onClick={() => handleAction("meeting")} active={activeAction === "meeting"} />
            <ActionButton icon={Mail} label="Draft Email" onClick={() => handleAction("email")} active={activeAction === "email"} />

            <p className="pt-3 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Analysis</p>
            <ActionButton icon={TrendingUp} label="Trends & Patterns" onClick={() => handleAction("trends")} active={activeAction === "trends"} />
            <ActionButton icon={FlaskConical} label="Scientific Summary" onClick={() => handleAction("scientific")} active={activeAction === "scientific"} />
            <ActionButton icon={DollarSign} label="Sales Insights" onClick={() => handleAction("sales")} active={activeAction === "sales"} />

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
