"use client";

import { useState, useRef } from "react";

type SheetRow = Record<string, string>;

type ProcessedItem = {
  gstin: string;
  rowNumber: number;
  /** idle = not started, skipped = already has data, fetching, success, error */
  status: "idle" | "skipped" | "fetching" | "success" | "error";
  gstr1Status?: string;
  gstr3bStatus?: string;
  gstr1Latest?: string;
  gstr3bLatest?: string;
  error?: string;
};

type ReturnStatus = {
  returnType: string;
  overallStatus: string;
  filedPeriods: string;
  notFiledPeriods: string;
  latestFiledOn: string | null;
  totalPeriods: number;
  filedCount: number;
};

type LookupResponse = {
  success: true;
  input: { gstin: string; fy: string; financialYearLabel: string };
  returnStatuses: ReturnStatus[];
};

const FY_OPTIONS = ["2022", "2023", "2024", "2025", "2026"];
const DEFAULT_FY = "2025";

/** Column names written to the sheet — keyed by FY label */
function colGstr1(fyLabel: string) { return `GSTR1 ${fyLabel}`; }
function colGstr3b(fyLabel: string) { return `GSTR3B ${fyLabel}`; }
function colGstr1Date(fyLabel: string) { return `GSTR1 Date ${fyLabel}`; }
function colGstr3bDate(fyLabel: string) { return `GSTR3B Date ${fyLabel}`; }

const APPS_SCRIPT_CODE = `function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const doc = SpreadsheetApp.openById(data.spreadsheetId);
    const sheet = data.gid
      ? doc.getSheets().find(s => s.getSheetId().toString() === data.gid)
      : doc.getSheets()[0];

    if (!sheet) {
      return json({ success: false, error: "Sheet not found" });
    }

    const sheetData = sheet.getDataRange().getValues();
    let headers = sheetData[0].map(h => h.toString().trim());

    // Ensure required columns exist (create if missing)
    const needed = data.columns; // array of column names to ensure
    needed.forEach(col => {
      if (!headers.includes(col)) {
        sheet.getRange(1, headers.length + 1).setValue(col);
        headers.push(col);
      }
    });

    // Build header → index map
    const hMap = {};
    headers.forEach((h, i) => { hMap[h.toLowerCase()] = i; });

    const gstinColIdx = hMap[data.gstinColumn.toLowerCase()];
    if (gstinColIdx === undefined) {
      return json({ success: false, error: "GSTIN column not found: " + data.gstinColumn });
    }

    // Re-read to get latest data (columns may have been added)
    const fresh = sheet.getDataRange().getValues();
    const freshHeaders = fresh[0].map(h => h.toString().trim());
    const fMap = {};
    freshHeaders.forEach((h, i) => { fMap[h.toLowerCase()] = i; });

    const updates = data.updates; // [{ gstin, values: { colName: value } }]

    for (let i = 1; i < fresh.length; i++) {
      const rowGstin = (fresh[i][gstinColIdx] || "").toString().trim().toUpperCase().replace(/\\s+/g, "");
      if (!rowGstin) continue;
      const match = updates.find(u => u.gstin.trim().toUpperCase().replace(/\\s+/g, "") === rowGstin);
      if (!match) continue;

      for (const [colName, value] of Object.entries(match.values)) {
        const colIdx = fMap[colName.toLowerCase()];
        if (colIdx === undefined) continue;
        // Only write if cell is empty (skip already-synced)
        if (fresh[i][colIdx] !== "" && fresh[i][colIdx] !== null && fresh[i][colIdx] !== undefined) continue;
        sheet.getRange(i + 1, colIdx + 1).setValue(value);
      }
    }

    return json({ success: true });
  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}`;

export default function Home() {
  // Config
  const [sheetUrl, setSheetUrl] = useState("");
  const [appsScriptUrl, setAppsScriptUrl] = useState("");
  const [fy, setFy] = useState(DEFAULT_FY);
  const [gstinColumn, setGstinColumn] = useState("");
  const [delayMs, setDelayMs] = useState(1000);

  // Sheet state
  const [headers, setHeaders] = useState<string[]>([]);
  const [sheetRows, setSheetRows] = useState<SheetRow[]>([]);
  const [spreadsheetId, setSpreadsheetId] = useState("");
  const [gid, setGid] = useState("");

  // Processing state
  const [items, setItems] = useState<ProcessedItem[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isPaused, setIsPaused] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(0);

  // UI state
  const [loadingSheet, setLoadingSheet] = useState(false);
  const [sheetError, setSheetError] = useState<string | null>(null);
  const [showScriptSetup, setShowScriptSetup] = useState(false);
  const [scriptCopied, setScriptCopied] = useState(false);
  const [writeStatus, setWriteStatus] = useState<{ success?: boolean; message?: string } | null>(null);
  const [writingToSheet, setWritingToSheet] = useState(false);

  const processingRef = useRef(false);
  const indexRef = useRef(0);
  const abortRef = useRef<AbortController | null>(null);

  // Derived FY label e.g. "2025-26"
  const fyLabel = `${fy}-${String(Number(fy) + 1).slice(-2)}`;

  async function handleLoadSheet(e: React.FormEvent | null) {
    if (e) e.preventDefault();
    if (!sheetUrl.trim()) return;

    setLoadingSheet(true);
    setSheetError(null);
    setWriteStatus(null);
    try {
      const res = await fetch("/api/sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sheetUrl }),
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.message || "Failed to parse Google Sheet.");

      setHeaders(data.headers);
      setSheetRows(data.rows);
      setSpreadsheetId(data.spreadsheetId);
      setGid(data.gid || "");

      // Auto-detect GSTIN column
      const gstinCol = data.headers.find((h: string) => {
        const n = h.toLowerCase().replace(/[^a-z0-9]/g, "");
        return n.includes("gstin") || n.includes("gstnumber") || n.includes("gstno") || n === "gst";
      });
      const colKey = gstinCol || data.headers[0] || "";
      setGstinColumn(colKey);

      // Build items — detect which rows already have this FY's data so we can skip them
      const g1Col = colGstr1(fyLabel);
      const g3bCol = colGstr3b(fyLabel);

      const newItems: ProcessedItem[] = (data.rows as SheetRow[]).map((row, idx) => {
        const rawGstin = (row[colKey] || "").trim().toUpperCase().replace(/\s+/g, "");
        const alreadyHasData = !!(row[g1Col]?.trim() || row[g3bCol]?.trim());
        return {
          gstin: rawGstin,
          rowNumber: idx + 2,
          status: alreadyHasData ? "skipped" : "idle",
          gstr1Status: row[g1Col] || undefined,
          gstr3bStatus: row[g3bCol] || undefined,
          gstr1Latest: row[colGstr1Date(fyLabel)] || undefined,
          gstr3bLatest: row[colGstr3bDate(fyLabel)] || undefined,
        };
      });
      setItems(newItems);
      setCurrentIndex(0);
      setIsPaused(false);
      setIsProcessing(false);
    } catch (err) {
      setSheetError(err instanceof Error ? err.message : "Unable to read Google Sheet.");
      setHeaders([]);
      setSheetRows([]);
      setItems([]);
    } finally {
      setLoadingSheet(false);
    }
  }

  async function runBatch() {
    processingRef.current = true;
    setIsProcessing(true);
    setIsPaused(false);

    while (indexRef.current < items.length && processingRef.current) {
      const idx = indexRef.current;
      const item = items[idx];

      // Skip rows already marked as skipped (already synced)
      if (item.status === "skipped") {
        indexRef.current++;
        setCurrentIndex(indexRef.current);
        continue;
      }

      if (!item.gstin || item.gstin.length !== 15) {
        setItems(prev => {
          const next = [...prev];
          next[idx] = { ...next[idx], status: "error", error: !item.gstin ? "Empty GSTIN" : "Invalid GSTIN (must be 15 chars)" };
          return next;
        });
        indexRef.current++;
        setCurrentIndex(indexRef.current);
        continue;
      }

      setItems(prev => { const n = [...prev]; n[idx] = { ...n[idx], status: "fetching" }; return n; });

      try {
        const ctrl = new AbortController();
        abortRef.current = ctrl;

        const res = await fetch("/api/gst", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ gstin: item.gstin, fy }),
          signal: ctrl.signal,
        });

        const result: LookupResponse = await res.json();

        if (res.ok && result.success) {
          const gstr1 = result.returnStatuses.find(s => s.returnType === "GSTR1");
          const gstr3b = result.returnStatuses.find(s => s.returnType === "GSTR3B");

          setItems(prev => {
            const n = [...prev];
            n[idx] = {
              ...n[idx],
              status: "success",
              gstr1Status: gstr1?.overallStatus ?? "No Data",
              gstr3bStatus: gstr3b?.overallStatus ?? "No Data",
              gstr1Latest: gstr1?.latestFiledOn ?? undefined,
              gstr3bLatest: gstr3b?.latestFiledOn ?? undefined,
            };
            return n;
          });
        } else {
          throw new Error((result as any).message || "Failed to fetch GST details.");
        }
      } catch (err: any) {
        if (err.name === "AbortError") break;
        setItems(prev => { const n = [...prev]; n[idx] = { ...n[idx], status: "error", error: err.message || "API request failed" }; return n; });
      }

      indexRef.current++;
      setCurrentIndex(indexRef.current);

      if (indexRef.current < items.length && processingRef.current) {
        await new Promise(r => setTimeout(r, delayMs));
      }
    }

    setIsProcessing(false);
    processingRef.current = false;
  }

  function handleStart() {
    // Only reset idle/error rows; keep already-skipped rows as skipped
    setItems(prev => prev.map(item =>
      item.status === "skipped" ? item : { ...item, status: "idle", error: undefined }
    ));
    indexRef.current = 0;
    setCurrentIndex(0);
    void runBatch();
  }

  function handlePause() {
    processingRef.current = false;
    setIsPaused(true);
    abortRef.current?.abort();
  }

  function handleResume() {
    indexRef.current = currentIndex;
    void runBatch();
  }

  function handleReset() {
    handlePause();
    indexRef.current = 0;
    setCurrentIndex(0);
    setItems(prev => prev.map(item =>
      item.status === "skipped" ? item : { ...item, status: "idle", error: undefined }
    ));
    setIsPaused(false);
    setIsProcessing(false);
    setWriteStatus(null);
  }

  async function handlePushToSheet() {
    if (!appsScriptUrl.trim()) {
      setWriteStatus({ success: false, message: "Please paste your Google Apps Script Web App URL first." });
      return;
    }

    setWritingToSheet(true);
    setWriteStatus(null);

    const g1Col = colGstr1(fyLabel);
    const g3bCol = colGstr3b(fyLabel);
    const g1DateCol = colGstr1Date(fyLabel);
    const g3bDateCol = colGstr3bDate(fyLabel);

    const updates = items
      .filter(item => item.status === "success")
      .map(item => ({
        gstin: item.gstin,
        values: {
          [g1Col]: item.gstr1Status ?? "",
          [g3bCol]: item.gstr3bStatus ?? "",
          ...(item.gstr1Latest ? { [g1DateCol]: item.gstr1Latest } : {}),
          ...(item.gstr3bLatest ? { [g3bDateCol]: item.gstr3bLatest } : {}),
        },
      }));

    if (updates.length === 0) {
      setWriteStatus({ success: false, message: "No successful records to push. Run the batch first." });
      setWritingToSheet(false);
      return;
    }

    try {
      await fetch(appsScriptUrl, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          spreadsheetId,
          gid,
          gstinColumn,
          columns: [g1Col, g3bCol, g1DateCol, g3bDateCol],
          updates,
        }),
      });

      setWriteStatus({ success: true, message: `Sheet update sent for ${updates.length} rows (FY ${fyLabel}). Check your sheet in a few seconds.` });
    } catch (err) {
      setWriteStatus({ success: false, message: err instanceof Error ? err.message : "Failed to connect to Apps Script." });
    } finally {
      setWritingToSheet(false);
    }
  }

  function copyScript() {
    void navigator.clipboard.writeText(APPS_SCRIPT_CODE);
    setScriptCopied(true);
    setTimeout(() => setScriptCopied(false), 2000);
  }

  const total = items.length;
  const skippedCount = items.filter(i => i.status === "skipped").length;
  const successCount = items.filter(i => i.status === "success").length;
  const errorCount = items.filter(i => i.status === "error").length;
  const toProcess = items.filter(i => i.status !== "skipped").length;
  const processed = items.filter(i => i.status === "success" || i.status === "error").length;
  const progressPercent = toProcess > 0 ? Math.round((processed / toProcess) * 100) : 0;

  function statusBadge(status: string) {
    const s = status.toLowerCase();
    if (s === "filed" || s.startsWith("filed")) return "bg-emerald-50 text-emerald-700 ring-emerald-600/20";
    if (s === "not filed") return "bg-red-50 text-red-700 ring-red-600/20";
    if (s.startsWith("partial")) return "bg-amber-50 text-amber-700 ring-amber-600/20";
    return "bg-slate-50 text-slate-500 ring-slate-200";
  }

  return (
    <main className="relative isolate min-h-screen overflow-hidden px-4 py-8 sm:px-6 lg:px-8">
      {/* Background */}
      <div className="pointer-events-none absolute inset-0">
        <div className="absolute left-[-8rem] top-[-6rem] h-72 w-72 rounded-full bg-[color:var(--accent)]/15 blur-3xl" />
        <div className="absolute right-[-7rem] top-24 h-80 w-80 rounded-full bg-[#d4a24a]/20 blur-3xl" />
        <div className="absolute inset-0 bg-[linear-gradient(rgba(16,33,58,0.04)_1px,transparent_1px),linear-gradient(90deg,rgba(16,33,58,0.04)_1px,transparent_1px)] bg-[size:56px_56px]" />
      </div>

      <div className="relative mx-auto flex w-full max-w-7xl flex-col gap-6">

        {/* ── Header ── */}
        <header className="overflow-hidden rounded-[2.5rem] border border-[color:var(--border)] bg-[color:var(--surface)] px-6 py-6 shadow-[0_24px_90px_rgba(16,33,58,0.12)] backdrop-blur-xl sm:px-10 sm:py-10">
          <div className="flex flex-wrap items-center gap-3 text-xs font-semibold uppercase tracking-[0.28em] text-[color:var(--muted)]">
            <span className="rounded-full border border-[color:var(--border)] bg-white/70 px-3 py-1">Bulk Processor</span>
            <span className="rounded-full border border-[color:var(--border)] bg-white/70 px-3 py-1">Google Sheet Sync</span>
            <span className="rounded-full border border-[color:var(--border)] bg-white/70 px-3 py-1">Year-wise Status</span>
          </div>

          <div className="mt-5 grid gap-6 lg:grid-cols-[1.2fr_0.8fr] lg:items-end">
            <div className="space-y-3">
              <p className="text-sm font-semibold uppercase tracking-[0.24em] text-[color:var(--accent)]">GST Filing Status Tracker</p>
              <h1 className="max-w-3xl text-4xl font-semibold tracking-tight text-[color:var(--foreground)] sm:text-5xl">
                Sync year-wise GST status directly to your sheet.
              </h1>
              <p className="max-w-2xl text-base leading-7 text-[color:var(--muted)]">
                Load your Google Sheet, pick the FY, run the batch — GSTR-1 and GSTR-3B status columns are added automatically. Already-synced rows are skipped.
              </p>
            </div>

            {/* Metrics */}
            <div className="grid gap-3 rounded-[2rem] border border-[color:var(--border)] bg-[color:var(--surface-strong)] p-5 shadow-[0_18px_50px_rgba(16,33,58,0.08)]">
              <div className="flex items-center justify-between gap-4">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-[0.24em] text-[color:var(--muted)]">Progress</p>
                  <p className="mt-1 text-lg font-semibold text-[color:var(--foreground)]">{processed} of {toProcess} processed</p>
                  {skippedCount > 0 && <p className="text-[11px] text-[color:var(--muted)]">{skippedCount} already synced → skipped</p>}
                </div>
                <span className={`rounded-full px-3 py-1 text-xs font-semibold ring-1 ${
                  isProcessing ? "bg-[color:var(--warning)]/10 text-[color:var(--warning)] ring-[color:var(--warning)]/20"
                  : isPaused ? "bg-amber-50 text-amber-700 ring-amber-200"
                  : successCount > 0 && processed === toProcess && toProcess > 0 ? "bg-[color:var(--success)]/10 text-[color:var(--success)] ring-[color:var(--success)]/20"
                  : "bg-slate-100 text-slate-600 ring-slate-200"
                }`}>
                  {isProcessing ? "Processing…" : isPaused ? "Paused" : processed === toProcess && toProcess > 0 ? "Done" : "Idle"}
                </span>
              </div>

              <div className="relative h-2.5 w-full overflow-hidden rounded-full bg-slate-200">
                <div className="h-full bg-[color:var(--accent)] transition-all duration-300" style={{ width: `${progressPercent}%` }} />
              </div>

              <div className="grid grid-cols-4 gap-2 text-center text-sm mt-1">
                {[
                  { label: "Total", value: total, cls: "" },
                  { label: "Skipped", value: skippedCount, cls: "text-slate-400" },
                  { label: "Success", value: successCount, cls: "text-[color:var(--success)]" },
                  { label: "Failed", value: errorCount, cls: "text-[color:var(--danger)]" },
                ].map(({ label, value, cls }) => (
                  <div key={label} className="rounded-xl bg-white/70 p-2 border border-[color:var(--border)]">
                    <p className={`text-[0.6rem] font-semibold uppercase tracking-wider text-[color:var(--muted)]`}>{label}</p>
                    <p className={`text-base font-semibold ${cls}`}>{value}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </header>

        {/* ── Config + Controls ── */}
        <section className="grid gap-6 lg:grid-cols-[1.1fr_0.9fr]">
          <div className="flex flex-col gap-6">

            {/* Sheet loader */}
            <form onSubmit={handleLoadSheet} className="overflow-hidden rounded-[2.5rem] border border-[color:var(--border)] bg-[color:var(--surface)] p-6 shadow-[0_24px_90px_rgba(16,33,58,0.12)] backdrop-blur-xl sm:p-8">
              <h2 className="text-xl font-semibold text-[color:var(--foreground)] mb-4">Google Sheet Connection</h2>
              <div className="space-y-4">
                <label className="block space-y-1.5">
                  <span className="text-xs font-semibold uppercase tracking-wider text-[color:var(--muted)]">Google Sheet URL</span>
                  <div className="flex gap-2">
                    <input
                      type="url"
                      value={sheetUrl}
                      onChange={e => setSheetUrl(e.target.value)}
                      className="w-full rounded-2xl border border-[color:var(--border)] bg-white/85 px-4 py-3 text-sm font-medium text-[color:var(--foreground)] outline-none transition focus:border-[color:var(--accent)] focus:bg-white"
                      placeholder="https://docs.google.com/spreadsheets/d/.../edit"
                    />
                    <button
                      type="submit"
                      disabled={loadingSheet || isProcessing}
                      className="inline-flex items-center justify-center rounded-2xl bg-[color:var(--accent)] px-5 text-sm font-semibold text-white shadow-[0_10px_20px_rgba(23,58,109,0.18)] transition hover:bg-[color:var(--accent-strong)] disabled:opacity-60 disabled:cursor-not-allowed"
                    >
                      {loadingSheet ? "Loading…" : "Load"}
                    </button>
                  </div>
                </label>

                {sheetError && (
                  <div className="rounded-xl bg-[color:var(--danger)]/8 border border-[color:var(--danger)]/20 p-3 text-xs text-[color:var(--danger)]">{sheetError}</div>
                )}

                <div className="grid gap-4 sm:grid-cols-2 pt-1">
                  {headers.length > 0 && (
                    <label className="block space-y-1.5">
                      <span className="text-xs font-semibold uppercase tracking-wider text-[color:var(--muted)]">GSTIN Column</span>
                      <select
                        value={gstinColumn}
                        onChange={e => setGstinColumn(e.target.value)}
                        className="w-full rounded-2xl border border-[color:var(--border)] bg-white/85 px-4 py-3 text-sm font-medium text-[color:var(--foreground)] outline-none transition focus:border-[color:var(--accent)]"
                      >
                        {headers.map(h => <option key={h} value={h}>{h}</option>)}
                      </select>
                    </label>
                  )}
                  <label className="block space-y-1.5">
                    <span className="text-xs font-semibold uppercase tracking-wider text-[color:var(--muted)]">Financial Year</span>
                    <select
                      value={fy}
                      onChange={e => setFy(e.target.value)}
                      className="w-full rounded-2xl border border-[color:var(--border)] bg-white/85 px-4 py-3 text-sm font-medium text-[color:var(--foreground)] outline-none transition focus:border-[color:var(--accent)]"
                    >
                      {FY_OPTIONS.map(y => <option key={y} value={y}>{`${y}-${String(Number(y) + 1).slice(-2)}`}</option>)}
                    </select>
                  </label>
                </div>

                {/* FY columns info */}
                {fy && (
                  <div className="rounded-2xl bg-[color:var(--accent)]/5 border border-[color:var(--accent)]/15 px-4 py-3 text-xs text-[color:var(--accent)] space-y-1">
                    <p className="font-semibold">Columns that will be written to sheet:</p>
                    <p className="font-mono">{colGstr1(fyLabel)}, {colGstr3b(fyLabel)}, {colGstr1Date(fyLabel)}, {colGstr3bDate(fyLabel)}</p>
                    <p className="text-[color:var(--muted)] text-[11px] mt-1">Rows where these columns are already filled will be automatically skipped.</p>
                  </div>
                )}
              </div>
            </form>

            {/* Run controls */}
            <div className="overflow-hidden rounded-[2.5rem] border border-[color:var(--border)] bg-[color:var(--surface)] p-6 shadow-[0_24px_90px_rgba(16,33,58,0.12)] backdrop-blur-xl sm:p-8">
              <h2 className="text-xl font-semibold text-[color:var(--foreground)] mb-4">Run Controls & Write-Back</h2>
              <div className="space-y-4">
                <label className="block space-y-1.5">
                  <span className="text-xs font-semibold uppercase tracking-wider text-[color:var(--muted)]">Google Apps Script Web App URL</span>
                  <input
                    type="url"
                    value={appsScriptUrl}
                    onChange={e => setAppsScriptUrl(e.target.value)}
                    className="w-full rounded-2xl border border-[color:var(--border)] bg-white/85 px-4 py-3 text-sm font-medium text-[color:var(--foreground)] outline-none transition focus:border-[color:var(--accent)]"
                    placeholder="https://script.google.com/macros/s/.../exec"
                  />
                  <button
                    type="button"
                    onClick={() => setShowScriptSetup(!showScriptSetup)}
                    className="text-xs font-semibold text-[color:var(--accent)] hover:underline mt-1"
                  >
                    {showScriptSetup ? "Hide setup guide" : "How to set up Google Apps Script? →"}
                  </button>
                </label>

                {showScriptSetup && (
                  <div className="rounded-2xl border border-[color:var(--border)] bg-white/70 p-4 text-xs space-y-3 leading-relaxed text-[color:var(--muted)]">
                    <p className="font-semibold text-[color:var(--foreground)] text-sm">Step-by-Step Google Sheet Write Setup:</p>
                    <ol className="list-decimal pl-4 space-y-1.5">
                      <li>In your Google Sheet → <strong>Extensions › Apps Script</strong>.</li>
                      <li>Delete existing code and paste the snippet below.</li>
                      <li>Save → <strong>Deploy › New deployment › Web app</strong>.</li>
                      <li>Execute as: <strong>Me</strong> · Who has access: <strong>Anyone</strong>.</li>
                      <li>Click Deploy, authorize, copy the <strong>Web App URL</strong> → paste above.</li>
                    </ol>
                    <div className="relative mt-2">
                      <pre className="max-h-48 overflow-y-auto bg-slate-900 text-slate-100 p-3 rounded-xl font-mono text-[10px] leading-relaxed">{APPS_SCRIPT_CODE}</pre>
                      <button
                        type="button"
                        onClick={copyScript}
                        className="absolute right-2 top-2 rounded bg-slate-800 hover:bg-slate-700 px-2 py-1 text-[10px] text-white font-medium border border-slate-700 transition"
                      >
                        {scriptCopied ? "Copied!" : "Copy Code"}
                      </button>
                    </div>
                  </div>
                )}

                <label className="block space-y-1">
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-[color:var(--muted)]">Delay between requests</span>
                  <select
                    value={delayMs}
                    onChange={e => setDelayMs(Number(e.target.value))}
                    className="w-full rounded-xl border border-[color:var(--border)] bg-white px-3 py-2 text-xs font-medium text-[color:var(--foreground)] outline-none"
                  >
                    <option value={200}>200 ms (Fast)</option>
                    <option value={500}>500 ms</option>
                    <option value={1000}>1 second (Safe)</option>
                    <option value={2000}>2 seconds</option>
                  </select>
                </label>

                <div className="flex flex-wrap gap-2 pt-1">
                  {!isProcessing && !isPaused && (
                    <button
                      type="button"
                      onClick={handleStart}
                      disabled={items.filter(i => i.status !== "skipped").length === 0}
                      className="flex-1 rounded-2xl bg-[color:var(--accent)] py-3 px-4 text-sm font-semibold text-white shadow-lg transition hover:bg-[color:var(--accent-strong)] disabled:opacity-60 disabled:cursor-not-allowed"
                    >
                      {skippedCount > 0 ? `Start Batch (${toProcess} remaining)` : "Start Lookup Batch"}
                    </button>
                  )}
                  {isProcessing && (
                    <button type="button" onClick={handlePause}
                      className="flex-1 rounded-2xl bg-[color:var(--warning)] py-3 px-4 text-sm font-semibold text-white shadow-lg transition hover:bg-[color:var(--warning)]/90">
                      Pause Batch
                    </button>
                  )}
                  {isPaused && (
                    <button type="button" onClick={handleResume}
                      className="flex-1 rounded-2xl bg-[color:var(--success)] py-3 px-4 text-sm font-semibold text-white shadow-lg transition hover:bg-[color:var(--success)]/90">
                      Resume Batch
                    </button>
                  )}
                  <button type="button" onClick={handleReset} disabled={items.length === 0}
                    className="rounded-2xl bg-white/70 py-3 px-4 text-sm font-semibold text-[color:var(--foreground)] border border-[color:var(--border)] hover:bg-white disabled:opacity-60">
                    Reset
                  </button>
                  <button
                    type="button"
                    onClick={handlePushToSheet}
                    disabled={writingToSheet || successCount === 0 || !appsScriptUrl}
                    className="w-full rounded-2xl bg-[color:var(--success)] py-3 px-4 text-sm font-semibold text-white shadow-lg transition hover:bg-[color:var(--success)]/90 disabled:opacity-50 disabled:cursor-not-allowed mt-1"
                  >
                    {writingToSheet ? "Updating Sheet…" : `Push ${successCount} Rows to Sheet (FY ${fyLabel})`}
                  </button>
                </div>

                {writeStatus && (
                  <div className={`rounded-xl border p-3 text-xs ${writeStatus.success ? "bg-[color:var(--success)]/8 border-[color:var(--success)]/20 text-[color:var(--success)]" : "bg-[color:var(--danger)]/8 border-[color:var(--danger)]/20 text-[color:var(--danger)]"}`}>
                    {writeStatus.message}
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Sheet preview */}
          <aside className="overflow-hidden rounded-[2.5rem] border border-[color:var(--border)] bg-[color:var(--surface)] p-6 shadow-[0_24px_90px_rgba(16,33,58,0.12)] backdrop-blur-xl sm:p-8 flex flex-col max-h-[600px]">
            <div className="mb-4">
              <h2 className="text-xl font-semibold text-[color:var(--foreground)]">Sheet Preview</h2>
              <p className="text-xs text-[color:var(--muted)] mt-1">{sheetRows.length} rows loaded</p>
            </div>
            <div className="overflow-auto border border-[color:var(--border)] rounded-2xl bg-white/60 flex-1">
              {sheetRows.length > 0 ? (
                <table className="min-w-full border-collapse text-left text-xs">
                  <thead className="bg-slate-100/80 sticky top-0 border-b border-[color:var(--border)]">
                    <tr>
                      <th className="px-3 py-2 font-semibold text-slate-700">Row</th>
                      {headers.slice(0, 4).map(h => <th key={h} className="px-3 py-2 font-semibold text-slate-700">{h}</th>)}
                      {headers.length > 4 && <th className="px-3 py-2 text-slate-400">…</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {sheetRows.slice(0, 12).map((row, idx) => (
                      <tr key={idx} className="border-t border-[color:var(--border)] hover:bg-white/40">
                        <td className="px-3 py-2 text-slate-400 font-mono">{idx + 2}</td>
                        {headers.slice(0, 4).map(h => (
                          <td key={h} className="px-3 py-2 text-[color:var(--foreground)] truncate max-w-[120px]">{row[h] || "—"}</td>
                        ))}
                        {headers.length > 4 && <td className="px-3 py-2 text-slate-300">…</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="h-full flex items-center justify-center p-8 text-center text-slate-400 text-xs">
                  No sheet loaded. Paste a URL and click Load.
                </div>
              )}
            </div>
            {sheetRows.length > 12 && (
              <p className="text-[10px] text-center text-[color:var(--muted)] mt-2">Showing first 12 of {sheetRows.length} rows</p>
            )}
          </aside>
        </section>

        {/* ── Results Table ── */}
        <section className="overflow-hidden rounded-[2.5rem] border border-[color:var(--border)] bg-[color:var(--surface-strong)] shadow-[0_24px_90px_rgba(16,33,58,0.12)] backdrop-blur-xl">
          <div className="border-b border-[color:var(--border)] px-6 py-5 sm:px-8 flex justify-between items-center flex-wrap gap-4">
            <div>
              <h3 className="text-xl font-semibold text-[color:var(--foreground)]">
                GST Filing Status — FY {fyLabel}
              </h3>
              <p className="text-xs text-[color:var(--muted)] mt-0.5">Year-wise GSTR-1 &amp; GSTR-3B status per GSTIN</p>
            </div>
            {skippedCount > 0 && (
              <span className="rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-500 ring-1 ring-slate-200">
                {skippedCount} already synced
              </span>
            )}
          </div>

          <div className="overflow-x-auto max-h-[520px]">
            <table className="min-w-full border-collapse text-left text-sm">
              <thead className="bg-slate-50 sticky top-0 border-b border-[color:var(--border)] text-xs font-semibold uppercase tracking-wider text-[color:var(--muted)]">
                <tr>
                  <th className="px-5 py-3.5">Row</th>
                  <th className="px-5 py-3.5">GSTIN</th>
                  <th className="px-5 py-3.5">GSTR-1 {fyLabel}</th>
                  <th className="px-5 py-3.5">GSTR-3B {fyLabel}</th>
                  <th className="px-5 py-3.5">Latest GSTR-1</th>
                  <th className="px-5 py-3.5">Latest GSTR-3B</th>
                  <th className="px-5 py-3.5">Status</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[color:var(--border)] bg-white/50">
                {items.length > 0 ? items.map((item, idx) => (
                  <tr
                    key={idx}
                    className={`hover:bg-slate-50/50 transition-colors ${
                      item.status === "fetching" ? "bg-blue-50/30" :
                      item.status === "error" ? "bg-red-50/10" :
                      item.status === "skipped" ? "bg-slate-50/60 opacity-60" : ""
                    }`}
                  >
                    <td className="px-5 py-4 font-mono text-xs text-slate-400">{item.rowNumber}</td>
                    <td className="px-5 py-4 font-mono text-xs font-medium text-[color:var(--foreground)]">{item.gstin || "—"}</td>

                    {/* GSTR-1 status */}
                    <td className="px-5 py-4 text-xs">
                      {item.gstr1Status ? (
                        <span className={`inline-flex rounded-full px-2 py-0.5 font-semibold ring-1 text-[10px] ${statusBadge(item.gstr1Status)}`}>
                          {item.gstr1Status}
                        </span>
                      ) : item.status === "fetching" ? <span className="text-blue-400 animate-pulse">…</span> : "—"}
                    </td>

                    {/* GSTR-3B status */}
                    <td className="px-5 py-4 text-xs">
                      {item.gstr3bStatus ? (
                        <span className={`inline-flex rounded-full px-2 py-0.5 font-semibold ring-1 text-[10px] ${statusBadge(item.gstr3bStatus)}`}>
                          {item.gstr3bStatus}
                        </span>
                      ) : item.status === "fetching" ? <span className="text-blue-400 animate-pulse">…</span> : "—"}
                    </td>

                    <td className="px-5 py-4 text-xs font-mono text-[color:var(--muted)]">{item.gstr1Latest || "—"}</td>
                    <td className="px-5 py-4 text-xs font-mono text-[color:var(--muted)]">{item.gstr3bLatest || "—"}</td>

                    {/* Row process status */}
                    <td className="px-5 py-4">
                      <span className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ${
                        item.status === "success" ? "bg-emerald-50 text-emerald-700 ring-emerald-600/20" :
                        item.status === "fetching" ? "bg-blue-50 text-blue-700 ring-blue-600/20 animate-pulse" :
                        item.status === "error" ? "bg-red-50 text-red-700 ring-red-600/20" :
                        item.status === "skipped" ? "bg-slate-100 text-slate-500 ring-slate-200" :
                        "bg-slate-50 text-slate-400 ring-slate-200"
                      }`}>
                        {item.status === "success" && "✓ Done"}
                        {item.status === "fetching" && "Fetching…"}
                        {item.status === "error" && (item.error || "Error")}
                        {item.status === "skipped" && "Already Synced"}
                        {item.status === "idle" && "Pending"}
                      </span>
                    </td>
                  </tr>
                )) : (
                  <tr>
                    <td colSpan={7} className="px-5 py-14 text-center text-sm text-[color:var(--muted)]">
                      Load a Google Sheet to view and process GSTINs.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}
