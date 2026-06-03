"use client";

import { useState, useRef, useEffect } from "react";

type SheetRow = Record<string, string>;

type ProcessedItem = {
  gstin: string;
  rowNumber: number;
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

    if (!sheet) return json({ success: false, error: "Sheet not found" });

    const sheetData = sheet.getDataRange().getValues();
    let headers = sheetData[0].map(h => h.toString().trim());

    data.columns.forEach(col => {
      if (!headers.includes(col)) {
        sheet.getRange(1, headers.length + 1).setValue(col);
        headers.push(col);
      }
    });

    const fresh = sheet.getDataRange().getValues();
    const freshHeaders = fresh[0].map(h => h.toString().trim());
    const fMap = {};
    freshHeaders.forEach((h, i) => { fMap[h.toLowerCase()] = i; });

    const gstinColIdx = fMap[data.gstinColumn.toLowerCase()];
    if (gstinColIdx === undefined) return json({ success: false, error: "GSTIN column not found" });

    for (let i = 1; i < fresh.length; i++) {
      const rowGstin = (fresh[i][gstinColIdx] || "").toString().trim().toUpperCase().replace(/\\s+/g, "");
      if (!rowGstin) continue;
      const match = data.updates.find(u => u.gstin.trim().toUpperCase().replace(/\\s+/g, "") === rowGstin);
      if (!match) continue;

      for (const [colName, value] of Object.entries(match.values)) {
        const colIdx = fMap[colName.toLowerCase()];
        if (colIdx === undefined) continue;
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
  const [sheetUrl, setSheetUrl] = useState("");
  const [appsScriptUrl, setAppsScriptUrl] = useState("");
  const [fy, setFy] = useState(DEFAULT_FY);
  const [gstinColumn, setGstinColumn] = useState("");
  const [delayMs, setDelayMs] = useState(1000);

  const [headers, setHeaders] = useState<string[]>([]);
  const [sheetRows, setSheetRows] = useState<SheetRow[]>([]);
  const [spreadsheetId, setSpreadsheetId] = useState("");
  const [gid, setGid] = useState("");

  const [items, setItems] = useState<ProcessedItem[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isPaused, setIsPaused] = useState(false);
  const [currentIndex, setCurrentIndex] = useState(0);

  const [loadingSheet, setLoadingSheet] = useState(false);
  const [sheetError, setSheetError] = useState<string | null>(null);
  const [showScriptSetup, setShowScriptSetup] = useState(false);
  const [scriptCopied, setScriptCopied] = useState(false);
  const [writeStatus, setWriteStatus] = useState<{ success?: boolean; message?: string } | null>(null);
  const [writingToSheet, setWritingToSheet] = useState(false);

  const processingRef = useRef(false);
  const indexRef = useRef(0);
  const abortRef = useRef<AbortController | null>(null);

  const fyLabel = `${fy}-${String(Number(fy) + 1).slice(-2)}`;

  // ── Load saved URLs from localStorage on mount ──
  useEffect(() => {
    const savedSheet = localStorage.getItem("gst_sheet_url") ?? "";
    const savedScript = localStorage.getItem("gst_apps_script_url") ?? "";
    if (savedSheet) setSheetUrl(savedSheet);
    if (savedScript) setAppsScriptUrl(savedScript);
    if (savedSheet) void loadSheet(savedSheet);
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // ── Persist URLs whenever they change ──
  const handleSheetUrlChange = (v: string) => {
    setSheetUrl(v);
    localStorage.setItem("gst_sheet_url", v);
  };
  const handleAppsScriptUrlChange = (v: string) => {
    setAppsScriptUrl(v);
    localStorage.setItem("gst_apps_script_url", v);
  };

  async function loadSheet(url: string) {
    if (!url.trim()) return;
    setLoadingSheet(true);
    setSheetError(null);
    setWriteStatus(null);
    try {
      const res = await fetch("/api/sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ sheetUrl: url }),
      });
      const data = await res.json();
      if (!data.success) throw new Error(data.message || "Failed to parse Google Sheet.");

      setHeaders(data.headers);
      setSheetRows(data.rows);
      setSpreadsheetId(data.spreadsheetId);
      setGid(data.gid || "");

      const gstinCol = data.headers.find((h: string) => {
        const n = h.toLowerCase().replace(/[^a-z0-9]/g, "");
        return n.includes("gstin") || n.includes("gstnumber") || n.includes("gstno") || n === "gst";
      });
      const colKey = gstinCol || data.headers[0] || "";
      setGstinColumn(colKey);

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

  async function handleLoadSheet(e: React.FormEvent) {
    e.preventDefault();
    await loadSheet(sheetUrl);
  }

  async function runBatch() {
    processingRef.current = true;
    setIsProcessing(true);
    setIsPaused(false);

    while (indexRef.current < items.length && processingRef.current) {
      const idx = indexRef.current;
      const item = items[idx];

      if (item.status === "skipped") {
        indexRef.current++;
        setCurrentIndex(indexRef.current);
        continue;
      }

      if (!item.gstin || item.gstin.length !== 15) {
        setItems(prev => {
          const n = [...prev];
          n[idx] = { ...n[idx], status: "error", error: !item.gstin ? "Empty GSTIN" : "Invalid GSTIN" };
          return n;
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
          // Match returnType regardless of hyphens: "GSTR1", "GSTR-1", "gstr1" all match
          const normalize = (s: string) => s.replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
          const gstr1 = result.returnStatuses.find(s => normalize(s.returnType) === "GSTR1");
          const gstr3b = result.returnStatuses.find(s => normalize(s.returnType) === "GSTR3B");

          setItems(prev => {
            const n = [...prev];
            n[idx] = {
              ...n[idx],
              status: "success",
              gstr1Status: gstr1?.overallStatus ?? "Not Filed",
              gstr3bStatus: gstr3b?.overallStatus ?? "Not Filed",
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
        setItems(prev => { const n = [...prev]; n[idx] = { ...n[idx], status: "error", error: err.message || "API error" }; return n; });
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
      setWriteStatus({ success: false, message: "Paste your Apps Script Web App URL in Step 3 first." });
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
      setWriteStatus({ success: false, message: "No records fetched yet. Run the batch first." });
      setWritingToSheet(false);
      return;
    }

    try {
      await fetch(appsScriptUrl, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ spreadsheetId, gid, gstinColumn, columns: [g1Col, g3bCol, g1DateCol, g3bDateCol], updates }),
      });
      setWriteStatus({ success: true, message: `Sent ${updates.length} rows to sheet (FY ${fyLabel}). Check your sheet.` });
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
    if (s === "filed") return "badge-filed";
    if (s === "not filed") return "badge-notfiled";
    return "badge-nodata";
  }

  return (
    <main className="gst-root">
      {/* ── Blobs ── */}
      <div className="blob blob-1" />
      <div className="blob blob-2" />

      <div className="gst-container">

        {/* ── Header ── */}
        <header className="card hero-card">
          <div className="hero-chips">
            <span className="chip">Auto-save Settings</span>
            <span className="chip">GSTR-1 &amp; GSTR-3B</span>
            <span className="chip">Skip Already-Synced</span>
          </div>
          <div className="hero-body">
            <div>
              <p className="hero-eyebrow">GST Filing Status Tracker</p>
              <h1 className="hero-title">Check &amp; sync year-wise GST filing status to your sheet.</h1>
              <p className="hero-sub">Your Google Sheet URL and Apps Script URL are saved automatically — no re-entry needed.</p>
            </div>
            {/* Progress widget */}
            <div className="progress-card">
              <div className="progress-top">
                <div>
                  <p className="progress-label">Progress</p>
                  <p className="progress-count">{processed} / {toProcess} processed</p>
                  {skippedCount > 0 && <p className="progress-skip">{skippedCount} already synced, skipped</p>}
                </div>
                <span className={`status-pill ${isProcessing ? "pill-processing" : isPaused ? "pill-paused" : processed === toProcess && toProcess > 0 ? "pill-done" : "pill-idle"}`}>
                  {isProcessing ? "Running…" : isPaused ? "Paused" : processed === toProcess && toProcess > 0 ? "Done ✓" : "Idle"}
                </span>
              </div>
              <div className="progress-bar-track">
                <div className="progress-bar-fill" style={{ width: `${progressPercent}%` }} />
              </div>
              <div className="progress-stats">
                {[
                  { label: "Total", value: total, cls: "" },
                  { label: "Skipped", value: skippedCount, cls: "stat-skip" },
                  { label: "Done", value: successCount, cls: "stat-done" },
                  { label: "Failed", value: errorCount, cls: "stat-fail" },
                ].map(({ label, value, cls }) => (
                  <div key={label} className="stat-box">
                    <p className="stat-label">{label}</p>
                    <p className={`stat-value ${cls}`}>{value}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </header>

        {/* ── 3-Step Grid ── */}
        <section className="steps-grid">

          {/* Step 1 */}
          <div className="card step-card">
            <div className="step-header">
              <span className="step-num">1</span>
              <h2 className="step-title">Connect Sheet</h2>
            </div>
            <form onSubmit={handleLoadSheet} className="step-body">
              <label className="field-label">Google Sheet URL
                <div className="input-row">
                  <input
                    type="url"
                    value={sheetUrl}
                    onChange={e => handleSheetUrlChange(e.target.value)}
                    className="text-input"
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                  />
                  <button type="submit" disabled={loadingSheet || isProcessing} className="btn btn-primary btn-sm">
                    {loadingSheet ? "…" : "Load"}
                  </button>
                </div>
              </label>
              {sheetError && <div className="error-box">{sheetError}</div>}
              <div className="two-col">
                <label className="field-label">GSTIN Column
                  {headers.length > 0 ? (
                    <select value={gstinColumn} onChange={e => setGstinColumn(e.target.value)} className="select-input">
                      {headers.map(h => <option key={h} value={h}>{h}</option>)}
                    </select>
                  ) : <div className="placeholder-select">Load sheet first</div>}
                </label>
                <label className="field-label">Financial Year
                  <select value={fy} onChange={e => setFy(e.target.value)} className="select-input">
                    {FY_OPTIONS.map(y => <option key={y} value={y}>{y}-{String(Number(y)+1).slice(-2)}</option>)}
                  </select>
                </label>
              </div>
            </form>
            <div className="step-footer">
              <span className="footer-label">Columns that will be written</span>
              <code className="footer-code">GSTR1 {fyLabel} &nbsp;·&nbsp; GSTR3B {fyLabel}</code>
            </div>
          </div>

          {/* Step 2 */}
          <div className="card step-card">
            <div className="step-header">
              <span className="step-num">2</span>
              <h2 className="step-title">Fetch Filing Status</h2>
            </div>
            <div className="step-body">
              <label className="field-label">Delay Between Requests
                <select value={delayMs} onChange={e => setDelayMs(Number(e.target.value))} className="select-input">
                  <option value={200}>Fast — 200 ms</option>
                  <option value={500}>Balanced — 500 ms</option>
                  <option value={1000}>Safe — 1 second (recommended)</option>
                  <option value={2000}>Slow — 2 seconds</option>
                </select>
              </label>
              <div className="btn-group">
                {!isProcessing && !isPaused && (
                  <button onClick={handleStart} disabled={toProcess === 0} className="btn btn-primary">
                    {skippedCount > 0 ? `Run Batch (${toProcess} rows)` : "Run Batch"}
                  </button>
                )}
                {isProcessing && <button onClick={handlePause} className="btn btn-warning">Pause</button>}
                {isPaused && <button onClick={handleResume} className="btn btn-success">Resume</button>}
                <button onClick={handleReset} disabled={items.length === 0} className="btn btn-ghost">Reset</button>
              </div>
            </div>
            <div className="step-footer">
              <span className="footer-label">Progress</span>
              <span className="footer-value">{processed} done · {progressPercent}%</span>
            </div>
          </div>

          {/* Step 3 */}
          <div className="card step-card">
            <div className="step-header">
              <span className="step-num">3</span>
              <h2 className="step-title">Push to Sheet</h2>
            </div>
            <div className="step-body">
              <label className="field-label">Apps Script Web App URL
                <input
                  type="url"
                  value={appsScriptUrl}
                  onChange={e => handleAppsScriptUrlChange(e.target.value)}
                  className="text-input"
                  placeholder="https://script.google.com/macros/s/.../exec"
                />
              </label>
              <button type="button" onClick={() => setShowScriptSetup(true)} className="link-btn">
                How to set up Apps Script? →
              </button>
              <button
                onClick={handlePushToSheet}
                disabled={writingToSheet || successCount === 0 || !appsScriptUrl}
                className="btn btn-success"
              >
                {writingToSheet ? "Pushing…" : `Push ${successCount} Records to Sheet`}
              </button>
              {writeStatus && (
                <div className={writeStatus.success ? "success-box" : "error-box"}>{writeStatus.message}</div>
              )}
            </div>
            <div className="step-footer">
              <span className="footer-label">Ready to push</span>
              <span className="footer-value">{successCount} rows · FY {fyLabel}</span>
            </div>
          </div>
        </section>

        {/* ── Results Table ── */}
        <section className="card results-card">
          <div className="results-header">
            <div>
              <h3 className="results-title">Filing Status — FY {fyLabel}</h3>
              <p className="results-sub">Live status for each GSTIN from the taxpayer portal</p>
            </div>
            {skippedCount > 0 && <span className="chip">{skippedCount} already synced</span>}
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Row</th>
                  <th>GSTIN</th>
                  <th>GSTR-1 ({fyLabel})</th>
                  <th>GSTR-3B ({fyLabel})</th>
                  <th>Latest Filed</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {items.length > 0 ? items.map((item, idx) => (
                  <tr key={idx} className={
                    item.status === "fetching" ? "row-fetching" :
                    item.status === "error" ? "row-error" :
                    item.status === "skipped" ? "row-skipped" : ""
                  }>
                    <td className="td-mono muted">{item.rowNumber}</td>
                    <td className="td-mono bold">{item.gstin || "—"}</td>
                    <td>
                      {item.gstr1Status
                        ? <span className={`badge ${statusBadge(item.gstr1Status)}`}>{item.gstr1Status}</span>
                        : item.status === "fetching" ? <span className="fetching-dot">…</span> : <span className="muted">—</span>}
                    </td>
                    <td>
                      {item.gstr3bStatus
                        ? <span className={`badge ${statusBadge(item.gstr3bStatus)}`}>{item.gstr3bStatus}</span>
                        : item.status === "fetching" ? <span className="fetching-dot">…</span> : <span className="muted">—</span>}
                    </td>
                    <td className="td-mono muted small">
                      {item.gstr1Latest || item.gstr3bLatest
                        ? `${item.gstr1Latest ?? "—"} / ${item.gstr3bLatest ?? "—"}`
                        : "—"}
                    </td>
                    <td>
                      <span className={`row-status-pill ${
                        item.status === "success" ? "pill-done" :
                        item.status === "fetching" ? "pill-processing" :
                        item.status === "error" ? "pill-error" :
                        item.status === "skipped" ? "pill-skipped" : "pill-idle"
                      }`}>
                        {item.status === "success" && "✓ Done"}
                        {item.status === "fetching" && "Fetching…"}
                        {item.status === "error" && (item.error || "Error")}
                        {item.status === "skipped" && "Synced"}
                        {item.status === "idle" && "Pending"}
                      </span>
                    </td>
                  </tr>
                )) : (
                  <tr><td colSpan={6} className="empty-row">Load a Google Sheet in Step 1 to get started.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>

      {/* ── Apps Script Modal ── */}
      {showScriptSetup && (
        <div className="modal-overlay" onClick={() => setShowScriptSetup(false)}>
          <div className="modal-box" onClick={e => e.stopPropagation()}>
            <div className="modal-head">
              <h3>Google Apps Script Setup</h3>
              <button onClick={() => setShowScriptSetup(false)} className="modal-close">✕</button>
            </div>
            <ol className="modal-steps">
              <li>Open your Google Sheet → <strong>Extensions → Apps Script</strong></li>
              <li>Replace all existing code with the snippet below</li>
              <li>Click <strong>Deploy → New deployment → Web app</strong></li>
              <li>Set <em>Execute as</em>: <strong>Me</strong> · <em>Who has access</em>: <strong>Anyone</strong></li>
              <li>Click Deploy, authorize, and copy the <strong>Web App URL</strong></li>
              <li>Paste the URL into Step 3 above</li>
            </ol>
            <div className="script-wrap">
              <pre className="script-code">{APPS_SCRIPT_CODE}</pre>
              <button onClick={copyScript} className="copy-btn">{scriptCopied ? "Copied!" : "Copy"}</button>
            </div>
          </div>
        </div>
      )}

      <style>{`
        /* ── Reset & Base ── */
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        .gst-root {
          position: relative;
          min-height: 100vh;
          overflow-x: hidden;
          padding: 24px 16px 48px;
          font-family: var(--font-geist-sans, 'Inter', system-ui, sans-serif);
          background: linear-gradient(145deg, #f8f4ec 0%, #f0e8da 100%);
          color: #10213a;
        }

        /* ── Decorative blobs ── */
        .blob {
          position: fixed;
          border-radius: 50%;
          pointer-events: none;
          filter: blur(80px);
          opacity: .55;
          z-index: 0;
        }
        .blob-1 { width: 420px; height: 420px; top: -100px; left: -120px; background: radial-gradient(circle, #d4e4f7 0%, #b8d0ef 100%); }
        .blob-2 { width: 360px; height: 360px; top: 80px; right: -100px; background: radial-gradient(circle, #f0d9a8 0%, #e8c47e 100%); }

        /* ── Layout ── */
        .gst-container {
          position: relative;
          z-index: 1;
          max-width: 1200px;
          margin: 0 auto;
          display: flex;
          flex-direction: column;
          gap: 20px;
        }

        /* ── Card ── */
        .card {
          background: rgba(255,255,255,0.78);
          border: 1px solid rgba(16,33,58,0.1);
          border-radius: 24px;
          backdrop-filter: blur(12px);
          box-shadow: 0 4px 32px rgba(16,33,58,0.08);
          overflow: hidden;
        }

        /* ── Hero ── */
        .hero-card { padding: 28px 24px; }
        @media (min-width: 640px) { .hero-card { padding: 40px 44px; } }

        .hero-chips { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 20px; }
        .chip {
          background: rgba(255,255,255,0.7);
          border: 1px solid rgba(16,33,58,0.12);
          border-radius: 999px;
          padding: 4px 12px;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: .06em;
          color: #5d6a7d;
        }

        .hero-body { display: grid; gap: 28px; }
        @media (min-width: 900px) { .hero-body { grid-template-columns: 1.3fr .9fr; align-items: end; } }

        .hero-eyebrow { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .18em; color: #173a6d; margin-bottom: 8px; }
        .hero-title { font-size: clamp(22px, 4vw, 40px); font-weight: 700; line-height: 1.22; color: #10213a; margin-bottom: 10px; }
        .hero-sub { font-size: 14px; line-height: 1.7; color: #5d6a7d; }

        /* Progress card */
        .progress-card {
          background: rgba(255,255,255,0.92);
          border: 1px solid rgba(16,33,58,0.1);
          border-radius: 20px;
          padding: 20px;
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .progress-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
        .progress-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .14em; color: #5d6a7d; }
        .progress-count { font-size: 18px; font-weight: 700; color: #10213a; margin-top: 2px; }
        .progress-skip { font-size: 11px; color: #5d6a7d; margin-top: 2px; }

        .progress-bar-track { height: 8px; border-radius: 999px; background: #e2e8f0; overflow: hidden; }
        .progress-bar-fill { height: 100%; border-radius: 999px; background: #173a6d; transition: width .35s ease; }

        .progress-stats { display: grid; grid-template-columns: repeat(4,1fr); gap: 8px; }
        .stat-box { background: rgba(255,255,255,.7); border: 1px solid rgba(16,33,58,0.09); border-radius: 12px; padding: 8px 4px; text-align: center; }
        .stat-label { font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: .1em; color: #5d6a7d; }
        .stat-value { font-size: 18px; font-weight: 700; color: #10213a; }
        .stat-skip { color: #94a3b8; }
        .stat-done { color: #1d6d4f; }
        .stat-fail { color: #a13c3c; }

        /* Status pills */
        .status-pill {
          flex-shrink: 0;
          border-radius: 999px;
          padding: 4px 10px;
          font-size: 11px;
          font-weight: 700;
          white-space: nowrap;
        }
        .pill-idle { background: #f1f5f9; color: #64748b; outline: 1px solid #e2e8f0; }
        .pill-processing { background: rgba(168,107,17,.1); color: #a86b11; outline: 1px solid rgba(168,107,17,.2); }
        .pill-paused { background: #fef3c7; color: #92400e; outline: 1px solid #fde68a; }
        .pill-done { background: rgba(29,109,79,.1); color: #1d6d4f; outline: 1px solid rgba(29,109,79,.2); }
        .pill-error { background: rgba(161,60,60,.1); color: #a13c3c; outline: 1px solid rgba(161,60,60,.2); }
        .pill-skipped { background: #f1f5f9; color: #94a3b8; outline: 1px solid #e2e8f0; }

        /* ── Steps grid ── */
        .steps-grid { display: grid; gap: 16px; }
        @media (min-width: 768px) { .steps-grid { grid-template-columns: repeat(3,1fr); } }

        .step-card { display: flex; flex-direction: column; }
        .step-header {
          display: flex;
          align-items: center;
          gap: 10px;
          padding: 18px 20px 0;
        }
        .step-num {
          width: 28px; height: 28px;
          border-radius: 50%;
          background: #173a6d;
          color: #fff;
          font-size: 12px;
          font-weight: 800;
          display: flex; align-items: center; justify-content: center;
          flex-shrink: 0;
        }
        .step-title { font-size: 16px; font-weight: 700; color: #10213a; }

        .step-body { flex: 1; padding: 16px 20px; display: flex; flex-direction: column; gap: 12px; }
        .step-footer {
          padding: 12px 20px;
          border-top: 1px solid rgba(16,33,58,0.07);
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 8px;
          flex-wrap: wrap;
        }
        .footer-label { font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: .1em; color: #94a3b8; }
        .footer-code { font-size: 10px; font-family: monospace; color: #173a6d; background: rgba(23,58,109,0.07); padding: 3px 6px; border-radius: 6px; }
        .footer-value { font-size: 12px; font-weight: 600; color: #10213a; }

        /* ── Form elements ── */
        .field-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .08em; color: #5d6a7d; display: flex; flex-direction: column; gap: 5px; }
        .input-row { display: flex; gap: 6px; }
        .text-input {
          flex: 1;
          border: 1.5px solid rgba(16,33,58,0.14);
          border-radius: 10px;
          padding: 8px 12px;
          font-size: 12px;
          font-weight: 500;
          color: #10213a;
          background: rgba(255,255,255,.85);
          outline: none;
          transition: border-color .15s;
          min-width: 0;
        }
        .text-input:focus { border-color: #173a6d; background: #fff; }
        .select-input {
          width: 100%;
          border: 1.5px solid rgba(16,33,58,0.14);
          border-radius: 10px;
          padding: 7px 10px;
          font-size: 12px;
          font-weight: 500;
          color: #10213a;
          background: rgba(255,255,255,.85);
          outline: none;
          transition: border-color .15s;
        }
        .select-input:focus { border-color: #173a6d; }
        .placeholder-select {
          border: 1.5px dashed rgba(16,33,58,0.12);
          border-radius: 10px;
          padding: 7px 10px;
          font-size: 12px;
          color: #94a3b8;
        }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }

        /* ── Buttons ── */
        .btn {
          border: none;
          border-radius: 10px;
          padding: 9px 16px;
          font-size: 13px;
          font-weight: 700;
          cursor: pointer;
          transition: all .15s;
          white-space: nowrap;
        }
        .btn:disabled { opacity: .45; cursor: not-allowed; }
        .btn-primary { background: #173a6d; color: #fff; box-shadow: 0 4px 12px rgba(23,58,109,.22); }
        .btn-primary:hover:not(:disabled) { background: #102b52; }
        .btn-warning { background: #a86b11; color: #fff; }
        .btn-warning:hover:not(:disabled) { background: #8c5a0d; }
        .btn-success { background: #1d6d4f; color: #fff; box-shadow: 0 4px 12px rgba(29,109,79,.2); }
        .btn-success:hover:not(:disabled) { background: #165a40; }
        .btn-ghost { background: rgba(255,255,255,.7); border: 1.5px solid rgba(16,33,58,0.12); color: #10213a; }
        .btn-ghost:hover:not(:disabled) { background: #fff; }
        .btn-sm { padding: 8px 14px; font-size: 12px; }
        .btn-group { display: flex; flex-wrap: wrap; gap: 8px; }
        .btn-group > .btn { flex: 1; min-width: 80px; }

        .link-btn { background: none; border: none; font-size: 12px; font-weight: 600; color: #173a6d; cursor: pointer; padding: 0; text-align: left; }
        .link-btn:hover { text-decoration: underline; }

        /* ── Feedback boxes ── */
        .error-box { background: rgba(161,60,60,.07); border: 1px solid rgba(161,60,60,.2); border-radius: 10px; padding: 10px 12px; font-size: 12px; color: #a13c3c; }
        .success-box { background: rgba(29,109,79,.07); border: 1px solid rgba(29,109,79,.2); border-radius: 10px; padding: 10px 12px; font-size: 12px; color: #1d6d4f; }

        /* ── Results table ── */
        .results-card { }
        .results-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          flex-wrap: wrap;
          gap: 12px;
          padding: 20px 20px 16px;
          border-bottom: 1px solid rgba(16,33,58,0.08);
        }
        .results-title { font-size: 17px; font-weight: 700; color: #10213a; }
        .results-sub { font-size: 12px; color: #5d6a7d; margin-top: 2px; }
        .table-wrap { overflow-x: auto; }
        .data-table { width: 100%; border-collapse: collapse; min-width: 600px; font-size: 13px; }
        .data-table thead tr { background: rgba(248,250,252,.95); }
        .data-table th {
          padding: 11px 16px;
          text-align: left;
          font-size: 10px;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: .1em;
          color: #64748b;
          white-space: nowrap;
          border-bottom: 1px solid rgba(16,33,58,0.08);
        }
        .data-table td {
          padding: 12px 16px;
          border-bottom: 1px solid rgba(16,33,58,0.05);
          vertical-align: middle;
        }
        .data-table tbody tr:hover { background: rgba(248,250,252,.5); }
        .row-fetching { background: rgba(59,130,246,.04) !important; }
        .row-error { background: rgba(161,60,60,.04) !important; }
        .row-skipped { opacity: .55; }
        .td-mono { font-family: monospace; }
        .bold { font-weight: 600; }
        .muted { color: #94a3b8; }
        .small { font-size: 11px; }
        .empty-row { padding: 48px 16px !important; text-align: center; color: #94a3b8; font-size: 13px; }

        /* ── Badges ── */
        .badge { display: inline-flex; border-radius: 6px; padding: 2px 8px; font-size: 11px; font-weight: 700; outline: 1px solid; }
        .badge-filed { background: #ecfdf5; color: #1d6d4f; outline-color: rgba(29,109,79,.25); }
        .badge-notfiled { background: #fef2f2; color: #a13c3c; outline-color: rgba(161,60,60,.25); }
        .badge-nodata { background: #f8fafc; color: #94a3b8; outline-color: #e2e8f0; }

        .row-status-pill { display: inline-flex; border-radius: 999px; padding: 3px 10px; font-size: 11px; font-weight: 700; }
        .fetching-dot { color: #93c5fd; font-weight: 700; }

        /* ── Modal ── */
        .modal-overlay {
          position: fixed; inset: 0; z-index: 100;
          background: rgba(0,0,0,.45);
          backdrop-filter: blur(4px);
          display: flex; align-items: center; justify-content: center;
          padding: 16px;
        }
        .modal-box {
          background: #fff;
          border-radius: 24px;
          width: 100%; max-width: 560px;
          max-height: 90vh;
          overflow-y: auto;
          padding: 24px;
          display: flex;
          flex-direction: column;
          gap: 16px;
          box-shadow: 0 24px 80px rgba(0,0,0,.2);
        }
        .modal-head { display: flex; align-items: center; justify-content: space-between; }
        .modal-head h3 { font-size: 17px; font-weight: 700; color: #10213a; }
        .modal-close { background: none; border: none; font-size: 18px; color: #94a3b8; cursor: pointer; padding: 4px; line-height: 1; }
        .modal-close:hover { color: #10213a; }
        .modal-steps { list-style: decimal; padding-left: 20px; display: flex; flex-direction: column; gap: 8px; font-size: 13px; color: #5d6a7d; line-height: 1.6; }
        .modal-steps strong, .modal-steps em { color: #10213a; }
        .script-wrap { position: relative; margin-top: 4px; }
        .script-code {
          display: block;
          background: #0f172a;
          color: #94a3b8;
          border-radius: 14px;
          padding: 14px;
          font-size: 10px;
          font-family: monospace;
          line-height: 1.6;
          max-height: 240px;
          overflow-y: auto;
          white-space: pre-wrap;
          word-break: break-all;
        }
        .copy-btn {
          position: absolute;
          top: 8px; right: 8px;
          background: rgba(255,255,255,.12);
          border: 1px solid rgba(255,255,255,.18);
          color: #fff;
          border-radius: 8px;
          padding: 4px 10px;
          font-size: 11px;
          font-weight: 600;
          cursor: pointer;
        }
        .copy-btn:hover { background: rgba(255,255,255,.22); }
      `}</style>
    </main>
  );
}
