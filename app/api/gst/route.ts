import { type NextRequest } from "next/server";
import https from "node:https";

type GstInput = {
  gstin: string;
  fy: string; // four-digit start year, e.g. "2025"
};

type FilingRow = {
  fy: string;
  taxp: string;
  mof: string;
  dof: string;
  rtntype: string;
  arn: string;
  status: string;
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
  input: {
    gstin: string;
    fy: string;
    financialYearLabel: string;
  };
  returnStatuses: ReturnStatus[];
  filingRows: FilingRow[];
  rawStatus: string;
};

type ErrorResponse = {
  success: false;
  message: string;
};

const GSTIN_PATTERN = /^[0-9A-Z]{15}$/;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function normalizeGstin(value: string): string {
  return value.trim().toUpperCase().replace(/\s+/g, "");
}

function normalizeFy(value: string): string {
  const match = value.match(/\d{4}/);
  return match ? match[0] : "";
}

function financialYearLabel(fy: string): string {
  const start = Number(fy);
  if (!Number.isFinite(start)) return fy;
  return `${start}-${String(start + 1).slice(-2)}`;
}

/** Full FY string required by the GST portal API e.g. "2025-2026" */
function portalFyString(fy: string): string {
  const start = Number(fy);
  if (!Number.isFinite(start)) return fy;
  return `${start}-${start + 1}`;
}

function pickString(record: Record<string, unknown>, keys: string[]): string | null {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return null;
}

function extractFilingRows(data: unknown): FilingRow[] {
  if (!isRecord(data)) return [];
  const fs = data["filingStatus"];
  if (!fs) return [];

  // Portal returns: { filingStatus: [[row, row, ...]] } — one level of nesting
  const flatRows: unknown[] = [];
  if (Array.isArray(fs)) {
    for (const item of fs) {
      if (Array.isArray(item)) flatRows.push(...item);
      else flatRows.push(item);
    }
  }

  return flatRows
    .filter(isRecord)
    .map((record) => ({
      fy: pickString(record, ["fy"]) ?? "",
      taxp: pickString(record, ["taxp"]) ?? "",
      mof: pickString(record, ["mof"]) ?? "",
      dof: pickString(record, ["dof"]) ?? "",
      rtntype: pickString(record, ["rtntype"]) ?? "",
      arn: pickString(record, ["arn"]) ?? "",
      status: pickString(record, ["status"]) ?? "",
    }))
    .filter((row) => row.rtntype.length > 0);
}

function latestDateLabel(values: string[]): string | null {
  let latestTs: number | null = null;
  let latestLabel: string | null = null;
  for (const value of values) {
    const [d, m, y] = value.split("/").map(Number);
    if (!d || !m || !y) continue;
    const ts = Date.UTC(y, m - 1, d);
    if (latestTs === null || ts > latestTs) { latestTs = ts; latestLabel = value; }
  }
  return latestLabel;
}

function buildReturnStatuses(rows: FilingRow[]): ReturnStatus[] {
  const grouped = new Map<string, FilingRow[]>();
  for (const row of rows) {
    const key = (row.rtntype || "Unknown").toUpperCase();
    const list = grouped.get(key) ?? [];
    list.push(row);
    grouped.set(key, list);
  }

  return Array.from(grouped.entries()).map(([returnType, groupedRows]) => {
    const filedRows = groupedRows.filter((r) => r.status.toLowerCase() === "filed");
    const notFiledRows = groupedRows.filter((r) => r.status.toLowerCase() !== "filed");

    const filedPeriods = filedRows.map((r) => r.taxp).filter(Boolean).join(", ");
    const notFiledPeriods = notFiledRows.map((r) => r.taxp).filter(Boolean).join(", ");
    const latest = latestDateLabel(filedRows.map((r) => r.dof).filter(Boolean));

    const total = groupedRows.length;
    const filedCount = filedRows.length;
    // "Filed" = at least one period filed (more lenient than requiring all periods)
    const overallStatus =
      filedCount > 0 ? "Filed" : "Not Filed";

    return { returnType, overallStatus, filedPeriods, notFiledPeriods, latestFiledOn: latest, totalPeriods: total, filedCount };
  }).sort((a, b) => a.returnType.localeCompare(b.returnType));
}

/**
 * Use Node's native https module instead of fetch (undici).
 * fetch/undici has a different TLS fingerprint that gets blocked by the GST portal WAF.
 * Node's https module matches the fingerprint our local tests use — which works.
 */
function nodeHttpsPost(
  hostname: string,
  path: string,
  body: string,
  headers: Record<string, string | number>,
): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = https.request(
      {
        hostname,
        path,
        method: "POST",
        headers: { ...headers, "Content-Length": Buffer.byteLength(body) },
      },
      (res) => {
        let data = "";
        res.on("data", (chunk: Buffer) => { data += chunk.toString(); });
        res.on("end", () => resolve(data));
        res.on("error", reject);
      },
    );
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

async function fetchReturnDetails(input: GstInput): Promise<{ rows: FilingRow[]; rawStatus: string }> {
  try {
    const fyForPortal = portalFyString(input.fy);
    const bodyStr = JSON.stringify({ gstin: input.gstin, fy: fyForPortal });

    console.log(`[GST] POST gstin=${input.gstin} fy=${fyForPortal}`);

    const text = await nodeHttpsPost(
      "services.gst.gov.in",
      "/services/api/search/taxpayerReturnDetails",
      bodyStr,
      {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Origin": "https://services.gst.gov.in",
        "Referer": "https://services.gst.gov.in/services/searchtp",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
        "X-Requested-With": "XMLHttpRequest",
      },
    );

    console.log(`[GST] Response length=${text.length} preview=${text.substring(0, 80)}`);

    // Guard: portal WAF sometimes returns HTML rejection page with HTTP 200
    if (text.trim().startsWith("<")) {
      console.error("[GST] Got HTML response instead of JSON — WAF/portal rejection");
      return { rows: [], rawStatus: "Portal rejected the request" };
    }

    let data: unknown;
    try { data = JSON.parse(text); } catch { return { rows: [], rawStatus: "Invalid JSON from portal" }; }

    const rows = extractFilingRows(data);
    console.log(`[GST] Extracted ${rows.length} filing rows`);

    let rawStatus = "No records found";
    if (isRecord(data)) {
      const s = data["status"];
      if (typeof s === "string" && s.trim()) rawStatus = s.trim();
    }
    if (rows.length > 0) rawStatus = "OK";
    return { rows, rawStatus };
  } catch (err) {
    console.error("[GST] Error:", err);
    return { rows: [], rawStatus: "Error fetching data" };
  }
}

export async function POST(request: NextRequest) {
  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return Response.json({ success: false, message: "Request body must be valid JSON." } as ErrorResponse, { status: 400 });
  }

  if (!isRecord(body)) {
    return Response.json({ success: false, message: "Request body must include gstin and fy." } as ErrorResponse, { status: 400 });
  }

  const gstin = normalizeGstin(typeof body.gstin === "string" ? body.gstin : "");
  const fy = normalizeFy(typeof body.fy === "string" ? body.fy : "");

  if (!GSTIN_PATTERN.test(gstin)) {
    return Response.json({ success: false, message: "GSTIN must be a 15-character uppercase alphanumeric value." } as ErrorResponse, { status: 400 });
  }
  if (!fy) {
    return Response.json({ success: false, message: "Financial year must contain a four-digit start year, such as 2025." } as ErrorResponse, { status: 400 });
  }

  const { rows, rawStatus } = await fetchReturnDetails({ gstin, fy });
  const returnStatuses = buildReturnStatuses(rows);

  const response: LookupResponse = {
    success: true,
    input: { gstin, fy, financialYearLabel: financialYearLabel(fy) },
    returnStatuses,
    filingRows: rows,
    rawStatus,
  };

  return Response.json(response);
}