import { type NextRequest } from "next/server";

type GstInput = {
  gstin: string;
  fy: string;
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
  /** "Filed" | "Not Filed" | "NA" */
  overallStatus: string;
  /** Comma-separated list of filed quarters/months */
  filedPeriods: string;
  /** Comma-separated list of unfiled quarters/months */
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

const GST_PORTAL_HEADERS = {
  Accept: "application/json, text/plain, */*",
  "Accept-Language": "en-US,en;q=0.9",
  Origin: "https://services.gst.gov.in",
  Pragma: "no-cache",
  Referer: "https://services.gst.gov.in/services/searchtp",
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
  "X-Requested-With": "XMLHttpRequest",
} as const;

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

function parseJsonText(text: string): unknown {
  if (!text) return null;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}

function pickString(record: Record<string, unknown>, keys: string[]): string | null {
  for (const key of keys) {
    const value = record[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return null;
}

function flattenRecords(value: unknown): Record<string, unknown>[] {
  const queue: unknown[] = [value];
  const records: Record<string, unknown>[] = [];
  while (queue.length > 0) {
    const current = queue.shift();
    if (Array.isArray(current)) {
      queue.unshift(...current);
      continue;
    }
    if (isRecord(current)) {
      records.push(current);
      queue.push(...Object.values(current));
    }
  }
  return records;
}

function findRecordWithKeys(value: unknown, keys: string[]): Record<string, unknown> | null {
  const queue: unknown[] = [value];
  while (queue.length > 0) {
    const current = queue.shift();
    if (Array.isArray(current)) { queue.unshift(...current); continue; }
    if (!isRecord(current)) continue;
    if (keys.some((k) => k in current)) return current;
    queue.push(...Object.values(current));
  }
  return null;
}

function extractFilingRows(value: unknown): FilingRow[] {
  const source = findRecordWithKeys(value, ["filingStatus"]);
  if (!source) return [];
  const rows = flattenRecords(source["filingStatus"]);
  return rows
    .map((record) => ({
      fy: pickString(record, ["fy"]) ?? "",
      taxp: pickString(record, ["taxp"]) ?? "",
      mof: pickString(record, ["mof"]) ?? "",
      dof: pickString(record, ["dof"]) ?? "",
      rtntype: pickString(record, ["rtntype"]) ?? "",
      arn: pickString(record, ["arn"]) ?? "",
      status: pickString(record, ["status"]) ?? "",
    }))
    .filter((row) => row.taxp.length > 0 || row.rtntype.length > 0);
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
    // Simple Filed / Not Filed — if even one period is missing it's Not Filed
    const overallStatus =
      total === 0 ? "Not Filed"
      : filedCount === total ? "Filed"
      : "Not Filed";

    return { returnType, overallStatus, filedPeriods, notFiledPeriods, latestFiledOn: latest, totalPeriods: total, filedCount };
  }).sort((a, b) => a.returnType.localeCompare(b.returnType));
}

async function fetchReturnDetails(input: GstInput): Promise<{ rows: FilingRow[]; rawStatus: string }> {
  try {
    const response = await fetch(
      "https://services.gst.gov.in/services/api/search/taxpayerReturnDetails",
      {
        method: "POST",
        cache: "no-store",
        headers: { ...GST_PORTAL_HEADERS, "Content-Type": "application/json" },
        body: JSON.stringify({ gstin: input.gstin, fy: input.fy }),
      },
    );
    const text = await response.text();
    const data = parseJsonText(text);
    const rows = extractFilingRows(data);

    // Extract top-level status message from portal response
    let rawStatus = "No records found";
    if (isRecord(data)) {
      const s = data["status"];
      if (typeof s === "string" && s.trim()) rawStatus = s.trim();
    }
    if (rows.length > 0) rawStatus = "OK";
    return { rows, rawStatus };
  } catch (err) {
    console.error("[returnDetails] Error:", err);
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