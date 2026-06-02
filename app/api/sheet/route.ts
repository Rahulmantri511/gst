import { type NextRequest } from "next/server";

export const dynamic = "force-dynamic";

function parseCsv(text: string): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (char === '"') {
        if (next === '"') {
          cell += '"';
          i++; // Skip next quote
        } else {
          inQuotes = false;
        }
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ",") {
        row.push(cell);
        cell = "";
      } else if (char === "\r" || char === "\n") {
        row.push(cell);
        if (row.some(c => c.trim().length > 0) || lines.length === 0) {
          lines.push(row);
        }
        row = [];
        cell = "";
        if (char === "\r" && next === "\n") {
          i++; // Skip \n
        }
      } else {
        cell += char;
      }
    }
  }

  // Handle final cell / row
  if (cell || row.length > 0) {
    row.push(cell);
    if (row.some(c => c.trim().length > 0)) {
      lines.push(row);
    }
  }

  return lines;
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { sheetUrl } = body;

    if (!sheetUrl || typeof sheetUrl !== "string") {
      return Response.json({ success: false, message: "A valid Google Sheet URL is required." }, { status: 400 });
    }

    // Extract spreadsheet ID
    const idMatch = sheetUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!idMatch) {
      return Response.json({ success: false, message: "Could not find a spreadsheet ID in the URL. Ensure it is a valid Google Sheets link." }, { status: 400 });
    }
    const spreadsheetId = idMatch[1];

    // Extract GID (tab ID)
    const gidMatch = sheetUrl.match(/[#&]gid=([0-9]+)/);
    const gid = gidMatch ? gidMatch[1] : null;

    let exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv`;
    if (gid) {
      exportUrl += `&gid=${gid}`;
    }

    const response = await fetch(exportUrl, {
      cache: "no-store",
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
      }
    });

    if (!response.ok) {
      if (response.status === 404) {
        return Response.json({
          success: false,
          message: "Sheet not found. Make sure the Google Sheet exists and you have shared it as 'Anyone with the link can view'."
        }, { status: 404 });
      }
      return Response.json({
        success: false,
        message: `Google returned an error (HTTP ${response.status}). Make sure the Google Sheet is shared publicly ("Anyone with the link can view").`
      }, { status: response.status });
    }

    const text = await response.text();
    const parsed = parseCsv(text);

    if (parsed.length === 0) {
      return Response.json({ success: false, message: "The sheet appears to be empty." }, { status: 400 });
    }

    const headers = parsed[0].map(h => h.trim());
    const rows = parsed.slice(1).map((row) => {
      const record: Record<string, string> = {};
      headers.forEach((header, idx) => {
        record[header] = row[idx] || "";
      });
      return record;
    });

    return Response.json({
      success: true,
      spreadsheetId,
      gid,
      headers,
      rowsCount: rows.length,
      rows
    });

  } catch (error) {
    return Response.json({
      success: false,
      message: error instanceof Error ? error.message : "An unexpected error occurred while parsing the sheet."
    }, { status: 500 });
  }
}
