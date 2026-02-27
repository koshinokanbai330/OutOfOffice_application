import { getAccessToken } from "./authService";
import { addDays, toISODateStr } from "./subjectHelper";

const APP_FOLDER = "/drive/special/approot";
const TEMPLATE_FILE = "BT-Allowance-Template.xlsx";

export interface ExcelRequest {
  familyName: string;
  startDate: Date;
  endDate: Date;
  destination: string;
}

interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
}

async function ensureTemplate(token: string, templateBytes: ArrayBuffer): Promise<string> {
  // Check if template already in app folder
  const checkResp = await fetch(
    `https://graph.microsoft.com/v1.0${APP_FOLDER}:/${TEMPLATE_FILE}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (checkResp.ok) {
    const item: DriveItem = await checkResp.json();
    return item.id;
  }

  // Upload template
  const uploadResp = await fetch(
    `https://graph.microsoft.com/v1.0${APP_FOLDER}:/${TEMPLATE_FILE}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      body: templateBytes,
    }
  );

  if (!uploadResp.ok) {
    throw new Error(`Failed to upload template: ${uploadResp.status}`);
  }

  const item: DriveItem = await uploadResp.json();
  return item.id;
}

async function copyTemplate(token: string, templateId: string, destName: string): Promise<string> {
  const copyResp = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${templateId}/copy`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Prefer: "respond-async",
      },
      body: JSON.stringify({
        name: destName,
        parentReference: { path: `/drive/special/approot` },
      }),
    }
  );

  if (copyResp.status === 202) {
    // Long-running operation — poll
    const location = copyResp.headers.get("Location");
    if (!location) throw new Error("No Location header for copy operation");
    return pollCopyOperation(token, location);
  }

  if (copyResp.ok) {
    const item: DriveItem = await copyResp.json();
    return item.id;
  }

  throw new Error(`Copy failed: ${copyResp.status}`);
}

async function pollCopyOperation(token: string, location: string): Promise<string> {
  for (let i = 0; i < 30; i++) {
    await new Promise((r) => setTimeout(r, 2000));
    const resp = await fetch(location, { headers: { Authorization: `Bearer ${token}` } });
    if (resp.status === 200) {
      const result = await resp.json();
      if (result.status === "completed") return result.resourceId as string;
      if (result.status === "failed") throw new Error("Copy operation failed");
    }
  }
  throw new Error("Copy operation timed out");
}

function getDates(startDate: Date, endDate: Date): Date[] {
  const dates: Date[] = [];
  let cur = new Date(startDate);
  while (cur <= endDate) {
    dates.push(new Date(cur));
    cur = addDays(cur, 1);
  }
  return dates;
}

async function findHeaderRow(
  token: string,
  driveItemId: string,
  worksheet: string,
  headerText: string
): Promise<number> {
  // Get used range
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/workbook/worksheets('${encodeURIComponent(worksheet)}')/usedRange`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!resp.ok) throw new Error(`Failed to get used range: ${resp.status}`);
  const data = await resp.json();
  const values: string[][] = data.values;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).includes(headerText)) {
        return r + 1; // 1-based
      }
    }
  }
  return -1;
}

async function writeRows(
  token: string,
  driveItemId: string,
  worksheet: string,
  startRow: number,
  dates: Date[],
  destination: string
): Promise<void> {
  for (let i = 0; i < dates.length; i++) {
    const row = startRow + i;
    const d = dates[i];
    const dateStr = d.toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" });
    const values = [[dateStr, destination, "7:00", "9:00", "18:00", "21:00"]];

    const resp = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${driveItemId}/workbook/worksheets('${encodeURIComponent(worksheet)}')/range(address='A${row}:F${row}')`,
      {
        method: "PATCH",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ values }),
      }
    );
    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(`Failed to write row ${row}: ${resp.status} ${err}`);
    }
  }
}

export async function createExcel(req: ExcelRequest, templateBytes: ArrayBuffer): Promise<string> {
  const token = await getAccessToken();
  const startStr = toISODateStr(req.startDate).replace(/-/g, "");
  const destName = `BT-Allowance-${req.familyName}-${startStr}.xlsx`;

  // Ensure template exists in OneDrive
  const templateId = await ensureTemplate(token, templateBytes);

  // Copy template to new file
  const newItemId = await copyTemplate(token, templateId, destName);

  // Determine dates
  const dates = getDates(req.startDate, req.endDate);
  const isMultiDay = dates.length > 1;

  // Find header row in "日帰り One-Day" sheet
  const oneDaySheet = "日帰り One-Day";
  const headerRow = await findHeaderRow(token, newItemId, oneDaySheet, "日付");
  const dataStartRow = headerRow > 0 ? headerRow + 1 : 2;

  await writeRows(token, newItemId, oneDaySheet, dataStartRow, dates, req.destination);

  if (isMultiDay) {
    const overnightSheet = "宿泊 Overnight";
    const headerRow2 = await findHeaderRow(token, newItemId, overnightSheet, "日付");
    const dataStartRow2 = headerRow2 > 0 ? headerRow2 + 1 : 2;
    await writeRows(token, newItemId, overnightSheet, dataStartRow2, dates, req.destination);
  }

  // Get webUrl
  const getResp = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${newItemId}?$select=webUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const item: DriveItem = await getResp.json();
  return item.webUrl || destName;
}
