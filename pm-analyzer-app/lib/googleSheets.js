import { google } from "googleapis";

// Client di-cache supaya tidak re-auth tiap request (auto refresh token ditangani library ini)
let cachedClient = null;

export function getSheetsClient() {
  if (cachedClient) return cachedClient;

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      // \n literal di env var perlu diubah jadi newline asli
      private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    },
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  cachedClient = google.sheets({ version: "v4", auth });
  return cachedClient;
}

export const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

// Cari nomor baris (1-indexed, sesuai posisi asli di sheet) berdasarkan ID di kolom pertama tabel.
// dataStartRow = baris pertama data (bukan header) -- di template kita ini row 5.
export async function findRowByID(sheetName, dataStartRow, idColumn, id) {
  const sheets = getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!${idColumn}${dataStartRow}:${idColumn}2000`,
  });
  const rows = res.data.values || [];
  const idx = rows.findIndex((r) => r[0] === id);
  return idx === -1 ? null : dataStartRow + idx;
}

export async function getSheetIdByName(sheetName) {
  const sheets = getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find((s) => s.properties.title === sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' tidak ditemukan`);
  return sheet.properties.sheetId;
}
