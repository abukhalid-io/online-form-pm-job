import { getSheetsClient, SPREADSHEET_ID, findRowByID, getSheetIdByName } from "../../lib/googleSheets";

const SHEET_NAME = "PM_Detail_Iterasi";
const DATA_START_ROW = 5; // baris 1-3 judul, baris 4 header, data mulai baris 5
const ID_COLUMN = "A";

// Kolom formula (Standar Zero, Standar Span, Deviasi Zero, Deviasi Span, Status, Peringatan)
// SENGAJA tidak pernah ditulis dari app -- biar tetap dihitung otomatis oleh rumus di sheet.
// Kalau app menimpa kolom ini dengan nilai statis, rumusnya akan hilang.
const EDITABLE_COLUMNS = {
  id_pm: "B",
  no_iterasi: "C",
  tahap: "D",
  waktu: "E",
  pembacaan_zero: "F",
  pembacaan_span: "G",
  toleransi_zero: "J",
  toleransi_span: "K",
  tindakan_kalibrasi: "O",
  catatan: "Q",
};

export default async function handler(req, res) {
  const sheets = getSheetsClient();

  try {
    if (req.method === "GET") {
      const { id_pm } = req.query;
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A${DATA_START_ROW}:Q2000`,
      });
      let rows = result.data.values || [];
      if (id_pm) rows = rows.filter((r) => r[1] === id_pm);
      return res.status(200).json({ data: rows });
    }

    if (req.method === "POST") {
      const body = req.body;
      const id_detail = body.id_detail || `DET-${Date.now()}`;

      // Susun 1 baris penuh (17 kolom A-Q), kolom formula dikosongkan agar rumus sheet yang isi
      const row = [
        id_detail,                      // A
        body.id_pm,                     // B
        body.no_iterasi,                // C
        body.tahap,                     // D
        body.waktu,                     // E
        body.pembacaan_zero ?? "",      // F
        body.pembacaan_span ?? "",      // G
        "", "",                         // H, I -> formula Standar Zero/Span
        body.toleransi_zero,            // J
        body.toleransi_span,            // K
        "", "", "",                     // L, M, N -> formula Deviasi & Status
        body.tindakan_kalibrasi ?? "",  // O
        "",                             // P -> formula Peringatan
        body.catatan ?? "",             // Q
      ];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A${DATA_START_ROW}`,
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: [row] },
      });

      return res.status(201).json({ message: "Data tersimpan", id_detail });
    }

    if (req.method === "PUT") {
      const { id_detail, ...fields } = req.body;
      if (!id_detail) return res.status(400).json({ message: "id_detail wajib diisi" });

      const rowNumber = await findRowByID(SHEET_NAME, DATA_START_ROW, ID_COLUMN, id_detail);
      if (!rowNumber) return res.status(404).json({ message: "ID tidak ditemukan" });

      const updates = [];
      for (const [key, col] of Object.entries(EDITABLE_COLUMNS)) {
        if (fields[key] !== undefined) {
          updates.push({ range: `${SHEET_NAME}!${col}${rowNumber}`, values: [[fields[key]]] });
        }
      }
      if (updates.length === 0) return res.status(400).json({ message: "Tidak ada field yang diupdate" });

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: "USER_ENTERED", data: updates },
      });

      return res.status(200).json({ message: "Data diperbarui" });
    }

    if (req.method === "DELETE") {
      const { id_detail } = req.query;
      if (!id_detail) return res.status(400).json({ message: "id_detail wajib diisi" });

      const rowNumber = await findRowByID(SHEET_NAME, DATA_START_ROW, ID_COLUMN, id_detail);
      if (!rowNumber) return res.status(404).json({ message: "ID tidak ditemukan" });

      const sheetId = await getSheetIdByName(SHEET_NAME);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: { sheetId, dimension: "ROWS", startIndex: rowNumber - 1, endIndex: rowNumber },
              },
            },
          ],
        },
      });

      return res.status(200).json({ message: "Data dihapus" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} tidak diizinkan`);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ message: "Terjadi kesalahan server", error: err.message });
  }
}
