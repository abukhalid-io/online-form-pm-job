import { getSheetsClient, SPREADSHEET_ID, findRowByID, getSheetIdByName } from "../../lib/googleSheets";

const SHEET_NAME = "PM_Header";
const DATA_START_ROW = 5;
const ID_COLUMN = "A";

// Kolom D (Jenis Gas) dan J (Jumlah Iterasi) adalah formula -- jangan ditulis dari app
const EDITABLE_COLUMNS = {
  tanggal: "B",
  id_analyzer: "C",
  teknisi: "E",
  id_tabung_zero: "F",
  id_tabung_span: "G",
  tekanan_zero: "H",
  tekanan_span: "I",
  status_akhir: "K",
  waktu_selesai: "L",
  catatan: "M",
};

export default async function handler(req, res) {
  const sheets = getSheetsClient();

  try {
    if (req.method === "GET") {
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A${DATA_START_ROW}:M2000`,
      });
      return res.status(200).json({ data: result.data.values || [] });
    }

    if (req.method === "POST") {
      const body = req.body;
      const id_pm = body.id_pm || `PM-${Date.now()}`;

      const row = [
        id_pm, body.tanggal, body.id_analyzer, "", body.teknisi,
        body.id_tabung_zero ?? "-", body.id_tabung_span ?? "-",
        body.tekanan_zero ?? "-", body.tekanan_span ?? "-",
        "", body.status_akhir ?? "Berjalan", body.waktu_selesai ?? "", body.catatan ?? "",
      ];

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: `${SHEET_NAME}!A${DATA_START_ROW}`,
        valueInputOption: "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: [row] },
      });

      return res.status(201).json({ message: "Sesi PM dibuat", id_pm });
    }

    if (req.method === "PUT") {
      const { id_pm, ...fields } = req.body;
      if (!id_pm) return res.status(400).json({ message: "id_pm wajib diisi" });

      const rowNumber = await findRowByID(SHEET_NAME, DATA_START_ROW, ID_COLUMN, id_pm);
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
      const { id_pm } = req.query;
      if (!id_pm) return res.status(400).json({ message: "id_pm wajib diisi" });

      const rowNumber = await findRowByID(SHEET_NAME, DATA_START_ROW, ID_COLUMN, id_pm);
      if (!rowNumber) return res.status(404).json({ message: "ID tidak ditemukan" });

      const sheetId = await getSheetIdByName(SHEET_NAME);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [
            { deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: rowNumber - 1, endIndex: rowNumber } } },
          ],
        },
      });

      return res.status(200).json({ message: "Sesi PM dihapus (data detail terkait tidak otomatis terhapus, cek manual)" });
    }

    res.setHeader("Allow", ["GET", "POST", "PUT", "DELETE"]);
    return res.status(405).end(`Method ${req.method} tidak diizinkan`);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ message: "Terjadi kesalahan server", error: err.message });
  }
}
