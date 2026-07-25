import { useState } from "react";
import VerificationTimer from "../components/VerificationTimer";

// Contoh alur input 1 iterasi verifikasi. Sesuaikan id_pm dengan sesi PM yang sedang berjalan
// (biasanya diambil dari halaman "Sesi PM Aktif" -- belum termasuk di scaffold ini).
export default function InputVerifikasi() {
  const [form, setForm] = useState({
    id_pm: "",
    no_iterasi: 1,
    tahap: "Verifikasi Awal",
    pembacaan_zero: "",
    pembacaan_span: "",
    toleransi_zero: 2,
    toleransi_span: 2,
  });
  const [siapCatat, setSiapCatat] = useState(false);
  const [status, setStatus] = useState("");

  const handleChange = (e) => setForm((f) => ({ ...f, [e.target.name]: e.target.value }));

  const handleSubmit = async (e) => {
    e.preventDefault();
    setStatus("Menyimpan...");
    try {
      const res = await fetch("/api/pm-detail-iterasi", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          ...form,
          waktu: new Date().toISOString(),
          tindakan_kalibrasi: "",
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.message);
      setStatus(`Tersimpan sebagai ${data.id_detail}`);
    } catch (err) {
      setStatus(`Gagal: ${err.message}`);
    }
  };

  return (
    <main style={{ maxWidth: 480, margin: "40px auto", fontFamily: "sans-serif" }}>
      <h1>Input Verifikasi PM Analyzer</h1>

      <VerificationTimer
        storageKey={`timer-${form.id_pm || "belum-dipilih"}-iterasi-${form.no_iterasi}`}
        onReady={() => setSiapCatat(true)}
      />

      <form onSubmit={handleSubmit} style={{ marginTop: 24, display: "grid", gap: 12 }}>
        <label>
          ID Sesi PM
          <input name="id_pm" value={form.id_pm} onChange={handleChange} required />
        </label>

        <label>
          No. Iterasi
          <input type="number" name="no_iterasi" min={1} max={10} value={form.no_iterasi} onChange={handleChange} required />
        </label>

        <label>
          Pembacaan Zero
          <input type="number" step="0.01" name="pembacaan_zero" value={form.pembacaan_zero} onChange={handleChange} />
        </label>

        <label>
          Pembacaan Span
          <input type="number" step="0.01" name="pembacaan_span" value={form.pembacaan_span} onChange={handleChange} />
        </label>

        <button type="submit" disabled={!siapCatat}>
          {siapCatat ? "Catat Pembacaan" : "Tunggu 5 menit dulu..."}
        </button>
      </form>

      {status && <p>{status}</p>}
    </main>
  );
}
