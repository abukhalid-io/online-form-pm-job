import { useEffect, useState, useCallback } from "react";

const WAIT_SECONDS = 5 * 60; // 5 menit stabilisasi gas sebelum boleh mencatat pembacaan

/**
 * storageKey: string unik per sesi verifikasi, misal `timer-${id_pm}-iterasi-${no_iterasi}`
 * onReady: dipanggil sekali saat 5 menit selesai, biasanya untuk enable tombol "Catat Pembacaan"
 *
 * CATATAN: waktu mulai disimpan di localStorage supaya timer tidak reset kalau halaman
 * di-refresh. Ini hanya bertahan di browser yang sama -- kalau butuh timer yang bisa
 * dipantau lintas device, waktu mulai perlu disimpan di server (mis. kolom "Waktu" di
 * PM_Detail_Iterasi) dan dihitung ulang dari situ.
 */
export default function VerificationTimer({ storageKey, onReady }) {
  const [startTime, setStartTime] = useState(null);
  const [remaining, setRemaining] = useState(WAIT_SECONDS);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    const saved = typeof window !== "undefined" ? localStorage.getItem(storageKey) : null;
    if (saved) setStartTime(Number(saved));
  }, [storageKey]);

  useEffect(() => {
    if (!startTime) return;
    const tick = () => {
      const elapsed = Math.floor((Date.now() - startTime) / 1000);
      const sisa = Math.max(WAIT_SECONDS - elapsed, 0);
      setRemaining(sisa);
      if (sisa === 0) {
        setReady(true);
      }
    };
    tick();
    const interval = setInterval(tick, 1000);
    return () => clearInterval(interval);
  }, [startTime]);

  useEffect(() => {
    if (ready) onReady?.();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ready]);

  const mulaiTimer = useCallback(() => {
    const now = Date.now();
    localStorage.setItem(storageKey, String(now));
    setStartTime(now);
    setReady(false);
  }, [storageKey]);

  const resetTimer = useCallback(() => {
    localStorage.removeItem(storageKey);
    setStartTime(null);
    setReady(false);
    setRemaining(WAIT_SECONDS);
  }, [storageKey]);

  const menit = String(Math.floor(remaining / 60)).padStart(2, "0");
  const detik = String(remaining % 60).padStart(2, "0");

  if (!startTime) {
    return (
      <button type="button" onClick={mulaiTimer}>
        Mulai Tunggu Stabilisasi Gas (5 menit)
      </button>
    );
  }

  return (
    <div>
      {ready ? (
        <p>Sudah 5 menit — silakan catat pembacaan.</p>
      ) : (
        <p>Menunggu stabilisasi gas: {menit}:{detik}</p>
      )}
      <button type="button" onClick={resetTimer}>
        Ulangi Timer
      </button>
    </div>
  );
}
