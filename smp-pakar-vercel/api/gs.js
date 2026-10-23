const ALLOWED_ACTIONS = new Set([
  "authenticateUser", "getAllUsers", "saveUser", "bulkImportUsers", "deleteUser",
  "getSiswaByKelas", "getAllKelasList", "getAllSiswaList", "getDashboardSummary",
  "generateQRPresensiToken", "getQRHistory", "processDirectQRScanWithSelfie",
  "getPresensiData", "submitIzinSakit", "getIzinSakitData", "updateStatusIzin",
  "getJurnalKBM", "submitJurnalKBM", "getNilaiSiswa", "submitNilaiSiswa"
]);

module.exports = async (req, res) => {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");

  if (req.method !== "POST") {
    res.status(405).json({ success: false, message: "Gunakan POST." });
    return;
  }

  const gasUrl = process.env.GAS_WEB_APP_URL;
  if (!gasUrl || !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(gasUrl)) {
    res.status(500).json({
      success: false,
      message: "GAS_WEB_APP_URL belum diatur atau bukan URL Web App Apps Script yang valid."
    });
    return;
  }

  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    const action = body && body.action;
    const payload = body && Array.isArray(body.payload) ? body.payload : [];
    if (!ALLOWED_ACTIONS.has(action)) {
      res.status(400).json({ success: false, message: "Aksi tidak diizinkan." });
      return;
    }

    const upstream = await fetch(gasUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action, payload }),
      redirect: "follow"
    });
    const raw = await upstream.text();
    let parsed;
    try { parsed = JSON.parse(raw); }
    catch (_) {
      throw new Error("Respons Apps Script bukan JSON. Periksa deployment Web App dan izin aksesnya.");
    }

    if (!upstream.ok || parsed.success === false) {
      res.status(502).json({
        success: false,
        message: parsed.message || "Apps Script gagal memproses permintaan."
      });
      return;
    }

    // Apps Script doPost wraps the called function result in `data`.
    res.status(200).json({ success: true, data: parsed.data });
  } catch (error) {
    res.status(502).json({
      success: false,
      message: error.message || "Tidak dapat terhubung ke Apps Script."
    });
  }
};
