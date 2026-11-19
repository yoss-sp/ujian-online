// =====================================================
// Apps Script untuk ujian-online.html
// Menyimpan data peserta ke Google Sheets:
//   Sheet "Peserta": Jam Mulai | Nama | Kelas | No Absen | Mata Pelajaran | Jam Selesai | Jumlah Pelanggaran | Status
//   Sheet "Log"    : semua kejadian (mulai, pelanggaran, buka-kunci, selesai, dll.)
// =====================================================

const TZ = "Asia/Jakarta";
const HEADER_PESERTA = ["Jam Mulai", "Nama", "Kelas", "No Absen", "Mata Pelajaran", "Jam Selesai", "Jumlah Pelanggaran", "Status"];
const HEADER_LOG = ["Waktu", "Nama", "Kelas", "No Absen", "Mata Pelajaran", "Jenis", "Keterangan"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // aman jika banyak siswa mengirim bersamaan
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const peserta = getSheet(ss, "Peserta", HEADER_PESERTA);
    const log = getSheet(ss, "Log", HEADER_LOG);
    const t = Utilities.formatDate(new Date(d.waktu), TZ, "dd/MM/yyyy HH:mm:ss");

    if (d.type === "mulai") {
      peserta.appendRow([t, d.nama, d.kelas, d.absen, d.mapel, "", 0, "Mengerjakan"]);
    } else {
      const row = findRow(peserta, d);
      if (row) {
        peserta.getRange(row, 7).setValue(d.pelanggaran);
        if (d.type === "selesai") {
          peserta.getRange(row, 6).setValue(t);
          peserta.getRange(row, 8).setValue("Selesai");
        } else if (d.type === "batas") {
          peserta.getRange(row, 8).setValue("Dihentikan (batas pelanggaran)");
        }
      }
    }
    log.appendRow([t, d.nama, d.kelas, d.absen, d.mapel, d.type, d.detail]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function getSheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange("A:H").setNumberFormat("@");  // teks, supaya no. absen "05" tidak jadi 5
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight("bold");
  }
  return sh;
}

// cari baris peserta terbaru dengan nama + kelas + absen + mapel yang sama
function findRow(sh, d) {
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (String(v[i][1]) === String(d.nama) && String(v[i][2]) === String(d.kelas) &&
        String(v[i][3]) === String(d.absen) && String(v[i][4]) === String(d.mapel)) return i + 1;
  }
  return 0;
}

// untuk mengetes URL di browser (harus tampil "ok")
function doGet() { return ContentService.createTextOutput("ok"); }
