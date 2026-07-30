/**
 * FORM INPUT TRANSAKSI - PIODALAN NADI 13 JANUARI 2026
 * Cara pasang:
 * 1. Buka Google Sheets kamu.
 * 2. Menu Extensions > Apps Script.
 * 3. Hapus isi file "Code.gs" bawaan, ganti dengan isi file INI.
 * 4. Klik "+" di samping "Files" > pilih "HTML" > kasih nama persis: Form
 * 5. Isi file Form.html itu dengan isi file Form.html yang saya kasih terpisah.
 * 6. Klik Save (ikon disket), lalu klik Run sekali pada fungsi onOpen (akan minta izin akses, klik Allow).
 * 7. Kembali ke Sheets, refresh halaman. Akan muncul menu baru "Input Transaksi" di menu bar atas.
 */

const SHEET_NAME = 'PIODALAN NADI 13 JANUARI 2026';
const FIRST_DATA_ROW = 8;     // baris pertama data transaksi
const COL = { ACARA: 1, NO: 2, TANGGAL: 3, KETERANGAN: 4, KRITERIA: 5, SUB: 6, DEBET: 7, KREDIT: 8, SALDO: 9 };

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Input Transaksi')
    .addItem('+ Tambah Transaksi Baru', 'showForm')
    .addToUi();
}

function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(420)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Input Transaksi Baru');
}

/** Cari baris kosong pertama setelah baris terakhir yang sudah terisi tanggal */
function findNextRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const dates = sheet.getRange(FIRST_DATA_ROW, COL.TANGGAL, Math.max(lastRow - FIRST_DATA_ROW + 1, 1), 1).getValues();
  let lastFilled = FIRST_DATA_ROW - 1;
  for (let i = 0; i < dates.length; i++) {
    if (dates[i][0] !== '' && dates[i][0] !== null) lastFilled = FIRST_DATA_ROW + i;
  }
  return lastFilled + 1;
}

/** Dipanggil dari Form.html saat tombol SIMPAN ditekan */
function submitTransaksi(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Tab "' + SHEET_NAME + '" tidak ditemukan.');

  const row = findNextRow_(sheet);

  // NO TRS lanjut dari jumlah tanggal yang sudah terisi
  const lastRow = Math.max(row - 1, FIRST_DATA_ROW - 1);
  const dateCount = sheet.getRange(FIRST_DATA_ROW, COL.TANGGAL, Math.max(lastRow - FIRST_DATA_ROW + 1, 1), 1)
    .getValues().flat().filter(v => v !== '' && v !== null).length;
  const noTrs = dateCount + 1;

  const debet = Number(data.debet) || 0;
  const kredit = Number(data.kredit) || 0;

  // saldo berjalan = saldo baris sebelumnya (kalau ada & angka) + debet - kredit
  let prevSaldo = 0;
  if (row > FIRST_DATA_ROW) {
    const prevVal = sheet.getRange(row - 1, COL.SALDO).getValue();
    if (typeof prevVal === 'number') prevSaldo = prevVal;
  }
  const saldo = prevSaldo + debet - kredit;

  sheet.getRange(row, COL.ACARA).setValue('PIODALAN NADI');
  sheet.getRange(row, COL.NO).setValue(noTrs);
  sheet.getRange(row, COL.TANGGAL).setValue(new Date(data.tanggal));
  sheet.getRange(row, COL.KETERANGAN).setValue(data.keterangan);
  sheet.getRange(row, COL.KRITERIA).setValue(data.kriteria);
  sheet.getRange(row, COL.SUB).setValue(data.subKriteria);
  sheet.getRange(row, COL.DEBET).setValue(debet || '');
  sheet.getRange(row, COL.KREDIT).setValue(kredit || '');
  sheet.getRange(row, COL.SALDO).setValue(saldo);

  return { row: row, noTrs: noTrs, saldo: saldo };
}
