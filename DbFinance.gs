// ==========================================
// KONFIGURASI DATABASE KEUANGAN
// ==========================================
const FINANCE_SHEET_ID = '12J3we8zzrNf1E1myiFXpHOnEsN5rDL8_m6W1IuK70fs'; 

// [OPTIMASI] Menyimpan instance Spreadsheet di memori agar tidak openById berulang kali
let cachedSS = null;

function getSS() {
  if (!cachedSS) {
    cachedSS = SpreadsheetApp.openById(FINANCE_SHEET_ID);
  }
  return cachedSS;
}

// [OPTIMASI] Fungsi cerdas untuk mengambil atau membuat Tab Sheet
function getSheetSafe(sheetName, autoCreate = false, headers = []) {
  const ss = getSS();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && autoCreate) {
    sheet = ss.insertSheet(sheetName);
    if (headers.length > 0) sheet.appendRow(headers);
  }
  return sheet;
}

// ==========================================
// MENGELOLA FOLDER GOOGLE DRIVE
// ==========================================
function getMainFolder() {
  const folderName = "Bukti Pengeluaran Laleh";
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

// ==========================================
// 1. MENGAMBIL STATISTIK ARUS KAS (SANGAT RINGAN)
// ==========================================
function getFinanceStats() {
  try {
    const ss = getSS();
    let totalHargaBarang = 0, totalOngkir = 0, totalDanaMasuk = 0, totalPengeluaran = 0;

    // A. Proses Data Order (Hanya baca jika sheet ada)
    const orderSheet = ss.getSheetByName('Orders');
    if (orderSheet) {
      const orderData = orderSheet.getDataRange().getValues();
      for (let i = 1; i < orderData.length; i++) {
        if (orderData[i][11] !== 'Dibatalkan') totalHargaBarang += (parseInt(orderData[i][6]) || 0);
      }
    }

    // B. Proses Data Keuangan Agen
    const userSheet = ss.getSheetByName('Users');
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        if (userData[i][3] && String(userData[i][3]).toLowerCase() === 'agen' && userData[i][6]) {
          try {
            let b = JSON.parse(userData[i][6]);
            totalOngkir += (parseInt(b.ongkir) || 0);
            totalDanaMasuk += (parseInt(b.dana_masuk) || 0);
          } catch(e) {}
        }
      }
    }

    // C. Proses Data Pengeluaran
    const expSheet = ss.getSheetByName('Expenses');
    if (expSheet) {
      const expData = expSheet.getDataRange().getValues();
      for (let i = 1; i < expData.length; i++) {
        totalPengeluaran += (parseInt(expData[i][3]) || 0);
      }
    }

    return {
      hargaBarang: totalHargaBarang, ongkir: totalOngkir,
      danaMasuk: totalDanaMasuk, pengeluaran: totalPengeluaran,
      labaBersih: totalDanaMasuk - totalPengeluaran
    };
  } catch (error) {
    return { hargaBarang: 0, ongkir: 0, danaMasuk: 0, pengeluaran: 0, labaBersih: 0 };
  }
}

// ==========================================
// 2. MANAJEMEN PENGELUARAN (READ & WRITE)
// ==========================================
function getExpenses() {
  const ss = getSS();
  const sheet = ss.getSheetByName('Expenses');
  if (!sheet) return []; // [OPTIMASI] Langsung kembalikan kosong jika sheet belum pernah dibuat

  const data = sheet.getDataRange().getValues();
  let expenses = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      let rawDate = data[i][1];
      let safeDate = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, "GMT+7", "yyyy-MM-dd") : String(rawDate);

      expenses.push({
        id: data[i][0], tanggal: safeDate, kategori: data[i][2],
        nominal: parseInt(data[i][3]) || 0, keterangan: data[i][4], bukti_url: data[i][5] || ''
      });
    }
  }
  return expenses.reverse(); 
}

function addExpenseWithFile(data, fileData) {
  try {
    let fileUrl = "";
    if (fileData && fileData.base64) {
      const mainFolder = getMainFolder();
      const blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, data.tanggal + " - Bukti - " + fileData.name);
      fileUrl = mainFolder.createFile(blob).getUrl();
    }
    
    // Auto-create sheet jika baru pertama kali input
    const sheet = getSheetSafe('Expenses', true, ["ID Pengeluaran", "Tanggal", "Kategori", "Nominal", "Keterangan", "Bukti URL", "Timestamp"]);
    sheet.appendRow(["EXP-" + new Date().getTime(), data.tanggal, data.kategori, data.nominal, data.keterangan, fileUrl, new Date().toISOString()]);
    
    return { success: true, message: "Pengeluaran berhasil dicatat!" };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function deleteExpense(id) {
  try {
    const sheet = getSS().getSheetByName('Expenses');
    if (!sheet) return { success: false, message: "Data kosong." };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data pengeluaran dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// ==========================================
// 3. MANAJEMEN TUTUP BUKU
// ==========================================
function closeFinanceBook(archiveData) {
  try {
    const mainFolder = getMainFolder();
    const subFolder = mainFolder.createFolder("Arsip " + archiveData.name);

    if (archiveData.expenses && archiveData.expenses.length > 0) {
      archiveData.expenses.forEach(exp => {
        if (exp.bukti_url) {
          const match = exp.bukti_url.match(/[-\w]{25,}/);
          if (match && match[0]) {
            try { DriveApp.getFileById(match[0]).moveTo(subFolder); } catch(e) {}
          }
        }
      });
    }

    const histSheet = getSheetSafe('FinanceHistory', true, ["ID Arsip", "Nama Periode", "Tanggal Tutup", "Summary (JSON)", "Expenses (JSON)"]);
    histSheet.appendRow([
      "HST-" + new Date().getTime(), archiveData.name, new Date().toISOString().slice(0, 10),
      JSON.stringify(archiveData.summary), JSON.stringify(archiveData.expenses)
    ]);

    const expSheet = getSS().getSheetByName('Expenses');
    if (expSheet) {
      const lastRow = expSheet.getLastRow();
      if (lastRow > 1) expSheet.getRange(2, 1, lastRow - 1, expSheet.getLastColumn()).clearContent();
    }

    return { success: true, message: "Tutup buku selesai! File bukti dipindahkan ke sub-folder: Arsip " + archiveData.name };
  } catch (error) { return { success: false, message: "Gagal: " + error.toString() }; }
}

function getFinanceHistory() {
  try {
    const sheet = getSS().getSheetByName('FinanceHistory');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    let history = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        let rawDate = data[i][2];
        let safeDate = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, "GMT+7", "yyyy-MM-dd") : String(rawDate);

        history.push({
          id: data[i][0], name: data[i][1], date: safeDate,
          summary: JSON.parse(data[i][3] || "{}"), expenses: JSON.parse(data[i][4] || "[]")
        });
      }
    }
    return history.reverse();
  } catch (error) { return []; }
}

function renameFinanceHistory(id, newName) {
  try {
    const sheet = getSS().getSheetByName('FinanceHistory');
    if (!sheet) return { success: false };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2).setValue(newName);
        return { success: true };
      }
    }
    return { success: false, message: "Arsip tidak ditemukan" };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// ==========================================
// 4. MASTER KATEGORI PENGELUARAN
// ==========================================
function getExpenseCategories() {
  try {
    const sheet = getSS().getSheetByName('Settings');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'FINANCE_CATEGORIES') return JSON.parse(data[i][1] || "[]");
      }
    }
    return ["HPP Produksi (Kain, Jahit)", "Biaya Ekspedisi (Ongkir Riil)", "Operasional & Gaji", "Marketing & Iklan", "Packaging"];
  } catch (error) {
    return ["HPP Produksi (Kain, Jahit)", "Biaya Ekspedisi (Ongkir Riil)", "Operasional & Gaji", "Marketing & Iklan", "Packaging"];
  }
}

function updateExpenseCategories(newList) {
  try {
    const sheet = getSheetSafe('Settings', true, ["Key", "Value"]);
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'FINANCE_CATEGORIES') {
        sheet.getRange(i + 1, 2).setValue(JSON.stringify(newList));
        found = true; break;
      }
    }
    if (!found) sheet.appendRow(['FINANCE_CATEGORIES', JSON.stringify(newList)]);
    return { success: true, message: "Kategori pengeluaran diperbarui." };
  } catch (error) { return { success: false, message: error.toString() }; }
}
