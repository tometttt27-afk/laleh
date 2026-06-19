/**
 * ROUTER: Navigasi Halaman
 */
function doGet(e) {
  let page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'signin';
  let template;
  
  try {
    if (page === 'signup') template = HtmlService.createTemplateFromFile('SignUp');
    else if (page === 'forgot') template = HtmlService.createTemplateFromFile('LupaPw');
    else if (page === 'reset') template = HtmlService.createTemplateFromFile('ResetPw');
    else if (page === 'home') template = HtmlService.createTemplateFromFile('AgenDsh');
    else if (page === 'admin') template = HtmlService.createTemplateFromFile('AdminDsh');
    else if (page === 'finance') template = HtmlService.createTemplateFromFile('FinanceDsh'); 
    else template = HtmlService.createTemplateFromFile('SignIn');
  } catch (error) {
    console.error("Router Error: ", error);
    template = HtmlService.createTemplateFromFile('SignIn');
  }
  
  template.url = ScriptApp.getService().getUrl();
  template.emailParam = (e && e.parameter && e.parameter.email) ? e.parameter.email : ''; 
  
  return template.evaluate()
    .setTitle('Laleh.id')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * FUNGSI AMBIL STATISTIK ADMIN
 */
function getAdminStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const wsUsers = ss.getSheetByName("Users");
  let totalAgen = 0;
  if (wsUsers) {
    const dataUsers = wsUsers.getDataRange().getValues();
    totalAgen = dataUsers.filter(row => row[3] && row[3].toString().trim().toLowerCase() === "agen").length;
  }
  
  const wsProducts = ss.getSheetByName("Products");
  let totalProduk = 0;
  if (wsProducts) {
    totalProduk = Math.max(0, wsProducts.getLastRow() - 1); 
  }
  
  const wsOrders = ss.getSheetByName("Orders");
  let totalPesanan = 0;
  if (wsOrders) {
    totalPesanan = Math.max(0, wsOrders.getLastRow() - 1); 
  }
  
  return {
    totalAgen: totalAgen,
    totalProduk: totalProduk,
    totalPesanan: totalPesanan 
  };
}

/**
 * FUNGSI TAMBAH PRODUK BARU (ADMIN)
 */
function addProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Products");
  
  if (!ws) {
    ws = ss.insertSheet("Products");
    ws.appendRow(["ID Produk", "Nama", "Harga Agen", "Foto (JSON)", "Ukuran", "Warna", "Deskripsi", "Pelengkap (JSON)", "PO Start", "PO End", "Status", "Harga Asli", "Diskon (%)", "Varian Ukuran JSON"]);
  }
  
  const id = "PRD-" + new Date().getTime();
  const photosStr = JSON.stringify(data.photos || []);
  const addonsStr = JSON.stringify(data.addons || []);
  const sizeDetailsStr = JSON.stringify(data.sizeDetails || []);
  
  ws.appendRow([
    id, 
    data.name, 
    data.price, 
    photosStr, 
    data.sizes, 
    data.colors, 
    data.desc, 
    addonsStr, 
    data.poStart, 
    data.poEnd, 
    "Aktif",
    data.originalPrice || data.price,
    data.discount || 0,
    sizeDetailsStr
  ]);
  
  return { success: true, message: "Produk berhasil ditambahkan ke database!" };
}

/**
 * FUNGSI UPDATE PRODUK (ADMIN)
 */
function updateProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Products");
  if (!ws) return { success: false, message: "Database Produk tidak ditemukan." };
  
  const sheetData = ws.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][0] === data.id) {
      rowIndex = i + 1; 
      break;
    }
  }
  
  if (rowIndex === -1) return { success: false, message: "Produk tidak ditemukan di database." };
  
  const photosStr = JSON.stringify(data.photos || []);
  const addonsStr = JSON.stringify(data.addons || []);
  const sizeDetailsStr = JSON.stringify(data.sizeDetails || []);
  const currentStatus = sheetData[rowIndex - 1][10] || "Aktif"; 
  
  const range = ws.getRange(rowIndex, 2, 1, 13); 
  range.setValues([[
    data.name, 
    data.price, 
    photosStr, 
    data.sizes, 
    data.colors, 
    data.desc, 
    addonsStr, 
    data.poStart, 
    data.poEnd,
    currentStatus,
    data.originalPrice || data.price,
    data.discount || 0,
    sizeDetailsStr
  ]]);
  
  return { success: true, message: "Perubahan produk berhasil disimpan!" };
}

/**
 * FUNGSI HAPUS PRODUK
 */
function deleteProductBackEnd(productId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Products");
  if (!ws) return { success: false, message: "Database Produk tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      ws.deleteRow(i + 1); 
      return { success: true, message: "Produk berhasil dihapus." };
    }
  }
  return { success: false, message: "Data produk tidak ditemukan untuk dihapus." };
}

/**
 * FUNGSI AMBIL DAFTAR PRODUK (AGEN & ADMIN)
 */
function getAvailableProducts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Products");
  
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const products = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][10] !== "Aktif") continue; 
    
    let photos = [];
    try { photos = JSON.parse(data[i][3]); } catch(e) {}
    
    let sizeDetails = [];
    try { 
      if(data[i][13]) sizeDetails = JSON.parse(data[i][13]); 
    } catch(e) {}

    let originalPrice = (data[i][11] !== undefined && data[i][11] !== "") ? data[i][11] : data[i][2];
    let discount = (data[i][12] !== undefined && data[i][12] !== "") ? data[i][12] : 0;
    
    if (!sizeDetails || sizeDetails.length === 0) {
        let sizeList = data[i][4] ? String(data[i][4]).split(',').map(s => s.trim()) : ['All Size'];
        sizeDetails = sizeList.map(s => ({
            size: s, weight: 0, originalPrice: originalPrice, discount: discount, finalPrice: data[i][2]
        }));
    }

    let priceFormatted = data[i][2];
    if(!isNaN(priceFormatted) && priceFormatted !== "") {
       priceFormatted = "Rp " + Number(data[i][2]).toLocaleString('id-ID');
    }

    let originalPriceFormatted = originalPrice;
    if(!isNaN(originalPriceFormatted) && originalPriceFormatted !== "") {
       originalPriceFormatted = "Rp " + Number(originalPrice).toLocaleString('id-ID');
    }
    
    let poStartVal = data[i][8];
    let poEndVal = data[i][9];
    if (poStartVal instanceof Date) poStartVal = poStartVal.toISOString();
    else if (poStartVal) poStartVal = String(poStartVal);
    else poStartVal = "";
    
    if (poEndVal instanceof Date) poEndVal = poEndVal.toISOString();
    else if (poEndVal) poEndVal = String(poEndVal);
    else poEndVal = "";

    products.push({
      id: data[i][0],
      name: data[i][1],
      price: priceFormatted,
      rawPrice: data[i][2],
      originalPrice: originalPrice,
      originalPriceFormatted: originalPriceFormatted,
      discount: discount,
      image: photos.length > 0 ? photos[0] : 'https://via.placeholder.com/300?text=No+Image',
      photos: photos,
      sizes: data[i][4],
      sizeDetails: sizeDetails, 
      colors: data[i][5],
      desc: data[i][6],
      addons: data[i][7],
      poStart: poStartVal,
      poEnd: poEndVal, 
      stock: 'PO',
      status: data[i][10]
    });
  }
  
  return products.reverse();
}

/**
 * ==========================================
 * FITUR TRANSAKSI / PESANAN (AGEN & ADMIN)
 * ==========================================
 */

function submitAgenOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Orders");
  
  if (!ws) {
    ws = ss.insertSheet("Orders");
    ws.appendRow(["ID Pesanan", "ID Produk", "Nama Produk", "Harga Satuan", "Qty", "Total Harga", "Ukuran", "Warna", "Catatan", "Email Agen", "Nama Agen", "Tanggal Pesanan", "Status", "Berat Total"]);
  }
  
  ws.appendRow([
    data.id_pesanan,
    data.id_produk,
    data.nama_produk,
    data.harga_satuan,
    data.qty,
    data.total_harga,
    data.ukuran,
    data.warna,
    data.catatan,
    data.email_agen,
    data.nama_agen,
    data.tanggal,
    "Menunggu",
    data.berat || 0 
  ]);
  
  return { success: true, message: "Pesanan berhasil dibuat." };
}

/**
 * Helper: Build weight lookup map from Products sheet (CACHED - baca sekali saja)
 * Returns object: { "PRD-xxx|L": 280, "PRD-xxx|M": 260, ... }
 */
function buildWeightMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName("Products");
  if (!productSheet) return {};
  
  const productData = productSheet.getDataRange().getValues();
  const weightMap = {};
  
  for (let i = 1; i < productData.length; i++) {
    if (!productData[i][0]) continue;
    const productId = productData[i][0];
    try {
      if (productData[i][13]) {
        const sizeDetails = JSON.parse(productData[i][13]);
        if (Array.isArray(sizeDetails)) {
          sizeDetails.forEach(d => {
            if (d.size && d.weight) {
              weightMap[productId + "|" + d.size] = parseInt(d.weight) || 0;
            }
          });
        }
      }
    } catch(e) {}
  }
  return weightMap;
}

/**
 * Helper: Get dynamic weight using pre-built weightMap (NO extra sheet reads)
 */
function calcDynamicWeight(weightMap, idProduk, ukuran, qty) {
  const key = idProduk + "|" + ukuran;
  const weightPerPcs = weightMap[key] || 0;
  return weightPerPcs * (parseInt(qty) || 0);
}

function getMyOrders(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Orders");
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const weightMap = buildWeightMap(); // Baca Products SEKALI saja
  const myOrders = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] && data[i][9].toString().trim().toLowerCase() === email.toString().trim().toLowerCase()) {
      const idProduk = data[i][1];
      const ukuran = data[i][6];
      const qty = data[i][4];
      
      myOrders.push({
        id_pesanan: data[i][0],
        id_produk: idProduk,
        nama_produk: data[i][2],
        harga_satuan: data[i][3], 
        qty: qty,
        total_harga: data[i][5],
        ukuran: ukuran,
        warna: data[i][7],
        catatan: data[i][8],      
        tanggal: data[i][11],
        status: data[i][12],
        berat: calcDynamicWeight(weightMap, idProduk, ukuran, qty)
      });
    }
  }
  return myOrders.reverse();
}

function updateAgenOrder(orderData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Orders");
  if (!ws) return { success: false, message: "Database Pesanan tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderData.id_pesanan) {
      if (data[i][12] !== "Menunggu") {
        return { success: false, message: "Pesanan yang sudah diproses Admin tidak dapat diubah." };
      }
      
      ws.getRange(i + 1, 5).setValue(orderData.qty);
      ws.getRange(i + 1, 6).setValue(orderData.total_harga);
      ws.getRange(i + 1, 7).setValue(orderData.ukuran);
      ws.getRange(i + 1, 8).setValue(orderData.warna);
      ws.getRange(i + 1, 9).setValue(orderData.catatan);
      ws.getRange(i + 1, 14).setValue(orderData.berat || 0); 
      
      return { success: true, message: "Perubahan pesanan berhasil disimpan!" };
    }
  }
  return { success: false, message: "Pesanan tidak ditemukan di database." };
}

function getAllOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Orders");
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const weightMap = buildWeightMap(); // Baca Products SEKALI saja
  const allOrders = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    const idProduk = data[i][1];
    const ukuran = data[i][6];
    const qty = data[i][4];
    
    allOrders.push({
      id_pesanan: data[i][0],
      id_produk: idProduk, 
      nama_produk: data[i][2],
      qty: qty,
      total_harga: data[i][5],
      ukuran: ukuran,
      warna: data[i][7],
      catatan: data[i][8],
      email_agen: data[i][9],
      nama_agen: data[i][10],
      tanggal: data[i][11],
      status: data[i][12],
      berat: calcDynamicWeight(weightMap, idProduk, ukuran, qty)
    });
  }
  return allOrders.reverse();
}

function updateOrderStatusBackEnd(orderIdString, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Orders");
  if (!ws) return { success: false, message: "Database Pesanan tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  const idsArray = orderIdString.split(','); 
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (idsArray.includes(data[i][0])) {
      ws.getRange(i + 1, 13).setValue(newStatus); 
      updatedCount++;
    }
  }
  
  if (updatedCount > 0) {
     return { success: true, message: `Status pesanan berhasil diperbarui!` };
  } else {
     return { success: false, message: "Pesanan tidak ditemukan." };
  }
}

function resetOrdersBackEnd() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsOrders = ss.getSheetByName("Orders");
  const wsUsers = ss.getSheetByName("Users");
  
  if (!wsOrders) return { success: false, message: "Database Pesanan tidak ditemukan." };
  
  let isReset = false;

  // 1. Menghapus semua data riwayat di sheet Orders
  const lastRow = wsOrders.getLastRow();
  if (lastRow > 1) {
    wsOrders.deleteRows(2, lastRow - 1);
    isReset = true;
  }
  
  // 2. Mereset data Tagihan (Ongkir & Dana Masuk) di sheet Users
  if (wsUsers) {
    const lastRowUsers = wsUsers.getLastRow();
    if (lastRowUsers > 1) {
      const dataUsers = wsUsers.getRange(2, 1, lastRowUsers - 1, 7).getValues();
      // UPDATE: Tambahkan Array pengiriman: [] saat mereset
      const emptyBillingStr = JSON.stringify({ pengiriman: [], ongkir: 0, dana_masuk: 0, status_pembayaran: 'Belum Lunas' });
      const newBillingColumn = [];
      
      for (let i = 0; i < dataUsers.length; i++) {
        if (dataUsers[i][3] && String(dataUsers[i][3]).trim().toLowerCase() === "agen") {
          newBillingColumn.push([emptyBillingStr]);
        } else {
          newBillingColumn.push([dataUsers[i][6]]); 
        }
      }
      
      wsUsers.getRange(2, 7, newBillingColumn.length, 1).setValues(newBillingColumn);
      isReset = true;
    }
  }

  if (isReset) {
    return { success: true, message: "Semua data pesanan & tagihan keuangan berhasil dikosongkan!" };
  } else {
    return { success: false, message: "Database pesanan memang sudah kosong." };
  }
}

function getAllAgents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Users");
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const agents = [];
  const noHpIdx = findNoHpIndex_(data[0]);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] && data[i][3].toString().trim().toLowerCase() === "agen") {
      let addressObj = null;
      let addressStr = '<span class="text-red-400 italic">Alamat belum diatur</span>';
      
      try {
        if (data[i][5]) {
          addressObj = JSON.parse(data[i][5]);
          if (addressObj && addressObj.jalan) {
            addressStr = `<span class="font-semibold text-gray-700">${addressObj.nama_penerima || data[i][0]}</span><br>${addressObj.jalan}, RT ${addressObj.rt}/RW ${addressObj.rw}, Kec. ${addressObj.kecamatan}, Kab. ${addressObj.kabupaten}, Prov. ${addressObj.provinsi}`;
          }
        }
      } catch(e) {}

      let billingObj = { pengiriman: [], ongkir: 0, dana_masuk: 0, status_pembayaran: 'Belum Lunas' };
      try {
        if (data[i][6]) {
            let parsed = JSON.parse(data[i][6]);
            billingObj.pengiriman = parsed.pengiriman || [];
            billingObj.ongkir = parsed.ongkir || 0;
            billingObj.dana_masuk = parsed.dana_masuk || 0;
            billingObj.status_pembayaran = parsed.status_pembayaran || 'Belum Lunas';
        }
      } catch(e) {}

      agents.push({
        id: "AGN-" + String(i).padStart(3, '0'), 
        name: data[i][0],
        email: data[i][1],
        registerDate: data[i][4],
        address: addressStr,
        no_hp: (noHpIdx >= 0 && data[i][noHpIdx]) ? String(data[i][noHpIdx]) : "",
        billing: billingObj
      });
    }
  }
  return agents.reverse(); 
}

// UPDATE: Sekarang menerima array pengirimanArr dan totalOngkir
function updateAgentBilling(email, pengirimanArr, totalOngkir, danaMasuk, statusPembayaran) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { success: false, message: "Database tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      let billingObj = { 
          pengiriman: pengirimanArr || [], 
          ongkir: parseInt(totalOngkir) || 0, 
          dana_masuk: parseInt(danaMasuk) || 0, 
          status_pembayaran: statusPembayaran 
      };
      ws.getRange(i + 1, 7).setValue(JSON.stringify(billingObj)); 
      return { success: true, message: "Tagihan agen berhasil diperbarui!" };
    }
  }
  return { success: false, message: "Akun agen tidak ditemukan." };
}

function getAgentBilling(email) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { pengiriman: [], ongkir: 0, dana_masuk: 0, status_pembayaran: 'Belum Lunas' };
  
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      let billingObj = { pengiriman: [], ongkir: 0, dana_masuk: 0, status_pembayaran: 'Belum Lunas' };
      try {
        if (data[i][6]) {
            let parsed = JSON.parse(data[i][6]);
            billingObj.pengiriman = parsed.pengiriman || [];
            billingObj.ongkir = parsed.ongkir || 0;
            billingObj.dana_masuk = parsed.dana_masuk || 0;
            billingObj.status_pembayaran = parsed.status_pembayaran || 'Belum Lunas';
        }
      } catch(e) {}
      return billingObj;
    }
  }
  return { pengiriman: [], ongkir: 0, dana_masuk: 0, status_pembayaran: 'Belum Lunas' };
}

function updateAgenAddress(email, addressData) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { success: false, message: "Database tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      ws.getRange(i + 1, 6).setValue(JSON.stringify(addressData)); 
      return { success: true, message: "Alamat berhasil disimpan!" };
    }
  }
  return { success: false, message: "Akun tidak ditemukan." };
}

/**
 * ==========================================
 * FITUR NOMOR HP / WHATSAPP AGEN
 * ==========================================
 * Kolom "no_hp" ditambahkan otomatis pada sheet Users jika belum ada.
 * Tidak mengubah/menggeser kolom lama (Nama, Email, Password, Role,
 * Tanggal Daftar, Alamat JSON, Keuangan JSON) sehingga semua logika
 * berbasis index kolom lama tetap aman.
 */

// Helper: pastikan header "no_hp" ada, kembalikan nomor kolom (1-based)
function getNoHpColIndex_(ws) {
  const lastCol = ws.getLastColumn();
  if (lastCol > 0) {
    const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim().toLowerCase() === "no_hp") return c + 1;
    }
  }
  // Header belum ada -> buat kolom baru di paling kanan (tidak merusak data lama)
  const newCol = (lastCol > 0 ? lastCol : 0) + 1;
  ws.getRange(1, newCol).setValue("no_hp");
  return newCol;
}

// Helper: cari index (0-based) kolom no_hp dari baris header. -1 jika belum ada.
function findNoHpIndex_(headerRow) {
  if (!headerRow) return -1;
  for (let c = 0; c < headerRow.length; c++) {
    if (String(headerRow[c]).trim().toLowerCase() === "no_hp") return c;
  }
  return -1;
}

// Helper: validasi & rapikan nomor HP/WhatsApp
function validateNoHp_(noHp) {
  if (noHp === undefined || noHp === null || String(noHp).trim() === "") {
    return { valid: false, message: "Nomor HP/WhatsApp wajib diisi." };
  }
  const raw = String(noHp).trim();
  if (!/^[0-9+\-\s]+$/.test(raw)) {
    return { valid: false, message: "Nomor HP/WhatsApp hanya boleh berisi angka, spasi, tanda + dan -." };
  }
  const digits = raw.replace(/\D/g, "");
  if (digits.length < 10) return { valid: false, message: "Nomor HP/WhatsApp minimal 10 digit angka." };
  if (digits.length > 15) return { valid: false, message: "Nomor HP/WhatsApp maksimal 15 digit angka." };
  return { valid: true, clean: digits };
}

// Simpan / update nomor HP agen ke sheet Users (kolom no_hp)
function updateAgenPhone(email, no_hp) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { success: false, message: "Database tidak ditemukan." };

  const validation = validateNoHp_(no_hp);
  if (!validation.valid) return { success: false, message: validation.message };

  const colIndex = getNoHpColIndex_(ws); // memastikan header no_hp tersedia
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      // Simpan sebagai TEKS agar angka 0 di depan tidak hilang (cth: 0812... tetap utuh)
      const cell = ws.getRange(i + 1, colIndex);
      cell.setNumberFormat("@");
      cell.setValue(validation.clean);
      SpreadsheetApp.flush(); // pastikan perubahan benar-benar tersimpan ke sheet
      return { success: true, message: "Nomor HP/WhatsApp berhasil disimpan!", no_hp: validation.clean };
    }
  }
  return { success: false, message: "Akun tidak ditemukan." };
}

// Cek status nomor HP agen (opsional, dipakai bila diperlukan frontend)
function getUserPhoneStatus(email) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { success: false, no_hp: "", hasPhone: false };

  const data = ws.getDataRange().getValues();
  const idx = findNoHpIndex_(data[0]);
  const cleanEmail = String(email).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      const noHp = (idx >= 0 && data[i][idx]) ? String(data[i][idx]) : "";
      return { success: true, no_hp: noHp, hasPhone: noHp.trim() !== "" };
    }
  }
  return { success: false, no_hp: "", hasPhone: false };
}

/**
 * ==========================================
 * FUNGSI HASHING & AUTENTIKASI UTAMA
 * ==========================================
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length === 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

/**
 * FUNGSI LOGIN (SIGN IN)
 */
function prosesLogin(email, password) {
  if (!email || !password) return { success: false, message: "Email dan Password wajib diisi." };
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!ws) return { success: false, message: "Database tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  const hashedInput = hashPassword(String(password));
  const noHpIdx = findNoHpIndex_(data[0]);
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    const dbEmail = String(data[i][1]).trim().toLowerCase();
    const dbPass = String(data[i][2]);
    
    if (dbEmail === cleanEmail) {
      if (dbPass === hashedInput || dbPass === String(password)) {
        if (dbPass === String(password)) {
          ws.getRange(i + 1, 3).setValue(hashedInput); // Auto Migrate ke Hash
        }
        let addressData = null;
        try { addressData = data[i][5] ? JSON.parse(data[i][5]) : null; } catch(e) {}
        
        let noHp = (noHpIdx >= 0 && data[i][noHpIdx]) ? String(data[i][noHpIdx]) : "";
        
        return { success: true, message: "Login berhasil!", user: { name: data[i][0], email: data[i][1], role: data[i][3] || "Agen", address: addressData, no_hp: noHp } };
      }
    }
  }
  return { success: false, message: "Email atau Password salah." };
}

/**
 * FUNGSI PENDAFTARAN (SIGN UP) 
 */
function prosesDaftar(nama, email, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Users");
  
  if (!ws) {
    ws = ss.insertSheet("Users");
    ws.appendRow(["Nama", "Email", "Password", "Role", "Tanggal Daftar", "Alamat (JSON)", "Keuangan (JSON)"]);
  }

  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      return { success: false, message: "Email sudah terdaftar." };
    }
  }
  
  const hashedPassword = hashPassword(String(password));
  const tanggal = new Date().toLocaleString('id-ID');
  ws.appendRow([nama, email, hashedPassword, "Agen", tanggal]);
  return { success: true, message: "Pendaftaran berhasil!" };
}

function cekEmailTerdaftar(email) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      const resetLink = ScriptApp.getService().getUrl() + "?page=reset&email=" + encodeURIComponent(cleanEmail);
      try {
        MailApp.sendEmail({ to: cleanEmail, subject: "Reset Password - Laleh", htmlBody: `<p>Silakan klik link berikut: <a href="${resetLink}">Atur Ulang Sandi</a></p>` });
        return { success: true, message: "Email berhasil dikirim!" };
      } catch (e) { return { success: false, message: "Gagal mengirim email." }; }
    }
  }
  return { success: false, message: "Email tidak ditemukan." };
}

function updatePasswordBaru(email, newPassword) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const data = ws.getDataRange().getValues();
  const cleanEmail = String(email).trim().toLowerCase();
  const hashedPassword = hashPassword(String(newPassword));
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      ws.getRange(i + 1, 3).setValue(hashedPassword); 
      return { success: true, message: "Sandi berhasil diperbarui!" };
    }
  }
  return { success: false, message: "Gagal!" };
}

/**
 * ==========================================
 * FITUR PENDAFTARAN BARU DENGAN OTP VERIFIKASI
 * ==========================================
 */
function requestOtpRegistration(nama, email, password) {
  const cleanEmail = String(email).trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let wsUsers = ss.getSheetByName("Users");
  if (!wsUsers) {
      wsUsers = ss.insertSheet("Users");
      wsUsers.appendRow(["Nama", "Email", "Password", "Role", "Tanggal Daftar", "Alamat (JSON)", "Keuangan (JSON)"]);
  }
  
  const data = wsUsers.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
      return { success: false, message: "Email ini sudah terdaftar sebagai agen. Silakan login." };
    }
  }
  
  const otpCode = Math.floor(100000 + Math.random() * 900000).toString();
  const hashedPassword = hashPassword(String(password));
  
  const cache = CacheService.getScriptCache();
  const cachedData = JSON.stringify({ nama: nama, password: hashedPassword, otp: otpCode });
  cache.put("OTP_REG_" + cleanEmail, cachedData, 600); 
  
  try {
    // [PERBAIKAN] Validasi format email sebelum mengirim agar tidak gagal diam-diam (penyebab OTP "tidak terkirim")
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(cleanEmail)) {
      return { success: false, message: "Format email tidak valid. Periksa kembali alamat email Anda." };
    }

    // [PERBAIKAN] Pastikan kuota pengiriman email harian masih tersedia (penyebab umum OTP tidak terkirim)
    if (MailApp.getRemainingDailyQuota() <= 0) {
      return { success: false, message: "Kuota pengiriman email harian telah habis. Silakan coba lagi besok." };
    }

    const subject = "Kode Verifikasi Pendaftaran Agen Laleh";

    // [PERBAIKAN ANTI-SPAM] Sediakan versi teks biasa (plain text) agar email tidak terdeteksi sebagai spam (aturan HTML-only)
    const plainBody =
      "Halo " + nama + ",\n\n" +
      "Terima kasih telah mendaftar menjadi Agen Laleh.id.\n\n" +
      "Kode OTP verifikasi Anda: " + otpCode + "\n\n" +
      "Kode ini hanya berlaku selama 10 menit. Jangan beritahukan kode ini kepada siapa pun.\n\n" +
      "Email otomatis, mohon tidak dibalas.\n" +
      "Salam,\nTim Laleh.id";

    const body = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 10px;">
        <h2 style="color: #825E4A;">Halo ${nama},</h2>
        <p>Terima kasih telah mendaftar menjadi Agen Laleh.id.</p>
        <p>Berikut adalah 6 digit kode OTP verifikasi Anda:</p>
        <div style="text-align: center; margin: 30px 0;">
          <h1 style="color: #825E4A; font-size: 36px; letter-spacing: 8px; background: #fdfaf8; padding: 15px; border-radius: 10px; display: inline-block; margin:0;">${otpCode}</h1>
        </div>
        <p>Kode ini <strong>hanya berlaku selama 10 menit</strong>. Jangan beritahukan kode ini kepada siapa pun.</p>
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        <p style="font-size: 12px; color: #999;">Email otomatis, mohon tidak dibalas.<br>Salam,<br><strong>Tim Laleh.id</strong></p>
      </div>
    `;
    // [PERBAIKAN ANTI-SPAM] Tambahkan nama pengirim (name) & versi teks (body) untuk meningkatkan deliverability
    MailApp.sendEmail({
      to: cleanEmail,
      subject: subject,
      body: plainBody,
      htmlBody: body,
      name: "Laleh.id"
    });
    return { success: true, message: "Kode OTP berhasil dikirim ke email Anda. Cek juga folder Spam/Promosi bila tidak ada di Kotak Masuk." };
  } catch (error) {
    return { success: false, message: "Gagal mengirim email verifikasi: " + error.message };
  }
}

function verifyOtpAndRegister(email, otpInput) {
  const cleanEmail = String(email).trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  
  const cachedString = cache.get("OTP_REG_" + cleanEmail);
  
  if (!cachedString) {
    return { success: false, message: "Kode OTP sudah kedaluwarsa atau belum dikirim. Silakan minta kode baru." };
  }
  
  const cachedData = JSON.parse(cachedString);
  
  if (cachedData.otp !== String(otpInput)) {
    return { success: false, message: "Kode OTP yang Anda masukkan salah. Silakan coba lagi." };
  }
  
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  const tanggal = new Date().toLocaleString('id-ID');
  ws.appendRow([cachedData.nama, cleanEmail, cachedData.password, "Agen", tanggal]);
  
  cache.remove("OTP_REG_" + cleanEmail);
  
  return { success: true, message: "Verifikasi berhasil! Akun Anda kini sudah aktif." };
}

function kirimOtpUbahPassword(email) {
  try {
    var cleanEmail = String(email).trim().toLowerCase();
    var otp = Math.floor(100000 + Math.random() * 900000).toString(); 
    CacheService.getScriptCache().put('OTP_PW_' + cleanEmail, otp, 300);

    var subject = "🔒 Kode OTP Perubahan Password - Laleh.id";
    var body = `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 10px;">
        <h2 style="color: #825E4A;">Permintaan Ubah Kata Sandi</h2>
        <p>Halo,</p>
        <p>Kami menerima permintaan untuk mengubah password pada akun Agen Anda.</p>
        <p>Berikut adalah kode OTP verifikasi Anda:</p>
        <div style="text-align: center; margin: 30px 0;">
          <h1 style="color: #825E4A; font-size: 36px; letter-spacing: 8px; background: #fdfaf8; padding: 15px; border-radius: 10px; display: inline-block; margin:0;">${otp}</h1>
        </div>
        <p>Kode ini <strong>hanya berlaku selama 5 menit</strong>. Jangan memberikan kode ini kepada siapapun.</p>
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        <p style="font-size: 12px; color: #999;">Jika Anda tidak merasa melakukan permintaan ini, abaikan email ini dan pastikan akun Anda aman.</p>
      </div>
    `;

    MailApp.sendEmail({ to: cleanEmail, subject: subject, htmlBody: body });
    return { success: true, message: "Kode OTP berhasil dikirim ke email Anda. Cek Kotak Masuk atau folder Spam." };
  } catch (e) {
    return { success: false, message: "Gagal mengirim OTP: " + e.message };
  }
}

function verifikasiOtpUbahPassword(email, inputOtp) {
  var cleanEmail = String(email).trim().toLowerCase();
  var savedOtp = CacheService.getScriptCache().get('OTP_PW_' + cleanEmail);
  
  if (savedOtp && savedOtp === String(inputOtp)) {
    return { success: true, message: "OTP Valid." };
  } else {
    return { success: false, message: "Kode OTP salah atau sudah kedaluwarsa (lewat dari 5 menit)." };
  }
}

function simpanPasswordBaruOTP(email, passwordBaru) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Database Users tidak ditemukan." };
    
    var data = sheet.getDataRange().getValues();
    var cleanEmail = String(email).trim().toLowerCase();
    var hashedPw = hashPassword(String(passwordBaru));

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === cleanEmail) {
        sheet.getRange(i + 1, 3).setValue(hashedPw);
        CacheService.getScriptCache().remove('OTP_PW_' + cleanEmail);
        return { success: true, message: "Password berhasil diperbarui! Silakan gunakan sandi baru Anda." };
      }
    }
    return { success: false, message: "Email tidak ditemukan di sistem." };
  } catch (e) {
    return { success: false, message: "Kesalahan sistem: " + e.message };
  }
}

function deleteAgenOrder(orderId, emailAgen) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Orders");
  if (!ws) return { success: false, message: "Database Pesanan tidak ditemukan." };
  
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      if (data[i][9].toString().trim().toLowerCase() !== emailAgen.toString().trim().toLowerCase()) {
        return { success: false, message: "Akses ditolak. Anda hanya bisa membatalkan pesanan milik Anda sendiri." };
      }
      if (data[i][12] !== "Menunggu") {
        return { success: false, message: "Pesanan ini sudah diproses oleh Admin dan tidak dapat dibatalkan." };
      }
      
      ws.deleteRow(i + 1); 
      return { success: true, message: "Pesanan berhasil dibatalkan dan dihapus." };
    }
  }
  return { success: false, message: "Pesanan tidak ditemukan di database." };
}

/**
 * ==========================================
 * FUNGSI KELOLA JASA KIRIM / EKSPEDISI (NEW)
 * ==========================================
 */

/**
 * FUNGSI AMBIL DAFTAR EKSPEDISI DARI SHEET "Settings"
 */
function getEkspedisiList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Settings");
  
  // Jika sheet Settings belum ada, buat otomatis dan isi nilai default
  if (!ws) {
    ws = ss.insertSheet("Settings");
    ws.appendRow(["Daftar Ekspedisi"]); 
    ws.getRange("A1").setFontWeight("bold");
    
    // Nilai default awal
    const defaultEkspedisi = ["JNE", "J&T Express", "Sicepat", "Shopee Express", "Pos Indonesia", "Paxel", "Indah Cargo", "Lalamove", "Kurir Pribadi/Toko", "Ambil ke Toko"];
    const dataToInsert = defaultEkspedisi.map(e => [e]);
    ws.getRange(2, 1, dataToInsert.length, 1).setValues(dataToInsert);
  }
  
  // Baca seluruh data di Kolom A (mulai baris ke-2)
  const data = ws.getRange("A2:A").getValues(); 
  const list = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && String(data[i][0]).trim() !== "") {
      list.push(String(data[i][0]).trim());
    }
  }
  return list;
}

/**
 * FUNGSI SIMPAN DAFTAR EKSPEDISI BARU DARI WEB (ADMIN)
 */
function updateEkspedisiList(newList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Settings");
  
  if (!ws) {
    ws = ss.insertSheet("Settings");
  }
  
  // Bersihkan seluruh isi Kolom A
  ws.getRange("A:A").clearContent();
  
  // Tulis ulang Header
  ws.getRange("A1").setValue("Daftar Ekspedisi").setFontWeight("bold");
  
  // Masukkan data baru jika ada
  if (newList && newList.length > 0) {
    const dataToInsert = newList.map(e => [e]);
    ws.getRange(2, 1, dataToInsert.length, 1).setValues(dataToInsert);
  }
  
  return { success: true, message: "Daftar Jasa Kirim berhasil diperbarui!" };
}



// ==========================================
// FITUR ULASAN / REVIEW AGEN
// ==========================================

/**
 * Submit ulasan baru dari agen
 */
function submitReview(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Reviews");
  
  if (!ws) {
    ws = ss.insertSheet("Reviews");
    ws.appendRow(["id", "email_agen", "nama_agen", "rating", "pesan", "tanggal"]);
  }
  
  const id = "REV-" + new Date().getTime();
  ws.appendRow([
    id,
    data.email_agen,
    data.nama_agen,
    data.rating,
    data.pesan,
    new Date().toISOString()
  ]);
  
  return { success: true, message: "Ulasan berhasil dikirim! Terima kasih atas feedback Anda." };
}

/**
 * Ambil semua ulasan (untuk Admin)
 */
function getAllReviews() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Reviews");
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const reviews = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    reviews.push({
      id: data[i][0],
      email_agen: data[i][1],
      nama_agen: data[i][2],
      rating: data[i][3],
      pesan: data[i][4],
      tanggal: data[i][5]
    });
  }
  
  return reviews.reverse();
}

/**
 * Ambil ulasan milik agen tertentu
 */
function getMyReviews(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Reviews");
  if (!ws) return [];
  
  const data = ws.getDataRange().getValues();
  const reviews = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      reviews.push({
        id: data[i][0],
        email_agen: data[i][1],
        nama_agen: data[i][2],
        rating: data[i][3],
        pesan: data[i][4],
        tanggal: data[i][5]
      });
    }
  }
  
  return reviews.reverse();
}
