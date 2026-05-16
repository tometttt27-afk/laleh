// ==========================================
// INVOICE GENERATOR SYSTEM - Noora Modest Wear
// ==========================================

// [OPTIMASI] Menggunakan Active Spreadsheet (sama seperti DbUsers.js)
function getInvoiceSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ==========================================
// KONFIGURASI INVOICE
// ==========================================
const INVOICE_CONFIG = {
  companyName: "LALEH.ID",
  companyTagline: "Pakaian Muslim Berkualitas Premium",
  companyAddress: "Jl. Sumpiuh, Sumpiuh, Banyumas, Jawa Tengah 53195",
  companyPhone: "+62 889-8063-9722",
  companyEmail: "lalehgroupid@gmail.com",
  companyWebsite: "www.laleh.com",
  brandColor: "#825E4A",
  accentColor: "#FDF8F5"
};

// ==========================================
// 1. GENERATE INVOICE NUMBER
// ==========================================
function generateInvoiceNumber() {
  const now = new Date();
  const year = now.getFullYear().toString().slice(-2);
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const time = now.getTime().toString().slice(-6);
  return `INV-${year}${month}${day}-${time}`;
}

// ==========================================
// 2. SAVE INVOICE RECORD
// ==========================================
function saveInvoiceRecord(invoiceData) {
  try {
    const ss = getInvoiceSS();
    let sheet = ss.getSheetByName('Invoices');
    
    if (!sheet) {
      sheet = ss.insertSheet('Invoices');
      sheet.appendRow([
        "Invoice Number", "Tanggal", "Nama Agen", "Email Agen", 
        "Total Harga Barang", "Total Ongkir", "Total Invoice", 
        "Status Pembayaran", "PDF URL", "Timestamp"
      ]);
    }
    
    sheet.appendRow([
      invoiceData.invoiceNumber,
      invoiceData.date,
      invoiceData.agentName,
      invoiceData.agentEmail,
      invoiceData.totalBarang,
      invoiceData.totalOngkir,
      invoiceData.grandTotal,
      invoiceData.statusPembayaran,
      invoiceData.pdfUrl || "",
      new Date().toISOString()
    ]);
    
    return { success: true, message: "Invoice record saved" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 3. GENERATE INVOICE DATA
// ==========================================
function generateInvoiceData(agentEmail) {
  try {
    const ss = getInvoiceSS();
    
    // Get Agent Data - Kolom sesuai DbUsers.js:
    // [0]=Nama, [1]=Email, [2]=Password, [3]=Role, [4]=RegisterDate, [5]=Address(JSON), [6]=Billing(JSON)
    const userSheet = ss.getSheetByName('Users');
    if (!userSheet) return { success: false, message: "Data user tidak ditemukan" };
    
    const userData = userSheet.getDataRange().getValues();
    let agent = null;
    const cleanEmail = String(agentEmail).trim().toLowerCase();
    
    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][1]).trim().toLowerCase() === cleanEmail) {
        let billing = {};
        try {
          billing = userData[i][6] ? JSON.parse(userData[i][6]) : {};
        } catch(e) {}
        
        // Parse address dari JSON
        let addressStr = "-";
        try {
          if (userData[i][5]) {
            let addrObj = JSON.parse(userData[i][5]);
            if (addrObj && addrObj.jalan) {
              addressStr = `${addrObj.nama_penerima || userData[i][0]}<br>${addrObj.jalan}, RT ${addrObj.rt || '-'}/RW ${addrObj.rw || '-'}, Kec. ${addrObj.kecamatan || '-'}, Kab. ${addrObj.kabupaten || '-'}, Prov. ${addrObj.provinsi || '-'}`;
            }
          }
        } catch(e) {}
        
        agent = {
          name: userData[i][0],
          email: userData[i][1],
          address: addressStr,
          billing: billing
        };
        break;
      }
    }
    
    if (!agent) return { success: false, message: "Agen tidak ditemukan" };
    
    // Get Orders Data - Kolom sesuai getAllOrders() di DbUsers.js:
    // [0]=id_pesanan, [1]=id_produk, [2]=nama_produk, [3]=???, [4]=qty, [5]=total_harga,
    // [6]=ukuran, [7]=warna, [8]=catatan, [9]=email_agen, [10]=nama_agen, [11]=tanggal, [12]=status, [13]=berat
    const orderSheet = ss.getSheetByName('Orders');
    if (!orderSheet) return { success: false, message: "Data pesanan tidak ditemukan" };
    
    const orderData = orderSheet.getDataRange().getValues();
    let orders = [];
    let totalHargaBarang = 0;
    let totalBerat = 0;
    
    // Build weight map ONCE for all orders (optimasi performa)
    const productSheet = ss.getSheetByName('Products');
    const weightMap = {};
    if (productSheet) {
      const prodData = productSheet.getDataRange().getValues();
      for (let p = 1; p < prodData.length; p++) {
        if (!prodData[p][0]) continue;
        try {
          if (prodData[p][13]) {
            const sizeDetails = JSON.parse(prodData[p][13]);
            if (Array.isArray(sizeDetails)) {
              sizeDetails.forEach(d => {
                if (d.size && d.weight) {
                  weightMap[prodData[p][0] + "|" + d.size] = parseInt(d.weight) || 0;
                }
              });
            }
          }
        } catch(e) {}
      }
    }
    
    for (let i = 1; i < orderData.length; i++) {
      if (!orderData[i][0]) continue; // Skip baris kosong
      
      let orderEmail = String(orderData[i][9] || '').trim().toLowerCase();
      let orderStatus = String(orderData[i][12] || '').trim();
      
      if (orderEmail === cleanEmail && orderStatus !== 'Dibatalkan') {
        const qty = parseInt(orderData[i][4]) || 0;
        const harga = parseInt(orderData[i][5]) || 0;
        const idProduk = orderData[i][1];
        const ukuran = orderData[i][6];
        
        // Calculate DYNAMIC weight from weightMap (fast lookup)
        const weightKey = idProduk + "|" + ukuran;
        const berat = (weightMap[weightKey] || 0) * qty;
        
        orders.push({
          id: orderData[i][0],
          produk: orderData[i][2] || 'Produk',
          ukuran: ukuran || '-',
          warna: orderData[i][7] || '-',
          qty: qty,
          harga: harga,
          subtotal: harga,
          berat: berat
        });
        
        totalHargaBarang += harga;
        totalBerat += berat;
      }
    }
    
    // Calculate Totals
    const totalOngkir = parseInt(agent.billing.ongkir) || 0;
    const grandTotal = totalHargaBarang + totalOngkir;
    const danaMasuk = parseInt(agent.billing.dana_masuk) || 0;
    const sisaTagihan = grandTotal - danaMasuk;
    const statusPembayaran = agent.billing.status_pembayaran || "Belum Lunas";
    
    // Generate Invoice Number
    const invoiceNumber = generateInvoiceNumber();
    const invoiceDate = new Date().toLocaleDateString('id-ID', { 
      day: '2-digit', 
      month: 'long', 
      year: 'numeric' 
    });
    
    return {
      success: true,
      invoice: {
        invoiceNumber: invoiceNumber,
        date: invoiceDate,
        agentName: agent.name,
        agentEmail: agent.email,
        agentAddress: agent.address,
        orders: orders,
        totalHargaBarang: totalHargaBarang,
        totalBerat: totalBerat,
        pengiriman: agent.billing.pengiriman || [],
        totalOngkir: totalOngkir,
        grandTotal: grandTotal,
        danaMasuk: danaMasuk,
        sisaTagihan: sisaTagihan,
        statusPembayaran: statusPembayaran
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 4. GENERATE HTML INVOICE TEMPLATE
// ==========================================
function generateInvoiceHTML(invoiceData) {
  const config = INVOICE_CONFIG;
  
  // Format Currency
  const formatRp = (num) => {
    return "Rp " + num.toLocaleString('id-ID');
  };
  
  // Parse HTML dari address
  const cleanAddress = invoiceData.agentAddress
    .replace(/<br\s*[\/]?>/gi, '<br>')
    .replace(/<span[^>]*>/gi, '<span>')
    .trim();
  
  // Generate Order Rows
  let orderRows = '';
  invoiceData.orders.forEach((order, index) => {
    orderRows += `
      <tr style="border-bottom: 1px solid #e5e7eb;">
        <td style="padding: 12px 8px; text-align: center; font-size: 13px;">${index + 1}</td>
        <td style="padding: 12px 8px; font-size: 13px;">
          <strong>${order.produk}</strong><br>
          <span style="color: #6b7280; font-size: 11px;">Size: ${order.ukuran} | Warna: ${order.warna}</span>
        </td>
        <td style="padding: 12px 8px; text-align: center; font-size: 13px;">${order.qty}</td>
        <td style="padding: 12px 8px; text-align: right; font-size: 13px;">${formatRp(order.harga)}</td>
        <td style="padding: 12px 8px; text-align: right; font-size: 13px; font-weight: 600;">${formatRp(order.subtotal)}</td>
      </tr>
    `;
  });
  
  // Generate Shipping Rows
  let shippingRows = '';
  if (invoiceData.pengiriman && invoiceData.pengiriman.length > 0) {
    invoiceData.pengiriman.forEach(p => {
      shippingRows += `
        <tr>
          <td style="padding: 6px 0; font-size: 13px; color: #6b7280;">${p.ekspedisi}</td>
          <td style="padding: 6px 0; text-align: right; font-size: 13px; font-weight: 500;">${formatRp(p.ongkir)}</td>
        </tr>
      `;
    });
  } else if (invoiceData.totalOngkir > 0) {
    shippingRows += `
      <tr>
        <td style="padding: 6px 0; font-size: 13px; color: #6b7280;">Ongkos Kirim</td>
        <td style="padding: 6px 0; text-align: right; font-size: 13px; font-weight: 500;">${formatRp(invoiceData.totalOngkir)}</td>
      </tr>
    `;
  }
  
  // Payment Status Badge
  let statusBadge = '';
  if (invoiceData.statusPembayaran === 'Lunas') {
    statusBadge = '<span style="background: #10b981; color: white; padding: 4px 12px; border-radius: 6px; font-size: 11px; font-weight: 600;">LUNAS</span>';
  } else if (invoiceData.statusPembayaran === 'DP') {
    statusBadge = '<span style="background: #f59e0b; color: white; padding: 4px 12px; border-radius: 6px; font-size: 11px; font-weight: 600;">DP</span>';
  } else {
    statusBadge = '<span style="background: #ef4444; color: white; padding: 4px 12px; border-radius: 6px; font-size: 11px; font-weight: 600;">BELUM LUNAS</span>';
  }
  
  const html = `
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Invoice ${invoiceData.invoiceNumber}</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Inter', 'Segoe UI', sans-serif; background: #f9fafb; padding: 20px; }
    .invoice-container { max-width: 800px; margin: 0 auto; background: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; }
    .invoice-header { background: ${config.brandColor}; color: white; padding: 30px; }
    .invoice-body { padding: 30px; }
    table { width: 100%; border-collapse: collapse; }
    .print-button { background: ${config.brandColor}; color: white; border: none; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; margin: 20px 0; }
    @media print {
      body { background: white; padding: 0; }
      .invoice-container { box-shadow: none; }
      .print-button { display: none; }
    }
  </style>
</head>
<body>
  <div class="invoice-container">
    <!-- HEADER -->
    <div class="invoice-header">
      <table style="width: 100%;">
        <tr>
          <td style="width: 60%;">
            <div style="font-size: 28px; font-weight: 700; margin-bottom: 4px;">${config.companyName}</div>
            <div style="font-size: 13px; opacity: 0.9; margin-bottom: 16px;">${config.companyTagline}</div>
            <div style="font-size: 12px; opacity: 0.85; line-height: 1.6;">
              ${config.companyAddress}<br>
              ${config.companyPhone}<br>
              ${config.companyEmail}
            </div>
          </td>
          <td style="width: 40%; text-align: right; vertical-align: top;">
            <div style="font-size: 24px; font-weight: 700; margin-bottom: 8px;">INVOICE</div>
            <div style="font-size: 13px; opacity: 0.9;">
              <strong>${invoiceData.invoiceNumber}</strong><br>
              ${invoiceData.date}
            </div>
          </td>
        </tr>
      </table>
    </div>

    <!-- BODY -->
    <div class="invoice-body">
      <!-- BILL TO -->
      <div style="margin-bottom: 30px;">
        <div style="font-size: 11px; color: #6b7280; text-transform: uppercase; font-weight: 600; margin-bottom: 8px;">Tagihan Kepada:</div>
        <div style="font-size: 16px; font-weight: 600; color: #111827; margin-bottom: 4px;">${invoiceData.agentName}</div>
        <div style="font-size: 13px; color: #6b7280; line-height: 1.6;">${cleanAddress}</div>
      </div>

      <!-- ORDER ITEMS -->
      <table style="margin-bottom: 20px;">
        <thead>
          <tr style="background: ${config.accentColor}; border-bottom: 2px solid ${config.brandColor};">
            <th style="padding: 12px 8px; text-align: center; font-size: 12px; font-weight: 600; width: 5%;">No</th>
            <th style="padding: 12px 8px; text-align: left; font-size: 12px; font-weight: 600; width: 40%;">Produk</th>
            <th style="padding: 12px 8px; text-align: center; font-size: 12px; font-weight: 600; width: 10%;">Qty</th>
            <th style="padding: 12px 8px; text-align: right; font-size: 12px; font-weight: 600; width: 20%;">Harga</th>
            <th style="padding: 12px 8px; text-align: right; font-size: 12px; font-weight: 600; width: 25%;">Subtotal</th>
          </tr>
        </thead>
        <tbody>
          ${orderRows}
        </tbody>
      </table>

      <!-- SUMMARY -->
      <div style="margin-top: 30px; border-top: 2px solid #e5e7eb; padding-top: 20px;">
        <table style="width: 100%; max-width: 350px; margin-left: auto;">
          <tr>
            <td style="padding: 6px 0; font-size: 13px; color: #6b7280;">Subtotal Produk</td>
            <td style="padding: 6px 0; text-align: right; font-size: 13px; font-weight: 500;">${formatRp(invoiceData.totalHargaBarang)}</td>
          </tr>
          ${shippingRows}
          <tr style="border-top: 2px solid #e5e7eb;">
            <td style="padding: 12px 0; font-size: 16px; font-weight: 700; color: #111827;">TOTAL</td>
            <td style="padding: 12px 0; text-align: right; font-size: 18px; font-weight: 700; color: ${config.brandColor};">${formatRp(invoiceData.grandTotal)}</td>
          </tr>
          <tr style="border-top: 1px solid #e5e7eb;">
            <td style="padding: 8px 0; font-size: 13px; color: #6b7280;">Dana Masuk</td>
            <td style="padding: 8px 0; text-align: right; font-size: 13px; font-weight: 500; color: #10b981;">-${formatRp(invoiceData.danaMasuk)}</td>
          </tr>
          <tr style="border-top: 2px solid ${config.brandColor}; background: ${config.accentColor};">
            <td style="padding: 12px 8px; font-size: 15px; font-weight: 700; color: #111827;">SISA TAGIHAN</td>
            <td style="padding: 12px 8px; text-align: right; font-size: 17px; font-weight: 700; color: ${invoiceData.sisaTagihan > 0 ? '#ef4444' : '#10b981'};">${formatRp(invoiceData.sisaTagihan)}</td>
          </tr>
        </table>
      </div>

      <!-- STATUS PEMBAYARAN -->
      <div style="margin-top: 30px; padding: 20px; background: #f9fafb; border-radius: 8px; border-left: 4px solid ${config.brandColor};">
        <table style="width: 100%;">
          <tr>
            <td style="font-size: 13px; color: #6b7280;">Status Pembayaran:</td>
            <td style="text-align: right;">${statusBadge}</td>
          </tr>
        </table>
      </div>

      <!-- FOOTER NOTE -->
      <div style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #e5e7eb; text-align: center; color: #6b7280; font-size: 11px;">
        <p style="margin-bottom: 4px;">Terima kasih atas kepercayaan Anda kepada ${config.companyName}</p>
        <p>Dokumen ini adalah invoice resmi yang sah secara elektronik</p>
      </div>
    </div>
  </div>

  <center>
    <button class="print-button" onclick="window.print()">📄 Cetak / Simpan PDF</button>
  </center>

  <script>
    // Auto-focus untuk print di beberapa browser
    setTimeout(() => {
      lucide.createIcons();
    }, 100);
  </script>
</body>
</html>
  `;
  
  return html;
}

// ==========================================
// 5. SEND INVOICE VIA EMAIL
// ==========================================
function sendInvoiceEmail(agentEmail, invoiceHTML, invoiceNumber) {
  try {
    const subject = `Invoice ${invoiceNumber} - ${INVOICE_CONFIG.companyName}`;
    const message = `
      <div style="font-family: Arial, sans-serif;">
        <p>Kepada Yth. Agen ${INVOICE_CONFIG.companyName},</p>
        <p>Terlampir invoice untuk pesanan Anda. Silakan lakukan pembayaran sesuai nominal yang tertera.</p>
        <p>Jika ada pertanyaan, hubungi kami di ${INVOICE_CONFIG.companyPhone}</p>
        <br>
        <p>Salam hangat,<br><strong>${INVOICE_CONFIG.companyName}</strong></p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: agentEmail,
      subject: subject,
      htmlBody: message + "<hr>" + invoiceHTML
    });
    
    return { success: true, message: "Invoice berhasil dikirim via email" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 6. MAIN FUNCTION: GET INVOICE HTML
// ==========================================
function getInvoiceHTML(agentEmail) {
  const result = generateInvoiceData(agentEmail);
  
  if (!result.success) {
    return result;
  }
  
  const html = generateInvoiceHTML(result.invoice);
  
  // Save invoice record
  saveInvoiceRecord({
    invoiceNumber: result.invoice.invoiceNumber,
    date: result.invoice.date,
    agentName: result.invoice.agentName,
    agentEmail: result.invoice.agentEmail,
    totalBarang: result.invoice.totalHargaBarang,
    totalOngkir: result.invoice.totalOngkir,
    grandTotal: result.invoice.grandTotal,
    statusPembayaran: result.invoice.statusPembayaran,
    pdfUrl: ""
  });
  
  return {
    success: true,
    html: html,
    invoiceNumber: result.invoice.invoiceNumber,
    invoiceData: result.invoice
  };
}

// ==========================================
// 7. GET INVOICE HISTORY
// ==========================================
function getInvoiceHistory() {
  try {
    const ss = getInvoiceSS();
    const sheet = ss.getSheetByName('Invoices');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let invoices = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        invoices.push({
          invoiceNumber: data[i][0],
          tanggal: data[i][1],
          namaAgen: data[i][2],
          emailAgen: data[i][3],
          totalBarang: parseInt(data[i][4]) || 0,
          totalOngkir: parseInt(data[i][5]) || 0,
          grandTotal: parseInt(data[i][6]) || 0,
          statusPembayaran: data[i][7],
          pdfUrl: data[i][8] || ""
        });
      }
    }
    
    return invoices.reverse();
  } catch (error) {
    return [];
  }
}

// ==========================================
// 8. GENERATE INVOICE TEXT FOR WHATSAPP
// ==========================================
function getInvoiceTextForWhatsApp(agentEmail) {
  const result = generateInvoiceData(agentEmail);
  
  if (!result.success) {
    return result;
  }
  
  const inv = result.invoice;
  const formatRp = (num) => "Rp " + num.toLocaleString('id-ID');
  
  let msg = `*INVOICE - ${INVOICE_CONFIG.companyName}*\n`;
  msg += `No: ${inv.invoiceNumber}\n`;
  msg += `Tanggal: ${inv.date}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  msg += `*Kepada:* ${inv.agentName}\n\n`;
  msg += `*RINCIAN PESANAN:*\n`;
  
  inv.orders.forEach((order, i) => {
    msg += `${i+1}. *${order.produk}*\n`;
    msg += `   Ukuran: ${order.ukuran} | Warna: ${order.warna}\n`;
    msg += `   Qty: ${order.qty} | ${formatRp(order.harga)}\n`;
  });
  
  msg += `\n━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `Subtotal Produk: ${formatRp(inv.totalHargaBarang)}\n`;
  
  if (inv.pengiriman && inv.pengiriman.length > 0) {
    inv.pengiriman.forEach(p => {
      msg += `Ongkir (${p.ekspedisi}): ${formatRp(p.ongkir)}\n`;
    });
  } else if (inv.totalOngkir > 0) {
    msg += `Ongkir: ${formatRp(inv.totalOngkir)}\n`;
  }
  
  msg += `━━━━━━━━━━━━━━━━━━━━\n`;
  msg += `*TOTAL: ${formatRp(inv.grandTotal)}*\n`;
  msg += `Dana Masuk: -${formatRp(inv.danaMasuk)}\n`;
  msg += `*SISA TAGIHAN: ${formatRp(inv.sisaTagihan)}*\n`;
  msg += `Status: ${inv.statusPembayaran}\n`;
  msg += `━━━━━━━━━━━━━━━━━━━━\n\n`;
  msg += `Terima kasih atas kepercayaan Anda.\n`;
  msg += `_${INVOICE_CONFIG.companyName}_\n`;
  msg += `${INVOICE_CONFIG.companyPhone}`;
  
  // Save invoice record
  saveInvoiceRecord({
    invoiceNumber: inv.invoiceNumber,
    date: inv.date,
    agentName: inv.agentName,
    agentEmail: inv.agentEmail,
    totalBarang: inv.totalHargaBarang,
    totalOngkir: inv.totalOngkir,
    grandTotal: inv.grandTotal,
    statusPembayaran: inv.statusPembayaran,
    pdfUrl: "via-whatsapp"
  });
  
  return {
    success: true,
    text: msg,
    invoiceNumber: inv.invoiceNumber
  };
}

// ==========================================
// 9. SEND INVOICE VIA EMAIL (Legacy - masih bisa dipakai)
// ==========================================
function sendInvoiceToAgent(agentEmail) {
  const result = getInvoiceHTML(agentEmail);
  
  if (!result.success) {
    return result;
  }
  
  return sendInvoiceEmail(agentEmail, result.html, result.invoiceNumber);
}
