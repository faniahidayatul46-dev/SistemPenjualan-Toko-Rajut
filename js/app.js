const config = window.SUPABASE_CONFIG || {};
const state = { products: [], customers: [], sales: [], saleItems: [], page: location.hash.slice(1) || "dashboard", customerSort: "newest", salesSort: "newest", salesCustomerFilter: "" };
const $ = (selector) => document.querySelector(selector);
const rupiah = (value) => new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(Number(value || 0));
const dateText = (value) => value ? new Intl.DateTimeFormat("id-ID", { day: "2-digit", month: "short", year: "numeric" }).format(new Date(value)) : "—";
const localIsoDate = () => {
  const date = new Date();
  date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
  return date.toISOString().slice(0, 10);
};
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" }[char]));
const parseMoney = (value) => Number(String(value || "").replace(/[^\d]/g, "")) || 0;
const moneyInput = (value) => value === "" || value === null || value === undefined ? "" : rupiah(value).replace(/\u00a0/g, " ");

function excelColumnName(index) {
  let name = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    name = String.fromCharCode(65 + (value - 1) % 26) + name;
  }
  return name;
}

function xmlEscape(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;"
  }[character]));
}

function crc32(bytes) {
  let crc = -1;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ -1) >>> 0;
}

function zipFiles(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name);
    const data = encoder.encode(file.content);
    const checksum = crc32(data);
    const localHeader = new Uint8Array(30 + name.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, name.length, true);
    localHeader.set(name, 30);
    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + name.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(name, 46);
    centralParts.push(centralHeader);
    offset += localHeader.length + data.length;
  }
  const centralDirectory = new Uint8Array(centralParts.reduce((sum, part) => sum + part.length, 0));
  let centralOffset = 0;
  for (const part of centralParts) {
    centralDirectory.set(part, centralOffset);
    centralOffset += part.length;
  }
  const endRecord = new Uint8Array(22);
  const endView = new DataView(endRecord.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralDirectory.length, true);
  endView.setUint32(16, offset, true);
  return new Blob([...localParts, centralDirectory, endRecord], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function worksheetXml(sheet) {
  const rows = [sheet.columns.map((column) => column.label), ...sheet.rows.map((row) =>
    sheet.columns.map((column) => typeof column.value === "function" ? column.value(row) : row[column.value])
  )];
  const lastColumn = excelColumnName(sheet.columns.length - 1);
  const rowXml = rows.map((values, rowIndex) => {
    const cells = values.map((value, columnIndex) => {
      const reference = `${excelColumnName(columnIndex)}${rowIndex + 1}`;
      if (rowIndex > 0 && typeof value === "number" && Number.isFinite(value)) {
        return `<c r="${reference}" s="2"><v>${value}</v></c>`;
      }
      return `<c r="${reference}" s="${rowIndex === 0 ? 1 : 0}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
    }).join("");
    return `<row r="${rowIndex + 1}"${rowIndex === 0 ? ' ht="32" customHeight="1"' : ""}>${cells}</row>`;
  }).join("");
  const columns = sheet.columns.map((column, index) =>
    `<col min="${index + 1}" max="${index + 1}" width="${column.width || 18}" customWidth="1"/>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${columns}</cols><sheetData>${rowXml}</sheetData><autoFilter ref="A1:${lastColumn}${rows.length}"/></worksheet>`;
}

function createExcelWorkbook(sheets) {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts><fonts count="2"><font><sz val="10"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF572D45"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFEADFD3"/></left><right style="thin"><color rgb="FFEADFD3"/></right><top style="thin"><color rgb="FFEADFD3"/></top><bottom style="thin"><color rgb="FFEADFD3"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const files = [
    { name: "[Content_Types].xml", content: contentTypes },
    { name: "_rels/.rels", content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", content: workbookXml },
    { name: "xl/_rels/workbook.xml.rels", content: workbookRels },
    { name: "xl/styles.xml", content: styles },
    ...sheets.map((sheet, index) => ({ name: `xl/worksheets/sheet${index + 1}.xml`, content: worksheetXml(sheet) }))
  ];
  return zipFiles(files);
}

function exportAllDataToExcel() {
  const salesById = new Map(state.sales.map((sale) => [sale.id, sale]));
  const customersById = new Map(state.customers.map((customer) => [customer.id, customer]));
  const productsById = new Map(state.products.map((product) => [product.id, product]));
  const typeLabels = { sale: "Penjualan", purchase: "Pembelian", stock_adjustment: "Penyesuaian stok", settlement: "Pelunasan" };
  const channelLabels = { in_store: "Beli langsung di toko", whatsapp_online: "WhatsApp / Online" };
  const deliveryLabels = { delivery: "Delivery", pickup: "Ambil di tempat" };
  const movementLabels = { in: "Stok masuk", damaged: "Produk rusak", lost: "Produk hilang" };
  const paymentLabels = { transfer: "Transfer Bank", qris: "QRIS" };
  const sheets = [
    {
      name: "Produk",
      columns: [
        { label: "ID", value: "id", width: 38 }, { label: "SKU", value: "sku", width: 16 }, { label: "Nama produk", value: "name", width: 28 },
        { label: "Kategori", value: "category", width: 20 }, { label: "Deskripsi", value: "description", width: 42 },
        { label: "Harga (Rp)", value: "price", width: 16 }, { label: "Stok", value: "stock", width: 10 },
        { label: "Minimum stok", value: "minimum_stock", width: 14 }, { label: "URL gambar", value: "image_url", width: 36 },
        { label: "Dibuat", value: "created_at", width: 24 }, { label: "Diperbarui", value: "updated_at", width: 24 }
      ],
      rows: state.products
    },
    {
      name: "Pelanggan",
      columns: [
        { label: "ID", value: "id", width: 38 }, { label: "Nama", value: "name", width: 26 }, { label: "Telepon", value: "phone", width: 20 },
        { label: "Email", value: "email", width: 32 }, { label: "Alamat", value: "address", width: 44 },
        { label: "Dibuat", value: "created_at", width: 24 }
      ],
      rows: state.customers
    },
    {
      name: "Transaksi & Jurnal",
      columns: [
        { label: "ID", value: "id", width: 38 }, { label: "No. bukti", value: "invoice_no", width: 20 },
        { label: "Jenis transaksi", value: (entry) => typeLabels[entry.entry_type || "sale"] || entry.entry_type, width: 22 },
        { label: "Tanggal", value: "sale_date", width: 14 },
        { label: "ID pelanggan", value: "customer_id", width: 38 },
        { label: "Pelanggan / relasi", value: (entry) => customersById.get(entry.customer_id)?.name || "", width: 25 },
        { label: "ID invoice terkait", value: "related_sale_id", width: 38 },
        { label: "Invoice terkait", value: (entry) => salesById.get(entry.related_sale_id)?.invoice_no || "", width: 20 },
        { label: "Keterangan", value: (entry) => entry.journal_description || (entry.entry_type === "sale" ? "Penjualan produk" : ""), width: 38 },
        { label: "Produk mutasi", value: (entry) => productsById.get(entry.stock_product_id)?.name || "", width: 25 },
        { label: "Mutasi stok", value: (entry) => movementLabels[entry.stock_movement] || "", width: 18 },
        { label: "Jumlah unit", value: (entry) => entry.stock_quantity == null ? "" : Number(entry.stock_quantity), width: 13 },
        { label: "Stok sebelum", value: (entry) => entry.stock_before == null ? "" : Number(entry.stock_before), width: 13 },
        { label: "Stok sesudah", value: (entry) => entry.stock_after == null ? "" : Number(entry.stock_after), width: 13 },
        { label: "Kanal", value: (entry) => channelLabels[entry.sales_channel] || "", width: 22 },
        { label: "Penerimaan", value: (entry) => deliveryLabels[entry.delivery_method] || "", width: 16 },
        { label: "Metode pembayaran", value: (entry) => paymentLabels[entry.payment_method] || "", width: 20 },
        { label: "Referensi pembayaran", value: "payment_reference", width: 24 },
        { label: "URL bukti pembayaran", value: "payment_proof_url", width: 40 },
        { label: "Status verifikasi", value: "payment_status", width: 20 },
        { label: "Status tagihan", value: (entry) => entry.entry_type === "settlement"
          ? entry.payment_status === "verified" ? "Pelunasan terverifikasi" : entry.payment_status === "rejected" ? "Pelunasan ditolak" : "Menunggu verifikasi"
          : entry.entry_type === "sale" ? paymentStatusLabel(entry.payment_status, entry.paid_amount, entry.total) : entry.payment_status, width: 30 },
        { label: "Subtotal (Rp)", value: (entry) => Number(entry.subtotal || 0), width: 16 },
        { label: "Ongkir (Rp)", value: (entry) => Number(entry.shipping_fee || 0), width: 16 },
        { label: "Total (Rp)", value: (entry) => Number(entry.total || 0), width: 16 },
        { label: "Dibayar (Rp)", value: (entry) => Number(entry.paid_amount || 0), width: 16 },
        { label: "Sisa tagihan (Rp)", value: (entry) => entry.entry_type === "sale" ? salePaymentSummary(entry).outstanding : "", width: 18 },
        { label: "Catatan", value: "notes", width: 40 },
        { label: "Dibuat", value: "created_at", width: 24 }
      ],
      rows: state.sales
    },
    {
      name: "Detail Item",
      columns: [
        { label: "ID detail", value: "id", width: 38 }, { label: "ID transaksi", value: "sale_id", width: 38 },
        { label: "ID produk", value: "product_id", width: 38 }, { label: "No. invoice", value: (item) => salesById.get(item.sale_id)?.invoice_no || "", width: 20 },
        { label: "SKU", value: (item) => item.products?.sku || productsById.get(item.product_id)?.sku || "", width: 16 },
        { label: "Nama produk", value: (item) => item.products?.name || productsById.get(item.product_id)?.name || "", width: 28 },
        { label: "Jumlah", value: (item) => Number(item.quantity || 0), width: 12 },
        { label: "Harga satuan (Rp)", value: (item) => Number(item.unit_price || 0), width: 20 },
        { label: "Subtotal (Rp)", value: (item) => Number(item.subtotal ?? Number(item.quantity || 0) * Number(item.unit_price || 0)), width: 18 }
      ],
      rows: state.saleItems
    }
  ];
  const blob = createExcelWorkbook(sheets);
  const link = document.createElement("a");
  const date = localIsoDate();
  const objectUrl = URL.createObjectURL(blob);
  link.href = objectUrl;
  link.download = `Little-Stitches-Data-${date}.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
  notify("File Excel seluruh data berhasil diunduh");
}

function normalizeWhatsAppNumber(phone) {
  let number = String(phone || "").replace(/\D/g, "");
  if (number.startsWith("0")) number = `62${number.slice(1)}`;
  else if (number.startsWith("8")) number = `62${number}`;
  return number.length >= 9 && number.length <= 15 ? number : "";
}

async function receiptImageBlob(sale, items, paymentSummary) {
  const width = 511;
  const rows = items.length ? items : [{ products: { name: "Detail produk tidak tersedia" }, quantity: 0, unit_price: 0 }];
  const rowGap = 40;
  const extraHeight = Math.max(0, rows.length - 1) * rowGap;
  const height = 581 + extraHeight;
  const canvas = document.createElement("canvas");
  canvas.width = width * 2;
  canvas.height = height * 2;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Browser tidak dapat membuat gambar struk.");
  context.scale(2, 2);

  const text = (value, x, y, options = {}) => {
    context.font = options.font || "11px 'Trebuchet MS', Arial, sans-serif";
    context.fillStyle = options.color || "#4a302c";
    context.textAlign = options.align || "left";
    context.textBaseline = "alphabetic";
    if ("letterSpacing" in context) context.letterSpacing = options.letterSpacing || "0px";
    context.fillText(String(value ?? ""), x, y, options.maxWidth || width);
  };
  const line = (y, dashed = false) => {
    context.beginPath();
    context.moveTo(37, y);
    context.lineTo(478, y);
    context.strokeStyle = "#eadbd0";
    context.lineWidth = 0.8;
    context.setLineDash(dashed ? [1, 3] : []);
    context.stroke();
    context.setLineDash([]);
  };
  const rightText = (value, y, options = {}) => text(value, 475, y, { ...options, align: "right" });
  const totalStart = 344 + extraHeight;
  const statusLine = 440 + extraHeight;
  const thanksLine = 480 + extraHeight;
  const customerName = sale.customers?.name || state.customers.find((entry) => entry.id === sale.customer_id)?.name || "Pelanggan umum";
  const paperReceiptNumber = $("#receipt-paper .receipt-meta strong")?.textContent || sale.invoice_no || "—";
  const paymentMethod = sale.payment_method === "qris" ? "QRIS" : "Transfer Bank";
  const subtotal = Number(sale.subtotal || rows.reduce((sum, item) => sum + Number(item.quantity || 0) * Number(item.unit_price || 0), 0));
  const totals = [
    ["Subtotal produk", rupiah(subtotal)],
    ["Ongkir flat", rupiah(sale.shipping_fee)],
    ["Total tagihan", rupiah(sale.total)],
    ["Jumlah pembayaran", rupiah(paymentSummary.paid)],
    ["Sisa tagihan", rupiah(paymentSummary.outstanding)]
  ];

  context.fillStyle = "#f8f3ec";
  context.fillRect(0, 0, width, height);
  context.fillStyle = "#fffdfa";
  context.fillRect(7.5, 10.5, 496, height - 33);
  context.strokeStyle = "#eadbd0";
  context.lineWidth = 1;
  context.strokeRect(7.5, 10.5, 496, height - 33);

  const logoGradient = context.createLinearGradient(139, 38, 173, 72);
  logoGradient.addColorStop(0, "#8f4d6c");
  logoGradient.addColorStop(1, "#e5a078");
  context.beginPath();
  context.roundRect(139, 38, 34, 34, 11);
  context.fillStyle = logoGradient;
  context.fill();
  text("✦", 156, 61, { font: "19px Georgia, serif", color: "#fff7ec", align: "center" });
  text("LITTLE STITCHES", 183, 54, { font: "bold 19px Georgia, serif", color: "#4a302c", letterSpacing: "1.5px" });
  text("Handmade Crochet Products", 280, 70, { font: "10px 'Trebuchet MS', Arial, sans-serif", color: "#a27c6d", align: "center" });
  text("Jl. Benang Ceria No. 12, Bandung · +62 812-3456-7890", width / 2, 92, { font: "10px 'Trebuchet MS', Arial, sans-serif", color: "#99847b", align: "center" });

  text("No. struk", 37, 120, { color: "#a18d83" });
  rightText("Tanggal", 120, { color: "#a18d83" });
  text(paperReceiptNumber, 37, 137, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });
  rightText(dateText(sale.sale_date || sale.created_at || new Date()), 137, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });
  context.beginPath();
  context.roundRect(37, 139, 441, 50, 9);
  context.fillStyle = "#fbf3eb";
  context.fill();
  text("Pelanggan", 48, 160, { color: "#997869" });
  text(customerName, 48, 176, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });
  text("Invoice", 413, 160, { color: "#997869" });
  rightText(sale.invoice_no || "—", 176, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });

  line(207, true);
  text("Metode pembayaran", 37, 233, { color: "#a18d83" });
  rightText("Referensi", 233, { color: "#a18d83" });
  text(paymentMethod, 37, 249, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });
  rightText(sale.payment_reference || "Tidak tersedia", 249, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });

  const headerY = 273;
  text("ITEM", 40, headerY, { font: "9px 'Trebuchet MS', Arial, sans-serif", color: "#a27c6d", letterSpacing: "1px" });
  text("QTY", 261, headerY, { font: "9px 'Trebuchet MS', Arial, sans-serif", color: "#a27c6d", letterSpacing: "1px", align: "right" });
  text("HARGA", 337, headerY, { font: "9px 'Trebuchet MS', Arial, sans-serif", color: "#a27c6d", letterSpacing: "1px", align: "right" });
  rightText("SUBTOTAL", headerY, { font: "9px 'Trebuchet MS', Arial, sans-serif", color: "#a27c6d", letterSpacing: "1px" });
  line(284);
  rows.forEach((item, index) => {
    const y = 302 + index * rowGap;
    const product = item.products || state.products.find((entry) => entry.id === item.product_id);
    text(product?.name || "Produk rajut", 40, y, { maxWidth: 195 });
    text(Number(item.quantity || 0), 261, y, { align: "right" });
    text(rupiah(item.unit_price), 337, y, { align: "right" });
    rightText(rupiah(Number(item.quantity || 0) * Number(item.unit_price || 0)), y);
    line(324 + index * rowGap);
  });

  totals.forEach(([label, amount], index) => {
    const y = totalStart + index * 20;
    text(label, 37, y, { color: "#957b70" });
    rightText(amount, y, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif" });
  });
  line(statusLine, true);
  text("Status pembayaran", 37, statusLine + 20, { font: "bold 11px 'Trebuchet MS', Arial, sans-serif", color: "#8b4e6a" });
  rightText(paymentSummary.status.toUpperCase(), statusLine + 20, { font: "bold 10px 'Trebuchet MS', Arial, sans-serif", color: "#4d806a", maxWidth: 255 });
  line(thanksLine, true);
  text("Terima kasih telah mendukung karya handmade kami ♡", width / 2, thanksLine + 30, { font: "italic 15px Georgia, serif", color: "#885c70", align: "center" });
  text("little stitches · made slowly, loved deeply", width / 2, thanksLine + 50, { font: "10px 'Trebuchet MS', Arial, sans-serif", color: "#aa9186", align: "center" });

  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("Gambar struk gagal dibuat.")), "image/png");
  });
}

function downloadReceiptImage(blob, sale) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const invoice = String(sale.invoice_no || "struk").replace(/[^a-zA-Z0-9_-]/g, "-");
  link.href = objectUrl;
  link.download = `Little-Stitches-${invoice}.png`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}

async function sendReceiptToWhatsApp(sale, items, paymentSummary) {
  const customer = sale.customers || state.customers.find((entry) => entry.id === sale.customer_id);
  const phone = normalizeWhatsAppNumber(customer?.phone);
  if (!phone) {
    notify(customer?.phone ? "Nomor telepon pelanggan tidak valid" : "Pelanggan ini belum memiliki nomor telepon");
    return;
  }
  const message = $("#receipt-message").value.trim();
  if (!message) {
    notify("Tulis pesan pengantar sebelum membagikan struk");
    $("#receipt-message").focus();
    return;
  }
  const shareButton = $("#send-whatsapp");
  shareButton.disabled = true;
  try {
    const imageBlob = await receiptImageBlob(sale, items, paymentSummary);
    const invoice = String(sale.invoice_no || "struk").replace(/[^a-zA-Z0-9_-]/g, "-");
    const imageFile = new File([imageBlob], `Little-Stitches-${invoice}.png`, { type: "image/png" });
    if (navigator.share && navigator.canShare?.({ files: [imageFile] })) {
      await navigator.share({ files: [imageFile], title: `Struk ${sale.invoice_no || "Little Stitches"}`, text: message });
      notify("Gambar struk berhasil dibagikan");
      return;
    }

    downloadReceiptImage(imageBlob, sale);
    const whatsappUrl = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
    const link = document.createElement("a");
    link.href = whatsappUrl;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.click();
    notify("Gambar struk diunduh. Lampirkan file PNG pada chat WhatsApp yang terbuka.");
  } catch (error) {
    if (error.name === "AbortError") return;
    notify("Gambar struk gagal dibagikan");
    console.error(error);
  } finally {
    shareButton.disabled = false;
  }
}

async function responseError(response) {
  const body = await response.text();
  let detail = body;
  let code = "";
  try {
    const parsed = JSON.parse(body);
    detail = parsed.message || parsed.details || parsed.hint || body;
    code = parsed.code || "";
  } catch {}
  const error = new Error(detail || `Permintaan gagal (${response.status})`);
  error.code = code;
  return error;
}

function explainSaveError(error) {
  if (error?.code === "PGRST202" || /could not find the function/i.test(error?.message || "")) {
    return "Fungsi database belum tersedia. Jalankan ulang database/schema.sql di Supabase, lalu coba lagi.";
  }
  const message = error?.message || "Periksa koneksi dan konfigurasi Supabase.";
  return message.length > 180 ? `${message.slice(0, 177)}...` : message;
}

async function supabase(table, options = {}) {
  if (!config.url || !config.publishableKey) {
    throw new Error("Supabase belum dikonfigurasi. Isi Project URL dan Publishable key di frontend/js/config.js.");
  }
  const response = await fetch(`${config.url.replace(/\/$/, "")}/rest/v1/${table}${options.query || ""}`, {
    method: options.method || "GET",
    headers: {
      apikey: config.publishableKey,
      Authorization: `Bearer ${config.publishableKey}`,
      "Content-Type": "application/json",
      Prefer: ["POST", "PATCH"].includes(options.method) ? "return=representation" : undefined
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });
  if (!response.ok) throw await responseError(response);
  return response.status === 204 ? null : response.json();
}

async function supabaseRpc(functionName, payload) {
  if (!config.url || !config.publishableKey) {
    throw new Error("Supabase belum dikonfigurasi. Isi Project URL dan Publishable key di frontend/js/config.js.");
  }
  const response = await fetch(`${config.url.replace(/\/$/, "")}/rest/v1/rpc/${functionName}`, {
    method: "POST",
    headers: {
      apikey: config.publishableKey,
      Authorization: `Bearer ${config.publishableKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) throw await responseError(response);
  return response.json();
}

async function uploadPaymentProof(file) {
  if (!file) return "";
  if (!config.url || !config.publishableKey) throw new Error("Supabase belum dikonfigurasi");
  const allowedTypes = ["image/jpeg", "image/png", "application/pdf"];
  if (!allowedTypes.includes(file.type)) throw new Error("Format bukti pembayaran harus JPG, JPEG, PNG, atau PDF");
  const safeName = file.name.toLowerCase().replace(/[^a-z0-9._-]/g, "-");
  const objectPath = `${Date.now()}-${safeName}`;
  const response = await fetch(`${config.url.replace(/\/$/, "")}/storage/v1/object/payment-proofs/${objectPath}`, {
    method: "POST",
    headers: { apikey: config.publishableKey, Authorization: `Bearer ${config.publishableKey}`, "Content-Type": file.type, "x-upsert": "true" },
    body: file
  });
  if (!response.ok) throw await responseError(response);
  return `${config.url.replace(/\/$/, "")}/storage/v1/object/public/payment-proofs/${objectPath}`;
}

function bindMoneyInputs(form) {
  form.querySelectorAll(".money-input").forEach((input) => {
    input.addEventListener("focus", () => { input.value = parseMoney(input.value) ? String(parseMoney(input.value)) : ""; });
    input.addEventListener("input", () => { input.value = input.value ? moneyInput(parseMoney(input.value)) : ""; input.dispatchEvent(new Event("moneychange", { bubbles: true })); });
    input.addEventListener("blur", () => { if (input.value) input.value = moneyInput(parseMoney(input.value)); });
    if (input.value) input.value = moneyInput(parseMoney(input.value));
  });
}

function bindProofUpload(form) {
  const fileInput = form.querySelector("[name=payment_proof_file]");
  const fileName = form.querySelector("[data-proof-name]");
  const removeButton = form.querySelector("[data-remove-proof]");
  if (!fileInput) return;
  const storedUrl = form.querySelector("[name=payment_proof_url]");
  fileInput.required = !storedUrl?.value;
  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    fileName.textContent = file ? file.name : "Belum ada file dipilih";
    removeButton.hidden = !file;
  });
  removeButton.onclick = () => {
    fileInput.value = "";
    if (storedUrl) storedUrl.value = "";
    fileInput.required = true;
    fileName.textContent = "Belum ada file dipilih";
    removeButton.hidden = true;
  };
}

function bindSaleCalculations(form) {
  const productSelect = form.querySelector("[name=cart_product_id]");
  const quantityInput = form.querySelector("[name=cart_quantity]");
  const deliverySelect = form.querySelector("[name=delivery_method]");
  const totalInput = form.querySelector("[name=total]");
  const subtotalInput = form.querySelector("[name=subtotal]");
  const paymentPlanSelect = form.querySelector("[name=payment_plan]");
  const paidAmountInput = form.querySelector("[name=paid_amount]");
  const cartBody = form.querySelector("[data-cart-body]");
  const cart = [];
  try {
    const savedCart = JSON.parse(form.dataset.existingCart || "[]");
    cart.push(...savedCart);
  } catch (error) {
    console.warn("Keranjang transaksi tersimpan tidak dapat dibaca", error);
  }
  const existingProduct = form.dataset.existingProduct;
  if (!cart.length && existingProduct) cart.push({ productId: existingProduct, quantity: Number(form.dataset.existingQuantity || 1) });
  const calculate = () => {
    const product = state.products.find((item) => item.id === productSelect.value);
    if (!product || !quantityInput.value) return;
    const found = cart.find((item) => item.productId === product.id);
    if (found) found.quantity += Number(quantityInput.value);
    else cart.push({ productId: product.id, quantity: Number(quantityInput.value) });
    productSelect.value = "";
    quantityInput.value = 1;
    renderCart();
  };
  const renderCart = () => {
    const cartSection = form.querySelector(".cart-section");
    const addButton = form.querySelector("[data-add-cart]");
    cartSection.hidden = !deliverySelect.value;
    addButton.disabled = !deliverySelect.value;
    let subtotal = 0;
    cartBody.innerHTML = cart.map((item, index) => {
      const product = state.products.find((candidate) => candidate.id === item.productId);
      const lineTotal = Number(product?.price || 0) * item.quantity;
      subtotal += lineTotal;
      return `<tr><td>${escapeHtml(product?.name || "Produk")}</td><td>${item.quantity}</td><td>${rupiah(product?.price)}</td><td>${rupiah(lineTotal)}</td><td><button type="button" class="icon-btn delete" data-remove-cart="${index}">⌫</button></td></tr>`;
    }).join("") || `<tr><td colspan="5" class="empty">Keranjang masih kosong.</td></tr>`;
    const shipping = deliverySelect.value === "delivery" && subtotal < 50000 ? 15000 : 0;
    subtotalInput.value = moneyInput(subtotal);
    form.querySelector("[name=shipping_fee]").value = moneyInput(shipping);
    totalInput.value = moneyInput(subtotal + shipping);
    if (paymentPlanSelect && paidAmountInput) {
      paidAmountInput.value = moneyInput(paymentPlanSelect.value === "full" ? subtotal + shipping : Math.round((subtotal + shipping) / 2));
    }
    form.querySelector("[data-shipping-note]").textContent = deliverySelect.value === "pickup"
      ? "Ambil di tempat · Gratis ongkir"
      : subtotal >= 50000
        ? "Delivery · Gratis ongkir untuk minimal pembelian Rp50.000"
        : "Delivery · Ongkir Rp15.000 · Gratis ongkir mulai Rp50.000";
    cartBody.querySelectorAll("[data-remove-cart]").forEach((button) => button.onclick = () => { cart.splice(Number(button.dataset.removeCart), 1); renderCart(); });
  };
  form.querySelector("[data-add-cart]").onclick = calculate;
  deliverySelect.addEventListener("change", renderCart);
  if (paymentPlanSelect) paymentPlanSelect.addEventListener("change", renderCart);
  form.querySelector("[name=cart_product_id]").onkeydown = (event) => { if (event.key === "Enter") { event.preventDefault(); calculate(); } };
  form.querySelector("[name=cart_quantity]").onkeydown = (event) => { if (event.key === "Enter") { event.preventDefault(); calculate(); } };
  form.getCart = () => cart;
  renderCart();
  form.addEventListener("moneychange", () => {
    if (subtotalInput && totalInput) {
      subtotalInput.value = moneyInput(parseMoney(subtotalInput.value));
      totalInput.value = moneyInput(parseMoney(subtotalInput.value) + parseMoney(form.querySelector("[name=shipping_fee]").value));
    }
  });
}

async function loadData() {
  if (!config.url || !config.publishableKey) {
    $("#connection-text").textContent = "Isi frontend/js/config.js";
    return;
  }
  $("#connection-text").textContent = "Menghubungkan...";
  try {
    const [products, customers, sales, saleItems] = await Promise.all([
      supabase("products", { query: "?select=*&order=created_at.desc" }),
      supabase("customers", { query: "?select=*&order=created_at.desc" }),
      supabase("sales", { query: "?select=*,customers(id,name,phone,address)&order=sale_date.desc,created_at.desc" }),
      supabase("sale_items", { query: "?select=*,products(name,sku)&order=id.asc" })
    ]);
    if (products) state.products = products;
    if (customers) state.customers = customers;
    if (sales) state.sales = sales;
    if (saleItems) state.saleItems = saleItems;
    if (config.url && config.publishableKey) {
      $("#connection-text").textContent = "Terhubung";
      $(".pulse").style.background = "#8bc2a5";
    }
  } catch (error) {
    $("#connection-text").textContent = "Koneksi gagal";
    notify("Gagal memuat data Supabase. Periksa konfigurasi dan kebijakan akses.");
    console.error(error);
  }
}

function notify(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  setTimeout(() => toast.classList.remove("show"), 2800);
}

function statCard(label, value, foot, icon) {
  return `<div class="stat-card"><span class="stat-icon">${icon}</span><div class="stat-label">${label}</div><div class="stat-value">${value}</div><div class="stat-foot">${foot}</div></div>`;
}

function dataBanner(type, kicker, title, description) {
  return `<section class="data-banner ${type}-banner"><img src="assets/banner-${type}-photo.jpg" alt="" aria-hidden="true"><div class="data-banner-copy"><span>${kicker}</span><h2>${title}</h2><p>${description}</p></div><div class="banner-stitch">✦ <i>✦</i> ✦</div></section>`;
}

function productRows(products) {
  return products.length ? products.slice(0, 5).map((p) => `<tr><td><div class="product-name">${escapeHtml(p.name)}</div><div class="product-sub">${escapeHtml(p.sku)} · ${escapeHtml(p.category)}</div></td><td>${rupiah(p.price)}</td><td><span class="pill ${p.stock <= p.minimum_stock ? "red" : "green"}">${p.stock} unit</span></td></tr>`).join("") : `<tr><td colspan="3" class="empty">Belum ada data produk.</td></tr>`;
}

function renderDashboard() {
  const sales = state.sales.filter((item) => (item.entry_type || "sale") === "sale");
  const revenue = sales.reduce((sum, item) => sum + Number(item.total || 0), 0);
  const low = state.products.filter((p) => Number(p.stock) <= Number(p.minimum_stock)).length;
  const trend = getSalesTrend();
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Overview · ${dateText(new Date())}</div><h1>Selamat datang di Little Stitches.</h1><p>Pantau denyut Little Stitches dengan lebih hangat dan menyenangkan.</p></div></div>
    <section class="welcome-hero"><div class="hero-copy"><span class="hero-kicker">✦ THE COZY CROCHET DESK</span><h2>Dirajut dengan hati.<br><i>Dijual dengan cerita.</i></h2><p>Kelola koleksi handmade, pembayaran, dan persediaan dari satu ruang kerja yang ceria.</p><button class="hero-btn" data-page="products">Jelajahi koleksi <b>↗</b></button></div><div class="hero-sticker"><span>✦</span><b>Handmade<br>with love</b></div></section>
    <div class="stats-grid">${statCard("Total penjualan", rupiah(revenue), `${sales.length} transaksi tercatat`, "↗")}${statCard("Produk aktif", state.products.length, "Koleksi dalam katalog", "✦")}${statCard("Pelanggan", state.customers.length, "Relasi yang terjaga", "♧")}${statCard("Stok menipis", low, low ? "Perlu segera dipantau" : "Semua stok aman", "◌")}</div>
    <section class="panel trend-panel"><div class="panel-head"><div><h3>Tren penjualan</h3><span class="panel-subtitle">Performa pemasukan 6 bulan terakhir</span></div><span class="trend-total">${rupiah(trend.total)}</span></div><div class="trend-chart">${trend.points.map((point) => `<div class="trend-column"><span class="trend-value">${point.value ? rupiah(point.value).replace(",00", "").replace("Rp", "Rp ") : "—"}</span><div class="trend-bar-wrap"><i style="height:${point.height}%"></i></div><small>${point.label}</small></div>`).join("")}</div></section>
    <div class="grid-2"><section class="panel"><div class="panel-head"><h3>Produk paling siap dijual</h3><a class="text-link" href="#products">Lihat semua →</a></div><div class="table-wrap"><table><thead><tr><th>Produk</th><th>Harga</th><th>Stok</th></tr></thead><tbody>${productRows(state.products)}</tbody></table></div></section>
    <section class="panel"><div class="panel-head"><h3>Perhatian persediaan</h3><a class="text-link" href="#inventory">Kelola →</a></div>${state.products.filter((p) => p.stock <= p.minimum_stock).slice(0, 4).map((p) => `<div class="bar-row"><div class="bar-label"><span>${escapeHtml(p.name)}</span><span>${p.stock}/${p.minimum_stock} min</span></div><div class="bar"><i style="width:${Math.min(100, (p.stock / Math.max(p.minimum_stock, 1)) * 100)}%"></i></div></div>`).join("") || `<div class="empty">Stok semua produk dalam kondisi baik ✦</div>`}</section></div><footer class="dashboard-footer"><span>Fakultas Ekonomika dan Bisnis</span><span>Universitas Diponegoro</span></footer>`;
}

function getSalesTrend() {
  const points = [];
  const now = new Date();
  for (let offset = 5; offset >= 0; offset -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - offset, 1);
    const month = date.getMonth();
    const year = date.getFullYear();
    const value = state.sales.filter((sale) => (sale.entry_type || "sale") === "sale" && (() => {
      const saleDate = new Date(sale.sale_date || sale.created_at);
      return saleDate.getMonth() === month && saleDate.getFullYear() === year;
    })()).reduce((sum, sale) => sum + Number(sale.total || 0), 0);
    points.push({ label: new Intl.DateTimeFormat("id-ID", { month: "short" }).format(date), value });
  }
  const max = Math.max(...points.map((point) => point.value), 1);
  return { points: points.map((point) => ({ ...point, height: point.value ? Math.max(12, Math.round((point.value / max) * 100)) : 5 })), total: points.reduce((sum, point) => sum + point.value, 0) };
}

function renderProducts() {
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Koleksi rajutan</div><h1>Produk rajut.</h1><p>Rawat katalog karya tangan Little Stitches.</p></div><button class="btn" id="add-product">＋ Tambah produk</button></div>${dataBanner("products", "HANDMADE COLLECTION", "Setiap karya punya cerita.", "Katalog rajutan pilihan yang dibuat dengan penuh perhatian.")}<section class="panel"><div class="table-wrap"><table class="products-table"><thead><tr><th>Produk</th><th>Kategori</th><th>Harga</th><th>Stok</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${state.products.map((p) => `<tr><td><div class="product-name" title="${escapeHtml(p.name)}">${escapeHtml(p.name)}</div><div class="product-sub" title="${escapeHtml(p.sku)}">${escapeHtml(p.sku)}</div></td><td title="${escapeHtml(p.category)}">${escapeHtml(p.category)}</td><td>${rupiah(p.price)}</td><td>${p.stock} unit</td><td><span class="pill ${p.stock <= p.minimum_stock ? "red" : "green"}">${p.stock <= p.minimum_stock ? "Menipis" : "Aman"}</span></td><td><div class="row-actions"><button class="icon-btn edit" data-edit-product="${p.id}" title="Edit produk">✎</button><button class="icon-btn delete" data-delete-product="${p.id}" title="Hapus produk">⌫</button></div></td></tr>`).join("") || `<tr><td colspan="6" class="empty">Belum ada produk.</td></tr>`}</tbody></table></div></section>`;
  $("#add-product").onclick = showProductForm;
  document.querySelectorAll("[data-edit-product]").forEach((button) => button.onclick = () => showProductForm(state.products.find((product) => product.id === button.dataset.editProduct)));
  document.querySelectorAll("[data-delete-product]").forEach((button) => button.onclick = () => deleteProduct(button.dataset.deleteProduct));
}

function renderCustomers() {
  const customers = [...state.customers].sort((a, b) => {
    if (state.customerSort === "name_asc" || state.customerSort === "name_desc") {
      const direction = state.customerSort === "name_asc" ? 1 : -1;
      return direction * String(a.name || "").localeCompare(String(b.name || ""), "id", { sensitivity: "base" });
    }
    const direction = state.customerSort === "oldest" ? 1 : -1;
    return direction * (new Date(a.created_at || 0) - new Date(b.created_at || 0));
  });
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Customer circle</div><h1>Pelanggan.</h1><p>Bangun hubungan hangat dengan setiap pelanggan.</p></div><button class="btn" id="add-customer">＋ Tambah pelanggan</button></div>${dataBanner("customers", "OUR LOVELY CUSTOMERS", "Cerita indah dimulai bersama.", "Kenali pelanggan yang tumbuh bersama Little Stitches.")}<section class="panel"><div class="panel-head sort-panel-head"><h3>Daftar pelanggan</h3><label class="sort-control"><span>Urutkan</span><select id="customer-sort" aria-label="Urutkan pelanggan"><option value="newest" ${state.customerSort === "newest" ? "selected" : ""}>Tanggal terbaru</option><option value="oldest" ${state.customerSort === "oldest" ? "selected" : ""}>Tanggal terlama</option><option value="name_asc" ${state.customerSort === "name_asc" ? "selected" : ""}>Nama A–Z</option><option value="name_desc" ${state.customerSort === "name_desc" ? "selected" : ""}>Nama Z–A</option></select></label></div><div class="table-wrap"><table class="customers-table"><thead><tr><th>Nama</th><th>Kontak</th><th>Email</th><th>Alamat</th><th>Aksi</th></tr></thead><tbody>${customers.map((c) => `<tr><td class="product-name" title="${escapeHtml(c.name)}">${escapeHtml(c.name)}</td><td title="${escapeHtml(c.phone || "")}">${escapeHtml(c.phone || "—")}</td><td title="${escapeHtml(c.email || "")}">${escapeHtml(c.email || "—")}</td><td title="${escapeHtml(c.address || "")}">${escapeHtml(c.address || "—")}</td><td><div class="row-actions"><button class="icon-btn edit" data-edit-customer="${c.id}" title="Edit pelanggan">✎</button><button class="icon-btn delete" data-delete-customer="${c.id}" title="Hapus pelanggan">⌫</button></div></td></tr>`).join("") || `<tr><td colspan="5" class="empty">Belum ada pelanggan.</td></tr>`}</tbody></table></div></section>`;
  $("#customer-sort").onchange = (event) => {
    state.customerSort = event.currentTarget.value;
    renderCustomers();
  };
  $("#add-customer").onclick = showCustomerForm;
  document.querySelectorAll("[data-edit-customer]").forEach((button) => button.onclick = () => showCustomerForm(state.customers.find((customer) => customer.id === button.dataset.editCustomer)));
  document.querySelectorAll("[data-delete-customer]").forEach((button) => button.onclick = () => deleteCustomer(button.dataset.deleteCustomer));
}

function paymentStatusLabel(status, paidAmount, totalAmount) {
  const paid = Number(paidAmount || 0);
  const total = Number(totalAmount || 0);
  if (status === "rejected") return paid < total ? "Ditolak · Belum lunas" : "Ditolak";
  if (status === "verified" && paid >= total) return "Lunas";
  if (status === "verified" && paid > 0) return "DP terverifikasi · Belum lunas";
  if (status === "verified") return "Terverifikasi · Belum lunas";
  return "Menunggu verifikasi · Belum lunas";
}

function salePaymentSummary(sale) {
  const settlements = state.sales.filter((entry) =>
    entry.entry_type === "settlement" &&
    entry.related_sale_id === sale.id &&
    entry.payment_status !== "rejected"
  );
  const settledAmount = settlements.reduce((sum, entry) => sum + Number(entry.paid_amount || 0), 0);
  const initialPaid = sale.payment_status === "rejected" ? 0 : Number(sale.paid_amount || 0);
  const paid = initialPaid + settledAmount;
  const verified = (sale.payment_status === "verified" ? initialPaid : 0) +
    settlements.reduce((sum, entry) => sum + (entry.payment_status === "verified" ? Number(entry.paid_amount || 0) : 0), 0);
  const total = Number(sale.total || 0);
  let status;
  if (paid >= total && total > 0 && verified >= total) status = "Lunas";
  else if (settlements.some((entry) => entry.payment_status === "pending")) status = "Pelunasan menunggu verifikasi · Belum lunas";
  else if (sale.payment_status === "rejected" && verified === 0) status = "Ditolak · Belum lunas";
  else if (verified > 0) status = "DP terverifikasi · Belum lunas";
  else status = "Menunggu verifikasi · Belum lunas";
  return { paid, outstanding: Math.max(0, total - paid), status };
}

function renderSales() {
  const sales = state.sales.filter((sale) => (sale.entry_type || "sale") === "sale");
  const filteredSales = state.salesCustomerFilter
    ? sales.filter((sale) => sale.customer_id === state.salesCustomerFilter)
    : sales;
  const sortedSales = [...filteredSales].sort((a, b) => {
    if (state.salesSort === "customer_asc" || state.salesSort === "customer_desc") {
      const direction = state.salesSort === "customer_asc" ? 1 : -1;
      return direction * String(a.customers?.name || "").localeCompare(String(b.customers?.name || ""), "id", { sensitivity: "base" });
    }
    const direction = state.salesSort === "oldest" ? 1 : -1;
    return direction * (new Date(a.sale_date || a.created_at || 0) - new Date(b.sale_date || b.created_at || 0));
  });
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Cash desk</div><h1>Pembayaran.</h1><p>Catat, verifikasi, dan cetak bukti pembayaran pelanggan.</p></div><button class="btn" id="add-sale">＋ Catat pembayaran</button></div>${dataBanner("payments", "A LITTLE THANK-YOU", "Setiap pesanan, penuh makna.", "Pembayaran yang rapi untuk karya yang dirajut dengan cinta.")}<section class="panel"><div class="panel-head sort-panel-head"><h3>Riwayat pembayaran</h3><div class="table-controls"><span class="soft-count">${filteredSales.length} dari ${sales.length} transaksi</span><label class="sort-control"><span>Pelanggan</span><select id="sales-customer-filter" aria-label="Filter pembayaran berdasarkan pelanggan"><option value="">Semua pelanggan</option>${state.customers.slice().sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "id", { sensitivity: "base" })).map((customer) => `<option value="${customer.id}" ${state.salesCustomerFilter === customer.id ? "selected" : ""}>${escapeHtml(customer.name)}</option>`).join("")}</select></label><label class="sort-control"><span>Urutkan</span><select id="sales-sort" aria-label="Urutkan pembayaran"><option value="newest" ${state.salesSort === "newest" ? "selected" : ""}>Tanggal terbaru</option><option value="oldest" ${state.salesSort === "oldest" ? "selected" : ""}>Tanggal terlama</option><option value="customer_asc" ${state.salesSort === "customer_asc" ? "selected" : ""}>Pelanggan A–Z</option><option value="customer_desc" ${state.salesSort === "customer_desc" ? "selected" : ""}>Pelanggan Z–A</option></select></label></div></div><div class="table-wrap"><table class="sales-table"><thead><tr><th>Invoice</th><th>Pelanggan</th><th>Kanal pembelian</th><th>Tanggal</th><th>Metode</th><th>Total</th><th>Dibayar</th><th>Status</th><th>Aksi</th></tr></thead><tbody>${sortedSales.map((s) => {const total = Number(s.total || 0);const summary = salePaymentSummary(s);const hasSettlementHistory = state.sales.some((entry) => entry.entry_type === "settlement" && entry.related_sale_id === s.id);return `<tr><td class="product-name" title="${escapeHtml(s.invoice_no)}">${escapeHtml(s.invoice_no)}</td><td title="${escapeHtml(s.customers?.name || "")}">${escapeHtml(s.customers?.name || "—")}</td><td><span class="pill ${s.sales_channel === "in_store" ? "green" : "orange"}" title="${s.sales_channel === "in_store" ? "Beli langsung di toko" : "WhatsApp / Online"}">${s.sales_channel === "in_store" ? "Beli langsung di toko" : "WhatsApp / Online"}</span></td><td>${dateText(s.sale_date)}</td><td><span class="method-dot ${s.payment_method}"></span>${s.payment_method === "qris" ? "QRIS" : "Transfer"}</td><td>${rupiah(total)}</td><td>${rupiah(summary.paid)}${summary.outstanding ? `<div class="product-sub">Sisa ${rupiah(summary.outstanding)}</div>` : ""}</td><td><span class="pill ${summary.status === "Lunas" ? "green" : summary.status.startsWith("Ditolak") ? "red" : "orange"}" title="${escapeHtml(summary.status)}">${escapeHtml(summary.status)}</span></td><td><div class="row-actions">${summary.outstanding > 0 ? `<button class="icon-btn settle" data-settle-sale="${s.id}" title="Catat pelunasan ${rupiah(summary.outstanding)}">₊</button>` : ""}${hasSettlementHistory ? "" : `<button class="icon-btn edit" data-edit-sale="${s.id}" title="Edit pembayaran">✎</button>`}<button class="icon-btn delete" data-delete-sale="${s.id}" title="Hapus pembayaran">⌫</button><button class="icon-btn print" data-print-sale="${s.id}" title="Cetak struk">⎙</button></div></td></tr>`;}).join("") || `<tr><td colspan="9" class="empty">${state.salesCustomerFilter ? "Belum ada pembayaran untuk pelanggan ini." : "Belum ada pembayaran."}</td></tr>`}</tbody></table></div></section>`;
  $("#sales-customer-filter").onchange = (event) => {
    state.salesCustomerFilter = event.currentTarget.value;
    renderSales();
  };
  $("#sales-sort").onchange = (event) => {
    state.salesSort = event.currentTarget.value;
    renderSales();
  };
  $("#add-sale").onclick = () => showSaleForm();
  document.querySelectorAll("[data-print-sale]").forEach((button) => button.onclick = () => openReceipt(button.dataset.printSale));
  document.querySelectorAll("[data-edit-sale]").forEach((button) => button.onclick = () => showSaleForm(state.sales.find((sale) => sale.id === button.dataset.editSale)));
  document.querySelectorAll("[data-delete-sale]").forEach((button) => button.onclick = () => deleteSale(button.dataset.deleteSale));
  document.querySelectorAll("[data-settle-sale]").forEach((button) => button.onclick = () => showSettlementForm(state.sales.find((sale) => sale.id === button.dataset.settleSale)));
}

function renderInventory() {
  const total = state.products.reduce((sum, p) => sum + Number(p.stock || 0), 0);
  const purchases = state.sales.filter((sale) => sale.entry_type === "purchase");
  const sales = state.sales.filter((sale) => (sale.entry_type || "sale") === "sale");
  const settlements = state.sales.filter((sale) => sale.entry_type === "settlement");
  const adjustments = state.sales.filter((sale) => sale.entry_type === "stock_adjustment");
  const journalRows = [
    ...sales.map((sale) => {
      const items = state.saleItems.filter((item) => item.sale_id === sale.id);
      const description = items.map((item) => `${item.products?.name || "Produk rajut"} × ${item.quantity}`).join(", ") || "Detail produk belum tersedia";
      const total = Number(sale.total || 0);
      const cash = sale.payment_status === "verified" ? Math.min(Number(sale.paid_amount || 0), total) : 0;
      const receivable = Math.max(0, total - cash);
      return {
        date: sale.sale_date,
        reference: sale.invoice_no,
        description,
        party: sale.customers?.name || "Pelanggan",
        debitLines: cash > 0 || receivable > 0 ? [
          ...(cash > 0 ? [{ account: "Kas", amount: cash }] : []),
          ...(receivable > 0 ? [{ account: "Piutang Usaha", amount: receivable }] : [])
        ] : [{ account: "Piutang Usaha", amount: 0 }],
        creditLines: [{ account: "Penjualan", amount: total }],
        status: sale.payment_status,
        paidAmount: Number(sale.paid_amount || 0),
        total,
        paymentProof: sale.payment_proof_url,
        type: "sale",
        id: sale.id
      };
    }),
    ...purchases.map((entry) => {
      const total = Number(entry.total || 0);
      const cash = Math.min(Number(entry.paid_amount || 0), total);
      const payable = Math.max(0, total - cash);
      return {
        date: entry.sale_date,
        reference: entry.invoice_no,
        description: entry.journal_description || "Pembelian persediaan",
        party: entry.customers?.name || "Pemasok",
        debitLines: [{ account: "Persediaan", amount: total }],
        creditLines: cash > 0 || payable > 0 ? [
          ...(cash > 0 ? [{ account: "Kas", amount: cash }] : []),
          ...(payable > 0 ? [{ account: "Utang Usaha", amount: payable }] : [])
        ] : [{ account: "Kas", amount: 0 }],
        status: Number(entry.paid_amount || 0) >= total ? "verified" : "pending",
        type: "purchase",
        id: entry.id
      };
    }),
    ...settlements.map((entry) => ({
      date: entry.sale_date,
      reference: entry.invoice_no,
      description: entry.journal_description || "Pelunasan piutang",
      party: entry.customers?.name || "Pelanggan",
      debitLines: [{ account: "Kas", amount: Number(entry.paid_amount || 0) }],
      creditLines: [{ account: "Piutang Usaha", amount: Number(entry.paid_amount || 0) }],
      status: entry.payment_status,
      paymentProof: entry.payment_proof_url,
      type: "settlement",
      id: entry.id
    }))
  ].sort((a, b) => new Date(b.date) - new Date(a.date));
  const journalMarkup = journalRows.map((entry, entryIndex) => {
    const lines = [
      ...entry.debitLines.map((line) => ({ ...line, side: "debit" })),
      ...entry.creditLines.map((line) => ({ ...line, side: "credit" }))
    ];
    const lineCount = lines.length;
    const statusClass = entry.status === "rejected" ? "red" : entry.status === "verified" && (entry.type !== "sale" || entry.paidAmount >= entry.total) ? "green" : "orange";
    const statusLabel = entry.type === "sale"
      ? paymentStatusLabel(entry.status, entry.paidAmount, entry.total)
      : entry.type === "settlement"
        ? entry.status === "verified" ? "Pelunasan terverifikasi" : entry.status === "rejected" ? "Pelunasan ditolak" : "Menunggu verifikasi pelunasan"
        : entry.status === "verified" ? "Lunas" : "DP / Hutang";
    const actions = entry.type === "sale" && entry.paymentProof
      ? `<div class="row-actions journal-actions"><button class="icon-btn proof-view" data-proof-sale="${entry.id}" title="Lihat bukti pembayaran">▧</button>${entry.status === "pending" ? `<button class="icon-btn edit" data-verify-sale="${entry.id}" title="Verifikasi pembayaran">✓</button><button class="icon-btn delete" data-reject-sale="${entry.id}" title="Tolak pembayaran">×</button>` : ""}</div>`
      : entry.type === "settlement"
        ? `<div class="row-actions journal-actions">${entry.paymentProof ? `<button class="icon-btn proof-view" data-proof-sale="${entry.id}" title="Lihat bukti pelunasan">▧</button>` : ""}${entry.status === "pending" ? `<button class="icon-btn edit" data-verify-sale="${entry.id}" title="Verifikasi pelunasan">✓</button><button class="icon-btn delete" data-reject-sale="${entry.id}" title="Tolak pelunasan">×</button>` : ""}</div>`
      : entry.type === "purchase"
        ? `<div class="row-actions journal-actions"><button class="icon-btn edit" data-edit-purchase="${entry.id}" title="Edit jurnal">✎</button><button class="icon-btn delete" data-delete-purchase="${entry.id}" title="Hapus jurnal">⌫</button></div>`
        : "";
    return Array.from({ length: lineCount }, (_, index) => {
      const line = lines[index];
      return `<tr class="${entryIndex % 2 === 0 ? "journal-transaction-even" : "journal-transaction-odd"} ${index === 0 ? "journal-group-start" : "journal-continuation"} ${line.side === "credit" ? "journal-credit-line" : "journal-debit-line"}">${index === 0 ? `<td rowspan="${lineCount}">${dateText(entry.date)}</td><td rowspan="${lineCount}">${escapeHtml(entry.reference)}</td>` : ""}<td>${escapeHtml(line.account)}</td><td class="journal-amount debit-amount">${line.side === "debit" ? rupiah(line.amount) : ""}</td><td class="journal-amount credit-amount">${line.side === "credit" ? rupiah(line.amount) : ""}</td>${index === 0 ? `<td rowspan="${lineCount}"><div class="product-name">${escapeHtml(entry.description)}</div><div class="product-sub">${escapeHtml(entry.party)}</div></td><td rowspan="${lineCount}"><span class="pill ${statusClass}">${statusLabel}</span>${actions}</td>` : ""}</tr>`;
    }).join("");
  }).join("");
  const stockAdjustmentMarkup = adjustments.map((entry, index) => {
    const product = state.products.find((item) => item.id === entry.stock_product_id);
    const movementLabel = { in: "Stok masuk", damaged: "Rusak", lost: "Hilang" }[entry.stock_movement] || "Penyesuaian";
    return `<tr class="${index % 2 === 0 ? "journal-transaction-even" : "journal-transaction-odd"}"><td>${dateText(entry.sale_date)}</td><td>${escapeHtml(entry.invoice_no)}</td><td><div class="product-name">${escapeHtml(product?.name || "Produk tidak ditemukan")}</div><div class="product-sub">${escapeHtml(product?.sku || "")}</div></td><td><span class="pill ${entry.stock_movement === "in" ? "green" : "orange"}">${movementLabel}</span></td><td>${entry.stock_quantity} unit</td><td>${entry.stock_before} → ${entry.stock_after} unit</td><td>${escapeHtml(entry.notes || "—")}</td></tr>`;
  }).join("");
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Catatan stok & keuangan</div><h1>Jurnal persediaan.</h1><p>Transaksi pelanggan dan pembelian dicatat dengan akun debit-kredit yang seimbang.</p></div><div class="page-actions"><button class="btn light" id="add-stock-adjustment">＋ Penyesuaian stok</button><button class="btn" id="add-purchase">＋ Catat pembelian</button></div></div><div class="stats-grid">${statCard("Total unit", total, "Di seluruh produk", "▦")}${statCard("Nilai persediaan", rupiah(state.products.reduce((s,p)=>s + p.stock*p.price, 0)), "Harga katalog", "◈")}${statCard("Pembelian DP", purchases.filter(s => Number(s.paid_amount || 0) < Number(s.total || 0)).length, "Masih memiliki sisa", "◌")}${statCard("Pesanan pelanggan", sales.length, "Tercatat otomatis dari penjualan", "↘")}</div>${dataBanner("inventory", "THREADS IN MOTION", "Setiap karya, tercatat.", "Riwayat jurnal Little Stitches berdasarkan transaksi yang tersimpan.")}<section class="panel"><div class="panel-head"><div><h3>Riwayat penyesuaian stok</h3><span class="panel-subtitle">Stok masuk, produk rusak, dan produk hilang tercatat terpisah dari jurnal nilai keuangan.</span></div></div><div class="table-wrap"><table class="stock-adjustment-table"><thead><tr><th>Tanggal</th><th>No. Bukti</th><th>Produk</th><th>Jenis</th><th>Jumlah</th><th>Stok sebelum → sesudah</th><th>Catatan</th></tr></thead><tbody>${stockAdjustmentMarkup || `<tr><td colspan="7" class="empty">Belum ada penyesuaian stok.</td></tr>`}</tbody></table></div></section><section class="panel"><div class="panel-head"><div><h3>Jurnal umum</h3><span class="panel-subtitle">Setiap transaksi ditampilkan per baris: akun debit di atas dan akun kredit di bawah.</span></div></div><div class="table-wrap"><table class="journal-table"><thead><tr><th>Tanggal</th><th>No. Bukti</th><th>Akun</th><th>Debit</th><th>Kredit</th><th>Keterangan</th><th>Status / Aksi</th></tr></thead><tbody>${journalMarkup || `<tr><td colspan="7" class="empty">Belum ada transaksi jurnal.</td></tr>`}</tbody></table></div></section>`;
  $("#add-stock-adjustment").onclick = () => showStockAdjustmentForm();
  $("#add-purchase").onclick = () => showPurchaseForm();
  document.querySelectorAll("[data-edit-purchase]").forEach((button) => button.onclick = () => showPurchaseForm(state.sales.find((entry) => entry.id === button.dataset.editPurchase)));
  document.querySelectorAll("[data-delete-purchase]").forEach((button) => button.onclick = () => deleteSale(button.dataset.deletePurchase, "Jurnal pembelian"));
  document.querySelectorAll("[data-proof-sale]").forEach((button) => button.onclick = () => openPaymentProof(button.dataset.proofSale));
  document.querySelectorAll("[data-verify-sale]").forEach((button) => button.onclick = () => updatePaymentStatus(button.dataset.verifySale, "verified"));
  document.querySelectorAll("[data-reject-sale]").forEach((button) => button.onclick = () => updatePaymentStatus(button.dataset.rejectSale, "rejected"));
}

function showProductForm(product = null) {
  const editing = Boolean(product?.id);
  const categories = ["Aksesoris", "Tas", "Topi", "Bookmark", "Holder", "Busana", "Hiasan Rumah", "Lainnya"];
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">${editing ? "Edit produk" : "Produk baru"}</div><h1>${editing ? "Perbarui karya." : "Tambah karya."}</h1></div></div><form class="form-card" id="product-form"><div class="form-grid"><div class="field"><label>SKU</label><input name="sku" value="${escapeHtml(product?.sku || "")}" placeholder="LS-004" required></div><div class="field"><label>Nama produk</label><input name="name" value="${escapeHtml(product?.name || "")}" placeholder="Nama karya rajut" required></div><div class="field"><label>Kategori</label><select name="category" required>${categories.map((category) => `<option ${category === product?.category ? "selected" : ""}>${category}</option>`).join("")}</select></div><div class="field"><label>Harga (Rp)</label><input name="price" class="money-input" inputmode="numeric" value="${moneyInput(product?.price || "")}" placeholder="Rp0" required></div><div class="field"><label>Stok awal</label><input name="stock" type="number" min="0" value="${product?.stock ?? ""}" required></div><div class="field"><label>Batas minimum</label><input name="minimum_stock" type="number" min="0" value="${product?.minimum_stock ?? 3}" required></div><div class="field full"><label>Deskripsi</label><textarea name="description" rows="3">${escapeHtml(product?.description || "")}</textarea></div></div><div class="form-actions"><button type="button" class="btn light" data-page="products">Batal</button><button class="btn">${editing ? "Simpan perubahan" : "Simpan produk"}</button></div></form>`;
  bindMoneyInputs($("#product-form"));
  $("#product-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    data.price = parseMoney(data.price);
    ["stock", "minimum_stock"].forEach((key) => data[key] = Number(data[key]));
    try {
      if (editing) {
        const result = await supabase("products", { method: "PATCH", query: `?id=eq.${product.id}`, body: data });
        if (!result?.[0]) throw new Error("Supabase tidak mengembalikan produk yang diperbarui.");
        Object.assign(product, result[0]);
        notify("Produk berhasil diperbarui");
      } else {
        const result = await supabase("products", { method: "POST", body: data });
        if (!result?.[0]) throw new Error("Supabase tidak mengembalikan produk yang disimpan.");
        state.products.unshift(result[0]);
        notify("Produk berhasil disimpan");
      }
      renderProducts();
    } catch (error) { notify("Produk gagal disimpan"); console.error(error); }
  };
}

function showCustomerForm(customer = null) {
  const editing = Boolean(customer?.id);
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">${editing ? "Edit pelanggan" : "Pelanggan baru"}</div><h1>${editing ? "Perbarui relasi." : "Tambah relasi."}</h1></div></div><form class="form-card" id="customer-form"><div class="form-grid"><div class="field"><label>Nama lengkap</label><input name="name" value="${escapeHtml(customer?.name || "")}" required></div><div class="field"><label>No. telepon</label><input name="phone" value="${escapeHtml(customer?.phone || "")}"></div><div class="field"><label>Email</label><input name="email" type="email" value="${escapeHtml(customer?.email || "")}"></div><div class="field full"><label>Alamat</label><textarea name="address" rows="3">${escapeHtml(customer?.address || "")}</textarea></div></div><div class="form-actions"><button type="button" class="btn light" data-page="customers">Batal</button><button class="btn">${editing ? "Simpan perubahan" : "Simpan pelanggan"}</button></div></form>`;
  $("#customer-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    try {
      if (editing) {
        const result = await supabase("customers", { method: "PATCH", query: `?id=eq.${customer.id}`, body: data });
        if (!result?.[0]) throw new Error("Supabase tidak mengembalikan pelanggan yang diperbarui.");
        Object.assign(customer, result[0]);
        notify("Pelanggan berhasil diperbarui");
      } else {
        const result = await supabase("customers", { method: "POST", body: data });
        if (!result?.[0]) throw new Error("Supabase tidak mengembalikan pelanggan yang disimpan.");
        state.customers.unshift(result[0]);
        notify("Pelanggan berhasil disimpan");
      }
      renderCustomers();
    } catch (error) { notify("Pelanggan gagal disimpan"); console.error(error); }
  };
}

function showPurchaseForm(entry = null) {
  const editing = Boolean(entry?.id);
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">${editing ? "Edit jurnal pembelian" : "Jurnal pembelian"}</div><h1>${editing ? "Perbarui pembelian." : "Catat pembelian."}</h1><p>Simpan pembelian bahan atau produk beserta status DP dan pelunasannya.</p></div></div><form class="form-card" id="purchase-form"><div class="form-grid"><div class="field"><label>Nomor referensi</label><input name="invoice_no" value="${escapeHtml(entry?.invoice_no || `PUR-${Date.now().toString().slice(-6)}`)}" required></div><div class="field"><label>Pemasok / relasi</label><select name="customer_id" required><option value="">Pilih pemasok / relasi</option>${state.customers.map((customer) => `<option value="${customer.id}" ${customer.id === entry?.customer_id ? "selected" : ""}>${escapeHtml(customer.name)}</option>`).join("")}</select></div><div class="field full"><label>Keterangan pembelian</label><input name="journal_description" value="${escapeHtml(entry?.journal_description || "")}" placeholder="Contoh: Pembelian benang katun premium" required></div><div class="field"><label>Total pembelian (Rp)</label><input name="total" class="money-input" inputmode="numeric" value="${moneyInput(entry?.total || "")}" placeholder="Rp0" required></div><div class="field"><label>Nominal DP / pembayaran (Rp)</label><input name="paid_amount" class="money-input" inputmode="numeric" value="${moneyInput(entry?.paid_amount || "")}" placeholder="Rp0" required></div><div class="field"><label>Metode pembayaran</label><select name="payment_method"><option value="transfer" ${entry?.payment_method === "transfer" ? "selected" : ""}>Transfer Bank</option><option value="qris" ${entry?.payment_method === "qris" ? "selected" : ""}>QRIS</option></select></div><div class="field"><label>Nomor referensi bank</label><input name="payment_reference" value="${escapeHtml(entry?.payment_reference || "")}" placeholder="Opsional"></div><div class="field full"><label>Catatan jurnal</label><textarea name="notes" rows="3" placeholder="Catatan tambahan untuk pembelian ini">${escapeHtml(entry?.notes || "")}</textarea></div></div><div class="form-actions"><button type="button" class="btn light" data-page="inventory">Batal</button><button class="btn">${editing ? "Simpan perubahan" : "Simpan jurnal"}</button></div></form>`;
  bindMoneyInputs($("#purchase-form"));
  $("#purchase-form").onsubmit = async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.target));
    data.entry_type = "purchase";
    data.sale_date = new Date().toISOString().slice(0, 10);
    data.total = parseMoney(data.total);
    data.paid_amount = parseMoney(data.paid_amount);
    data.payment_status = data.paid_amount >= data.total ? "verified" : "pending";
    try {
      if (editing) {
        const saved = (await supabase("sales", { method: "PATCH", query: `?id=eq.${entry.id}`, body: data }))?.[0];
        if (!saved) throw new Error("Supabase tidak mengembalikan jurnal pembelian yang diperbarui.");
        Object.assign(entry, saved);
        entry.customers = state.customers.find((customer) => customer.id === entry.customer_id);
        notify("Jurnal pembelian berhasil diperbarui");
      } else {
        const saved = (await supabase("sales", { method: "POST", body: data }))?.[0];
        if (!saved) throw new Error("Supabase tidak mengembalikan jurnal pembelian yang disimpan.");
        saved.customers = state.customers.find((customer) => customer.id === saved.customer_id);
        state.sales.unshift(saved);
        notify("Jurnal pembelian berhasil disimpan");
      }
      renderInventory();
    } catch (error) {
      notify("Jurnal pembelian gagal disimpan");
      console.error(error);
    }
  };
}

function showStockAdjustmentForm() {
  const productOptions = state.products
    .map((product) => `<option value="${product.id}">${escapeHtml(product.name)} · stok ${Number(product.stock || 0)} unit</option>`)
    .join("");
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Mutasi persediaan</div><h1>Penyesuaian stok.</h1><p>Catat penambahan stok atau unit yang rusak maupun hilang.</p></div></div><form class="form-card" id="stock-adjustment-form"><div class="form-grid"><div class="field full"><label>Produk</label><select name="product_id" required><option value="">Pilih produk</option>${productOptions}</select></div><div class="field"><label>Jenis penyesuaian</label><select name="movement" required><option value="in">Stok masuk</option><option value="damaged">Produk rusak</option><option value="lost">Produk hilang</option></select></div><div class="field"><label>Jumlah unit</label><input name="quantity" type="number" min="1" step="1" inputmode="numeric" placeholder="Contoh: 3" required></div><div class="field full"><label>Catatan (opsional)</label><textarea name="notes" rows="3" maxlength="500" placeholder="Contoh: tambahan hasil produksi atau ditemukan saat stok opname"></textarea></div></div><div class="config-note stock-adjustment-note">Stok produk akan langsung diperbarui dan catatan perubahannya disimpan di riwayat penyesuaian stok.</div><div class="form-actions"><button type="button" class="btn light" data-page="inventory">Batal</button><button class="btn">Simpan penyesuaian</button></div></form>`;
  $("#stock-adjustment-form").onsubmit = async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const formData = new FormData(form);
    const productId = String(formData.get("product_id") || "");
    const movement = String(formData.get("movement") || "");
    const quantity = Number(formData.get("quantity"));
    const notes = String(formData.get("notes") || "").trim();
    const product = state.products.find((item) => item.id === productId);
    if (!product) return notify("Pilih produk yang valid");
    if (!Number.isInteger(quantity) || quantity <= 0) return notify("Jumlah unit harus bilangan bulat lebih dari 0");
    if (movement !== "in" && quantity > Number(product.stock || 0)) return notify("Jumlah penyesuaian melebihi stok yang tersedia");
    const submitButton = form.querySelector("button[type=submit],button:not([type])");
    submitButton.disabled = true;
    try {
      const result = await supabaseRpc("record_stock_movement", {
        p_product_id: productId,
        p_movement: movement,
        p_quantity: quantity,
        p_notes: notes || null
      });
      const saved = Array.isArray(result) ? result[0] : result;
      if (!saved?.id || !Number.isInteger(Number(saved.stock_after))) {
        throw new Error("Supabase tidak mengembalikan catatan penyesuaian dan stok terbaru.");
      }
      product.stock = Number(saved.stock_after);
      product.updated_at = new Date().toISOString();
      state.sales.unshift(saved);
      notify("Penyesuaian stok berhasil disimpan");
      renderInventory();
    } catch (error) {
      submitButton.disabled = false;
      notify(`Penyesuaian stok gagal: ${explainSaveError(error)}`);
      console.error(error);
    }
  };
}

function showSettlementForm(sale) {
  if (!sale) return notify("Transaksi tidak ditemukan");
  const total = Number(sale.total || 0);
  const summary = salePaymentSummary(sale);
  const outstanding = summary.outstanding;
  if (!outstanding) return notify("Tagihan transaksi ini sudah lunas");
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Pembayaran lanjutan</div><h1>Catat pelunasan.</h1><p>${escapeHtml(sale.invoice_no)} · ${escapeHtml(sale.customers?.name || "Pelanggan")}</p></div></div><form class="form-card" id="settlement-form"><div class="form-grid"><div class="field"><label>Total tagihan</label><input class="money-input" value="${moneyInput(total)}" readonly></div><div class="field"><label>Sudah dibayarkan</label><input class="money-input" value="${moneyInput(summary.paid)}" readonly></div><div class="field"><label>Sisa tagihan</label><input class="money-input total-highlight" value="${moneyInput(outstanding)}" readonly></div><div class="field"><label>Tanggal pelunasan</label><input name="payment_date" type="date" value="${localIsoDate()}" required></div><div class="field"><label>Pembayaran</label><select name="settlement_type" required><option value="full">Lunasi sisa tagihan · ${escapeHtml(rupiah(outstanding))}</option></select></div><div class="field"><label>Metode pembayaran</label><select name="payment_method" required><option value="transfer">Transfer Bank</option><option value="qris">QRIS</option></select></div><div class="field"><label>Nomor referensi</label><input name="payment_reference" placeholder="Opsional"></div><div class="field full"><label>Bukti pelunasan *</label><div class="upload"><input name="payment_proof_file" type="file" accept=".jpg,.jpeg,.png,.pdf,image/jpeg,image/png,application/pdf" required><small class="proof-file-name" data-proof-name>Belum ada file dipilih</small><button type="button" class="btn light proof-remove" data-remove-proof hidden>Hapus file</button><small style="display:block;color:#a08f83;margin-top:8px">Wajib diisi. Format JPG, JPEG, PNG, atau PDF.</small></div></div></div><div class="form-actions"><button type="button" class="btn light" data-page="sales">Batal</button><button class="btn">Simpan pelunasan</button></div></form>`;
  const form = $("#settlement-form");
  bindProofUpload(form);
  form.onsubmit = async (event) => {
    event.preventDefault();
    const formData = new FormData(form);
    const proofFile = formData.get("payment_proof_file");
    if (!proofFile?.size) {
      notify("Bukti pelunasan wajib diunggah");
      form.querySelector("[name=payment_proof_file]").reportValidity();
      return;
    }
    const submitButton = form.querySelector(".form-actions .btn:last-child");
    submitButton.disabled = true;
    try {
      const currentSale = state.sales.find((item) => item.id === sale.id);
      if (!currentSale) throw new Error("Transaksi tidak lagi tersedia.");
      const currentOutstanding = salePaymentSummary(currentSale).outstanding;
      if (!currentOutstanding) throw new Error("Tagihan transaksi ini sudah lunas.");
      const paymentProofUrl = proofFile?.size ? await uploadPaymentProof(proofFile) : null;
      const result = await supabaseRpc("record_sale_settlement", {
        p_sale_id: currentSale.id,
        p_amount: currentOutstanding,
        p_payment_method: String(formData.get("payment_method")),
        p_payment_reference: String(formData.get("payment_reference") || ""),
        p_payment_proof_url: paymentProofUrl || "",
        p_payment_date: String(formData.get("payment_date"))
      });
      const saved = Array.isArray(result) ? result[0] : result;
      if (!saved?.id) throw new Error("Supabase tidak mengembalikan catatan pelunasan.");
      saved.customers = currentSale.customers;
      state.sales.unshift(saved);
      notify("Pelunasan berhasil dicatat");
      renderSales();
    } catch (error) {
      submitButton.disabled = false;
      notify(`Pelunasan gagal dicatat: ${explainSaveError(error)}`);
      console.error(error);
    }
  };
}

function saleFormMarkup(sale = {}) {
  const existingProduct = sale.product_id || "";
  const deliveryMethod = sale.delivery_method || "";
  const paymentPlan = Number(sale.paid_amount || 0) >= Number(sale.total || 0) && Number(sale.total || 0) > 0 ? "full" : "half";
  return `<div class="form-grid"><div class="field"><label>No. invoice</label><input name="invoice_no" value="${escapeHtml(sale.invoice_no || `LS-${Date.now().toString().slice(-6)}`)}" required></div><div class="field"><label>Tanggal penjualan</label><input name="sale_date" type="date" value="${escapeHtml(sale.sale_date || localIsoDate())}" required></div><div class="field"><label>Pelanggan</label><select name="customer_id" required><option value="">Pilih pelanggan</option>${state.customers.map((c) => `<option value="${c.id}" ${c.id === sale.customer_id ? "selected" : ""}>${escapeHtml(c.name)}</option>`).join("")}</select></div><div class="field full"><label>Kanal pembelian</label><select name="sales_channel" required><option value="in_store" ${sale.sales_channel === "in_store" ? "selected" : ""}>Beli langsung di toko</option><option value="whatsapp_online" ${!sale.sales_channel || sale.sales_channel === "whatsapp_online" ? "selected" : ""}>WhatsApp / Online</option></select></div><div class="field full"><label>Metode penerimaan</label><select name="delivery_method" required><option value="">Pilih metode penerimaan terlebih dahulu</option><option value="delivery" ${deliveryMethod === "delivery" ? "selected" : ""}>Delivery</option><option value="pickup" ${deliveryMethod === "pickup" ? "selected" : ""}>Ambil di tempat</option></select><small class="field-help">Keranjang akan digunakan setelah metode penerimaan dipilih.</small></div><div class="field full cart-section"><label>Tambah produk ke keranjang</label><div class="cart-add-row"><select name="cart_product_id"><option value="">Pilih produk</option>${state.products.map((p) => `<option value="${p.id}">${escapeHtml(p.name)} · ${rupiah(p.price)}</option>`).join("")}</select><input name="cart_quantity" type="number" min="1" value="1"><button type="button" class="btn" data-add-cart>＋ Tambahkan</button></div><div class="cart-box"><div class="cart-title">Keranjang transaksi</div><div class="table-wrap"><table class="cart-table"><thead><tr><th>Produk</th><th>Qty</th><th>Harga</th><th>Subtotal</th><th></th></tr></thead><tbody data-cart-body></tbody></table></div></div></div><div class="field"><label>Subtotal produk</label><input name="subtotal" class="money-input" inputmode="numeric" readonly value="${moneyInput(sale.subtotal || "")}"></div><div class="field"><label>Ongkir</label><input name="shipping_fee" class="money-input" inputmode="numeric" readonly value="${moneyInput(sale.shipping_fee ?? 0)}"></div><div class="field"><label>Total tagihan</label><input name="total" class="money-input total-highlight" inputmode="numeric" readonly value="${moneyInput(sale.total || "")}" required></div><div class="field"><label>Pembayaran saat ini</label><select name="payment_plan"><option value="half" ${paymentPlan === "half" ? "selected" : ""}>DP 50% dari total</option><option value="full" ${paymentPlan === "full" ? "selected" : ""}>Lunas 100%</option></select></div><div class="field"><label>Nominal dibayarkan</label><input name="paid_amount" class="money-input" inputmode="numeric" value="${moneyInput(sale.total ? paymentPlan === "full" ? sale.total : Math.round(Number(sale.total) / 2) : "")}" readonly required></div><div class="field full"><div class="shipping-note" data-shipping-note>Pilih metode penerimaan untuk melihat ongkir.</div></div><div class="field"><label>Metode pembayaran</label><select name="payment_method" required><option value="transfer" ${sale.payment_method === "transfer" ? "selected" : ""}>Transfer Bank</option><option value="qris" ${sale.payment_method === "qris" ? "selected" : ""}>QRIS</option></select></div><div class="field"><label>Nomor referensi</label><input name="payment_reference" value="${escapeHtml(sale.payment_reference || "")}" placeholder="Opsional"></div><div class="field full"><label>Bukti pembayaran *</label><div class="upload"><input name="payment_proof_file" type="file" accept=".jpg,.jpeg,.png,.pdf,image/jpeg,image/png,application/pdf" ${sale.payment_proof_url ? "" : "required"}><input type="hidden" name="payment_proof_url" value="${escapeHtml(sale.payment_proof_url || "")}"><small class="proof-file-name" data-proof-name>${sale.payment_proof_url ? "Bukti pembayaran tersimpan" : "Belum ada file dipilih"}</small><button type="button" class="btn light proof-remove" data-remove-proof ${sale.payment_proof_url ? "" : "hidden"}>Hapus file</button><small style="display:block;color:#a08f83;margin-top:8px">Wajib diisi. Format JPG, JPEG, PNG, atau PDF. File disimpan ke Supabase Storage.</small></div></div><div class="field full"><label>Catatan</label><textarea name="notes" rows="3">${escapeHtml(sale.notes || "")}</textarea></div></div>`;
}

function showSaleForm(sale = null) {
  const editing = Boolean(sale?.id);
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">${editing ? "Edit pembayaran" : "Pembayaran baru"}</div><h1>${editing ? "Perbarui transaksi." : "Catat penjualan."}</h1><p>Masukkan beberapa produk ke keranjang sebelum menyimpan pembayaran.</p></div></div><form class="form-card" id="sale-form" data-existing-product="${escapeHtml(sale?.product_id || "")}" data-existing-quantity="${sale?.quantity || 1}">${saleFormMarkup(sale || {})}<div class="form-actions"><button type="button" class="btn light" data-page="sales">Batal</button><button class="btn">${editing ? "Simpan perubahan" : "Simpan pembayaran"}</button></div></form>`;
  const saleForm = $("#sale-form");
  if (editing) {
    saleForm.dataset.existingCart = JSON.stringify(state.saleItems.filter((item) => item.sale_id === sale.id).map((item) => ({ productId: item.product_id, quantity: Number(item.quantity) })));
  }
  bindMoneyInputs(saleForm);
  bindProofUpload(saleForm);
  bindSaleCalculations(saleForm);
  saleForm.onsubmit = async (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const data = Object.fromEntries(formData);
    ["subtotal", "shipping_fee", "total", "paid_amount"].forEach((key) => data[key] = parseMoney(data[key]));
    const paymentPlan = data.payment_plan;
    delete data.payment_plan;
    data.paid_amount = paymentPlan === "full" ? data.total : Math.round(data.total / 2);
    const cart = e.target.getCart();
    if (!cart.length || data.subtotal < 50000) {
      notify("Minimal pembelian produk adalah Rp50.000");
      return;
    }
    const proofFile = formData.get("payment_proof_file");
    const storedProofUrl = String(formData.get("payment_proof_url") || "");
    if (!proofFile?.size && !storedProofUrl) {
      notify("Bukti pembayaran wajib diunggah");
      saleForm.querySelector("[name=payment_proof_file]").reportValidity();
      return;
    }
    delete data.payment_proof_file;
    data.payment_status = editing && data.paid_amount === Number(sale.paid_amount || 0) ? sale.payment_status : "pending";
    const { cart_product_id, cart_quantity, payment_proof_file, ...salePayload } = data;
    try {
      if (proofFile && proofFile.size) salePayload.payment_proof_url = await uploadPaymentProof(proofFile);
      let saved;
      if (editing) {
        saved = (await supabase("sales", { method: "PATCH", query: `?id=eq.${sale.id}`, body: salePayload }))?.[0];
        if (!saved) throw new Error("Supabase tidak mengembalikan pembayaran yang diperbarui.");
        Object.assign(sale, saved);
        sale.customers = state.customers.find((customer) => customer.id === sale.customer_id);
        await supabase("sale_items", { method: "DELETE", query: `?sale_id=eq.${sale.id}` });
        state.saleItems = state.saleItems.filter((line) => line.sale_id !== sale.id);
        const itemRows = cart.map((cartItem) => ({ sale_id: sale.id, product_id: cartItem.productId, quantity: cartItem.quantity, unit_price: state.products.find((product) => product.id === cartItem.productId)?.price || 0 }));
        const savedItems = await supabase("sale_items", { method: "POST", body: itemRows });
        if (!savedItems) throw new Error("Supabase tidak mengembalikan detail item pembayaran.");
        state.saleItems.push(...savedItems);
      } else {
        saved = (await supabase("sales", { method: "POST", body: salePayload }))?.[0];
        if (!saved) throw new Error("Supabase tidak mengembalikan pembayaran yang disimpan.");
        saved.customers = state.customers.find((customer) => customer.id === saved.customer_id);
        state.sales.unshift(saved);
        const itemRows = cart.map((cartItem) => ({ sale_id: saved.id, product_id: cartItem.productId, quantity: cartItem.quantity, unit_price: state.products.find((product) => product.id === cartItem.productId)?.price || 0 }));
        const savedItems = await supabase("sale_items", { method: "POST", body: itemRows });
        if (!savedItems) throw new Error("Supabase tidak mengembalikan detail item pembayaran.");
        state.saleItems.push(...savedItems);
      }
      notify(editing ? "Pembayaran berhasil diperbarui" : "Pembayaran berhasil disimpan");
      if (!editing) {
        await openReceipt(saved, true);
      } else renderSales();
    } catch (error) { notify("Pembayaran gagal disimpan"); console.error(error); }
  };
}

async function deleteProduct(id) {
  if (!window.confirm("Hapus produk ini?")) return;
  try {
    await supabase("products", { method: "DELETE", query: `?id=eq.${id}` });
    state.products = state.products.filter((product) => product.id !== id);
    notify("Produk dihapus");
    renderProducts();
  } catch (error) { notify("Produk gagal dihapus"); console.error(error); }
}

async function deleteCustomer(id) {
  if (!window.confirm("Hapus pelanggan ini?")) return;
  try {
    await supabase("customers", { method: "DELETE", query: `?id=eq.${id}` });
    state.customers = state.customers.filter((customer) => customer.id !== id);
    notify("Pelanggan dihapus");
    renderCustomers();
  } catch (error) { notify("Pelanggan gagal dihapus. Pastikan belum memiliki transaksi."); console.error(error); }
}

async function deleteSale(id, label = "Pembayaran") {
  if (!window.confirm(`Hapus ${label.toLowerCase()} ini?`)) return;
  try {
    await supabase("sales", { method: "DELETE", query: `?id=eq.${id}` });
    state.sales = state.sales.filter((sale) => sale.id !== id);
    notify(`${label} dihapus`);
    state.page === "inventory" ? renderInventory() : renderSales();
  } catch (error) { notify("Pembayaran gagal dihapus"); console.error(error); }
}

async function openPaymentProof(saleId) {
  const sale = state.sales.find((record) => record.id === saleId);
  if (!sale?.payment_proof_url) {
    notify("Bukti pembayaran tidak tersedia");
    return;
  }
  const proofUrl = sale.payment_proof_url;
  if (!/^https?:\/\//i.test(proofUrl)) {
    notify("URL bukti pembayaran tidak valid");
    return;
  }
  const isPdf = /\.pdf(?:$|[?#])/i.test(proofUrl);
  $("#proof-title").textContent = `Bukti · ${sale.invoice_no}`;
  $("#proof-content").innerHTML = isPdf
    ? `<iframe class="proof-preview" src="${escapeHtml(proofUrl)}" title="Bukti pembayaran PDF"></iframe>`
    : `<img class="proof-preview" src="${escapeHtml(proofUrl)}" alt="Bukti pembayaran ${escapeHtml(sale.invoice_no)}">`;
  const statusText = sale.entry_type === "settlement"
    ? sale.payment_status === "verified" ? "Pelunasan terverifikasi" : sale.payment_status === "rejected" ? "Pelunasan ditolak" : "Menunggu verifikasi pelunasan"
    : paymentStatusLabel(sale.payment_status, sale.paid_amount, sale.total);
  $("#proof-actions").innerHTML = `<span class="proof-current-status">Status: <strong>${escapeHtml(statusText)}</strong></span><a class="btn light" href="${escapeHtml(proofUrl)}" target="_blank" rel="noopener noreferrer">Buka file</a>${sale.payment_status === "pending" ? `<button class="btn light reject-payment" data-proof-reject="${sale.id}">Tolak</button><button class="btn" data-proof-verify="${sale.id}">✓ Verifikasi pembayaran</button>` : ""}`;
  $("#proof-modal").classList.add("open");
  $("#proof-modal").setAttribute("aria-hidden", "false");
  const verifyButton = $("#proof-actions [data-proof-verify]");
  const rejectButton = $("#proof-actions [data-proof-reject]");
  if (verifyButton) verifyButton.onclick = () => updatePaymentStatus(sale.id, "verified");
  if (rejectButton) rejectButton.onclick = () => updatePaymentStatus(sale.id, "rejected");
}

async function updatePaymentStatus(saleId, paymentStatus) {
  const sale = state.sales.find((record) => record.id === saleId);
  if (!sale) return notify("Transaksi tidak ditemukan");
  const action = paymentStatus === "verified" ? "memverifikasi" : "menolak";
  if (!window.confirm(`Yakin ingin ${action} pembayaran ${sale.invoice_no}?`)) return;
  try {
    const updated = (await supabase("sales", {
      method: "PATCH",
      query: `?id=eq.${saleId}`,
      body: { payment_status: paymentStatus }
    }))?.[0];
    if (!updated) throw new Error("Supabase tidak mengembalikan status pembayaran yang diperbarui.");
    Object.assign(sale, updated);
    closePaymentProof();
    notify(paymentStatus === "verified" ? "Pembayaran berhasil diverifikasi" : "Pembayaran ditolak");
    if (state.page === "inventory") renderInventory();
    if (state.page === "sales") renderSales();
  } catch (error) {
    notify("Status pembayaran gagal diperbarui");
    console.error(error);
  }
}

function closePaymentProof() {
  $("#proof-modal").classList.remove("open");
  $("#proof-modal").setAttribute("aria-hidden", "true");
  $("#proof-content").replaceChildren();
  $("#proof-actions").replaceChildren();
}

async function openReceipt(saleOrId, afterSave = false) {
  const sale = typeof saleOrId === "object" ? saleOrId : state.sales.find((item) => item.id === saleOrId);
  if (!sale) return notify("Transaksi tidak ditemukan");
  let items = [];
  try {
    if (sale.id) items = await supabase("sale_items", { query: `?select=*,products(name,sku)&sale_id=eq.${sale.id}` }) || [];
  } catch (error) { notify("Detail produk tidak dapat dimuat"); console.error(error); }
  const fallbackProduct = sale.product_id ? state.products.find((product) => product.id === sale.product_id) : null;
  if (!items.length && fallbackProduct) items = [{ quantity: sale.quantity || 1, unit_price: fallbackProduct.price, products: fallbackProduct }];
  const paymentSummary = salePaymentSummary(sale);
  const paid = paymentSummary.paid;
  const total = Number(sale.total || 0);
  const receiptNumber = `LS-${String(sale.id || Date.now()).replace(/[^a-zA-Z0-9]/g, "").slice(-8).toUpperCase()}`;
  const subtotal = Number(sale.subtotal || items.reduce((sum, item) => sum + Number(item.quantity || 0) * Number(item.unit_price || 0), 0));
  const shipping = Number(sale.shipping_fee || 0);
  const customerName = sale.customers?.name || "Kak";
  $("#receipt-paper").innerHTML = `<div class="receipt-brand"><span class="receipt-logo">✦</span><div><strong>LITTLE STITCHES</strong><small>Handmade Crochet Products</small></div></div><div class="receipt-address">Jl. Benang Ceria No. 12, Bandung · +62 812-3456-7890</div><div class="receipt-meta"><span>No. struk<strong>${receiptNumber}</strong></span><span>Tanggal<strong>${dateText(sale.sale_date || new Date())}</strong></span></div><div class="receipt-customer"><span>Pelanggan<strong>${escapeHtml(sale.customers?.name || "Pelanggan umum")}</strong></span><span>Invoice<strong>${escapeHtml(sale.invoice_no || "—")}</strong></span></div><div class="receipt-line"></div><div class="receipt-method"><span>Metode pembayaran<strong>${sale.payment_method === "qris" ? "QRIS" : "Transfer Bank"}</strong></span><span>Referensi<strong>${escapeHtml(sale.payment_reference || "Tidak tersedia")}</strong></span></div><table class="receipt-items"><thead><tr><th>Item</th><th>Qty</th><th>Harga</th><th>Subtotal</th></tr></thead><tbody>${items.length ? items.map((item) => `<tr><td>${escapeHtml(item.products?.name || "Produk rajut")}</td><td>${item.quantity}</td><td>${rupiah(item.unit_price)}</td><td>${rupiah(item.quantity * item.unit_price)}</td></tr>`).join("") : `<tr><td colspan="4" class="receipt-empty">Detail produk tidak tersedia</td></tr>`}</tbody></table><div class="receipt-totals"><span>Subtotal produk<strong>${rupiah(subtotal)}</strong></span><span>Ongkir flat<strong>${rupiah(shipping)}</strong></span><span>Total tagihan<strong>${rupiah(total)}</strong></span><span>Jumlah pembayaran<strong>${rupiah(paid)}</strong></span><span>Sisa tagihan<strong>${rupiah(paymentSummary.outstanding)}</strong></span><span class="receipt-grand">Status pembayaran<strong>${escapeHtml(paymentSummary.status.toUpperCase())}</strong></span></div><div class="receipt-thanks">Terima kasih telah mendukung karya handmade kami ♡<small>little stitches · made slowly, loved deeply</small></div>`;
  $("#receipt-title").textContent = afterSave ? "Pembayaran tersimpan" : "Struk pembayaran";
  $("#receipt-message").value = `Halo ${customerName}, berikut struk pembayaran untuk pesanan ${sale.invoice_no || "Little Stitches"}. Terima kasih sudah berbelanja!`;
  const customerPhone = sale.customers?.phone || state.customers.find((entry) => entry.id === sale.customer_id)?.phone;
  $("#receipt-share-hint").textContent = customerPhone
    ? `Tujuan: ${customerPhone}. Di ponsel, pilih WhatsApp dan chat pelanggan tersebut. Browser yang tidak mendukung berbagi gambar akan mengunduh PNG agar dapat dilampirkan manual.`
    : "Nomor telepon pelanggan belum tersedia. Tambahkan nomor pelanggan agar struk dapat dibagikan ke WhatsApp.";
  $("#receipt-modal").classList.add("open");
  $("#receipt-modal").setAttribute("aria-hidden", "false");
  $("#print-receipt").onclick = () => window.print();
  $("#send-whatsapp").onclick = () => sendReceiptToWhatsApp(sale, items, paymentSummary);
}

function closeReceipt() {
  $("#receipt-modal").classList.remove("open");
  $("#receipt-modal").setAttribute("aria-hidden", "true");
}

function renderSettings() {
  const configured = Boolean(config.url && config.publishableKey);
  $("#view").innerHTML = `<div class="page-heading"><div><div class="eyebrow">Workspace status</div><h1>Koneksi Supabase.</h1><p>Konfigurasi aplikasi dikelola oleh pemilik project.</p></div></div><section class="form-card"><div class="config-note">${configured ? "Project URL dan Publishable key sudah dimuat dari frontend/js/config.js. Pengunjung dapat menggunakan aplikasi tanpa mengisi kredensial." : "Aplikasi belum terhubung. Pemilik project perlu mengisi Project URL dan Publishable key satu kali di frontend/js/config.js sebelum aplikasi dipublikasikan. Pengunjung tidak perlu mengisi konfigurasi."}</div><div class="connection-summary"><span class="connection-status ${configured ? "ready" : "missing"}"></span><div><strong>${configured ? "Konfigurasi tersedia" : "Konfigurasi diperlukan"}</strong><small>${configured ? escapeHtml(config.url) : "Cek frontend/js/config.js"}</small></div></div><div class="form-actions"><button class="btn" id="retry-connection" ${configured ? "" : "disabled"}>Muat ulang data</button></div></section>`;
  $("#retry-connection").onclick = async () => { await loadData(); render(); };
}

function render() {
  document.querySelectorAll(".nav-link").forEach((link) => link.classList.toggle("active", link.dataset.page === state.page));
  ({ dashboard: renderDashboard, products: renderProducts, customers: renderCustomers, sales: renderSales, inventory: renderInventory, settings: renderSettings }[state.page] || renderDashboard)();
}

document.addEventListener("click", (event) => {
  const target = event.target.closest("[data-page]");
  if (target) { event.preventDefault(); state.page = target.dataset.page; location.hash = state.page; $(".sidebar").classList.remove("open"); render(); }
  if (event.target.closest("[data-close-receipt]")) closeReceipt();
  if (event.target.closest("[data-close-proof]")) closePaymentProof();
});
window.addEventListener("hashchange", () => { state.page = location.hash.slice(1) || "dashboard"; render(); });
$("#mobile-menu").onclick = () => $(".sidebar").classList.toggle("open");
$("#export-excel").onclick = () => {
  if (!config.url || !config.publishableKey) {
    notify("Konfigurasikan Supabase dan muat data sebelum mengekspor");
    return;
  }
  if (!state.products.length && !state.customers.length && !state.sales.length && !state.saleItems.length) {
    notify("Belum ada data untuk diekspor");
    return;
  }
  try {
    exportAllDataToExcel();
  } catch (error) {
    notify("Ekspor Excel gagal dibuat");
    console.error(error);
  }
};
$("#today").textContent = new Intl.DateTimeFormat("id-ID", { weekday: "long", day: "numeric", month: "long" }).format(new Date());
loadData().finally(render);
