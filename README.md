<div align="center">
  <img src="frontend/assets/little-stitches-logo.png" alt="Logo Little Stitches" width="104" height="104">
  <h1>Little Stitches</h1>
  <p><strong>Sistem Penjualan & Persediaan Produk Rajut</strong></p>
  <p>Ruang kerja sederhana untuk mengelola karya handmade, pelanggan, pembayaran, dan persediaan.</p>
  <p><a href="https://faniahidayatul46-dev.github.io/Sistem-Penjualan-Toko-Rajut/frontend/index.html"><strong>🧶 Buka Aplikasi Little Stitches</strong></a></p>

  <p>
    <img alt="Frontend" src="https://img.shields.io/badge/Frontend-HTML%20%7C%20CSS%20%7C%20JavaScript-8f4d6c?style=flat-square">
    <img alt="Backend" src="https://img.shields.io/badge/Backend-Node.js-e5a078?style=flat-square">
    <img alt="Database" src="https://img.shields.io/badge/Database-Supabase%20%7C%20PostgreSQL-4d806a?style=flat-square">
    <img alt="License" src="https://img.shields.io/badge/License-Private-4a2038?style=flat-square">
  </p>
</div>

<div align="center">
  <sub><code>#481D37</code> Plum &nbsp;·&nbsp; <code>#8F4D6C</code> Berry &nbsp;·&nbsp; <code>#E5A078</code> Peach &nbsp;·&nbsp; <code>#F8F3EC</code> Cream</sub>
</div>

---

## Tentang aplikasi

Little Stitches adalah aplikasi web untuk operasional penjualan dan persediaan produk rajut. Aplikasi menggunakan frontend HTML, CSS, dan JavaScript tanpa framework; backend Node.js menyajikan file web; dan Supabase menyediakan REST API, PostgreSQL, serta penyimpanan bukti pembayaran.

### Fitur

- **Dashboard** untuk melihat ringkasan penjualan, jumlah produk dan pelanggan, tren penjualan, serta stok yang perlu diperhatikan.
- **Produk rajut** untuk mengelola katalog, harga, kategori, dan jumlah persediaan.
- **Pelanggan** untuk menyimpan informasi kontak dan mengurutkan daftar berdasarkan tanggal maupun alfabet.
- **Pembayaran** untuk mencatat pembelian langsung di toko atau melalui WhatsApp/online, memilih QRIS atau transfer, dan merekam DP 50% atau pembayaran penuh.
- **Struk** untuk melihat dan mencetak transaksi serta membuat gambar struk PNG yang bisa dibagikan lewat menu berbagi perangkat, bersama pesan yang dapat disunting.
- **Jurnal persediaan** untuk melihat transaksi debit-kredit, mencatat pembelian, dan mengoreksi stok masuk, rusak, atau hilang.
- **Ekspor Excel** untuk mengunduh data dalam workbook dengan sheet Produk, Pelanggan, Transaksi & Jurnal, dan Detail Item.

## Teknologi

| Bagian | Teknologi |
| --- | --- |
| Antarmuka | HTML, CSS, JavaScript |
| Server web | Node.js bawaan (`http`, tanpa framework) |
| Data dan API | Supabase REST API |
| Basis data | PostgreSQL |
| Penyimpanan file | Supabase Storage |

## Struktur proyek

```text
.
├── index.html                 # Entry point GitHub Pages; mengarahkan ke frontend
├── backend/
│   └── app.js                  # Server Node.js dan endpoint health check
├── database/
│   └── schema.sql              # Skema Supabase, RPC, kebijakan, dan data awal
├── frontend/
│   ├── assets/                 # Logo, ilustrasi, foto banner, dan motif rajut
│   ├── js/
│   │   ├── app.js              # Logika aplikasi dan integrasi Supabase
│   │   └── config.js           # URL project dan publishable key Supabase
│   ├── index.html              # Halaman utama aplikasi
│   └── style.css               # Gaya visual responsif Little Stitches
└── README.md
```

## Menjalankan aplikasi

### Prasyarat

- Node.js 18 atau lebih baru.
- Project Supabase dengan akses ke SQL Editor dan API.

### 1. Siapkan database Supabase

1. Buka **SQL Editor** pada project Supabase.
2. Jalankan isi [`database/schema.sql`](database/schema.sql).
3. Skrip menyiapkan empat entitas utama: `customers`, `products`, `sales`, dan `sale_items`; termasuk fungsi RPC dan bucket `payment-proofs`.

### 2. Atur koneksi Supabase

Edit [`frontend/js/config.js`](frontend/js/config.js), lalu isi URL project dan **publishable key** dari pengaturan API Supabase:

```js
window.SUPABASE_CONFIG = {
  url: "https://PROJECT-REF.supabase.co",
  publishableKey: "YOUR_SUPABASE_PUBLISHABLE_KEY"
};
```

Jangan masukkan `service_role` key atau secret key ke frontend. Publishable key memang digunakan aplikasi browser; batasi akses data dengan kebijakan keamanan yang sesuai kebutuhan deployment.

### 3. Jalankan server lokal

Dari folder root proyek, jalankan:

```powershell
node backend\app.js
```

Buka aplikasi di [http://localhost:3000](http://localhost:3000). Pemeriksaan server tersedia di [http://localhost:3000/api/health](http://localhost:3000/api/health).

Server menggunakan modul bawaan Node.js, sehingga tidak memerlukan `package.json` atau instalasi dependency untuk dijalankan.

## Model data

ERD dipertahankan dengan **empat tabel utama**:

| Entitas | Kegunaan |
| --- | --- |
| `customers` | Informasi pelanggan. |
| `products` | Katalog produk rajut dan stok terkini. |
| `sales` | Penjualan, pembelian, penyesuaian stok, dan pelunasan; transaksi jurnal disimpan di entitas ini. |
| `sale_items` | Rincian produk dan kuantitas untuk setiap penjualan. |

Hubungan utama: satu pelanggan dapat memiliki banyak transaksi; satu transaksi penjualan dapat memiliki banyak detail item; setiap detail item merujuk ke satu produk.

## Catatan deployment

- Backend saat ini adalah server file statis kecil; operasi data aplikasi dilakukan dari browser ke Supabase REST API.
- Untuk deployment berbasis Node.js, jalankan `node backend\app.js` pada layanan hosting yang mendukung Node.js dan atur `PORT` melalui environment layanan bila diperlukan.
- Jika hanya menerbitkan frontend statis, host isi folder `frontend/` dan pastikan `frontend/js/config.js` berisi konfigurasi project yang benar.
- **GitHub Pages:** pada repository, buka **Settings → Pages**, pilih **Deploy from a branch**, pilih branch tempat proyek disimpan dan folder **/(root)**. File [`index.html`](index.html) di root akan meneruskan pengunjung ke aplikasi di folder `frontend/`. Setelah deployment selesai, buka URL Pages yang ditampilkan pada pengaturan tersebut.
- `schema.sql` mengaktifkan akses publik pada tabel dan bucket sesuai pola aplikasi saat ini. Tinjau kebijakan Row Level Security dan Storage sebelum menggunakan data produksi.
- Aplikasi tidak menyimpan atau membutuhkan kredensial server di browser.

## Identitas visual

Logo benang dan palet warna README mengikuti identitas Little Stitches:

<div align="center">
  <img src="frontend/assets/little-stitches-logo.png" alt="Logo Little Stitches — gulungan benang minimalis" width="88">
  <br>
  <strong>Hangat, lembut, dan dirajut dengan hati.</strong>
</div>
