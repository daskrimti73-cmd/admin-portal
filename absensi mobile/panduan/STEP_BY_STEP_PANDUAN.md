# PANDUAN STEP BY STEP: Menu Data Statistik dengan Grafik Absensi
## Website Kejaksaan Tinggi Kepulauan Riau

---

## 📌 STEP 1: LOGIN KE WORDPRESS ADMIN

1. Buka browser dan akses: `https://kejati-kepulauanriau.kejaksaan.go.id/wp-admin`
2. Masukkan username dan password admin
3. Anda akan masuk ke Dashboard WordPress

---

## 📌 STEP 2: MEMBUAT STRUKTUR MENU "DATA STATISTIK"

### A. Membuat Halaman-Halaman Baru

1. Di sidebar kiri, klik **"Laman"** (atau "Pages")
2. Klik **"Tambah Baru"** (atau "Add New")

**Buat 3 halaman berikut:**

#### Halaman 1: ABSENSI MOBILE
- Judul: `ABSENSI MOBILE`
- Permalink: `absensi-mobile`
- Klik **"Terbitkan"** (Publish)

#### Halaman 2: SIPEDE
- Judul: `SIPEDE`
- Permalink: `sipede`
- Klik **"Terbitkan"** (Publish)

#### Halaman 3: CMS
- Judul: `CMS`
- Permalink: `cms`
- Klik **"Terbitkan"** (Publish)

### B. Mengatur Menu Navigasi

1. Pergi ke **Tampilan** → **Menu** (Appearance → Menus)
2. Pilih menu utama yang sedang aktif
3. Di bagian **"Laman"** (Pages), centang ketiga halaman yang baru dibuat:
   - ✅ ABSENSI MOBILE
   - ✅ SIPEDE
   - ✅ CMS
4. Klik **"Tambah ke Menu"**

### C. Membuat Sub-Menu di Bawah "DATA STATISTIK"

1. Di struktur menu, **drag** halaman-halaman tersebut ke bawah "DATA STATISTIK"
2. Geser sedikit ke kanan untuk menjadikannya sub-menu
3. Atur urutan:
   ```
   DATA STATISTIK
      └── ABSENSI MOBILE
      └── SIPEDE
      └── CMS
   ```
4. Klik **"Simpan Menu"** (Save Menu)

---

## 📌 STEP 3: INSTALL PLUGIN UNTUK GRAFIK

Untuk menampilkan grafik batang dengan filter, kita perlu install plugin.

### Rekomendasi Plugin (Pilih Salah Satu):

#### Opsi A: wpDataTables (Rekomendasi Terbaik)
1. Pergi ke **Plugin** → **Tambah Baru**
2. Cari: `wpDataTables`
3. Klik **Install** lalu **Aktifkan**
4. Fitur: Import Excel, Grafik interaktif, Filter

#### Opsi B: TablePress + Extension
1. Install plugin `TablePress`
2. Install extension `TablePress Extension: Charts`

#### Opsi C: Visualizer Charts (Gratis)
1. Pergi ke **Plugin** → **Tambah Baru**
2. Cari: `Visualizer Charts`
3. Klik **Install** lalu **Aktifkan**

**Untuk panduan ini, saya akan menggunakan wpDataTables karena paling cocok untuk kebutuhan Anda.**

---

## 📌 STEP 4: MEMBUAT GRAFIK DENGAN wpDataTables

### A. Mengupload Data Excel

1. Siapkan file Excel absensi dengan format yang sudah saya buat di folder `data-absensi`
2. Di WordPress Admin, pergi ke **wpDataTables** → **Create a Table**
3. Pilih **"Create a table linked to an existing data source"**
4. Pilih sumber data: **Excel file**
5. Upload file Excel absensi
6. Klik **"Create the table"**

### B. Membuat Grafik Batang

1. Pergi ke **wpDataTables** → **Create a Chart**
2. Pilih tipe chart: **Bar Chart** (Grafik Batang)
3. Pilih tabel yang sudah diupload tadi
4. Konfigurasi:
   - **Axis X**: Tanggal/Hari
   - **Axis Y**: Jumlah (Hadir, Izin, Sakit, dll)
5. Aktifkan opsi **Filter**
6. Simpan chart

### C. Mendapatkan Shortcode

Setelah chart dibuat, Anda akan mendapat shortcode seperti:
```
[wpdatatable id=1]
[wpdatachart id=1]
```

---

## 📌 STEP 5: MEMASUKKAN GRAFIK KE HALAMAN ABSENSI MOBILE

1. Pergi ke **Laman** → **Semua Laman**
2. Klik **Edit** pada halaman "ABSENSI MOBILE"
3. Di editor, tambahkan konten:

```html
<h2>Dashboard Absensi Pegawai</h2>
<p>Pilih bulan dan tahun untuk melihat data absensi:</p>

<!-- Filter Bulan -->
[wpdatachart id=1]

<!-- Tabel Detail -->
[wpdatatable id=1]
```

4. Klik **Update** untuk menyimpan

---

## 📌 STEP 6: MEMBUAT FILTER BULAN DAN TAHUN

Untuk membuat filter dropdown bulan dan tahun, ada 2 cara:

### Cara 1: Menggunakan wpDataTables Filter (Built-in)

1. Edit tabel di wpDataTables
2. Di tab **"Filtering"**, aktifkan filter untuk kolom "Tanggal"
3. Simpan perubahan

### Cara 2: Membuat Halaman Terpisah per Bulan

Jika filter built-in tidak cukup, buat sub-halaman:
```
ABSENSI MOBILE
   └── September 2025
   └── Oktober 2025
   └── November 2025
   └── Desember 2025
   └── dst...
```

---

## 📌 LANGKAH SELANJUTNYA

Setelah Step 1-6 selesai, konfirmasi kepada saya:

1. ✅ Apakah Anda sudah bisa login ke WordPress Admin?
2. ✅ Apakah halaman sudah berhasil dibuat?
3. ✅ Plugin apa yang Anda pilih untuk diinstall?
4. ✅ Apakah data Excel sudah siap dengan format yang benar?

Saya akan membantu Anda di step berikutnya setelah ini selesai.

---

## 📞 TROUBLESHOOTING

### Masalah: Tidak bisa install plugin
- **Solusi**: Hubungi admin IT untuk mendapatkan akses install plugin, atau minta mereka yang menginstall

### Masalah: Format Excel tidak sesuai
- **Solusi**: Gunakan template yang saya buat di folder `data-absensi`

### Masalah: Grafik tidak muncul
- **Solusi**: Pastikan shortcode sudah benar dan plugin sudah aktif

---

*Dokumen ini dibuat untuk membantu pembuatan menu Data Statistik di Website Kejati Kepri*
*Dibuat: Januari 2026*
