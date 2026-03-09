# Troubleshooting Format Excel yang Berbeda

## Format yang Didukung Saat Ini

Sistem sudah dirancang untuk fleksibel menangani berbagai format Excel:

### 1. Format Persentase
- ✅ String dengan %: "84.62%", "8.06%"
- ✅ Desimal: 0.8462, 0.0806
- ✅ Angka biasa: 84.62, 8.06
- ❌ Format dengan koma: "84,62%" (belum didukung)

### 2. Posisi Kolom
- ✅ Kolom persentase di posisi 2 setelah "JENIS KEHADIRAN" (prioritas)
- ✅ Kolom persentase di posisi 3 atau lebih (fallback dengan deteksi otomatis)
- ❌ Kolom persentase sebelum kolom "JENIS KEHADIRAN" (belum didukung)

### 3. Nama Header
- ✅ "KEHADIRAN KESELURUHAN"
- ✅ "PERSENTASE KEHADIRAN"
- ✅ "PERSENTASE KEHADIRAN KESELURUHAN"
- ❌ Nama lain seperti "SUMMARY", "TOTAL KEHADIRAN" (belum didukung)

### 4. Nama Label Jenis Kehadiran
- ✅ WFO
- ✅ DINAS LUAR, DINAS, LUAR
- ✅ UNDANGAN
- ✅ SAKIT
- ✅ IJIN, IZIN, IIJN, IJN
- ✅ CUTI
- ✅ ISOMAN (akan digabung ke IJIN)
- ✅ TIDAK ABSEN, ALPHA, ALPA
- ❌ Nama lain seperti "WORK FROM OFFICE", "PERJALANAN DINAS" (belum didukung)

## Cara Cek Apakah Format Didukung

### Langkah 1: Upload Excel di Admin Panel
1. Buka admin panel dan upload Excel
2. Buka Console (F12)
3. Cari log berikut:

```
✅ Using column 2 (labelCol+2) for percentage
✅ Reading percentage from column 2
  ✅ WFO: 84.6%
  ✅ DINAS LUAR: 8.1%
  ...
```

### Langkah 2: Cek Database
Jalankan query SQL:

```sql
SELECT 
  bulan, tahun, bidang,
  wfo, wfo_pct,
  dinas_luar, dinas_luar_pct
FROM rekapitulasi 
WHERE tahun = 2025 AND bulan = 'Agustus' AND bidang = 'GABUNGAN';
```

Pastikan kolom `wfo_pct`, `dinas_luar_pct`, dll tidak NULL dan nilainya sesuai dengan Excel.

### Langkah 3: Cek Dashboard
1. Buka dashboard dan pilih bulan yang diupload
2. Buka Console (F12)
3. Cari log:

```
✅ Using REKAPITULASI data for donut chart
✅ Using PERCENTAGE from Excel (not calculated)
📊 Persentase dari Excel: WFO=84.6%, Dinas=8.1%, ...
```

## Jika Format Tidak Didukung

### Skenario 1: Kolom Persentase Tidak Terdeteksi

**Gejala:**
```
⚠️ Persentase tidak ditemukan di Excel, menghitung dari angka mentah...
```

**Solusi:**
1. Pastikan sheet REKAPITULASI ada
2. Pastikan ada header "KEHADIRAN KESELURUHAN" atau "PERSENTASE KEHADIRAN"
3. Pastikan kolom persentase ada di kolom 2 atau 3 setelah "JENIS KEHADIRAN"
4. Pastikan format persentase adalah string ("84.62%") atau desimal (0.8462)

**Jika masih gagal:**
- Screenshot sheet REKAPITULASI di Excel
- Screenshot console log dari admin panel
- Hubungi developer untuk update parser

### Skenario 2: Label Jenis Kehadiran Tidak Dikenali

**Gejala:**
Beberapa jenis kehadiran tidak muncul di dashboard atau persentasenya 0%.

**Solusi:**
1. Cek nama label di Excel, pastikan sesuai dengan daftar yang didukung
2. Jika nama berbeda, ubah di Excel atau hubungi developer untuk tambah pattern baru

**Contoh:**
Jika Excel menggunakan "WORK FROM OFFICE" bukan "WFO", ubah menjadi "WFO" atau minta developer tambah pattern:

```javascript
{ patterns: ['WFO', 'WORK FROM OFFICE'], key: 'wfo_pct' }
```

### Skenario 3: Format Persentase dengan Koma

**Gejala:**
Excel menggunakan format "84,62%" (dengan koma) tapi sistem tidak bisa baca.

**Solusi Sementara:**
1. Ubah format di Excel menjadi "84.62%" (dengan titik)
2. Atau ubah menjadi desimal: 0.8462

**Solusi Permanen:**
Hubungi developer untuk update parser agar support format dengan koma.

### Skenario 4: Struktur Tabel Berubah Total

**Gejala:**
Sistem tidak menemukan tabel KEHADIRAN KESELURUHAN sama sekali.

**Solusi:**
1. Pastikan sheet REKAPITULASI ada
2. Pastikan ada tabel dengan header "KEHADIRAN KESELURUHAN"
3. Jika struktur benar-benar berbeda, hubungi developer untuk update parser

## Tips Agar Format Tetap Konsisten

1. **Gunakan Template Excel yang Sama**
   - Simpan template Excel yang sudah berhasil diupload
   - Gunakan template yang sama untuk bulan-bulan berikutnya
   - Hanya ubah data, jangan ubah struktur

2. **Jangan Ubah Nama Header**
   - Tetap gunakan "KEHADIRAN KESELURUHAN"
   - Tetap gunakan "JENIS KEHADIRAN"
   - Tetap gunakan nama label yang sama (WFO, DINAS LUAR, dll)

3. **Jangan Ubah Posisi Kolom**
   - Kolom persentase tetap di posisi 2 atau 3 setelah "JENIS KEHADIRAN"
   - Jangan tambah kolom di tengah-tengah

4. **Gunakan Format Persentase yang Konsisten**
   - Jika menggunakan string "84.62%", tetap gunakan format itu
   - Jika menggunakan desimal 0.8462, tetap gunakan format itu
   - Jangan campur-campur format

## Kontak Developer

Jika mengalami masalah dengan format Excel yang berbeda:

1. Screenshot sheet REKAPITULASI di Excel (bagian tabel KEHADIRAN KESELURUHAN)
2. Screenshot console log dari admin panel (saat upload)
3. Screenshot dashboard yang menampilkan persentase salah
4. Kirim ke developer dengan deskripsi masalah

Developer akan update parser untuk support format baru.

---

Dibuat: 13 Februari 2026
Versi: 1.0

