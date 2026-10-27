# Dashboard Evaluasi Operasi Sistem — UP2B Jawa Barat

Prototipe final project training Claude. Dashboard web untuk menyajikan hasil Evaluasi Operasi Bulanan (EOB) sistem kelistrikan PLN UP2B Jawa Barat kepada Manajemen Atas.

## Cara pakai

1. Buka halaman dashboard. Isi bawaannya adalah **data contoh (samaran)**: nama instalasi disamarkan dan angka diskalakan.
2. Klik **Unggah Excel EOB** dan pilih file *Lampiran EOB UP2B Jabar* bulan berjalan (.xlsx).
3. File dibaca **di browser saja**. Tidak ada data yang dikirim atau disimpan ke server mana pun, termasuk GitHub.

## Isi dashboard

Ringkasan dan temuan utama, tren beban puncak, kesiapan pasokan per sub sistem (status waspada/defisit, ambang 400 MW), keandalan, mutu tegangan, pembebanan penghantar dan trafo, realisasi pekerjaan ROB, dan catatan kualitas data otomatis.

## Prinsip

- Data dicari berdasarkan label (judul tabel, nama bulan, judul kolom), bukan nomor baris.
- Bila satu bagian tidak terbaca, hanya modul itu yang ditandai; modul lain tetap tampil.
- Rumus yang dihitung ulang: Daya Mampu Pasok = DMN − Outage & Derating; Cadangan = DMP − Beban.
- Analisis kondisi normal; skenario N-1/kontingensi belum tercakup.
- Angka tetap diverifikasi penyusun sebelum dipaparkan.

## File

| File | Fungsi |
|---|---|
| `index.html` | Halaman dashboard |
| `xlsx-lite.js` | Pembaca file .xlsx di browser |
| `eob-reader.js` | Pembaca struktur Lampiran EOB dan pemeriksaan kualitas data |
| `samaran.js` | Data contoh (samaran) |
