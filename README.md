# Dashboard Evaluasi Operasi Sistem — UP2B Jawa Barat

Prototipe final project training Claude. Dashboard web yang mengubah file Excel **Lampiran Evaluasi Operasi Bulanan (EOB)** sistem kelistrikan PLN UP2B Jawa Barat menjadi gambaran risiko sistem yang bisa dibaca Manajemen Atas dalam hitungan menit.

**Buka dashboard:** https://haidarali98821.github.io/dashboard-eob-up2b-jabar/

---

## Untuk siapa

| Pengguna | Kebutuhan | Yang dilakukan di dashboard |
|---|---|---|
| **Penyusun EOB** (Tim Perencanaan dan Evaluasi Operasi Sistem) | Menyajikan hasil EOB tanpa menyusun ulang tabel dan grafik setiap bulan, sekaligus menangkap kesalahan data sebelum laporan disahkan | Mengunggah file Excel, membaca Catatan Data, memverifikasi angka |
| **Manajemen Atas PLN wilayah Jawa Barat** | Mengetahui kondisi dan risiko sistem tanpa membuka ±50 sheet Excel | Membaca Temuan utama dan modul yang relevan saat paparan rapat |

## Masalah yang diselesaikan

- Laporan EOB berbentuk satu file Excel dengan ±50 sheet. Informasi penting, misalnya rekor beban puncak atau sub sistem dengan cadangan tipis, tersebar di banyak tabel.
- Template diwariskan dari bulan ke bulan, sehingga label dan data lama mudah tertinggal tanpa disadari.
- Data operasi bersifat internal dan tidak boleh tersimpan di tempat publik.

## Cara mulai (3 langkah)

1. **Buka link dashboard** di laptop atau HP. Tidak perlu instalasi atau akun.
   Tampilan awal berlabel **DATA CONTOH**: nama instalasi disamarkan dan angka dikalikan 0,9. Angka ini bukan angka asli.
2. **Klik "Unggah Excel EOB"** dan pilih file *Lampiran EOB UP2B Jabar* bulan berjalan (format `.xlsx`). File juga bisa diseret langsung ke halaman.
   Periode laporan terbaca otomatis dari Lembar Pengesahan. Label di atas berubah menjadi **Data dari file**.
3. **Baca dari atas ke bawah:**
   1. **Temuan utama**: 4–6 kalimat risiko terpenting bulan ini.
   2. **Modul 1–7**: detail per topik.
   3. **Catatan Data**: hal yang perlu dicek penyusun sebelum dipaparkan.

> File Excel hanya dibaca di browser perangkat yang membukanya. Tidak ada data yang dikirim atau disimpan ke server mana pun, termasuk GitHub. Menutup halaman berarti datanya hilang dari dashboard.

## Hasil yang didapat

**Temuan utama dalam kalimat**, misalnya dari data contoh (samaran) di tampilan awal:
- Beban puncak 6.976 MW menjadi rekor baru, 2,8% di atas rekor sebelumnya.
- 7 dari 12 sub sistem berstatus waspada dengan cadangan di bawah 400 MW.
- 4 gangguan transmisi, 2 di antaranya menyebabkan padam.

**Tujuh modul evaluasi:**

| No | Modul | Isi | Sumber sheet |
|---|---|---|---|
| 1 | Ringkasan | Beban puncak, faktor beban, ENS, HTGP, ekskursi, pembebanan, SCADATEL | Rangkuman |
| 2 | Tren beban puncak | Siang/malam per bulan, tahun ini vs tahun lalu, pertumbuhan, faktor beban | Tabel 1, Load Factor |
| 3 | Pasokan per sub sistem | Daya Mampu Pasok, beban, cadangan, status **waspada** (< 400 MW) dan **defisit** (< 0), pilihan bulan, simulasi ambang | Tabel BP Koinsiden APB Per SS |
| 4 | Keandalan | Gangguan penghantar/trafo, penyebab gangguan, SRT, kalender HTGP, daftar gangguan transmisi | Tabel 3, 4, 7.1, 8, HTGP |
| 5 | Mutu tegangan | Tegangan min/maks realisasi vs rencana per level, ekskursi | Tabel 9, 10 |
| 6 | Pembebanan | Penghantar > 60%, trafo > 80%, sebaran pembebanan IBT dan trafo | Gambar 2.1/2.2, Data Python |
| 7 | Pekerjaan ROB | Realisasi per jenis rencana dan level tegangan | Evaluasi Pekerjaan ROB |

**Catatan kualitas data otomatis.** Contoh temuan nyata saat diuji dengan file Agustus dan September 2026:
- realisasi tegangan tercatat untuk bulan yang belum terjadi
- kejadian bertahun 2025 di tabel gangguan 2026
- waktu masuk lebih awal dari waktu keluar
- status pekerjaan ROB belum diisi
- jenis rencana yang tidak valid sehingga tidak terhitung di ringkasan Excel

## Rumus dan batasan

- **Daya Mampu Pasok (DMP)** = DMN IBT + KIT − Outage & Derating
- **Cadangan** = DMP − Beban Puncak Koinsiden
- **Waspada**: cadangan < 400 MW · **Defisit**: cadangan < 0. Slider ambang hanya untuk simulasi.
- Analisis kondisi normal; skenario N-1/kontingensi belum tercakup.
- Angka tetap diverifikasi penyusun sebelum dipaparkan ke manajemen.

## Jika ada modul yang tidak terbaca

Data dicari berdasarkan label (judul tabel, nama bulan, judul kolom), bukan nomor baris, sehingga tahan terhadap baris yang bertambah antar bulan. Bila struktur sheet berubah terlalu jauh, **hanya modul itu** yang ditandai "tidak terbaca" beserta alasannya. Modul lain tetap tampil, dan angka yang salah tidak pernah ditampilkan. Periksa sheet yang disebut, lalu unggah ulang.

## Isi repositori

| File | Fungsi |
|---|---|
| `index.html` | Halaman dashboard |
| `xlsx-lite.js` | Pembaca file .xlsx di browser |
| `eob-reader.js` | Pembaca struktur Lampiran EOB dan pemeriksaan kualitas data |
| `samaran.js` | Data contoh (samaran) untuk tampilan awal |

Dibuat dengan bantuan Claude sebagai final project training *Claude for Non-Tech Professionals*.
