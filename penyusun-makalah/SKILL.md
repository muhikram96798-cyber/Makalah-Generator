---
name: penyusun-makalah
description: >-
  Merancang, meriset, memverifikasi keterbukaan jurnal (open-access), menyusun
  kerangka, menulis, memvalidasi, dan mengompilasi makalah akademik berstandar
  perguruan tinggi (DOCX & PDF) dengan memori eksternal dan tata letak baku
  (A4, margin 4-3-3-3, Times New Roman 12pt, spasi 1.5/1.0, TOC Word native).
  Gunakan ketika pengguna meminta pembuatan makalah kuliah atau kajian ilmiah.
---

# Penyusun Makalah Akademik Standar (`penyusun-makalah`)

## Trigger Conditions
Aktifkan skill ini ketika pengguna meminta:
- Pembuatan makalah kuliah, paper akademik, tugas mata kuliah, atau kajian literatur.
- Penyusunan dokumen ilmiah terstruktur (Sampul, Kata Pengantar, Daftar Isi TOC Word, Bab I–III, Daftar Pustaka).
- Konversi atau kompilasi draf kajian akademik menjadi file `.docx` dan `.pdf` siap cetak berstandar kampus.

## Decision Tree
1. **Konfigurasi Atribut Sampul (Wajib Tanya via Chat), Logo Kampus, & Margin**:
   - *Komponen Cover Otomatis via Chat*: Jika pengguna belum memberikan identitas lengkap, **wajib minta pengguna mengisi sendiri seluruh komponen sampul via chat** sebelum kompilasi akhir: **(1) Judul Makalah**, **(2) Mata Kuliah** (`subtitle` opsional/dikosongkan secara *default*), **(3) Nama Dosen Pengampu + Gelar**, **(4) Nama & NIM/NPM Penulis** (tunggal/kelompok), **(5) Program Studi**, **(6) Fakultas**, **(7) Nama Perguruan Tinggi**, **(8) Kota & Tahun / Tahun Akademik**, serta **(9) Logo Kampus**. Dilarang menyisakan placeholder `[...]` pada `blueprint.json`.
   - *Margin & Tata Letak*: Terapkan **Struktur 3 Bab Utama**, **A4 margin isi 4-3-3-3** (khusus halaman sampul otomatis simetris `3-3-3-3` agar teks dan logo tepat di tengah kertas), **Times New Roman 12 pt**, dan tabel **Microsoft Word Grid**.
   - *Logo Cover*: **Wajib arahkan pengguna** meletakkan foto logo kampus (`.png`/`.jpg`) ke `./makalah_workspace/<slug-judul>/assets/logo_kampus.png` (atau kirim path-nya); `logo_default.png` hanya fallback bila tidak ada file logo.
2. **Sistem Penomoran Sub-bab**:
   - *Default*: Desimal konsisten (`1.1`, `1.2`, `2.1`, `2.1.1`), **maksimal 3 tingkat** (dilarang `2.1.1.1`). Alternatif: kombinasi (`A.`, `1.`, `a.`) maksimal 3 tingkat.
3. **Alur Riset & Gerbang Persetujuan (Wajib Default)**:
   - *Default*: Wajib **cari & uji keterbukaan Daftar Pustaka dulu** (`OPEN_ACCESS`, bebas *paywall* berbayar), lalu susun **kerangka pembahasan** yang memetakan sub-bab ke pustaka teruji tersebut, dan **tanyakan apakah ada tambahan atau perubahan** sebelum menulis naskah.
   - *Pengecualian*: Lewati gerbang persetujuan kerangka hanya jika pengguna secara eksplisit meminta *"langsung buat tanpa konfirmasi kerangka"* (uji keterbukaan pustaka tetap wajib).

## Core Workflow Steps

### Langkah 1: Pencarian Daftar Pustaka & Uji Keterbukaan (Anti-Paywall)
1. Baca aturan tata tulis pada [standar-akademik.md](./references/standar-akademik.md) dan skema data pada [schema-blueprint.md](./references/schema-blueprint.md).
2. Cari kandidat referensi ilmiah kredibel (**di atas 10 sumber / minimal 11 entri**, terbitan **5–10 tahun terakhir**: jurnal SINTA/DOAJ/GARUDA, prosiding, buku terbuka, laporan `.go.id`; dilarang Wikipedia, blog pribadi, atau medsos).
3. Buat folder `./makalah_workspace/<slug-judul>/` dan simpan kandidat ke `references_candidates.json`, lalu **uji keterbukaan & ekstrak inti isi PDF-nya**:
   ```powershell
   python "C:\Users\Zyrex D-Tech Pro\.gemini\config\skills\penyusun-makalah\scripts\cek_referensi.py" --input "./makalah_workspace/<slug-judul>/references_candidates.json" --output "./makalah_workspace/<slug-judul>/references_verified.json" --extract-notes "./makalah_workspace/<slug-judul>/references_notes.json"
   ```
4. **Tolak jurnal berbayar & penuhi kuota `> 10`**: Setiap entri `PAYWALLED`/`UNREACHABLE` **wajib diganti** dengan sumber `OPEN_ACCESS` hingga tersedia **minimal 11 sumber terbuka** yang inti isinya terekstrak di `references_notes.json`.

### Langkah 2: Poin-Poin Pembahasan Berbasis Jurnal & Persetujuan Pengguna
1. Baca `references_notes.json`, buat `blueprint.json` dengan `references` dari sumber `OPEN_ACCESS`, lalu petakan **poin-poin pembahasan spesifik** yang diambil dari isi nyata masing-masing jurnal (metode, data, temuan, evaluasi, kendala/solusi) selaras **1-to-1** dengan Rumusan Masalah & Kesimpulan.
2. **Wajib Konfirmasi Pengguna & Formulir Sampul via Chat**: Tampilkan rancangan **Kata Pengantar**, **Bab I**, **poin-poin pembahasan tiap sub-bab Bab II beserta rujukan jurnalnya + rencana Tabel/Gambar**, **Bab III**, hingga **Daftar Pustaka teruji**. Sekaligus **ajukan formulir ringkas di chat** agar pengguna mengisi sendiri komponen sampul yang masih kosong (**Judul, Mata Kuliah, Dosen Pengampu, Nama & NIM Penulis, Program Studi, Fakultas, Kampus, Kota/Tahun Akademik**) serta meletakkan logo kampus ke `./makalah_workspace/<slug-judul>/assets/logo_kampus.png` (atau mengirim path-nya) sebelum menulis draf/mengompilasi.

### Langkah 3: Penulisan Naskah yang Mudah Dipahami & Selaras Jurnal
Setelah poin-poin pembahasan disetujui, tulis draf ke `./makalah_workspace/<slug-judul>/drafts/`:
1. `kata_pengantar.md`: Puji syukur, konteks makalah, apresiasi dosen pengampu (tanggal & `"Penulis"` otomatis dari `preface_date`).
2. `bab_01.md`: `## 1.1 Latar Belakang` (piramida terbalik), `## 1.2 Rumusan Masalah`, `## 1.3 Tujuan Penulisan`.
3. `bab_02.md`: Tulis minimal **1.400 kata** (tiap sub-bab utama minimal **3 paragraf** seimbang) dengan **penjelasan yang jernih, mengalir, dan mudah dipahami** (jelaskan istilah teknis dengan contoh/analogi konkret) namun **100% selaras fakta/metrik dari `references_notes.json`**; bedah temuan tiap sumber tanpa menumpuk sitasi dalam satu kurung, lengkapi kutipan blok (`>`), tabel (`Tabel X.Y` & `Sumber:`), dan gambar (`Gambar X.Y` & `Sumber:`).
4. `bab_03.md`: `## 3.1 Kesimpulan` (1-to-1 dengan Rumusan Masalah) dan `## 3.2 Saran` (eksplisit untuk Instansi/Regulator, Masyarakat/Praktisi, dan Peneliti Selanjutnya).

### Langkah 4: Kompilasi Dokumen, Audit Kualitas, & Smoke Test
1. Kompilasi `.docx` (dengan native Word TOC) dan `.pdf` (two-pass dotted TOC) sekaligus jalankan *Academic Linter*:
   ```powershell
   python "C:\Users\Zyrex D-Tech Pro\.gemini\config\skills\penyusun-makalah\scripts\buat_makalah.py" --workspace "./makalah_workspace/<slug-judul>" --output-dir "./makalah_workspace/<slug-judul>/output"
   ```
2. Jika memodifikasi script di `scripts/`, wajib jalankan `--smoke-test` pada `cek_referensi.py` dan `buat_makalah.py`.
