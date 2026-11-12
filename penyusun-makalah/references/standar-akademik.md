# Standar Penulisan Makalah Akademik Perguruan Tinggi

Dokumen referensi ini memuat aturan baku penulisan makalah ilmiah yang wajib dipatuhi pada saat penyusunan substansi draf (`drafts/*.md`) maupun kompilasi dokumen akhir (`.docx` dan `.pdf`).

---

## 1. Konfigurasi Default & Struktur Dokumen
Jika pengguna tidak memberikan instruksi khusus, terapkan konfigurasi bawaan berikut:

- **Halaman Sampul (Cover)**:
  - **Wajib arahkan pengguna** untuk meletakkan foto/gambar logo resmi kampus (`assets/logo_kampus.png` atau path file `.png`/`.jpg` melalui `logo_path`) agar dicetak proporsional menjaga rasio asli (sisi terpanjang maks. `4,5 cm`) di tengah halaman dengan jarak vertikal seimbang penuh 1 halaman A4. Jika pengguna belum memiliki file logo, sistem otomatis menggunakan `assets/logo_default.png`.
  - Pastikan akurasi penulisan atribut:
    1. Judul Makalah (HURUF KAPITAL, *Bold*, 14 pt, rata tengah).
    2. Subjudul / Maksud Penulisan (contoh: *"Makalah Ini Disusun untuk Memenuhi Tugas pada Mata Kuliah ... Semester ... Tahun Akademik ..."*).
    3. Nama Dosen Pengampu (beserta gelar akademik lengkap dan akurat tanpa *typo*).
    4. Nama Mahasiswa & NIM (ditulis lengkap atau dalam tabel/daftar simetris).
    5. Nama Program Studi, Fakultas, Nama Institusi/Universitas, Kota, serta Tahun Akademik.
- **Struktur Bab Default (3 Bab Utama)**:
  - **Bagian Awal**: Halaman Sampul, Kata Pengantar, Daftar Isi (serta Daftar Tabel dan Daftar Gambar otomatis jika terdapat tabel/gambar di dalam naskah).
  - **BAB I PENDAHULUAN**:
    - `1.1 Latar Belakang`
    - `1.2 Rumusan Masalah`
    - `1.3 Tujuan Penulisan`
  - **BAB II PEMBAHASAN**:
    - Sub-bab analisis mendalam (`2.1`, `2.2`, dst.) sesuai topik dan butir rumusan masalah.
  - **BAB III PENUTUP**:
    - `3.1 Kesimpulan`
    - `3.2 Saran`
  - **Bagian Akhir**: `DAFTAR PUSTAKA`.

---

## 2. Spesifikasi Tata Letak & Tipografi (Page Setup)
- **Ukuran Kertas & Margin**:
  - Kertas **A4** (`21,0 x 29,7 cm`).
  - Batas margin default **4-3-3-3**: Kiri `4 cm`, Atas `3 cm`, Kanan `3 cm`, Bawah `3 cm` (dapat diubah ke `4-4-3-3` via flag CLI `--margin 4,4,3,3` bila ketentuan kampus meminta Atas `4 cm`).
- **Font**: **Times New Roman**, ukuran **12 pt** untuk seluruh teks utama.
- **Spasi Baris (Line Spacing)**:
  - **1,5 spasi**: Seluruh isi paragraf utama.
  - **1,0 spasi (Tunggal)**: Kutipan langsung panjang ($\ge 4$ baris), isi tabel, keterangan tabel/gambar, dan entri Daftar Pustaka.
- **Perataan & Indensi Paragraf**:
  - Seluruh paragraf isi wajib **Rata Kanan-Kiri (Justify)**.
  - Baris pertama setiap paragraf baru menjorok ke dalam (**First Line Indent**) sejauh **1,25 cm** (1x Tab).
- **Hierarki Judul & Penomoran Sub-bab**:
  - **Judul Bab**: HURUF KAPITAL SEMUA, **Bold**, rata tengah (*Center*).
  - **Judul Sub-bab**: *Title Case* (Huruf Kapital di Setiap Awal Kata selain kata tugas/preposisi seperti *di, ke, dari, dan, pada, dalam, untuk, dengan*), **Bold**, rata kiri (*Align Left*).
  - **Konsistensi Penomoran**: Gunakan satu sistem secara konsisten di seluruh dokumen (sistem desimal `1.1`, `1.1.1` ATAU kombinasi `A.`, `1.`, `a.`).
  - **Batas Kedalaman**: Maksimal **3 tingkat** sub-bab (misal: `2.1.1`). **Dilarang keras** membuat percabangan hingga 4–5 tingkat (seperti `2.1.3.4` atau `2.1.3.4.1`). Gunakan poin rincian di dalam paragraf bila diperlukan.
- **Penomoran Halaman (Pagination)**:
  - **Halaman Sampul**: Tidak diberi nomor halaman.
  - **Bagian Awal (Kata Pengantar s.d. Daftar Isi/Tabel/Gambar)**: Angka Romawi kecil (`i, ii, iii, iv`) di bagian **bawah tengah**.
  - **Bagian Isi & Akhir (Bab I s.d. Daftar Pustaka)**: Angka Arab (`1, 2, 3, 4, ...`) dimulai dari **`1`** dan seragam di bagian **bawah tengah** pada seluruh halaman (dapat diubah ke pojok kanan atas untuk halaman lanjutan via `cfg["pagination"]["body_other_page_pos"]`).
- **Kontrol Halaman (Page Break & Orphan/Widow)**:
  - Setiap bab utama (Bab I, Bab II, Bab III, Daftar Pustaka) wajib dimulai pada halaman baru menggunakan **Section / Page Break** (bukan *Enter* berulang).
  - **Dilarang ada judul sub-bab menggantung (Orphan Heading)** sendirian di bagian paling bawah halaman tanpa diikuti paragraf isinya (wajib mengaktifkan properti `keep_with_next = True`).

---

## 3. Kaidah Kebahasaan, Ejaan, & Penulisan Angka
- **Bahasa & Ejaan**: Gunakan Bahasa Indonesia baku, lugas, dan objektif sesuai **EYD V / PUEBI**. Hindari kata ganti orang pertama/kedua subjektif (*aku, kamu, kita*); gunakan kalimat pasif akademis atau istilah *penulis*.
- **Istilah Asing**: Semua istilah bahasa asing (misal: *artificial intelligence*, *supply chain*, *peer-review*) atau kata daerah/serapan yang belum baku wajib dicetak miring (`*italic*`).
- **Aturan Angka & Bilangan**:
  - **Larangan di Awal Kalimat**: Dilarang mengawali kalimat dengan angka. Tuliskan dengan huruf (contoh: *"Dua puluh responden terlibat dalam..."*, **bukan** *"20 responden terlibat dalam..."*).
  - **Bilangan 1–2 Kata**: Nyatakan dengan huruf bila dapat ditulis dengan satu atau dua kata (contoh: *"tiga kali"*, *"seratus lima puluh orang"*), kecuali jika merupakan rincian berurutan (contoh: *"terdiri atas 12 pria, 15 wanita, dan 4 anak-anak"*).
  - **Satuan Ukur**: Wajib menggunakan angka jika diikuti satuan berat, panjang, luas, waktu, mata uang, atau persentase (contoh: `5 cm`, `10 kg`, `25%`, `Rp50.000,00`).
- **Singkatan & Akronim**:
  - **Kemunculan Pertama**: Wajib ditulis kepanjangannya terlebih dahulu diikuti singkatan dalam kurung, contoh: `Badan Pusat Statistik (BPS)` atau `World Health Organization (WHO)`. Kemunculan berikutnya boleh menggunakan singkatan saja (`BPS`, `WHO`).
  - **Singkatan Umum**: Singkatan baku seperti `dll.`, `dsb.`, `dkk.`, `yth.`, dan `hlm.` dapat langsung digunakan sesuai tanda baca EYD V.

---

## 4. Substansi & Koherensi Antarbagian
- **Kepadatan Paragraf (Satu Ide Pokok)**:
  - Setiap paragraf wajib memuat **1 gagasan utama** yang utuh (kalimat utama + kalimat penjelas/bukti sitasi + kalimat penegas/transisi) dan terdiri dari **3 hingga 7 kalimat**.
  - **Dilarang** membuat paragraf yang hanya berisi 1–2 kalimat tunggal atau paragraf raksasa yang memakan hampir 1 halaman penuh.
- **Latar Belakang (Piramida Terbalik)**:
  - Susun dengan alur deduktif **Umum ke Khusus**:
    1. *Paragraf Pembuka (Konteks Makro)*: Fenomena global/nasional dan konsep ideal (*das sollen*).
    2. *Paragraf Tengah (Fakta Empiris & Masalah Spesifik)*: Data statistik, temuan terdahulu, dan kesenjangan di lapangan (*das sein*).
    3. *Paragraf Penutup (Urgensi)*: Alasan akademis dan praktis mengapa topik tersebut mendesak untuk dikaji dalam makalah ini.
- **Kedalaman & Gaya Penulisan Pembahasan (Bab II)**:
  - **Ekstraksi Poin Sebelum Menulis**: Ambil poin-poin inti (metode, temuan, metrik hasil, kendala) dari `references_notes.json` terlebih dahulu agar setiap uraian berakar pada isi nyata jurnal.
  - **Jernih, Mengalir, & Mudah Dipahami**: Jelaskan konsep dan istilah teknis dengan kalimat baku yang lugas, tidak berbelit-belit, serta sertakan contoh atau alur konkret agar mudah dipahami pembaca, tanpa mengurangi akurasi fakta jurnal.
  - **Bobot Seimbang**: Bab II minimal **1.400 kata**, setiap sub-bab utama (`2.1`, `2.2`, `2.3`) minimal **3 paragraf** analitis (ideal dipecah seimbang menjadi anak sub-bab `2.X.1`, `2.X.2`), mencakup konsep/mekanisme, komparasi studi (+ Tabel), evaluasi kinerja, serta kendala/keamanan/mitigasi, dan wajib membedah tiap referensi (hindari menumpuk banyak sitasi sekaligus tanpa uraian).
- **Sinkronisasi Rumusan Masalah & Kesimpulan**:
  - Jumlah dan urutan poin pada **Kesimpulan (Bab III, Sub-bab 3.1)** wajib menjawab secara langsung dan **sama persis jumlahnya** dengan poin pertanyaan pada **Rumusan Masalah (Bab I, Sub-bab 1.2)** serta **Tujuan Penulisan (Bab I, Sub-bab 1.3)**.
- **Kualitas Saran (Bab III, Sub-bab 3.2)**:
  - Saran harus **spesifik, konkret, dan operasional** (hindari kalimat klise seperti *"pembaca diharapkan lebih rajin belajar"* atau *"semoga makalah ini bermanfaat"*).
  - Tujukan saran secara eksplisit kepada 3 pihak terkait:
    1. **Bagi Instansi / Regulator / Pemangku Kebijakan**
    2. **Bagi Masyarakat / Praktisi / Pelaku Industri**
    3. **Bagi Peneliti / Penulis Selanjutnya** (untuk pengembangan metodologi atau ruang lingkup topik).

---

## 5. Aturan Kutipan, Tabel/Gambar, & Daftar Pustaka
- **Orisinalitas & Parafrase**:
  - Wajib melakukan **parafrase** (menulis ulang ide sumber dengan struktur kalimat sendiri yang akademis) untuk menjaga batas kemiripan (*similarity index*) di bawah **15%–25%**. Dilarang menyalin teks secara mentah (*copy-paste*).
- **Kutipan Langsung**:
  - **Kutipan Pendek (< 4 baris)**: Gabungkan langsung di dalam paragraf utama, apit dengan tanda petik dua (`"..."`), tetap menggunakan spasi **1,5**, disertai sitasi `(Penulis, Tahun: Halaman)`.
  - **Kutipan Panjang ($\ge$ 4 baris)**: Pisahkan menjadi blok paragraf tersendiri (gunakan sintaks `>` blockquote pada Markdown) **tanpa tanda petik**, menjorok ke dalam **1,25 cm** di sisi kiri, dengan jarak spasi **1,0 (tunggal)**.
- **Tabel dan Gambar**:
  - **Judul Tabel**: Letakkan di **atas** tabel dengan nomor urut bab (contoh: `Tabel 2.1 Perbandingan Efisiensi Algoritma`).
  - **Judul Gambar**: Letakkan di **bawah** gambar/grafik/diagram (contoh: `Gambar 2.1 Arsitektur Sistem Informasi`).
  - **Sumber Tabel/Gambar**: Wajib mencantumkan sumber referensi tepat di **bawah** tabel atau di bawah judul gambar (contoh: `Sumber: Diolah dari Badan Pusat Statistik (2024)`).
- **Kualitas & Seleksi Referensi**:
  - Hanya gunakan sumber kredibel: buku teks akademik, artikel jurnal ilmiah terakreditasi (SINTA/Scopus/DOAJ), prosiding konferensi, atau laporan resmi pemerintah/lembaga internasional.
  - **Rentang Tahun**: Utamakan referensi mutakhir terbitan **5–10 tahun terakhir** (kecuali teori klasik/grand theory).
  - **Sumber Terlarang**: Dilarang keras mengutip dari Wikipedia, blog pribadi (WordPress, Blogspot, Medium tanpa peninjauan sejawat), atau unggahan media sosial.
- **Penulisan Daftar Pustaka**:
  - **Jumlah Minimal**: Daftar Pustaka wajib berjumlah **di atas 10 referensi (minimal 11 sumber `OPEN_ACCESS` terverifikasi)**.
  - **Sinkronisasi 100%**: Setiap sitasi di dalam teks (`(Santoso, 2023)` atau `Menurut Santoso (2023)`) wajib tercantum di Daftar Pustaka, dan setiap entri di Daftar Pustaka wajib pernah dikutip di dalam naskah. Gunakan satu gaya sitasi secara konsisten (misal: APA Style).
  - **Urutan & Format**: Diurutkan secara **alfabetis (A–Z)** berdasarkan nama belakang penulis pertama, menggunakan format **Hanging Indent** (baris pertama di batas margin kiri, baris kedua dst. menjorok **1,25 cm**) dan spasi **1,0 (tunggal)**.
