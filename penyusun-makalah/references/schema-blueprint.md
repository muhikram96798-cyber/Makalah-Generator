# Skema `blueprint.json` & Konvensi Markdown Naskah Makalah

Dokumen ini mendefinisikan format data `blueprint.json` serta aturan penulisan file Markdown di dalam folder `drafts/` yang dibaca oleh kompilator `scripts/buat_makalah.py`.

---

## 1. Struktur Direktori Workspace Makalah

```text
./makalah_workspace/<slug-judul>/
├── blueprint.json              # Metadata sampul, daftar bab, dan daftar pustaka
├── references_candidates.json  # Kandidat sumber + URL/DOI sebelum diuji
├── references_verified.json    # Hasil uji keterbukaan dari cek_referensi.py
├── references_notes.json       # Ekstraksi inti teks PDF/HTML tiap jurnal terbuka
├── assets/                     # Menyimpan gambar naskah & logo universitas
└── drafts/
    ├── kata_pengantar.md       # Naskah Kata Pengantar
    ├── bab_01.md               # BAB I PENDAHULUAN
    ├── bab_02.md               # BAB II PEMBAHASAN
    └── bab_03.md               # BAB III PENUTUP
```

### Format `references_candidates.json`, `references_verified.json`, & `references_notes.json`
File `references_candidates.json` berisi array objek `[{"citation": "...", "url": "https://doi.org/..."}]`. Setelah diuji oleh `scripts/cek_referensi.py`, hasil tersimpan ke `references_verified.json` (`status`: `OPEN_ACCESS` / `PAYWALLED` / `UNREACHABLE`, `pdf_url`) dan jika dijalankan dengan `--extract-notes`, inti teks PDF/HTML masing-masing jurnal `OPEN_ACCESS` tersimpan ke `references_notes.json` untuk menyusun poin-poin pembahasan. Hanya referensi berstatus `OPEN_ACCESS` yang boleh dimasukkan ke `blueprint.json`.

---

## 2. Skema Lengkap `blueprint.json`

```json
{
  "title": "IMPLEMENTASI KECERDASAN BUATAN DALAM OPTIMALISASI RANTAI PASOK INDUSTRI MANUFAKTUR DI INDONESIA",
  "subtitle": "Makalah Ini Disusun untuk Memenuhi Tugas pada Mata Kuliah Manajemen Operasional",
  "course": "Manajemen Operasional",
  "lecturer": "Prof. Dr. Ir. Hendra Wijaya, M.T., IPU.",
  "authors": [
    {
      "name": "Budi Santoso",
      "id": "2301010045"
    },
    {
      "name": "Siti Aminah",
      "id": "2301010052"
    }
  ],
  "department": "Program Studi Teknik Industri",
  "faculty": "Fakultas Teknik",
  "institution": "Universitas Nusantara",
  "city": "Jakarta",
  "year": "2026",
  "academic_year": "2025/2026",
  "logo_path": "",
  "numbering_style": "decimal",
  "margin_preset": "4-3-3-3",
  "include_abstract": false,
  "abstract_file": "drafts/abstrak.md",
  "preface_date": "Jakarta, November 2026",
  "preface_file": "drafts/kata_pengantar.md",
  "problem_statements": [
    "Bagaimana arsitektur penerapan kecerdasan buatan pada manajemen rantai pasok industri manufaktur di Indonesia?",
    "Apa saja kendala struktural dan teknis dalam adopsi kecerdasan buatan pada rantai pasok industri manufaktur nasional?",
    "Bagaimana strategi optimalisasi rantai pasok berbasis kecerdasan buatan guna meningkatkan efisiensi operasional perusahaan?"
  ],
  "chapters": [
    {
      "number": "BAB I",
      "title": "PENDAHULUAN",
      "file": "drafts/bab_01.md",
      "status": "completed"
    },
    {
      "number": "BAB II",
      "title": "PEMBAHASAN",
      "file": "drafts/bab_02.md",
      "status": "completed"
    },
    {
      "number": "BAB III",
      "title": "PENUTUP",
      "file": "drafts/bab_03.md",
      "status": "completed"
    }
  ],
  "references": [
    "Christopher, M. (2022). *Logistics & supply chain management* (6th ed.). Pearson Education.",
    "Kementerian Perindustrian Republik Indonesia. (2024). *Laporan transformasi digital industri manufaktur nasional Making Indonesia 4.0*. Kemenperin RI.",
    "Pratama, R. A., & Nugroho, D. S. (2023). Integrasi *machine learning* dalam peramalan permintaan rantai pasok manufaktur. *Jurnal Teknik Industri Indonesia*, 18(2), 112-125."
  ]
}
```

### Catatan Atribut `blueprint.json` & Checklist Komponen Sampul Wajib via Chat:
Sebelum kompilasi final, agen wajib meminta pengguna mengisi sendiri seluruh komponen sampul melalui chat (kecuali sudah diberikan di awal) dan memastikan tidak ada nilai placeholder `[...]`:
1. `title` (Judul Makalah), `course` (Mata Kuliah; `subtitle` kosong/`""` secara *default* dan hanya dirender jika diisi eksplisit), `lecturer` (Nama Dosen Pengampu + Gelar).
2. `authors` (`name` & `id`/NIM; penulis tunggal otomatis dicetak dua baris `Nama` / `NIM. xxx`).
3. `department` (Program Studi), `faculty` (Fakultas), `institution` (Perguruan Tinggi), `city` (Kota — otomatis dideduplikasi jika akhir `institution` sudah menyebut kota tersebut), serta `year` & `academic_year`.
4. `logo_path` & tata letak sampul: Arahkan pengguna meletakkan foto logo kampus di `assets/logo_kampus.png` (maksimal `4,5 cm` dengan rasio asli terjaga); halaman sampul otomatis memakai margin simetris `3,0 cm` kiri–kanan agar tepat di tengah kertas A4, sementara bagian isi tetap mengikuti `margin_preset` (`4-3-3-3`).
- `problem_statements`: Daftar pertanyaan pada Rumusan Masalah. Digunakan oleh *Academic Linter* untuk memverifikasi sinkronisasi jumlah butir Rumusan Masalah (Bab I) dengan Kesimpulan (Bab III).
- `references`: Array string berisi entri Daftar Pustaka (**wajib di atas 10 / minimal 11 sumber `OPEN_ACCESS` terverifikasi**). Gunakan tanda `*...*` pada judul buku atau nama jurnal agar dicetak miring (*italic*) secara otomatis oleh kompilator. Kompilator akan otomatis mengurutkannya secara alfabetis (A–Z).

---

## 3. Konvensi Sintaks Markdown di Dalam `drafts/*.md`

1. **Jangan Tulis Ulang Judul Bab Utama di Baris Pertama**:
   - Judul `BAB I PENDAHULUAN` sudah otomatis dibuat oleh `buat_makalah.py` dari metadata `chapters` di `blueprint.json`. Langsung mulai file `bab_01.md` dengan sub-bab pertama (`## 1.1 Latar Belakang`).
2. **Hierarki Sub-bab (Maksimal 3 Tingkat)**:
   - Tingkat 2 (Sub-bab): `## 1.1 Latar Belakang` atau `## 2.1 Konsep Rantai Pasok`
   - Tingkat 3 (Anak Sub-bab): `### 2.1.1 Komponen Utama Rantai Pasok`
   - **Dilarang** menggunakan `####` (Tingkat 4+ seperti `2.1.1.1`).
3. **Paragraf Isi & Istilah Asing**:
   - Pisahkan antar-paragraf dengan 1 baris kosong.
   - Setiap paragraf berisi **3–7 kalimat** (1 ide pokok).
   - Cetak miring istilah asing dengan satu tanda bintang: `*supply chain management*`, `*artificial intelligence*`.
4. **Kutipan Langsung Panjang ($\ge 4$ Baris)**:
   - Gunakan blok `>` di awal baris (tanpa tanda petik dua `"..."`):
     ```markdown
     > Transformasi rantai pasok berbasis algoritma prediktif menuntut integrasi data lintas departemen secara waktu nyata, sehingga setiap distorsi permintaan di tingkat hilir dapat langsung dikompensasi oleh penjadwalan produksi di tingkat hulu tanpa menimbulkan efek cambuk stok yang merugikan (Christopher, 2022: 84).
     ```
   - Kompilator otomatis memformat blok `>` menjadi spasi **1,0 (tunggal)** dan menjorok kiri **1,25 cm**.
5. **Tabel Akademik (Judul di Atas, Sumber di Bawah)**:
   - Tulis baris judul tabel tepat di atas tabel Markdown dengan awalan `Tabel X.Y ...`, dan baris sumber tepat di bawah tabel dengan awalan `Sumber: ...`:
     ```markdown
     Tabel 2.1 Perbandingan Kinerja Rantai Pasok Konvensional dan Berbasis AI
     | Indikator Kinerja | Sistem Konvensional | Sistem Berbasis AI |
     | :--- | :--- | :--- |
     | Akurasi Peramalan | 68,5% | 91,2% |
     | Waktu Respons Stok | 14 hari | 3 hari |
     Sumber: Pratama & Nugroho (2023)
     ```
   - Kompilator otomatis menerapkan gaya **Tabel Standar Microsoft Word (*Full Grid* kotak penuh selebar area cetak, header bold dengan shading `#F2F2F2`, padding sel rapi, `tblHeader` berulang, dan `cantSplit`)**, memasang `keep_with_next` pada judul tabel, dan mendaftarkannya ke **Daftar Tabel**.
6. **Gambar / Ilustrasi (Judul & Sumber di Bawah)**:
   - Gunakan sintaks Markdown gambar standar diikuti baris `Gambar X.Y ...` dan `Sumber: ...`:
     ```markdown
     ![Arsitektur Sistem](assets/arsitektur.png)
     Gambar 2.1 Arsitektur Integrasi AI pada Rantai Pasok
     Sumber: Kementerian Perindustrian Republik Indonesia (2024)
     ```
   - Kompilator otomatis menaruh gambar di tengah halaman, meletakkan judul gambar & sumber tepat di bawahnya, dan mendaftarkannya ke **Daftar Gambar**.
