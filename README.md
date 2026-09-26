# Skill: Rubrik Penilaian OBE UNIROW Tuban

[![Antigravity Skill](https://img.shields.io/badge/Antigravity-Skill-blue.svg)](https://github.com/mariofahmi/Skillrublikpenilaian)
[![Kurikulum](https://img.shields.io/badge/Kurikulum-OBE%202026-green.svg)](https://unirow.ac.id)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Skill resmi otomatisasi penyusunan **Rubrik Penilaian Analitik 4 Pita** dan **Pedoman Nilai Akhir** standar mutu akademik Universitas PGRI Ronggolawe (UNIROW) Tuban berbasis kurikulum Outcome-Based Education (OBE) 2026.

---

## 🌟 Fitur Utama

1. **Auto-Extract Dokumen RPS**:
   - Membaca dokumen RPS (format Word `.docx` maupun Markdown `.md`).
   - Otomatis mengekstrak Identitas Mata Kuliah (Nama, Kode, SKS, Semester, Rumpun MK).
   - Memprioritaskan **Dosen Pengembang RPS** definitif dari tabel otorisasi sebagai Dosen Pengampu & Penandatangan.
2. **Peta 5 Kelompok Instrumen Penilaian (Kumulatif 100%)**:
   - Kuis Pemahaman Teori & Konsep Berkala: **25%**
   - Tugas Terstruktur (Paper Analisis & Studi Kasus): **15%**
   - Proyek Pembelajaran Berbasis Masalah / Aksi Nyata (PjBL/PBL): **10%**
   - Ujian Tengah Semester (UTS): **25%**
   - Ujian Akhir Semester (UAS): **25%**
3. **Rubrik Analitik 4 Pita Bergradasi Warna**:
   - **Sangat Baik (85–100)**: `#E2EFDA` (Hijau Muda)
   - **Baik (70–84)**: `#DDEBF7` (Biru Muda)
   - **Cukup (60–69)**: `#FFF2CC` (Kuning Muda)
   - **Kurang (<60)**: `#FCE4E4` (Merah Muda)
   - Proteksi tabel Word XML: `<w:cantSplit/>`, `<w:tblHeader/>`, dan lebar kolom terkunci ($9620\text{ dxa}$).
4. **Penegakan Pedoman Nilai Akhir Resmi UNIROW**:
   - Rumus Nilai Akhir:
     $$\text{NA} = \frac{1\text{P} + 2(\text{TGS}) + 3(\text{UTS}) + 4(\text{UAS})}{10}$$
   - Distribusi 7 Skala Huruf Mutu: **A (4.0)**, **AB (3.5)**, **B (3.0)**, **BC (2.5)**, **C (2.0)**, **D (1.0)**, **E (0.0)**.
   - Aturan kelulusan mata kuliah resmi sesuai Buku Pedoman Akademik UNIROW.
5. **Format Ganda Siap Pakai**:
   - Dokumen Word (`.docx`) ber-KOP resmi institusi.
   - Dokumen Markdown (`.md`) rapi dan terstandarisasi.
6. **Otorisasi Resmi**:
   - Menyetujui: Ketua Program Studi PPKn **Mario Fahmi Syahrial, M.Pd.**
   - Dosen Pengembang RPS / Dosen Pengampu mata kuliah.

---

## 📂 Struktur Direktori Skill

```text
Skillrublikpenilaian/
├── SKILL.md                             # Panduan dan instruksi utama agen AI
├── README.md                            # Dokumentasi repositori
├── .gitignore                           # Filter file binary/cache
├── assets/
│   └── build_rubrik_template.js         # Template referensi perancangan
├── references/
│   └── pedoman_nilai_akhir_unirow.md    # Ekstrak aturan penilaian resmi UNIROW
└── scripts/
    └── build_rubrik.py                  # Script Python generator otomatis
```

---

## 🚀 Cara Instalasi

### 1. Instalasi di Workspace Antigravity
Salin folder skill ke dalam folder `.agents/skills/` pada workspace Anda:
```bash
# Clone ke workspace
git clone https://github.com/mariofahmi/Skillrublikpenilaian.git .agents/skills/rubrik-penilaian-unirow
```

### 2. Instalasi Global Antigravity
Agar dapat digunakan di seluruh workspace komputer Anda:
```bash
git clone https://github.com/mariofahmi/Skillrublikpenilaian.git "%USERPROFILE%\.gemini\config\skills\rubrik-penilaian-unirow"
```

---

## 🛠️ Penggunaan CLI

### 1. Menghasilkan Rubrik dari Berkas RPS Tunggal
```bash
# Menghasilkan .docx dan .md di folder yang sama dengan file RPS
py scripts/build_rubrik.py "path/ke/RPS_Mata_Kuliah.docx"

# Menentukan folder output khusus:
py scripts/build_rubrik.py "path/ke/RPS_Mata_Kuliah.docx" -o "folder_tujuan"
```

### 2. Pemrosesan Seluruh RPS dalam Satu Folder (Batch Mode)
```bash
py scripts/build_rubrik.py --batch -o "RUBRIK_PENILAIAN"
```

---

## 🏛️ Institusi
- **Universitas**: Universitas PGRI Ronggolawe (UNIROW) Tuban
- **Fakultas**: Keguruan dan Ilmu Pendidikan (FKIP)
- **Program Studi**: Pendidikan Pancasila dan Kewarganegaraan (PPKn)
- **Koordinator Kurikulum / Kaprodi**: Mario Fahmi Syahrial, M.Pd.

---

## 📄 Lisensi
Hak Cipta (c) 2026 Program Studi PPKn, FKIP UNIROW Tuban.
Didistribusikan di bawah lisensi [MIT](LICENSE).
