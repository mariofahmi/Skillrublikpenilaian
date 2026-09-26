---
name: rubrik-penilaian-unirow
description: >
  Skill resmi penyusunan Rubrik Penilaian analitik 4 pita dan pedoman nilai akhir UNIROW berbasis RPS kurikulum OBE 2026 secara otomatis. Aktif saat mengetik /rubrik-penilaian-unirow atau menyebut rubrik penilaian.
---

# RUBRIK PENILAIAN DARI RPS — PANDUAN OTOMATISASI RESMI

Skill ini **secara otomatis** menyusun, memetakan, dan menerbitkan dokumen resmi **Rubrik Penilaian Mata Kuliah** berformat Word (`.docx`) dan Markdown (`.md`) yang diturunkan langsung dari dokumen Rencana Pembelajaran Semester (RPS) kurikulum Outcome-Based Education (OBE) Universitas PGRI Ronggolawe (UNIROW) Tuban.

---

## ⚡ Eksekusi Otomatis (Automatic One-Click Generation)

Ketika pengguna mengetik `/rubrik-penilaian-unirow` atau meminta pembuatan rubrik penilaian untuk suatu RPS:
**Agen WAJIB langsung mengeksekusi script generator otomatis** menggunakan Python:

```bash
# Otomatis untuk satu berkas RPS tertentu:
py ".agents/skills/rubrik-penilaian-unirow/scripts/build_rubrik.py" "path/ke/RPS.docx"

# Atau jika ingin menentukan folder penyimpanan output:
py ".agents/skills/rubrik-penilaian-unirow/scripts/build_rubrik.py" "path/ke/RPS.docx" --output "01_FIK/RPS OBE 2026/RUBRIK_PENILAIAN"

# Otomatis memproses seluruh dokumen RPS dalam satu folder:
py ".agents/skills/rubrik-penilaian-unirow/scripts/build_rubrik.py" --batch
```

Script generator di atas bekerja secara **100% otonom**:
1. **Auto-Extract**: Membaca otomatis file RPS (.docx maupun .md), mengambil Nama MK, Kode, SKS, Semester, Dosen, Kaprodi, dan matriks 16 minggu.
2. **Auto-Map & 100% Weight**: Mengelompokkan aktivitas evaluasi menjadi 5 kelompok instrumen terstandar dengan bobot kumulatif tepat 100%.
3. **Auto-Elaborate 4 Bands**: Mengembangkan indikator capaian ke dalam 4 pita analitik (Sangat Baik [85–100], Baik [70–84], Cukup [60–69], Kurang [<60]) dengan pewarnaan resmi dan lebar tabel terkunci ($9620\text{ dxa}$).
4. **Auto-Rule Pedoman Nilai Akhir**: Menegakkan rumus baku NA, tabel 7 skala huruf mutu UNIROW, dan klausul kelulusan resmi.
5. **Auto-Sign**: Menyertakan lembar pengesahan resmi Kaprodi PPKn (**Mario Fahmi Syahrial, M.Pd.**) dan Dosen Pengampu.

---

## 📋 Spesifikasi Format Dokumen Baku

| Parameter | Ketentuan Mutu Standar Kampus |
|---|---|
| **Format Kertas** | **A4 Portrait** (Margin terkunci $1134\text{ dxa}$ / $0.8\text{"}$) |
| **Tipografi** | **Times New Roman** (Judul 14pt bold, Subjudul 11pt, Isi & Tabel 8.5–9.5pt) |
| **KOP Resmi Institusi** | Memuat Universitas PGRI Ronggolawe Tuban, FKIP, Prodi PPKn, dan garis pembatas |
| **Garis Tabel & Margin** | Garis abu-abu `#D0D5DD` / biru navy header `#002060`, cell margin $70\text{ dxa}$, `<w:cantSplit/>`, `<w:tblHeader/>` |
| **Palet Shading 4 Pita** | Sangat Baik (`#E2EFDA`), Baik (`#DDEBF7`), Cukup (`#FFF2CC`), Kurang (`#FCE4E4`) |
| **Rumus Nilai Akhir** | **NA = (1P + 2TGS + 3UTS + 4UAS) / 10** |
| **Skala Huruf Mutu** | **7 Tingkat**: A (4), AB (3.5), B (3), BC (2.5), C (2), D (1), E (0) |
