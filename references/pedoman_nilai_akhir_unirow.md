# Pedoman Nilai Akhir Resmi UNIROW

Sumber: Buku Pedoman Akademik UNIROW (edisi terbaru yang tersedia), Bab II
"Kurikulum dan Pembelajaran", Bagian I "Penilaian", sub-bagian 1 "Nilai
Akhir" (± halaman 15).

**Gunakan isi file ini apa adanya, kata demi kata, setiap kali menyusun
bagian "Pedoman Nilai Akhir" pada dokumen rubrik penilaian.** Jangan
mengarang ulang rumus atau tabel dari ingatan, jangan mengganti dengan skala
generik (A/A-/B+/B/B-/C+/C/D/E) — itu BUKAN skala UNIROW. Jika pengguna
mengunggah edisi Buku Pedoman Akademik yang lebih baru dan isinya berbeda
dari file ini, ikuti dokumen yang diunggah dan tawarkan untuk memperbarui
file rujukan ini.

## 1. Rumus Nilai Akhir

Nilai Akhir (NA) diperoleh dari Penilaian Presensi (P), Tugas/Praktikum
(TGS), Ujian Tengah Semester (UTS), dan Ujian Akhir Semester (UAS), dengan
ketentuan perhitungan:

```
NA = (1×P + 2×TGS + 3×UTS + 4×UAS) / 10
```

Tuliskan di dokumen sebagai: **"NA = (1P + 2TGS + 3UTS + 4UAS) / 10"**

## 2. Distribusi Nilai Huruf

| Interval Nilai   | Nilai Huruf | Nilai Mutu |
|-------------------|:-----------:|:----------:|
| 85 < NA ≤ 100      | A           | 4          |
| 77,5 < NA ≤ 85     | AB          | 3,5        |
| 70 < NA ≤ 77,5     | B           | 3          |
| 65 < NA ≤ 70       | BC          | 2,5        |
| 55 < NA ≤ 65       | C           | 2          |
| 45 < NA ≤ 55       | D           | 1          |
| 0 < NA ≤ 45        | E           | 0          |

7 tingkat huruf (A, AB, B, BC, C, D, E) — bukan 9 tingkat seperti skala
umum. Jangan menambah atau mengurangi baris.

## 3. Kelulusan Mata Kuliah

1. Mahasiswa dinyatakan **lulus** mata kuliah jika mendapat nilai A, AB, B,
   BC, atau C.
2. Nilai **D** dinyatakan tidak lulus. Mahasiswa dengan nilai D dapat
   mengulang perkuliahan dengan kehadiran minimal 50% dan mengikuti ujian
   pada Ujian Akhir Semester (UAS) sesuai dengan ketentuan.
3. Nilai **E** dinyatakan tidak lulus, dan mahasiswa wajib mengikuti
   perkuliahan pada semester berikutnya sesuai ketentuan.

## Catatan implementasi di dokumen docx

Struktur bagian ini di dokumen rubrik (lihat `assets/build_rubrik_template.js`,
fungsi `buildPedomanNilaiAkhir`):

- **a. Perhitungan Nilai Akhir** — kalimat pengantar + rumus (paragraf rata
  tengah, bold).
- **b. Distribusi Nilai Huruf** — tabel 3 kolom (Interval Nilai, Nilai
  Huruf, Nilai Mutu), 7 baris data + 1 header.
- **c. Kelulusan Mata Kuliah** — tiga butir bernomor persis seperti di atas.
- Catatan kaki miring: *"Ketentuan nilai akhir mengikuti Buku Pedoman
  Akademik UNIROW (Bab II, Bagian I. Penilaian)."*
