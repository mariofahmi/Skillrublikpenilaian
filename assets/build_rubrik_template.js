/**
 * Template pembuat dokumen "Rubrik Penilaian Mata Kuliah" (UNIROW).
 *
 * CARA PAKAI: salin file ini ke folder kerja, isi objek CONFIG di bawah
 * sesuai mata kuliah yang sedang dikerjakan (identitas, peta instrumen,
 * kriteria rubrik per kelompok asesmen), lalu jalankan:
 *
 *     node build_rubrik_template.js
 *
 * Bagian "PEDOMAN NILAI AKHIR" (rumus NA, tabel huruf, aturan kelulusan)
 * SUDAH DIISI TETAP sesuai references/pedoman_nilai_akhir_unirow.md —
 * JANGAN diubah kecuali pengguna mengunggah edisi Buku Pedoman Akademik
 * UNIROW yang lebih baru dan isinya memang berbeda.
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, BorderStyle, VerticalAlign,
  PageBreak,
} = require("docx");

// ============================================================
// 1. CONFIG — edit bagian ini untuk tiap mata kuliah
// ============================================================

const CONFIG = {
  outputFile: "Rubrik_Penilaian_<Nama_MK>.docx",

  identitas: {
    universitas: "UNIVERSITAS PGRI RONGGOLAWE TUBAN",
    fakultas: "FAKULTAS ...",              // isi sesuai RPS
    prodi: "PROGRAM STUDI ...",            // isi sesuai RPS
    mataKuliah: "Nama Mata Kuliah",
    kodeMK: "KODE000",
    sks: "2 SKS",
    semester: "1",
    tglPenyusunan: "01 Januari 2026",
    dosenPengampu: "Nama Dosen, M.Pd.",
    rumpunMK: "MKK ...",
  },

  // Peta instrumen penilaian & bobot — WAJIB total 100%.
  // Ambil langsung dari tabel rencana mingguan RPS (kolom Teknik & Bobot).
  instrumen: [
    // { nama: "Kuis (lisan/tertulis) — N kali", pelaksanaan: "Minggu 1, 2, ...", bobot: "30%" },
  ],

  // Kelompok rubrik analitik 4 pita (Sangat Baik/Baik/Cukup/Kurang).
  // Buat satu grup per JENIS instrumen yang kriterianya berbeda (mis. Kuis,
  // Tugas Paper/Analisis, Observasi/Etnografi, Project, UTS/UAS, dst).
  // Turunkan "name" & isi levels dari kolom "Indikator" dan "Teknik &
  // Kriteria" pada tabel mingguan RPS — jangan mengarang kriteria yang
  // tidak berdasar pada RPS.
  rubrikGroups: [
    // {
    //   judul: "Rubrik Penilaian Kuis (Lisan/Tertulis)",
    //   catatan: "Digunakan pada Minggu ..., mengukur Sub-CPMK ranah ...",
    //   includeBobot: true, // false jika kriteria tidak berbobot per aspek
    //   criteria: [
    //     {
    //       name: "Ketepatan Konsep",
    //       bobot: "50%",
    //       levels: {
    //         "Sangat Baik": "...",
    //         "Baik": "...",
    //         "Cukup": "...",
    //         "Kurang": "...",
    //       },
    //     },
    //   ],
    // },
  ],

  pengesahan: {
    kiri: {
      jabatan: ["Menyetujui,", "Ketua Program Studi ..."],
      nama: "Nama Ketua Prodi, M.Pd.",
      nidn: "....................",
    },
    kanan: {
      jabatan: ["Dosen Pengembang RPS,"],
      nama: "Nama Dosen Pengembang, M.Pd.",
      nidn: "....................",
    },
  },
};

// ============================================================
// 2. Helper umum (tidak perlu diubah)
// ============================================================

const FONT = "Times New Roman";
const HEADER_SHADE = "D9E2F3";
const SUBHEADER_SHADE = "F2F2F2";
const BAND_COLORS = { "Sangat Baik": "E2EFDA", "Baik": "DDEBF7", "Cukup": "FFF2CC", "Kurang": "FCE4E4" };
// Lebar halaman A4 potret, margin 1134 DXA tiap sisi -> lebar guna ~9639 DXA.
const USABLE_WIDTH = 9639;

function p(text, opts = {}) {
  const { bold = false, italic = false, size = 22, align = AlignmentType.LEFT, spacingAfter = 120 } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { after: spacingAfter },
    children: [new TextRun({ text, bold, italics: italic, size, font: FONT })],
  });
}

function multiRunPara(runs, opts = {}) {
  const { align = AlignmentType.LEFT, spacingAfter = 120 } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { after: spacingAfter },
    children: runs.map(r => new TextRun({ font: FONT, size: 22, ...r })),
  });
}

function cell(text, opts = {}) {
  const {
    width, bold = false, shade = null, align = AlignmentType.LEFT, size = 20,
    verticalAlign = VerticalAlign.CENTER, colSpan = undefined, lines = null,
  } = opts;
  const paragraphs = lines
    ? lines.map(l => new Paragraph({ alignment: align, spacing: { after: 40 }, children: [new TextRun({ text: l, bold, size, font: FONT })] }))
    : [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, size, font: FONT })] })];
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: shade ? { fill: shade, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign, columnSpan: colSpan,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    children: paragraphs,
  });
}

function sectionTitle(num, title) {
  return new Paragraph({
    spacing: { before: 300, after: 150 },
    children: [new TextRun({ text: `${num}. ${title}`, bold: true, size: 24, font: FONT })],
  });
}

function sectionNote(text) {
  return new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text, italics: true, size: 20, font: FONT })] });
}

// Tabel rubrik analitik 4 pita — lebar kolom otomatis mengisi USABLE_WIDTH
// agar tidak pernah overflow ke luar margin halaman (lihat catatan di
// SKILL.md soal ini).
function buildRubricTable(criteria, includeBobot = true) {
  const kriteriaW = 1500;
  const bobotW = includeBobot ? 600 : 0;
  const levelW = Math.floor((USABLE_WIDTH - kriteriaW - bobotW) / 4);

  const headerCells = [cell("Kriteria", { width: kriteriaW, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER, size: 18 })];
  if (includeBobot) headerCells.push(cell("Bobot", { width: bobotW, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER, size: 18 }));
  ["Sangat Baik\n(85–100)", "Baik\n(70–84)", "Cukup\n(60–69)", "Kurang\n(<60)"].forEach(h => {
    headerCells.push(cell(h, { width: levelW, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER, size: 17, lines: h.split("\n") }));
  });
  const rows = [new TableRow({ children: headerCells, tableHeader: true })];

  criteria.forEach(c => {
    const rowCells = [cell(c.name, { width: kriteriaW, bold: true, shade: SUBHEADER_SHADE, size: 17 })];
    if (includeBobot) rowCells.push(cell(c.bobot, { width: bobotW, align: AlignmentType.CENTER, size: 17 }));
    ["Sangat Baik", "Baik", "Cukup", "Kurang"].forEach(level => {
      rowCells.push(cell(c.levels[level], { width: levelW, shade: BAND_COLORS[level], size: 15 }));
    });
    rows.push(new TableRow({ children: rowCells }));
  });

  const colWidths = includeBobot ? [kriteriaW, bobotW, levelW, levelW, levelW, levelW] : [kriteriaW, levelW, levelW, levelW, levelW];
  return new Table({ width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: colWidths, rows });
}

// ============================================================
// 3. Bagian TETAP — Pedoman Nilai Akhir resmi UNIROW
//    (lihat references/pedoman_nilai_akhir_unirow.md — JANGAN diubah
//    tanpa dasar dokumen resmi yang lebih baru)
// ============================================================

function buildPedomanNilaiAkhir(sectionNum) {
  const out = [];
  out.push(sectionTitle(sectionNum, "Pedoman Nilai Akhir"));

  out.push(multiRunPara([{ text: "a. Perhitungan Nilai Akhir", bold: true }], { spacingAfter: 100 }));
  out.push(p(
    "Nilai Akhir (NA) diperoleh dari Penilaian Presensi (P), Tugas/Praktikum (TGS), " +
    "Ujian Tengah Semester (UTS), dan Ujian Akhir Semester (UAS), dengan ketentuan perhitungan " +
    "sebagai berikut:",
    { spacingAfter: 150 }
  ));
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 250 },
    children: [new TextRun({ text: "NA = (1P + 2TGS + 3UTS + 4UAS) / 10", bold: true, size: 24, font: FONT })],
  }));

  out.push(multiRunPara([{ text: "b. Distribusi Nilai Huruf", bold: true }], { spacingAfter: 100 }));
  const konvHeader = new TableRow({
    children: [
      cell("Interval Nilai", { width: 3000, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
      cell("Nilai Huruf", { width: 2600, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
      cell("Nilai Mutu", { width: 2600, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
    ],
    tableHeader: true,
  });
  const konvRows = [
    ["85 < NA ≤ 100", "A", "4"],
    ["77,5 < NA ≤ 85", "AB", "3,5"],
    ["70 < NA ≤ 77,5", "B", "3"],
    ["65 < NA ≤ 70", "BC", "2,5"],
    ["55 < NA ≤ 65", "C", "2"],
    ["45 < NA ≤ 55", "D", "1"],
    ["0 < NA ≤ 45", "E", "0"],
  ].map(([a, b, c]) => new TableRow({
    children: [
      cell(a, { width: 3000, align: AlignmentType.CENTER, size: 19 }),
      cell(b, { width: 2600, align: AlignmentType.CENTER, size: 19, bold: true }),
      cell(c, { width: 2600, align: AlignmentType.CENTER, size: 19 }),
    ],
  }));
  out.push(new Table({ width: { size: 8200, type: WidthType.DXA }, columnWidths: [3000, 2600, 2600], rows: [konvHeader, ...konvRows] }));

  out.push(p("", { spacingAfter: 150 }));
  out.push(multiRunPara([{ text: "c. Kelulusan Mata Kuliah", bold: true }], { spacingAfter: 100 }));
  out.push(p("1. Mahasiswa dinyatakan lulus mata kuliah jika mendapat nilai A, AB, B, BC, atau C.", { spacingAfter: 80 }));
  out.push(p(
    "2. Nilai D dinyatakan tidak lulus. Mahasiswa dengan nilai D dapat mengulang perkuliahan dengan " +
    "kehadiran minimal 50% dan mengikuti ujian pada Ujian Akhir Semester (UAS) sesuai dengan ketentuan.",
    { spacingAfter: 80 }
  ));
  out.push(p(
    "3. Nilai E dinyatakan tidak lulus, dan mahasiswa wajib mengikuti perkuliahan pada semester " +
    "berikutnya sesuai ketentuan.",
    { spacingAfter: 100 }
  ));
  out.push(sectionNote("Ketentuan nilai akhir mengikuti Buku Pedoman Akademik UNIROW (Bab II, Bagian I. Penilaian)."));
  return out;
}

// ============================================================
// 4. Rakit dokumen
// ============================================================

function build() {
  const { identitas, instrumen, rubrikGroups, pengesahan } = CONFIG;
  const children = [];

  // Kop & judul
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: identitas.universitas, bold: true, size: 24, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: identitas.fakultas, bold: true, size: 22, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: identitas.prodi, bold: true, size: 22, font: FONT })] }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 60 },
    border: { bottom: { color: "auto", space: 1, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text: "", size: 2 })],
  }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300, after: 60 }, children: [new TextRun({ text: "RUBRIK PENILAIAN MATA KULIAH", bold: true, size: 30, font: FONT })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: identitas.mataKuliah.toUpperCase(), bold: true, size: 30, font: FONT })] }));

  // Identitas
  const idRows = [
    ["Mata Kuliah", identitas.mataKuliah, "Kode MK", identitas.kodeMK],
    ["Program Studi", identitas.prodi.replace(/^PROGRAM STUDI\s*/i, ""), "SKS", identitas.sks],
    ["Semester", identitas.semester, "Tgl Penyusunan", identitas.tglPenyusunan],
    ["Dosen Pengampu", identitas.dosenPengampu, "Rumpun MK", identitas.rumpunMK],
  ];
  children.push(new Table({
    width: { size: 9600, type: WidthType.DXA },
    columnWidths: [1800, 3600, 1800, 2400],
    rows: idRows.map(([l1, v1, l2, v2]) => new TableRow({
      children: [
        cell(l1, { width: 1800, bold: true, shade: SUBHEADER_SHADE, size: 20 }),
        cell(v1, { width: 3600, size: 20 }),
        cell(l2, { width: 1800, bold: true, shade: SUBHEADER_SHADE, size: 20 }),
        cell(v2, { width: 2400, size: 20 }),
      ],
    })),
  }));

  children.push(p("", { spacingAfter: 100 }));
  children.push(sectionNote(
    "Dokumen ini merupakan elaborasi rubrik penilaian dari RPS mata kuliah " + identitas.mataKuliah +
    ". Rubrik disusun per jenis instrumen penilaian dengan empat pita capaian: Sangat Baik, Baik, " +
    "Cukup, dan Kurang, agar selaras dengan prinsip penilaian berbasis kriteria (CPL/CPMK/Sub-CPMK)."
  ));

  // 1. Peta instrumen & bobot
  let secNum = 1;
  children.push(sectionTitle(secNum++, "Peta Instrumen Penilaian dan Bobot"));
  const bobotHeader = new TableRow({
    children: [
      cell("Instrumen Penilaian", { width: 4200, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
      cell("Pelaksanaan", { width: 2600, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
      cell("Bobot", { width: 1400, bold: true, shade: HEADER_SHADE, align: AlignmentType.CENTER }),
    ],
    tableHeader: true,
  });
  const bobotBody = instrumen.map(i => new TableRow({
    children: [
      cell(i.nama, { width: 4200, size: 19 }),
      cell(i.pelaksanaan, { width: 2600, size: 19, align: AlignmentType.CENTER }),
      cell(i.bobot, { width: 1400, size: 19, align: AlignmentType.CENTER, bold: true }),
    ],
  }));
  // Validasi total bobot = 100% — peringatkan di konsol jika tidak, jangan
  // diam-diam melanjutkan (bobot yang tidak total 100% adalah tanda RPS
  // belum lengkap dibaca dengan benar).
  const totalBobot = instrumen.reduce((sum, i) => sum + (parseFloat(String(i.bobot).replace("%", "").replace(",", ".")) || 0), 0);
  if (Math.round(totalBobot) !== 100) {
    console.warn(`PERINGATAN: total bobot instrumen = ${totalBobot}%, seharusnya 100%. Periksa kembali CONFIG.instrumen.`);
  }
  children.push(new Table({
    width: { size: 8200, type: WidthType.DXA },
    columnWidths: [4200, 2600, 1400],
    rows: [bobotHeader, ...bobotBody, new TableRow({
      children: [
        cell("TOTAL", { width: 4200, bold: true, shade: SUBHEADER_SHADE, colSpan: 2 }),
        cell(`${totalBobot}%`, { width: 1400, bold: true, shade: SUBHEADER_SHADE, align: AlignmentType.CENTER }),
      ],
    })],
  }));

  // 2..N. Rubrik per kelompok
  rubrikGroups.forEach((g, idx) => {
    if (idx > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(sectionTitle(secNum++, g.judul));
    if (g.catatan) children.push(sectionNote(g.catatan));
    children.push(buildRubricTable(g.criteria, g.includeBobot !== false));
  });

  // Pedoman Nilai Akhir (tetap)
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(...buildPedomanNilaiAkhir(secNum++));

  // Lembar Pengesahan
  children.push(p("", { spacingAfter: 200 }));
  children.push(sectionTitle(secNum++, "Lembar Pengesahan"));
  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  children.push(new Table({
    width: { size: 9600, type: WidthType.DXA },
    columnWidths: [4800, 4800],
    borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
    rows: [
      new TableRow({
        children: [
          cell("", { width: 4800, align: AlignmentType.CENTER, lines: pengesahan.kiri.jabatan }),
          cell("", { width: 4800, align: AlignmentType.CENTER, lines: pengesahan.kanan.jabatan }),
        ],
      }),
      new TableRow({
        children: [
          cell("", { width: 4800, align: AlignmentType.CENTER, lines: ["", "", "", pengesahan.kiri.nama, `NIDN. ${pengesahan.kiri.nidn}`] }),
          cell("", { width: 4800, align: AlignmentType.CENTER, lines: ["", "", "", pengesahan.kanan.nama, `NIDN. ${pengesahan.kanan.nidn}`] }),
        ],
      }),
    ],
  }));

  const doc = new Document({
    sections: [{
      properties: { page: { size: { width: 11907, height: 16840 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      children,
    }],
  });

  Packer.toBuffer(doc).then(buf => {
    require("fs").writeFileSync(CONFIG.outputFile, buf);
    console.log("done:", CONFIG.outputFile);
  });
}

build();
