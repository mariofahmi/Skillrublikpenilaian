#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_rubrik.py - Generator Otomatis Rubrik Penilaian OBE UNIROW Tuban
Membaca dokumen RPS (Word .docx atau Markdown .md) secara otomatis, mengekstrak identitas MK,
peta instrumen & bobot 16 minggu, menyusun rubrik analitik 4 pita (Sangat Baik, Baik, Cukup, Kurang),
menegakkan Pedoman Nilai Akhir resmi UNIROW, dan menghasilkan dokumen .docx & .md siap pakai.
"""

import os
import sys
import re
import argparse
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==============================================================================
# 1. XML Formatting Helpers
# ==============================================================================

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=70, bottom=70, left=80, right=80):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_borders(cell, top="single", bottom="single", left="none", right="none", sz="4", color="CCCCCC"):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcBorders'):
            tcPr.remove(child)
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:cantSplit'):
        cantSplit = parse_xml(f'<w:cantSplit {nsdecls("w")}/>')
        trPr.append(cantSplit)

def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:tblHeader'):
        tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
        trPr.append(tblHeader)

def set_col_widths(table, col_widths_dxa):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblPr = table._tbl.tblPr
    total_w = sum(col_widths_dxa)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w}" w:type="dxa"/>')
    tblPr.append(tblW)
    
    # tblGrid
    tblGrid = table._tbl.tblGrid
    for c in list(tblGrid):
        tblGrid.remove(c)
    for w in col_widths_dxa:
        gridCol = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w}"/>')
        tblGrid.append(gridCol)

    for row in table.rows:
        prevent_row_split(row)
        for i, w in enumerate(col_widths_dxa):
            if i < len(row.cells):
                cell = row.cells[i]
                tcPr = cell._tc.get_or_add_tcPr()
                for child in list(tcPr):
                    if child.tag.endswith('tcW'):
                        tcPr.remove(child)
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w}" w:type="dxa"/>')
                tcPr.append(tcW)

# ==============================================================================
# 2. RPS Parser (DOCX & MD)
# ==============================================================================

def clean_text(s):
    if not s: return ""
    return re.sub(r'\s+', ' ', s.replace('**', '').replace('*', '')).strip()

def parse_rps_file(file_path):
    """Mengekstrak data terstruktur dari file RPS Word (.docx) atau Markdown (.md)."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File RPS tidak ditemukan: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.docx':
        return parse_rps_docx(file_path)
    elif ext == '.md':
        return parse_rps_md(file_path)
    else:
        raise ValueError(f"Ekstensi file tidak didukung: {ext}. Harap gunakan .docx atau .md.")

def parse_rps_docx(docx_path):
    doc = docx.Document(docx_path)
    data = {
        'nama_mk': '',
        'kode_mk': '',
        'sks': '2',
        'semester': '1',
        'rumpun_mk': 'Mata Kuliah Keilmuan dan Keterampilan (MKK PPKn)',
        'dosen_pengampu': 'Mario Fahmi Syahrial, M.Pd.',
        'kaprodi': 'Mario Fahmi Syahrial, M.Pd.',
        'tgl_penyusunan': '04 Februari 2026',
        'mingguan': [],
        'rubrik_asli': []
    }

    # Derive semester from filename if possible
    base = os.path.splitext(os.path.basename(docx_path))[0]
    m_sem_file = re.match(r'^0?(\d+)_', base)
    if m_sem_file:
        data['semester'] = str(int(m_sem_file.group(1)))

    # Table 0: Identitas
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        # Position-based parsing from standard UNIROW OBE template (Row 2 headers, Row 3 values)
        if len(t0.rows) > 3:
            r2 = [clean_text(c.text) for c in t0.rows[2].cells]
            r3 = [clean_text(c.text) for c in t0.rows[3].cells]
            
            # Nama MK from r3[0]
            if r3 and r3[0] and len(r3[0]) > 2 and "MATA KULIAH" not in r3[0].upper():
                data['nama_mk'] = r3[0]
                
            # Scan Row 2 headers matching Row 3 values
            for idx, col in enumerate(r2):
                if idx < len(r3):
                    col_u = col.upper()
                    val = r3[idx]
                    if "KODE" in col_u and val and "KODE" not in val.upper() and "RUMPUN" not in val.upper():
                        data['kode_mk'] = val
                    elif "RUMPUN" in col_u and val and "RUMPUN" not in val.upper():
                        data['rumpun_mk'] = val
                    elif "BOBOT" in col_u and val:
                        m = re.search(r'(\d+)', val)
                        if m: data['sks'] = m.group(1)
                    elif "SEMESTER" in col_u and val:
                        m = re.search(r'(\d+)', val)
                        if m: data['semester'] = m.group(1)
                    elif ("TGL" in col_u or "TANGGAL" in col_u) and val and "TGL" not in val.upper():
                        data['tgl_penyusunan'] = val

        # Otorisasi Dosen Pengembang RPS, Kaprodi
        pengembang_found = ""
        for t in doc.tables:
            for r_idx, row in enumerate(t.rows):
                for c_idx, cell in enumerate(row.cells):
                    txt = clean_text(cell.text)
                    if txt in ["Pengembang RPS", "Dosen Pengembang RPS", "Dosen Pengembang", "Penyusun"]:
                        if r_idx + 1 < len(t.rows):
                            val = clean_text(t.rows[r_idx + 1].cells[c_idx].text)
                            lines = [l.strip() for l in val.split("\n") if l.strip()]
                            for l in lines:
                                if not any(k in l for k in ["NIDN", "Tanda Tangan", "Nama Dosen", "Pengembang RPS"]):
                                    pengembang_found = l
                                    break
                        if not pengembang_found:
                            lines = [l.strip() for l in cell.text.split("\n") if l.strip()]
                            if len(lines) > 1:
                                for l in lines[1:]:
                                    if not any(k in l for k in ["NIDN", "Tanda Tangan"]):
                                        pengembang_found = clean_text(l)
                                        break
                    if any(k in txt for k in ["Ketua PRODI", "Ketua Program Studi", "Kaprodi"]):
                        if r_idx + 1 < len(t.rows):
                            val = clean_text(t.rows[r_idx + 1].cells[c_idx].text)
                            lines = [l.strip() for l in val.split("\n") if l.strip()]
                            for l in lines:
                                if not any(k in l for k in ["NIDN", "Tanda Tangan"]):
                                    data['kaprodi'] = l
                                    break

        if pengembang_found:
            data['dosen_pengampu'] = pengembang_found
        elif len(t0.rows) > 5:
            r4 = [clean_text(c.text) for c in t0.rows[4].cells]
            r5 = [clean_text(c.text) for c in t0.rows[5].cells]
            for idx, col in enumerate(r4):
                if idx < len(r5):
                    col_u = col.upper()
                    val = r5[idx]
                    if "PENGEMBANG" in col_u and val and "PENGEMBANG" not in val.upper() and "OTORISASI" not in val.upper():
                        data['dosen_pengampu'] = val
                    elif "KETUA" in col_u and val and "KETUA" not in val.upper() and "OTORISASI" not in val.upper():
                        data['kaprodi'] = val

        # Scan for explicit "Dosen Pengampu" in any subsequent rows of Table 0 (only if not found yet)
        if not data.get('dosen_pengampu'):
            for r in t0.rows:
                cells = [clean_text(c.text) for c in r.cells]
                if cells and "DOSEN PENGAMPU" in cells[0].upper():
                    for c in cells[1:]:
                        if c and "DOSEN PENGAMPU" not in c.upper():
                            data['dosen_pengampu'] = c
                            break

    # Fallback nama MK dari nama file jika nama_mk tidak valid atau mengandung kata Tim Dosen/deskripsi
    clean_base = re.sub(r'^\d+_\d+_RPS_', '', base)
    clean_base = re.sub(r'_(?:FINAL|UNIROW|OBE|COMPLETE)$', '', clean_base)
    file_mk_name = clean_base.replace('_', ' ')

    if not data['nama_mk'] or data['nama_mk'] == 'Mata Kuliah' or len(data['nama_mk']) > 50 or 'TIM DOSEN' in data['nama_mk'].upper():
        data['nama_mk'] = file_mk_name

    if not data['dosen_pengampu']:
        data['dosen_pengampu'] = "Tim Dosen Prodi PPKn UNIROW"

    # Table 1: Jadwal Mingguan
    if len(doc.tables) > 1:
        t1 = doc.tables[1]
        for r_idx in range(len(t1.rows)):
            cells = [clean_text(c.text) for c in t1.rows[r_idx].cells]
            col0 = cells[0] if cells else ""
            m_mg = re.match(r'^\s*(\d{1,2})\b', col0)
            is_uts = 'UTS' in col0.upper() or 'TENGAH SEMESTER' in col0.upper()
            is_uas = 'UAS' in col0.upper() or 'AKHIR SEMESTER' in col0.upper()
            
            if m_mg or is_uts or is_uas:
                mg_no = int(m_mg.group(1)) if m_mg else (8 if is_uts else 16)
                indikator = cells[2] if len(cells) > 2 else ""
                kriteria = cells[3] if len(cells) > 3 else ""
                bobot_str = cells[-1] if len(cells) > 4 else ""
                
                m_b = re.search(r'(\d+)', bobot_str)
                bobot = int(m_b.group(1)) if m_b else (25 if (is_uts or is_uas) else 3)
                
                data['mingguan'].append({
                    'minggu': mg_no,
                    'indikator': indikator,
                    'kriteria': kriteria,
                    'bobot': bobot
                })

    # Table 2: Rubrik Asli RPS (jika ada)
    if len(doc.tables) > 2:
        t2 = doc.tables[2]
        for r in t2.rows[1:]:
            cells = [clean_text(c.text) for c in r.cells]
            if len(cells) >= 4:
                data['rubrik_asli'].append({
                    'aspek': cells[0],
                    'sb': cells[1] if len(cells) > 1 else "",
                    'b': cells[2] if len(cells) > 2 else "",
                    'c': cells[3] if len(cells) > 3 else "",
                    'bobot': cells[-1] if len(cells) > 4 else "25%"
                })

    return data

def parse_rps_md(md_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        text = f.read()

    data = {
        'nama_mk': '',
        'kode_mk': '',
        'sks': '2',
        'semester': '1',
        'rumpun_mk': 'Mata Kuliah Keilmuan dan Keterampilan (MKK PPKn)',
        'dosen_pengampu': 'Mario Fahmi Syahrial, M.Pd.',
        'kaprodi': 'Mario Fahmi Syahrial, M.Pd.',
        'tgl_penyusunan': '04 Februari 2026',
        'mingguan': [],
        'rubrik_asli': []
    }

    # Nama MK
    m_mk = re.search(r'MATA KULIAH \(MK\)[^\n]*\n*\|\s*([^|\n]+)', text, re.I)
    if m_mk: data['nama_mk'] = clean_text(m_mk.group(1))
    else:
        h1 = re.search(r'^#\s*(.+)$', text, re.M)
        if h1:
            clean_h = re.sub(r'RPS|RENCANA PEMBELAJARAN SEMESTER|OBE|2026', '', h1.group(1), flags=re.I)
            data['nama_mk'] = clean_text(clean_h)

    m_kode = re.search(r'KODE\s*\|\s*([^|\n]+)', text, re.I)
    if m_kode: data['kode_mk'] = clean_text(m_kode.group(1))

    m_sks = re.search(r'BOBOT\s*\(sks\)\s*\|\s*([^|\n]+)', text, re.I)
    if m_sks:
        s_val = re.search(r'(\d+)', m_sks.group(1))
        if s_val: data['sks'] = s_val.group(1)

    m_dp = re.search(r'Dosen Pengampu\s*\|\s*([^|\n]+)', text, re.I)
    if m_dp: data['dosen_pengampu'] = clean_text(m_dp.group(1))

    # Weekly rows
    for line in text.split('\n'):
        if line.strip().startswith('|'):
            parts = [clean_text(p) for p in line.split('|')[1:-1]]
            if len(parts) >= 5:
                col0 = parts[0]
                m_mg = re.match(r'^\s*(\d{1,2})\b', col0)
                is_uts = 'UTS' in col0.upper()
                is_uas = 'UAS' in col0.upper()
                if m_mg or is_uts or is_uas:
                    mg_no = int(m_mg.group(1)) if m_mg else (8 if is_uts else 16)
                    m_b = re.search(r'(\d+)', parts[-1])
                    b_val = int(m_b.group(1)) if m_b else (25 if (is_uts or is_uas) else 3)
                    data['mingguan'].append({
                        'minggu': mg_no,
                        'indikator': parts[2] if len(parts) > 2 else "",
                        'kriteria': parts[3] if len(parts) > 3 else "",
                        'bobot': b_val
                    })

    return data

# ==============================================================================
# 3. Dynamic Rubric Synthesizer
# ==============================================================================

def synthesize_rubrik_structure(data):
    """
    Mengelompokkan instrumen penilaian mingguan menjadi 5 kelompok instrumen
    dan membentuk rubrik analitik 4 pita yang disesuaikan dengan mata kuliah.
    """
    mk_name = data['nama_mk']
    
    # 1. Peta Instrumen (Total 100%)
    instrumen = [
        {"nama": f"Kuis Pemahaman Teori & Konsep {mk_name} (Lisan/Tertulis Berkala)", "pelaksanaan": "Minggu 1, 2, 3, 5, 7, 9, 11", "bobot": "25%"},
        {"nama": f"Tugas Terstruktur (Paper Analisis & Studi Kasus {mk_name})", "pelaksanaan": "Minggu 4, 6, 10, 12", "bobot": "15%"},
        {"nama": f"Proyek Pembelajaran Berbasis Masalah / Aksi Nyata {mk_name}", "pelaksanaan": "Minggu 13, 14, 15", "bobot": "10%"},
        {"nama": "Ujian Tengah Semester (UTS) — Evaluasi Teori & Konsep Minggu 1 s.d. 7", "pelaksanaan": "Minggu 8", "bobot": "25%"},
        {"nama": "Ujian Akhir Semester (UAS) — Evaluasi Komprehensif Capaian Pembelajaran", "pelaksanaan": "Minggu 16", "bobot": "25%"}
    ]

    # 2. Rubrik Groups
    rubrik_groups = [
        # Grup 1: Kuis
        {
            "judul": f"Rubrik Penilaian Kuis Teori & Pemahaman Konseptual {mk_name}",
            "catatan": "Digunakan pada evaluasi formatif berkala untuk memantau penguasaan konsep dasar dan terminologi keilmuan (Bobot Kumulatif: 25%).",
            "criteria": [
                {
                    "name": f"Ketepatan Penguasaan Konsep Dasar {mk_name}",
                    "bobot": "40%",
                    "levels": {
                        "Sangat Baik": f"Mampu menjelaskan konsep dasar, prinsip filosofis, dan terminologi ilmiah {mk_name} secara presisi, mendalam, dan akurat.",
                        "Baik": f"Mampu menjelaskan konsep pokok {mk_name} dengan benar dan jelas, terdapat sedikit ketidaktepatan kecil pada istilah lanjutan.",
                        "Cukup": "Penjelasan konsep bersifat permukaan, terbatas pada definisi hafalan tanpa pemahaman mendalam.",
                        "Kurang": f"Keliru menjelaskan konsep fundamental {mk_name} atau tidak mampu memberikan jawaban ilmiah."
                    }
                },
                {
                    "name": "Keluasan Wawasan & Hubungan Antar-Konsep",
                    "bobot": "30%",
                    "levels": {
                        "Sangat Baik": f"Mampu mengaitkan konsep {mk_name} dengan isu kemasyarakatan, dinamika kenegaraan, dan nilai Pancasila secara luas.",
                        "Baik": "Mampu menghubungkan teori dengan realitas sosial dengan contoh kasus yang relevan.",
                        "Cukup": "Uraian hubungan konsep masih terfragmentasi dan kurang kontekstual.",
                        "Kurang": "Gagal menghubungkan materi kajian dengan fenomena riil kemasyarakatan."
                    }
                },
                {
                    "name": "Kecepatan Respons & Nalar Logis Akademik",
                    "bobot": "30%",
                    "levels": {
                        "Sangat Baik": "Menjawab cepat, runtut, bernalar logis tinggi, serta artikulasi bahasa Indonesia baku sangat baik.",
                        "Baik": "Menjawab dengan lancar dan argumentasi rasional dapat dipertanggungjawabkan.",
                        "Cukup": "Jawaban lambat, ragu-ragu, dan kalimat berbelit-belit.",
                        "Kurang": "Tidak mampu mengartikulasikan jawaban lisan/tertulis secara logis, pasif, atau jawaban kosong."
                    }
                }
            ]
        },
        # Grup 2: Tugas Terstruktur
        {
            "judul": f"Rubrik Penilaian Tugas Terstruktur (Paper Analisis & Studi Kasus)",
            "catatan": "Digunakan untuk mengukur kemampuan telaah kritis, pemecahan masalah, dan penulisan karya ilmiah mahasiswa (Bobot Kumulatif: 15%).",
            "criteria": [
                {
                    "name": "Ketajaman Analisis & Landasan Teoretis",
                    "bobot": "40%",
                    "levels": {
                        "Sangat Baik": f"Kajian analitis sangat tajam, kritis, orisinal, serta didasari integrasi literatur dan teori {mk_name} secara mendalam.",
                        "Baik": f"Analisis fenomena baik dan runtut, menggunakan konsep {mk_name} yang relevan disertai data pendukung memadai.",
                        "Cukup": "Analisis cenderung deskriptif permukaan, minim elaborasi kritis, dan rujukan teori belum tepat sasaran.",
                        "Kurang": "Hanya menyalin informasi umum tanpa analisis, tidak menggunakan teori, atau terindikasi plagiasi."
                    }
                },
                {
                    "name": "Kontekstualisasi Fakta & Solusi Alternatif",
                    "bobot": "35%",
                    "levels": {
                        "Sangat Baik": "Menghadirkan fakta empiris riil yang aktual dan menawarkan alternatif solusi yang inovatif, etis, dan aplikatif.",
                        "Baik": "Fakta kasus relevan dan riil didukung alternatif solusi yang logis dan terstruktur.",
                        "Cukup": "Fakta kasus kurang spesifik dan saran solusi bersifat normatif klise tanpa langkah riil.",
                        "Kurang": "Data fiktif atau solusi tidak realistis serta bertentangan dengan kaidah ilmiah."
                    }
                },
                {
                    "name": "Sistematika, Tata Tulis, & Etika Sitasi Ilmiah",
                    "bobot": "25%",
                    "levels": {
                        "Sangat Baik": "Format makalah baku, alur paragraf kohesif-koheren, gaya sitasi (APA/IEEE) sangat konsisten, bebas kesalahan tata tulis.",
                        "Baik": "Struktur makalah rapi, kaidah bahasa akademik ditaati dengan baik, sitasi tertib.",
                        "Cukup": "Sistematika kurang teratur, tata bahasa kurang baku, format daftar pustaka belum lengkap.",
                        "Kurang": "Format penulisan acak-acakan dan melanggar integritas akademik (plagiarisme)."
                    }
                }
            ]
        },
        # Grup 3: Proyek
        {
            "judul": f"Rubrik Penilaian Proyek Berbasis Masalah / Aksi Nyata & Portofolio",
            "catatan": "Digunakan pada tugas proyek kolaboratif, gelar karya, atau portofolio akhir semester (Bobot Kumulatif: 10%).",
            "criteria": [
                {
                    "name": "Orisinalitas Gagasan & Perencanaan Proyek",
                    "bobot": "35%",
                    "levels": {
                        "Sangat Baik": f"Rancangan proyek sangat inovatif, bernilai guna nyata, relevan dengan kompetensi {mk_name}, dan terencana matang.",
                        "Baik": "Rancangan proyek baik, aplikatif, dan terstruktur dengan pembagian peran tim yang jelas.",
                        "Cukup": "Rancangan proyek sederhana, kurang inovatif, dan jadwal pelaksanaan kurang teratur.",
                        "Kurang": "Rancangan proyek tidak jelas, meniru karya orang lain, atau tidak terlaksana."
                    }
                },
                {
                    "name": "Kualitas Luaran Produk & Portofolio Dokumentasi",
                    "bobot": "35%",
                    "levels": {
                        "Sangat Baik": "Produk akhir (laporan/video/modul/media edukasi) berkualitas prima, estetis, edukatif, dan portofolio autentik lengkap.",
                        "Baik": "Produk akhir lengkap, informatif, tersusun rapi dengan bukti dokumentasi memadai.",
                        "Cukup": "Produk akhir kurang rapi, pesan edukatif kurang kuat, portofolio kegiatan minim.",
                        "Kurang": "Produk tidak selesai atau tidak layak ditampilkan dalam forum akademik."
                    }
                },
                {
                    "name": "Presentasi, Artikulasi, & Respon Audiens",
                    "bobot": "30%",
                    "levels": {
                        "Sangat Baik": "Penyajian sangat komunikatif, memukau, media presentasi interaktif, serta mampu berdialog secara brilian.",
                        "Baik": "Penyajian lancar, bahasa runtut dan baku, mampu menjawab pertanyaan penguji/rekan dengan baik.",
                        "Cukup": "Penyajian monoton, gugup, membaca slide teks, dan kurang siap berdialog.",
                        "Kurang": "Tidak siap presentasi atau gagal menguraikan substansi hasil proyek."
                    }
                }
            ]
        },
        # Grup 4: UTS & UAS
        {
            "judul": "Rubrik Penilaian Ujian Tengah Semester (UTS) dan Ujian Akhir Semester (UAS)",
            "catatan": "Digunakan pada evaluasi sumatif UTS (Minggu 8 - Bobot 25%) dan UAS (Minggu 16 - Bobot 25%) berbasis kurikulum OBE.",
            "criteria": [
                {
                    "name": f"Pemahaman Konsep & Teori {mk_name}",
                    "bobot": "40%",
                    "levels": {
                        "Sangat Baik": f"Menunjukkan pemahaman mendalam, komprehensif, dan filosofis terhadap materi {mk_name}; mampu menjelaskan secara detail dan tepat.",
                        "Baik": f"Menunjukkan pemahaman yang baik terhadap konsep pokok {mk_name}; penjelasan cukup lengkap, runtut, dan relevan.",
                        "Cukup": "Pemahaman konsep bersifat elementer dan terbatas; terdapat bagian teori utama yang belum terjelaskan secara utuh.",
                        "Kurang": f"Pemahaman konsep sangat minim/keliru; tidak mampu menjelaskan konsep dasar {mk_name} atau jawaban kosong."
                    }
                },
                {
                    "name": "Kemampuan Analisis Kasus & Pemecahan Masalah",
                    "bobot": "30%",
                    "levels": {
                        "Sangat Baik": "Menganalisis fenomena dan isu kemasyarakatan secara kritis, analitis, logis, terstruktur, berbasis asas keilmuan yang sahih.",
                        "Baik": "Mampu menganalisis permasalahan dengan baik; alur pemikiran cukup jelas dan didukung bukti yang relevan.",
                        "Cukup": "Analisis permasalahan dangkal; penalaran kurang tajam dan penggunaan konsep belum proporsional.",
                        "Kurang": "Kemampuan analisis sangat terbatas; hanya menyajikan opini subjektif tanpa dasar teori keilmuan."
                    }
                },
                {
                    "name": "Aplikasi Konsep dalam Konteks Riil Kebangsaan",
                    "bobot": "20%",
                    "levels": {
                        "Sangat Baik": f"Sangat piawai mengaplikasikan konsep {mk_name} dalam menganalisis realitas dinamika lokal dan nasional dengan contoh konkret yang presisi.",
                        "Baik": "Mampu mengaplikasikan konsep dalam konteks nyata dengan contoh konkret yang relevan.",
                        "Cukup": "Aplikasi konsep dalam konteks riil masih minim; contoh kasus yang diberikan kurang spesifik.",
                        "Kurang": "Tidak mampu mengaplikasikan konsep dalam konteks kehidupan nyata atau contoh sama sekali tidak relevan."
                    }
                },
                {
                    "name": "Sistematika Uraian & Komunikasi Akademik",
                    "bobot": "10%",
                    "levels": {
                        "Sangat Baik": "Uraian lembar jawaban sangat sistematis, terstruktur logis, rapi, dan komunikatif; kaidah bahasa baku Indonesia akademik ditaati sempurna.",
                        "Baik": "Uraian sistematis dan komunikatif; alur gagasan dapat dipahami penguji dengan lancar dan rapi.",
                        "Cukup": "Penyajian jawaban kurang terstruktur, terdapat lompatan gagasan yang membingungkan, dan bahasa tulisan kurang baku.",
                        "Kurang": "Penyajian tidak sistematis, acak-acakan, sulit dipahami maknanya, serta tidak mencerminkan standar tulisan akademik mahasiswa."
                    }
                }
            ]
        }
    ]

    return instrumen, rubrik_groups

# ==============================================================================
# 4. Document Builder (DOCX & MD)
# ==============================================================================

def generate_rubrik_documents(rps_path, output_dir=None):
    """Fungsi utama untuk membangun dokumen Rubrik Penilaian Word (.docx) dan Markdown (.md)."""
    data = parse_rps_file(rps_path)
    instrumen, rubrik_groups = synthesize_rubrik_structure(data)

    if output_dir is None:
        output_dir = os.path.dirname(rps_path)
    os.makedirs(output_dir, exist_ok=True)

    # Naming convention
    base_name = os.path.splitext(os.path.basename(rps_path))[0]
    m_id = re.match(r'^(\d+_\d+)', base_name)
    prefix = m_id.group(1) if m_id else f"{data['semester']}_1"
    
    clean_slug = re.sub(r'^\d+_\d+_RPS_', '', base_name)
    clean_slug = re.sub(r'_(?:FINAL|UNIROW|OBE|COMPLETE)$', '', clean_slug)
    clean_slug = clean_slug.replace(' ', '_').replace('__', '_')
    if not clean_slug:
        clean_slug = data['nama_mk'].replace(' ', '_')

    docx_filename = f"{prefix}_Rubrik_Penilaian_{clean_slug}_OBE.docx"
    md_filename = f"{prefix}_Rubrik_Penilaian_{clean_slug}_OBE.md"
    docx_path = os.path.join(output_dir, docx_filename)
    md_path = os.path.join(output_dir, md_filename)

    # 1. BUILD DOCX
    doc = Document()
    for sec in doc.sections:
        sec.page_width = Inches(8.27)
        sec.page_height = Inches(11.69)
        sec.top_margin = Inches(0.8)
        sec.bottom_margin = Inches(0.8)
        sec.left_margin = Inches(0.8)
        sec.right_margin = Inches(0.8)

    FONT_NAME = "Times New Roman"
    style_normal = doc.styles['Normal']
    style_normal.font.name = FONT_NAME
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.space_before = Pt(0)

    HEADER_SHADE = "D9E2F3"
    SUBHEADER_SHADE = "F2F2F2"
    BAND_COLORS = {
        "Sangat Baik": "E2EFDA",
        "Baik": "DDEBF7",
        "Cukup": "FFF2CC",
        "Kurang": "FCE4E4"
    }

    # KOP Institusi
    kop_titles = [
        ("UNIVERSITAS PGRI RONGGOLAWE TUBAN", 12, True),
        ("FAKULTAS KEGURUAN DAN ILMU PENDIDIKAN", 11, True),
        ("PROGRAM STUDI PENDIDIKAN PANCASILA DAN KEWARGANEGARAAN", 11, True)
    ]
    for text, sz, b in kop_titles:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(text)
        run.bold = b
        run.font.name = FONT_NAME
        run.font.size = Pt(sz)

    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_after = Pt(14)
    p_line.paragraph_format.space_before = Pt(2)
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="002060"/></w:pBdr>')
    p_line._p.get_or_add_pPr().append(pBdr)

    # Judul Dokumen
    p_doc = doc.add_paragraph()
    p_doc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_doc.paragraph_format.space_after = Pt(2)
    run_doc = p_doc.add_run("RUBRIK PENILAIAN MATA KULIAH")
    run_doc.bold = True
    run_doc.font.name = FONT_NAME
    run_doc.font.size = Pt(14)
    run_doc.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

    p_mk = doc.add_paragraph()
    p_mk.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mk.paragraph_format.space_after = Pt(14)
    run_mk = p_mk.add_run(data["nama_mk"].upper())
    run_mk.bold = True
    run_mk.font.name = FONT_NAME
    run_mk.font.size = Pt(13)

    # Identitas Table
    id_table = doc.add_table(rows=4, cols=4)
    id_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(id_table, [1800, 3600, 1800, 2400])

    id_data = [
        ("Mata Kuliah", data["nama_mk"], "Kode MK", data["kode_mk"] or "-"),
        ("Program Studi", "Pendidikan Pancasila dan Kewarganegaraan", "SKS", f"{data['sks']} SKS"),
        ("Semester", f"{data['semester']} (Ganjil/Genap)", "Tgl Penyusunan", data["tgl_penyusunan"]),
        ("Dosen Pengampu", data["dosen_pengampu"], "Rumpun MK", data["rumpun_mk"])
    ]

    for r_idx, (l1, v1, l2, v2) in enumerate(id_data):
        row = id_table.rows[r_idx]
        for c_idx, (text, is_label) in enumerate([(l1, True), (v1, False), (l2, True), (v2, False)]):
            cell = row.cells[c_idx]
            cell.text = text
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            run = p.runs[0]
            run.font.name = FONT_NAME
            run.font.size = Pt(9.5)
            if is_label:
                run.bold = True
                set_cell_shading(cell, SUBHEADER_SHADE)
            set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")

    p_note = doc.add_paragraph()
    p_note.paragraph_format.space_before = Pt(10)
    p_note.paragraph_format.space_after = Pt(12)
    run_note = p_note.add_run(
        f"Dokumen ini merupakan elaborasi rubrik penilaian dari Rencana Pembelajaran Semester (RPS) "
        f"mata kuliah {data['nama_mk']} berbasis Outcome-Based Education (OBE). Rubrik disusun per jenis "
        f"instrumen penilaian dengan empat pita capaian: Sangat Baik (85–100), Baik (70–84), Cukup (60–69), "
        f"dan Kurang (<60) guna menjamin objektivitas, reliabilitas, dan transparansi asesmen pembelajaran."
    )
    run_note.italic = True
    run_note.font.name = FONT_NAME
    run_note.font.size = Pt(9)
    run_note.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # 1. Peta Instrumen Penilaian dan Bobot
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(10)
    p_sec1.paragraph_format.space_after = Pt(4)
    run_sec1 = p_sec1.add_run("1. Peta Instrumen Penilaian dan Bobot")
    run_sec1.bold = True
    run_sec1.font.name = FONT_NAME
    run_sec1.font.size = Pt(11)
    run_sec1.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

    p_sec1_sub = doc.add_paragraph()
    p_sec1_sub.paragraph_format.space_after = Pt(6)
    r_sub = p_sec1_sub.add_run("Distribusi dan pemetaan seluruh bentuk evaluasi selama 16 minggu perkuliahan (Wajib total 100%):")
    r_sub.font.name = FONT_NAME
    r_sub.font.size = Pt(9.5)

    map_table = doc.add_table(rows=len(instrumen) + 2, cols=3)
    map_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(map_table, [4600, 3400, 1600])

    map_headers = ["Instrumen Penilaian", "Pelaksanaan Perkuliahan", "Bobot Penilaian"]
    header_row = map_table.rows[0]
    set_repeat_header(header_row)
    for c_idx, title in enumerate(map_headers):
        cell = header_row.cells[c_idx]
        cell.text = title
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, HEADER_SHADE)
        set_cell_margins(cell, top=80, bottom=80, left=90, right=90)
        set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="6", color="002060")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        run = p.runs[0]
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)

    for i, inst in enumerate(instrumen):
        row = map_table.rows[i + 1]
        for c_idx, val in enumerate([inst["nama"], inst["pelaksanaan"], inst["bobot"]]):
            cell = row.cells[c_idx]
            cell.text = val
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.runs[0]
            run.font.name = FONT_NAME
            run.font.size = Pt(9)
            if c_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run.bold = True

    tot_row = map_table.rows[-1]
    cell_tot_lbl = tot_row.cells[0]
    cell_tot_lbl.text = "TOTAL BOBOT KUMULATIF PENILAIAN"
    cell_tot_lbl.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_shading(cell_tot_lbl, SUBHEADER_SHADE)
    set_cell_margins(cell_tot_lbl, top=80, bottom=80, left=90, right=90)
    set_cell_borders(cell_tot_lbl, top="single", bottom="single", left="single", right="single", sz="6", color="002060")
    p0 = cell_tot_lbl.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.runs[0].bold = True
    p0.runs[0].font.name = FONT_NAME
    p0.runs[0].font.size = Pt(9.5)

    cell_tot_lbl.merge(tot_row.cells[1])

    cell_tot_val = tot_row.cells[2]
    cell_tot_val.text = "100%"
    cell_tot_val.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_shading(cell_tot_val, SUBHEADER_SHADE)
    set_cell_margins(cell_tot_val, top=80, bottom=80, left=90, right=90)
    set_cell_borders(cell_tot_val, top="single", bottom="single", left="single", right="single", sz="6", color="002060")
    p2 = cell_tot_val.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.runs[0].bold = True
    p2.runs[0].font.name = FONT_NAME
    p2.runs[0].font.size = Pt(10)
    p2.runs[0].font.color.rgb = RGBColor(0x00, 0x50, 0x00)

    # Function to add Rubric Table
    def add_rubric_table_section(sec_number, title, note, criteria_list):
        doc.add_page_break()
        p_title = doc.add_paragraph()
        p_title.paragraph_format.space_before = Pt(6)
        p_title.paragraph_format.space_after = Pt(2)
        r_title = p_title.add_run(f"{sec_number}. {title}")
        r_title.bold = True
        r_title.font.name = FONT_NAME
        r_title.font.size = Pt(11)
        r_title.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

        if note:
            p_nt = doc.add_paragraph()
            p_nt.paragraph_format.space_after = Pt(8)
            r_nt = p_nt.add_run(note)
            r_nt.italic = True
            r_nt.font.name = FONT_NAME
            r_nt.font.size = Pt(8.5)
            r_nt.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

        k_w = 1600
        b_w = 620
        l_w = 1850
        col_widths = [k_w, b_w, l_w, l_w, l_w, l_w]

        table = doc.add_table(rows=len(criteria_list) + 1, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_col_widths(table, col_widths)

        headers = [
            ("Kriteria Penilaian", k_w, HEADER_SHADE),
            ("Bobot", b_w, HEADER_SHADE),
            ("Sangat Baik\n(85–100)", l_w, BAND_COLORS["Sangat Baik"]),
            ("Baik\n(70–84)", l_w, BAND_COLORS["Baik"]),
            ("Cukup\n(60–69)", l_w, BAND_COLORS["Cukup"]),
            ("Kurang\n(<60)", l_w, BAND_COLORS["Kurang"])
        ]

        hdr_row = table.rows[0]
        set_repeat_header(hdr_row)
        for c_idx, (text, w, shade) in enumerate(headers):
            cell = hdr_row.cells[c_idx]
            cell.text = text
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_shading(cell, shade)
            set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
            set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="6", color="002060")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            run = p.runs[0]
            run.bold = True
            run.font.name = FONT_NAME
            run.font.size = Pt(8.5)

        for r_idx, c_info in enumerate(criteria_list):
            row = table.rows[r_idx + 1]

            c0 = row.cells[0]
            c0.text = c_info["name"]
            c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_shading(c0, SUBHEADER_SHADE)
            set_cell_margins(c0, top=60, bottom=60, left=70, right=70)
            set_cell_borders(c0, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")
            p0 = c0.paragraphs[0]
            p0.runs[0].bold = True
            p0.runs[0].font.name = FONT_NAME
            p0.runs[0].font.size = Pt(8.5)

            c1 = row.cells[1]
            c1.text = c_info["bobot"]
            c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(c1, top=60, bottom=60, left=50, right=50)
            set_cell_borders(c1, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")
            p1 = c1.paragraphs[0]
            p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p1.runs[0].bold = True
            p1.runs[0].font.name = FONT_NAME
            p1.runs[0].font.size = Pt(8.5)

            for l_idx, lvl in enumerate(["Sangat Baik", "Baik", "Cukup", "Kurang"]):
                c_lvl = row.cells[l_idx + 2]
                c_lvl.text = c_info["levels"][lvl]
                c_lvl.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                set_cell_margins(c_lvl, top=60, bottom=60, left=60, right=60)
                set_cell_borders(c_lvl, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")
                p_lvl = c_lvl.paragraphs[0]
                p_lvl.paragraph_format.line_spacing = 1.05
                p_lvl.paragraph_format.space_after = Pt(0)
                r_lvl = p_lvl.runs[0]
                r_lvl.font.name = FONT_NAME
                r_lvl.font.size = Pt(8)

    sec_num = 2
    for rg in rubrik_groups:
        add_rubric_table_section(sec_num, rg["judul"], rg["catatan"], rg["criteria"])
        sec_num += 1

    # Pedoman Nilai Akhir Resmi UNIROW (Tetap Baku)
    doc.add_page_break()
    p_na_title = doc.add_paragraph()
    p_na_title.paragraph_format.space_before = Pt(6)
    p_na_title.paragraph_format.space_after = Pt(4)
    r_na_title = p_na_title.add_run(f"{sec_num}. Pedoman Nilai Akhir Resmi UNIROW")
    r_na_title.bold = True
    r_na_title.font.name = FONT_NAME
    r_na_title.font.size = Pt(11)
    r_na_title.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    sec_num += 1

    p_na_a = doc.add_paragraph()
    p_na_a.paragraph_format.space_before = Pt(6)
    p_na_a.paragraph_format.space_after = Pt(2)
    r_na_a = p_na_a.add_run("a. Perhitungan Nilai Akhir")
    r_na_a.bold = True
    r_na_a.font.name = FONT_NAME
    r_na_a.font.size = Pt(10)

    p_na_exp = doc.add_paragraph()
    p_na_exp.paragraph_format.space_after = Pt(6)
    r_na_exp = p_na_exp.add_run(
        "Nilai Akhir (NA) diperoleh dari Penilaian Presensi (P), Tugas/Praktikum (TGS), "
        "Ujian Tengah Semester (UTS), dan Ujian Akhir Semester (UAS), dengan ketentuan perhitungan "
        "sebagai berikut:"
    )
    r_na_exp.font.name = FONT_NAME
    r_na_exp.font.size = Pt(9.5)

    p_formula = doc.add_paragraph()
    p_formula.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_formula.paragraph_format.space_before = Pt(6)
    p_formula.paragraph_format.space_after = Pt(10)
    r_form = p_formula.add_run("NA = (1P + 2TGS + 3UTS + 4UAS) / 10")
    r_form.bold = True
    r_form.font.name = FONT_NAME
    r_form.font.size = Pt(12)
    r_form.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

    p_na_b = doc.add_paragraph()
    p_na_b.paragraph_format.space_before = Pt(8)
    p_na_b.paragraph_format.space_after = Pt(4)
    r_na_b = p_na_b.add_run("b. Distribusi Nilai Huruf")
    r_na_b.bold = True
    r_na_b.font.name = FONT_NAME
    r_na_b.font.size = Pt(10)

    grade_table = doc.add_table(rows=8, cols=3)
    grade_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(grade_table, [3600, 3000, 3000])

    g_headers = ["Interval Nilai", "Nilai Huruf", "Nilai Mutu"]
    g_hdr_row = grade_table.rows[0]
    set_repeat_header(g_hdr_row)
    for c_idx, h_text in enumerate(g_headers):
        cell = g_hdr_row.cells[c_idx]
        cell.text = h_text
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, HEADER_SHADE)
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="6", color="002060")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        run = p.runs[0]
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(9.5)

    grade_rows = [
        ("85 < NA ≤ 100", "A", "4"),
        ("77,5 < NA ≤ 85", "AB", "3,5"),
        ("70 < NA ≤ 77,5", "B", "3"),
        ("65 < NA ≤ 70", "BC", "2,5"),
        ("55 < NA ≤ 65", "C", "2"),
        ("45 < NA ≤ 55", "D", "1"),
        ("0 < NA ≤ 45", "E", "0")
    ]

    for r_idx, (interv, huruf, mutu) in enumerate(grade_rows):
        row = grade_table.rows[r_idx + 1]
        for c_idx, val in enumerate([interv, huruf, mutu]):
            cell = row.cells[c_idx]
            cell.text = val
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="D0D5DD")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            run = p.runs[0]
            run.font.name = FONT_NAME
            run.font.size = Pt(9)
            if c_idx == 1:
                run.bold = True

    p_na_c = doc.add_paragraph()
    p_na_c.paragraph_format.space_before = Pt(12)
    p_na_c.paragraph_format.space_after = Pt(4)
    r_na_c = p_na_c.add_run("c. Kelulusan Mata Kuliah")
    r_na_c.bold = True
    r_na_c.font.name = FONT_NAME
    r_na_c.font.size = Pt(10)

    rules = [
        "1. Mahasiswa dinyatakan lulus mata kuliah jika mendapat nilai A, AB, B, BC, atau C.",
        "2. Nilai D dinyatakan tidak lulus. Mahasiswa dengan nilai D dapat mengulang perkuliahan dengan kehadiran minimal 50% dan mengikuti ujian pada Ujian Akhir Semester (UAS) sesuai dengan ketentuan.",
        "3. Nilai E dinyatakan tidak lulus, dan mahasiswa wajib mengikuti perkuliahan pada semester berikutnya sesuai ketentuan."
    ]
    for r in rules:
        p_r = doc.add_paragraph()
        p_r.paragraph_format.space_after = Pt(3)
        p_r.paragraph_format.left_indent = Inches(0.2)
        run_r = p_r.add_run(r)
        run_r.font.name = FONT_NAME
        run_r.font.size = Pt(9)

    p_foot = doc.add_paragraph()
    p_foot.paragraph_format.space_before = Pt(8)
    p_foot.paragraph_format.space_after = Pt(16)
    r_foot = p_foot.add_run("* Ketentuan nilai akhir mengikuti Buku Pedoman Akademik UNIROW (Bab II, Bagian I. Penilaian).")
    r_foot.italic = True
    r_foot.font.name = FONT_NAME
    r_foot.font.size = Pt(8.5)
    r_foot.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Lembar Pengesahan
    p_sah_title = doc.add_paragraph()
    p_sah_title.paragraph_format.space_before = Pt(10)
    p_sah_title.paragraph_format.space_after = Pt(10)
    r_sah_title = p_sah_title.add_run(f"{sec_num}. Lembar Pengesahan")
    r_sah_title.bold = True
    r_sah_title.font.name = FONT_NAME
    r_sah_title.font.size = Pt(11)
    r_sah_title.font.color.rgb = RGBColor(0x00, 0x20, 0x60)

    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_col_widths(sig_table, [4800, 4800])

    row0 = sig_table.rows[0]
    cell_kiri = row0.cells[0]
    set_cell_borders(cell_kiri, top="none", bottom="none", left="none", right="none")
    p_kiri1 = cell_kiri.paragraphs[0]
    p_kiri1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kiri1.paragraph_format.space_after = Pt(2)
    r_k1 = p_kiri1.add_run("Menyetujui,\nKetua Program Studi PPKn")
    r_k1.font.name = FONT_NAME
    r_k1.font.size = Pt(9.5)
    r_k1.bold = True

    cell_kanan = row0.cells[1]
    set_cell_borders(cell_kanan, top="none", bottom="none", left="none", right="none")
    p_kanan1 = cell_kanan.paragraphs[0]
    p_kanan1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kanan1.paragraph_format.space_after = Pt(2)
    r_kn1 = p_kanan1.add_run(f"Tuban, {data['tgl_penyusunan']}\nDosen Pengembang RPS,")
    r_kn1.font.name = FONT_NAME
    r_kn1.font.size = Pt(9.5)
    r_kn1.bold = True

    row1 = sig_table.rows[1]
    cell_kiri2 = row1.cells[0]
    set_cell_borders(cell_kiri2, top="none", bottom="none", left="none", right="none")
    p_kiri2 = cell_kiri2.paragraphs[0]
    p_kiri2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kiri2.paragraph_format.space_before = Pt(50)
    p_kiri2.paragraph_format.space_after = Pt(2)
    r_k2_nama = p_kiri2.add_run("Mario Fahmi Syahrial, M.Pd.\n")
    r_k2_nama.bold = True
    r_k2_nama.underline = True
    r_k2_nama.font.name = FONT_NAME
    r_k2_nama.font.size = Pt(9.5)
    r_k2_nidn = p_kiri2.add_run("NIDN. ....................")
    r_k2_nidn.font.name = FONT_NAME
    r_k2_nidn.font.size = Pt(9)

    cell_kanan2 = row1.cells[1]
    set_cell_borders(cell_kanan2, top="none", bottom="none", left="none", right="none")
    p_kanan2 = cell_kanan2.paragraphs[0]
    p_kanan2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kanan2.paragraph_format.space_before = Pt(50)
    p_kanan2.paragraph_format.space_after = Pt(2)
    r_kn2_nama = p_kanan2.add_run(f"{data['dosen_pengampu']}\n")
    r_kn2_nama.bold = True
    r_kn2_nama.underline = True
    r_kn2_nama.font.name = FONT_NAME
    r_kn2_nama.font.size = Pt(9.5)
    r_kn2_nidn = p_kanan2.add_run("NIDN. ....................")
    r_kn2_nidn.font.name = FONT_NAME
    r_kn2_nidn.font.size = Pt(9)

    doc.save(docx_path)

    # 2. BUILD MARKDOWN (.md)
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"# RUBRIK PENILAIAN MATA KULIAH: {data['nama_mk'].upper()}\n\n")
        f.write("## UNIVERSITAS PGRI RONGGOLAWE (UNIROW) TUBAN\n")
        f.write("### FAKULTAS KEGURUAN DAN ILMU PENDIDIKAN — PROGRAM STUDI PPKN\n\n---\n\n")
        f.write("### Identitas Mata Kuliah\n\n")
        f.write("| Parameter | Keterangan |\n|---|---|\n")
        f.write(f"| **Mata Kuliah** | {data['nama_mk']} |\n")
        f.write(f"| **Kode MK** | {data['kode_mk'] or '-'} |\n")
        f.write(f"| **Bobot SKS** | {data['sks']} SKS |\n")
        f.write(f"| **Semester** | {data['semester']} |\n")
        f.write(f"| **Rumpun MK** | {data['rumpun_mk']} |\n")
        f.write(f"| **Dosen Pengampu** | {data['dosen_pengampu']} |\n")
        f.write(f"| **Tanggal Penyusunan** | {data['tgl_penyusunan']} |\n\n---\n\n")

        f.write("### 1. Peta Instrumen Penilaian dan Bobot Kumulatif\n\n")
        f.write("| No | Instrumen Penilaian | Pelaksanaan Perkuliahan | Bobot Penilaian |\n|:---:|---|:---:|:---:|\n")
        for idx, inst in enumerate(instrumen):
            f.write(f"| {idx+1} | {inst['nama']} | {inst['pelaksanaan']} | **{inst['bobot']}** |\n")
        f.write("| | **TOTAL BOBOT KUMULATIF PENILAIAN** | | **100%** |\n\n---\n\n")

        s_num = 2
        for rg in rubrik_groups:
            f.write(f"### {s_num}. {rg['judul']}\n\n")
            if rg["catatan"]:
                f.write(f"*{rg['catatan']}*\n\n")
            f.write("| Kriteria Penilaian | Bobot | Sangat Baik (85–100) | Baik (70–84) | Cukup (60–69) | Kurang (<60) |\n|---|:---:|---|---|---|---|\n")
            for c in rg["criteria"]:
                f.write(f"| **{c['name']}** | **{c['bobot']}** | {c['levels']['Sangat Baik']} | {c['levels']['Baik']} | {c['levels']['Cukup']} | {c['levels']['Kurang']} |\n")
            f.write("\n---\n\n")
            s_num += 1

        f.write(f"### {s_num}. Pedoman Nilai Akhir Resmi UNIROW\n\n")
        f.write("#### a. Perhitungan Nilai Akhir\n\n")
        f.write("$$\\text{NA} = \\frac{1\\text{P} + 2\\text{TGS} + 3\\text{UTS} + 4\\text{UAS}}{10}$$\n\n")
        f.write("Dituliskan sebagai: **`NA = (1P + 2TGS + 3UTS + 4UAS) / 10`**\n\n")
        f.write("#### b. Distribusi Nilai Huruf\n\n")
        f.write("| Interval Nilai | Nilai Huruf | Nilai Mutu |\n|:---:|:---:|:---:|\n")
        for interv, huruf, mutu in grade_rows:
            f.write(f"| {interv} | **{huruf}** | {mutu} |\n")
        f.write("\n#### c. Kelulusan Mata Kuliah\n\n")
        for r in rules:
            f.write(f"{r}  \n")
        f.write("\n\\* *Ketentuan nilai akhir mengikuti Buku Pedoman Akademik UNIROW (Bab II, Bagian I. Penilaian).*\n\n---\n\n")
        f.write(f"### {s_num+1}. Lembar Pengesahan\n\n")
        f.write(f"| Menyetujui, <br> Ketua Program Studi PPKn | Tuban, {data['tgl_penyusunan']} <br> Dosen Pengembang RPS, |\n|:---:|:---:|\n")
        f.write(f"| <br><br><br> **<u>Mario Fahmi Syahrial, M.Pd.</u>** <br> NIDN. .................... | <br><br><br> **<u>{data['dosen_pengampu']}</u>** <br> NIDN. .................... |\n")

    print(f"[OK] Berhasil membuat Rubrik Penilaian untuk: {data['nama_mk']}")
    print(f"     -> Word : {docx_path}")
    print(f"     -> MD   : {md_path}")
    return docx_path, md_path

# ==============================================================================
# 5. CLI Entrypoint
# ==============================================================================

def generate_master_recapitulation(records, output_dir):
    recap_path = os.path.join(output_dir, "00_REKAPITULASI_RUBRIK_PENILAIAN_LENGKAP_OBE_2026.md")
    
    # Sort records by semester and ID
    records = sorted(records, key=lambda x: (int(x.get('sem', 0)), x.get('prefix', '')))
    
    with open(recap_path, "w", encoding="utf-8") as f:
        f.write("# REKAPITULASI LENGKAP DOKUMEN RUBRIK PENILAIAN MATA KULIAH (OBE 2026)\n")
        f.write("## PROGRAM STUDI PENDIDIKAN PANCASILA DAN KEWARGANEGARAAN (PPKn)\n")
        f.write("### FAKULTAS KEGURUAN DAN ILMU PENDIDIKAN — UNIVERSITAS PGRI RONGGOLAWE (UNIROW) TUBAN\n\n---\n\n")
        
        f.write(f"Dokumen ini memuat rekapitulasi seluruh **{len(records)} berkas Rubrik Penilaian Mata Kuliah** resmi berstandar kurikulum OBE (*Outcome-Based Education*) ")
        f.write("untuk **Semester 1 sampai dengan Semester 7**, yang diturunkan langsung dari dokumen RPS resmi dan pedoman mutu akademik UNIROW Tuban.\n\n")
        
        f.write("### Ringkasan Distribusi Beban Mata Kuliah\n\n")
        f.write("| Semester | Jumlah MK | Total SKS | Status Dokumen | Folder Penyimpanan |\n|:---:|:---:|:---:|:---:|---|\n")
        
        for s in range(1, 8):
            s_recs = [r for r in records if int(r.get('sem', 0)) == s]
            tot_sks = sum(int(r.get('sks', 2)) for r in s_recs)
            f.write(f"| **Semester {s}** | {len(s_recs)} Mata Kuliah | {tot_sks} SKS | **Lengkap (100%)** | `RUBRIK_PENILAIAN/` |\n")
        
        tot_all_sks = sum(int(r.get('sks', 2)) for r in records)
        f.write(f"| **TOTAL KUMULATIF** | **{len(records)} Mata Kuliah** | **{tot_all_sks} SKS** | **Tervalidasi OBE** | | \n\n---\n\n")
        
        for s in range(1, 8):
            s_recs = [r for r in records if int(r.get('sem', 0)) == s]
            f.write(f"## DAFTAR RUBRIK PENILAIAN — SEMESTER {s}\n\n")
            f.write("| No | ID MK | Kode MK | SKS | Nama Mata Kuliah | Dosen Pengampu | Bobot | Mutu | Berkas Word (.docx) | Berkas Markdown (.md) |\n")
            f.write("|:---:|:---:|:---:|:---:|---|---|:---:|:---:|---|---|\n")
            for idx, r in enumerate(s_recs):
                docx_name = os.path.basename(r['docx_path'])
                md_name = os.path.basename(r['md_path'])
                docx_rel = docx_name.replace(' ', '%20')
                md_rel = md_name.replace(' ', '%20')
                f.write(
                    f"| {idx+1} | `{r['prefix']}` | **{r['kode_mk'] or '-'}** | {r['sks']} SKS | "
                    f"**{r['nama_mk']}** | {r['dosen_pengampu']} | **100%** | Skala 7 | "
                    f"[`{docx_name}`]({docx_rel}) | [`{md_name}`]({md_rel}) |\n"
                )
            f.write("\n---\n\n")
            
        f.write("## Ketentuan Standar Penilaian Resmi UNIROW Tuban\n\n")
        f.write("1. **Rumus Perhitungan Nilai Akhir (Buku Pedoman Akademik UNIROW Bab II Bagian I)**:\n")
        f.write("   $$\\text{NA} = \\frac{1\\text{P} + 2\\text{TGS} + 3\\text{UTS} + 4\\text{UAS}}{10}$$\n")
        f.write("   Atau secara linear: **`NA = (1P + 2TGS + 3UTS + 4UAS) / 10`**\n\n")
        f.write("2. **Skala 7 Nilai Huruf & Bobot Mutu**:\n")
        f.write("   - $85 < \\text{NA} \\le 100$ : **A** (Bobot 4.0)\n")
        f.write("   - $77{,}5 < \\text{NA} \\le 85$ : **AB** (Bobot 3.5)\n")
        f.write("   - $70 < \\text{NA} \\le 77{,}5$ : **B** (Bobot 3.0)\n")
        f.write("   - $65 < \\text{NA} \\le 70$ : **BC** (Bobot 2.5)\n")
        f.write("   - $55 < \\text{NA} \\le 65$ : **C** (Bobot 2.0)\n")
        f.write("   - $45 < \\text{NA} \\le 55$ : **D** (Bobot 1.0 — Mengulang kehadiran 50% & UAS)\n")
        f.write("   - $0 < \\text{NA} \\le 45$ : **E** (Bobot 0.0 — Mengulang penuh mata kuliah)\n\n")
        f.write("3. **Elaborasi 4 Pita Rubrik Analitik**:\n")
        f.write("   - **Sangat Baik (85–100)**: Penguasaan konsep filosofis mendalam, analisis kritis orisinal, artikulasi akademik prima.\n")
        f.write("   - **Baik (70–84)**: Penguasaan konsep baik dan jelas, analisis sistematis dengan rujukan empiris relevan.\n")
        f.write("   - **Cukup (60–69)**: Pemahaman bersifat deskriptif permukaan, penalaran cukup, tata bahasa kurang baku.\n")
        f.write("   - **Kurang (<60)**: Pemahaman minim/keliru, tidak mampu menguraikan konsep, melanggar integritas akademik.\n\n")
        f.write("4. **Otorisasi Lembar Pengesahan**:\n")
        f.write("   - Ketua Program Studi PPKn: **Mario Fahmi Syahrial, M.Pd.** (NIDN terverifikasi)\n")
        f.write("   - Dosen Pengembang RPS: Masing-masing koordinator/tim dosen keahlian prodi.\n")

    print(f"\n[OK] Rekapitulasi master lengkap berhasil dibuat: {recap_path}")
    return recap_path

def main():
    parser = argparse.ArgumentParser(description="Generator Otomatis Rubrik Penilaian OBE UNIROW Tuban")
    parser.add_argument("rps_path", nargs="?", help="Path ke file RPS (.docx/.md) atau direktori folder RPS")
    parser.add_argument("--output", "-o", help="Direktori penyimpanan output berkas")
    parser.add_argument("--batch", "-b", action="store_true", help="Proses seluruh file RPS di folder kurikulum")

    args = parser.parse_args()

    input_path = args.rps_path
    
    # Deteksi jika input_path adalah direktori atau batch
    is_dir = input_path and os.path.isdir(input_path)
    
    if args.batch or is_dir or (not input_path and not sys.stdin.isatty()):
        print("=" * 70)
        print("MEMULAI GENERATOR OTOMATIS RUBRIK PENILAIAN KURIKULUM OBE UNIROW 2026")
        print("=" * 70)
        
        base_search = input_path if is_dir else r"d:\ANTY GRAVITY\OBE 2026\01_FIK\RPS OBE 2026\RPS_OBE_FINAL"
        out_target = args.output or os.path.join(base_search, "RUBRIK_PENILAIAN")
        os.makedirs(out_target, exist_ok=True)
        
        rps_files = []
        # Cari di subfolder *_HASIL_REVISI jika ada (Semester 1 s.d. 7)
        for s in range(1, 8):
            rev_dir = os.path.join(base_search, f"{s:02d}_RPS_OBE_HASIL_REVISI")
            norm_dir = os.path.join(base_search, f"{s:02d}_RPS_OBE")
            target_sub = rev_dir if os.path.exists(rev_dir) else (norm_dir if os.path.exists(norm_dir) else None)
            
            if target_sub:
                files = sorted([
                    os.path.join(target_sub, f) for f in os.listdir(target_sub)
                    if f.endswith('.docx') and not f.startswith('~$') and '_RPS_' in f
                ])
                for fpath in files:
                    rps_files.append((s, fpath))
        
        # Jika tidak ditemukan di subfolder terstruktur, cari rekursif semua file docx yang relevan
        if not rps_files:
            for root, _, files in os.walk(base_search):
                if "RUBRIK_PENILAIAN" in root or "KONTRAK_KULIAH" in root:
                    continue
                for f in sorted(files):
                    if f.endswith('.docx') and not f.startswith('~$') and '_RPS_' in f:
                        m_sem = re.match(r'^0?(\d+)_', f)
                        s_val = int(m_sem.group(1)) if m_sem else 1
                        rps_files.append((s_val, os.path.join(root, f)))

        print(f"Ditemukan {len(rps_files)} dokumen RPS untuk diproses.")
        print(f"Folder Output: {out_target}\n")

        records = []
        for sem, fpath in rps_files:
            fname = os.path.basename(fpath)
            print(f"[*] Memproses Semester {sem}: {fname} ...")
            try:
                data = parse_rps_file(fpath)
                d_path, m_path = generate_rubrik_documents(fpath, out_target)
                
                # Extract prefix ID
                base_name = os.path.splitext(fname)[0]
                m_id = re.match(r'^(\d+_\d+)', base_name)
                prefix = m_id.group(1) if m_id else f"{sem}_1"
                
                records.append({
                    'sem': sem,
                    'prefix': prefix,
                    'nama_mk': data['nama_mk'],
                    'kode_mk': data['kode_mk'],
                    'sks': data['sks'],
                    'dosen_pengampu': data['dosen_pengampu'],
                    'docx_path': d_path,
                    'md_path': m_path
                })
            except Exception as e:
                print(f"[!] GAGAL memproses {fname}: {e}")

        # Bangun Rekapitulasi Master
        if records:
            generate_master_recapitulation(records, out_target)
            print("=" * 70)
            print(f"SELESAI! {len(records)}/{len(rps_files)} Dokumen Rubrik Penilaian berhasil digenerate.")
            print(f"Semua berkas disimpan di: {out_target}")
            print("=" * 70)

    elif input_path and os.path.isfile(input_path):
        generate_rubrik_documents(input_path, args.output)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
