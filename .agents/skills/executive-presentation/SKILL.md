---
name: executive-presentation
description: Standar emas pembuatan paparan eksekutif pimpinan daerah & korporasi: zero-raster architecture, programmatic PptxGenJS, native Office charts, vector SVG icons, dan automasi ekspor headless COM.
---

# Executive Presentation AI Skill

Skill ini menyediakan panduan operasional standar tertinggi untuk memproduksi bahan paparan pimpinan daerah (Walikota, Bupati, Gubernur) dan direksi korporat dengan presisi data 100%, tata kelola visual Swiss Enterprise, dan automasi bebas eror.

---

## 1. Prinsip Utama: The Sovereign Enterprise Standard

1. **Zero-Raster AI Image Rule**:
   - JANGAN menyematkan foto AI generatif (fictional 3D/human) untuk paparan kebijakan publik atau finansial institusional. Gambar AI mengurangi wibawa dan keseriusan dokumen.
   - Keindahan visual diwujudkan melalui: *Swiss typography hierarchy, razor-sharp hairlines, semantic color accents, clean whitespace, dan vector SVG iconography*.
2. **True Native Microsoft Office Charts**:
   - Jangan menggambar grafik batang/lingkaran menggunakan penumpukan shape primitif.
   - Wajib menggunakan `slide.addChart(pptx.charts.BAR, ...)` dan `slide.addChart(pptx.charts.DOUGHNUT, ...)`. Hal ini menyematkan data Excel asli yang dapat diedit langsung oleh pengguna di PowerPoint.
3. **Pure Vector SVG Iconography**:
   - Gunakan icon vektor murni (Lucide / Heroicons) berukuran kecil (<2KB) yang di-encode ke Base64 (`data:image/svg+xml;base64,...`).
   - Bebas pecah di resolusi berapapun (LED Wall 8K hingga cetak kertas) dan warna dapat disesuaikan dengan kode heksadesimal tema.
4. **Programmatic PptxGenJS over Canva Generator**:
   - Canva AI generator TIDAK direkomendasikan untuk generasi dokumen awal karena sering memotong tabel anggaran dan merusak tata letak birokrasi.
   - Gunakan Node.js `pptxgenjs` sebagai engine utama generasi slide, dan simpan Canva hanya untuk kebutuhan presentasi panggung (*Magic Animate*).
5. **Strategic Slide Pacing & Cognitive Density Rule (15 vs 18 Slides)**:
   - Jangan memaksakan dual-entity kompleks dalam 1 slide jika memaksa font < 9.5pt.
   - **Formula Ambang Batas Kepadatan Slide**:
     * **15 Slide (The Executive Sweet Spot)**: Keseimbangan optimal untuk paparan pimpinan 15–20 menit (font 10.0–11.5pt, 3–4 konsep/slide, ruang kosong 30–35%).
     * **18 Slide (Zen Room Legibility Standard)**: Standar tertinggi kenyamanan visual untuk ruang sidang/layar prorector besar jarak 7–10m (font 12.0–13.5pt, maks 2 konsep/slide, ruang kosong 42–48%).
   - Tiga Pasangan Pemecahan Strategis Menuju 18 Slide:
     * *Tata Kelola Kelembagaan* (Kedudukan UPTD-BRIDA) ➔ *Chain of Impact* (Diagram Alur Dampak Warga).
     * *Problem Statement* (Diagnosa 5 GAP Struktural) ➔ *Solution Statement* (5 Terobosan & Quick Wins).
     * *Anchor Revenue* (Klaster A Lahan & Gedung Rp 12,82 M) ➔ *Service Revenue* (Klaster B–F Layanan Fasilitas Rp 10,65 M).
   - Wajib menjaga 100% data audit, nama mitra, dan rincian angka tanpa ada pemangkasan (Zero Data Loss Invariant).

---

## 2. Alur Kerja Produksi 5 Tahap

```
[1. Audit Dokumen Riil] ➔ [2. Narasi & Struktur Piramida] ➔ [3. Programmatic PPTX (PptxGenJS)] ➔ [4. Headless COM Export] ➔ [5. Visual Audit PNG 1080p]
```

### Tahap 1: Audit Rekonsiliasi Dokumen Sumber
- Ekstraksi dokumen kebijakan (visi, landasan hukum) dan dokumen anggaran teknis (rincian angka riil).
- Angka tidak boleh dibulatkan sembarangan jika terdapat rincian mata anggaran audit.

### Tahap 2: Struktur Piramida Eksekutif

#### A. Arsitektur 15 Slide (The Executive Sweet Spot — Paparan 15–20 Menit)
- Slide 01: Cover Masthead & 4 Metrik Lompatan Finansial-Dampak
- Slide 02: Kedudukan Kelembagaan & Chain of Impact Tripartit (BRIDA ➔ STP ➔ Warga)
- Slide 03: Landasan Yuridis & Fleksibilitas Pengelolaan BLUD
- Slide 04: Empat Pilar Layanan Terpadu KST Nasional (Perpres 106/2017)
- Slide 05: Pipeline Alur 5 Tahap Hilirisasi Riset Kampus ke Industri
- Slide 06: Portofolio Produk Inovasi Teruji & Infrastruktur Kawasan Canggih
- Slide 07: Ekosistem Hexahelix: 95+ Mitra Strategis Murni Tanpa Beban APBD
- Slide 08: Diagnosa 5 GAP Strategis & Matriks Solusi Terobosan (Quick Win 30 Hari)
- Slide 09: Roadmap 3 Fase Transformasi Kelembagaan (2026–2030)
- Slide 10: Re-Setting Target Finansial: Lompatan Kemandirian BLUD +212,8% (Column Chart)
- Slide 11: Struktur 6 Klaster Revenue Stream 2030 & Top 5 Katalisator (Doughnut Chart)
- Slide 12: Anatomi Mesin Pendapatan 2030: Rincian per Mata Anggaran (Financial Ledger)
- Slide 13: 6 Dimensi Keberdampakan Daerah: Solo Technopark untuk Warga Surakarta
- Slide 14: Tiga Permohonan Kebijakan Non-Anggaran kepada Walikota Surakarta
- Slide 15: Rencana Aksi 100 Hari Pertama Pasca-Arahan Walikota (Gantt Eksekusi)

#### B. Arsitektur 18 Slide (Zen Room Legibility — Layar Proyektor Besar / Sidang Pleno)
- Slide 01: Cover Masthead & 4 Metrik Lompatan Finansial-Dampak
- Slide 02: Kedudukan Kelembagaan UPTD di Bawah BRIDA & Model Operasional BLUD
- Slide 03: Chain of Impact: Transformasi Riset Menjadi Kesejahteraan Warga Surakarta
- Slide 04: Landasan Yuridis & 4 Fleksibilitas Tata Kelola Keuangan BLUD
- Slide 05: Empat Pilar Layanan Terpadu KST Nasional (Perpres 106/2017)
- Slide 06: Pipeline 5 Tahapan Hilirisasi Riset Kampus ke Industri (TRL 1–9)
- Slide 07: Showcase 5 Produk Teruji & Ekosistem 5 Fasilitas Kawasan Canggih
- Slide 08: Ekosistem Hexahelix: 95+ Mitra Strategis Murni Non-APBD
- Slide 09: Diagnosa 5 GAP Struktural & Bottleneck Saat Ini (Problem Statement)
- Slide 10: Lima Solusi Terobosan & Quick Wins Transformasi (Solution Statement)
- Slide 11: Roadmap 3 Fase Transformasi Kelembagaan Menuju Kemandirian (2026–2030)
- Slide 12: Re-Setting Target Finansial: Lintasan Kemandirian & Target BEP 2029 (Column Chart)
- Slide 13: Proporsi 6 Klaster Pendapatan 2030 & Top 5 Katalis Finansial (Doughnut Chart)
- Slide 14: Mesin Pendapatan Utama: Klaster A Kerjasama Pemanfaatan Lahan (Rp 12,82 M)
- Slide 15: Mesin Pendapatan Layanan: Klaster B–F Inovasi & Jasa Berbasis Fasilitas (Rp 10,65 M)
- Slide 16: Enam Dimensi Keberdampakan Pembangunan Daerah untuk Surakarta
- Slide 17: Tiga Permohonan Kebijakan Non-Anggaran kepada Walikota Surakarta (3 Hero Cards)
- Slide 18: Matriks Rencana Aksi 100 Hari Pertama Pasca-Arahan Walikota (Gantt Tabel Penuh)

### Tahap 3: Standar Kode PptxGenJS (Widescreen 16:9)
```javascript
import PptxGenJS from 'pptxgenjs';

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';
```

### Tahap 4: Injeksi Icon Vektor SVG
```javascript
function addSvgIcon(slide, { svgBase64, x, y, size = 0.35 }) {
  slide.addImage({
    data: `image/svg+xml;base64,${svgBase64}`,
    x, y, w: size, h: size
  });
}
```

### Tahap 5: Automasi Ekspor Headless COM & Inspeksi Mandiri
Gunakan skrip PowerShell untuk membuka COM PowerPoint lokal, menyimpan PDF vektor (`SaveAs(..., 32)`), dan mengekspor gambar per slide (`Export(..., "PNG", 1920, 1080)`). Agen wajib menginspeksi file PNG menggunakan tool visual sebelum menyerahkan ke pengguna.
