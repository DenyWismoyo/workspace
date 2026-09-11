# workspace

## Paparan Roadmap Solo Technopark (STP) 2026–2030
### Edisi Zen Enterprise Ultra (18 Slide Executive Master Deck)

Repository ini berisi bahan paparan, naskah akademik, analisis finansial, script generator otomatis PPTX/PDF, dan visualisasi presentasi interaktif untuk **Transformasi Kemandirian Finansial & Tata Kelola Solo Technopark (STP) Menuju Titik Impas (BEP 2029) dan Mandiri Penuh (2030)**.

---

### 🎯 Sorotan Kunci Paparan
* **Target Kemandirian Finansial**: Re-setting pendapatan BLUD dari Rp 8,20 M (2026) melompat +212,8% menjadi Rp 23,47 M (2030).
* **Titik Impas (BEP 2029)**: Tercapai di tahun 2029 pada level Rp 18,50 M (tanpa subsidi APBD untuk operasional rutin).
* **Efisiensi Anggaran Daerah**: Penghematan belanja Pemkot Surakarta hingga Rp 10,2 Miliar/tahun.
* **Top 5 Arus Kas Mandiri**: Sewa Kawasan & Lahan (Rp 8,5 M), Layanan Pengujian Lab (Rp 5,2 M), Pelatihan Vokasi & Sertifikasi (Rp 4,5 M), Royalti Paten (Rp 3,0 M), Layanan Inkubasi Startup (Rp 2,26 M).
* **Ekosistem Hexahelix**: Kolaborasi 95+ mitra strategis global & nasional tanpa belanja modal APBD.
* **3 Permohonan Kebijakan Walikota**: Instruksi Walikota E-Katalog Afirmatif, Surat Edaran Konsorsium Riset 5 Kampus, dan Revisi Perwali Tata Kelola Tarif Lab/Aset.
* **Rencana Aksi 100 Hari**: 6 tahapan terukur pasca-arahan Walikota dari harmonisasi draf regulasi hingga kick-off konsorsium riset terapan.

---

### 📁 Struktur Repositori
* **`bahan_paparan_walikota_v3/`**:
  * `generate_pptx_v9_18s.mjs`: Generator utama PowerPoint berbasis PptxGenJS dengan standar enterprise.
  * `export_v9_18s.ps1`: Script PowerShell otomatisasi PowerPoint COM untuk konversi PDF dan render preview PNG 1080p.
  * `PAPARAN_WALIKOTA_STP_V9_18S.pptx`: Master file PowerPoint 18 slide siap tayang.
  * `PAPARAN_WALIKOTA_STP_V9_18S.pdf`: Dokumen PDF resolusi tinggi resmi.
  * `slides_18s_preview/`: Galeri preview slide 1 s/d 18 format PNG kualitas HD.
* **`output_paparan/`**:
  * Folder distribusi presentasi eksekutif (`PAPARAN_WALIKOTA_STP_V9_18S.pptx` & `.pdf`).
  * `RINGKASAN_EKSEKUTIF_WALIKOTA.html`: Briefing eksekutif ringkas.
  * `presentasi_interaktif.html`: Web viewer presentasi interaktif modern.
* **`materi_paparan/`**:
  * Naskah materi, analisis GAP, traksi historis, dan data pendukung.
* **`bahan_paparan_walikota_v2/`**:
  * Arsip iterasi versi sebelumnya.

---

### 🚀 Cara Menjalankan Generator
```bash
# 1. Install dependencies
npm install

# 2. Generate Master PPTX 18 Slide
node bahan_paparan_walikota_v3/generate_pptx_v9_18s.mjs

# 3. Ekspor PDF & Render Preview PNG (Windows dengan MS Office PowerPoint terpasang)
powershell -ExecutionPolicy Bypass -File bahan_paparan_walikota_v3/export_v9_18s.ps1
```
