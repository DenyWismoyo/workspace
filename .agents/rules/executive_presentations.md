# 🎯 Standar Pembuatan & Pemolesan Presentasi Eksekutif (Executive WOW)

Aturan ini wajib dipatuhi pada setiap tugas pembuatan materi paparan, slide presentasi, dan dokumen visual tingkat pimpinan/kepala daerah di workspace ini.

## 1. Lokasi & Format Keluaran (Invariants)
- **Folder Khusus Luaran**: Semua file hasil generate presentasi (.pptx, .pdf, .html interaktif, panduan) **WAJIB** disimpan dalam folder terpisah: `output_paparan/`. Jangan mencampurnya dengan folder bahan mentah (`materi_paparan/`).
- **Wajib Berpasangan (PPTX & PDF)**: Setiap kali membuat file PowerPoint (`.pptx`), **WAJIB** secara otomatis membuat versi PDF pendamping dengan tata letak identik (menggunakan skrip konversi PowerPoint lokal atau headless renderer).
- **Dimensi Kanvas Widescreen 16:9 Sejati**:
  - Pada PptxGenJS, dilarang menggunakan default 720p (`10.0 x 5.625` inci).
  - Wajib mendefinisikan layout 1080p:
    ```javascript
    pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
    pptx.layout = 'LAYOUT_16x9_WIDE';
    ```

## 2. Prinsip Desain Visual (Executive WOW)
- **Kartu Bento & Hirarki Visual**: Hindari paragraf panjang di dalam slide. Gunakan sistem kartu bento (*bento grid cards*), ikon bertema, dan kontras warna modern (Dark Navy `#070d18` + Cyan Neon `#38bdf8` + Golden Accent `#f59e0b`).
- **Visualisasi Finansial 3D**: Untuk angka capaian atau lompatan target finansial (seperti proyeksi BLUD), wajib menggunakan aset visual 3D fotorealistik atau grafik bercahaya (*glowing neon bar/line chart*).
- **Web Slide Interaktif**:
  - Wajib menyertakan fitur narasi suara (*Text-to-Speech id-ID*).
  - Dilengkapi *ambient tech particle canvas* dan navigasi cepat (*Slide Overview Grid* tombol `O`).

## 3. Struktur Narasi Pimpinan Daerah (AI Storytelling)
Gunakan pendekatan komunikasi piramida strategis:
1. **The Hook**: Menyentuh legasi dan visi besar kepala daerah, bukan laporan administratif birokrasi.
2. **The Proof**: Bukti konkret terukur (contoh: mitra global tanpa menguras APBD).
3. **The Closing Call**: Permohonan kepemimpinan dan payung regulasi kebijakan, bukan sekadar meminta belanja anggaran.

## 4. Kompatibilitas Canva
- File PPTX yang dihasilkan harus bersih dan berbasis objek vektor/shape terpisah sehingga 100% kompatibel saat di-import langsung (*drag-and-drop*) ke Canva.com untuk dipoles lebih lanjut menggunakan Canva Magic Studio.
