# 🎨 PANDUAN DESAIN & TEKNIS PPTX — VERSI 3
## Solo Technopark: Paparan Walikota Surakarta
*Referensi Teknis untuk Konversi Naskah .md ke PPTX/PDF*

---

## A. SPESIFIKASI PRESENTASI

| Parameter | Nilai |
| :--- | :--- |
| **Format Slide** | Widescreen 16:9 |
| **Dimensi Canvas** | 13.333 in × 7.5 in (33.87 × 19.05 cm) |
| **Jumlah Slide** | 12 Slide |
| **Resolusi Target** | 1920 × 1080 px (Full HD) |
| **Font Utama** | Inter (dari Google Fonts) |
| **Font Fallback** | Calibri / Arial |

---

## B. PALET WARNA RESMI

| Nama | Hex | Penggunaan |
| :--- | :--- | :--- |
| **Deep Navy** | `#070D18` | Background utama slide |
| **Dark Blue** | `#0A192F` | Background kartu bento |
| **Steel Blue** | `#0F2D4E` | Border kartu, separator |
| **Electric Cyan** | `#0EA5E9` | Aksen heading, garis, highlight |
| **Cyan Light** | `#38BDF8` | Highlight sekunder, badge |
| **Gold Amber** | `#F59E0B` | Angka besar, metrik penting |
| **White** | `#FFFFFF` | Teks body utama |
| **Silver** | `#94A3B8` | Teks body sekunder, label |
| **Success Green** | `#10B981` | Ikon centang (✅), indikator positif |
| **Alert Red** | `#EF4444` | Indikator GAP, masalah |

---

## C. TIPOGRAFI & HIERARKI TEKS

| Elemen | Font | Size | Weight | Warna |
| :--- | :--- | :--- | :--- | :--- |
| Judul Slide | Inter | 32–36pt | Bold (700) | `#FFFFFF` |
| Subjudul | Inter | 18–20pt | Medium (500) | `#94A3B8` |
| Heading Bento | Inter | 14–16pt | SemiBold (600) | `#38BDF8` |
| Body Text | Inter | 11–13pt | Regular (400) | `#FFFFFF` |
| Angka Metrik Besar | Inter | 36–48pt | ExtraBold (800) | `#F59E0B` |
| Label/Badge | Inter | 9–10pt | Medium (500) | `#0EA5E9` |
| Footer/Attribution | Inter | 8pt | Regular | `#475569` |

---

## D. LAYOUT KARTU BENTO

### D.1 Spesifikasi Kartu
- **Corner Radius**: 8–12px
- **Background**: `#0A192F` atau `#0F2D4E`
- **Border**: 1px solid `#1E3A5F`
- **Border Accent (kiri)**: 3–4px solid `#0EA5E9`
- **Padding internal**: 16px semua sisi
- **Shadow**: `0 4px 6px rgba(0,0,0,0.3)`

### D.2 Grid Layout per Slide
- **3 Kartu Horizontal** (Slide 01, 02, 06): width ~4.0in, height ~2.5in
- **4 Kartu Grid 2×2** (Slide 04): width ~5.8in per kartu, height ~2.8in
- **Kartu Full-Width** (Slide 09, 10): width ~11.5in, height ~1.0–1.5in per baris
- **6 Kartu Grid 3×2** (Slide 11): width ~3.8in, height ~2.2in

---

## E. ELEMEN DESAIN KHUSUS PER SLIDE

| Slide | Elemen Khusus | Teknis |
| :---: | :--- | :--- |
| 01 | 3 Bento card horizontal + Badge | Badge = rectangle rounded, fill `#0EA5E9` opacity 20%, stroke `#0EA5E9` |
| 02 | Diagram chain of impact vertikal | Rectangle boxes + arrow connectors, stroke `#0EA5E9` |
| 03 | 3 Kartu regulasi + Quote | Text box italic dengan left border 3px `#F59E0B` |
| 04 | 4 Pilar kartu + tree list | Kartu dengan ikon dan daftar bullet `├──` menggunakan font Courier/Mono |
| 05 | Pipeline 5 tahap horizontal | Arrow shapes + connector, warna gradient cyan ke gold |
| 06 | Hexahelix diagram + tabel mitra | Hexagon shapes manual atau teks art diagram |
| 07 | Matriks 5 baris 2 kolom | Table dengan alternating row colors |
| 08 | Timeline roadmap 3 fase | Chevron shapes horizontal 3 segment |
| 09 | Tabel + bar chart tren | Bar chart dengan label angka gold, total di bawah |
| 10 | Breakdown 6 klaster + box total | Code-style text box monospace + highlight box total |
| 11 | 6 kartu grid 3×2 | Grid dengan ikon emoji + ukuran impact |
| 12 | 3 kartu kebijakan + timeline 100 hari | Kartu dengan number badge + timeline line bawah |

---

## F. CATATAN WARNA PER ELEMEN STATUS

- ✅ **Positif/Solusi** → `#10B981` (Green)
- ❌ **Gap/Masalah** → `#EF4444` (Red, opacity 70%)
- ⭐ **Prioritas** → `#F59E0B` (Gold)
- 🔵 **Informasi** → `#0EA5E9` (Cyan)
- ⚪ **Netral** → `#94A3B8` (Silver)

---

## G. FOOTER SLIDE (Setiap Slide)

```
[Logo STP kiri]  UPTD KST SOLO TECHNOPARK | BRIDA KOTA SURAKARTA  [Nomor Slide kanan]
www.solotechnopark.id                                               00/12
```
- **Tinggi footer**: 0.35 in
- **Warna latar footer**: `#030710`
- **Separator**: 1px line `#1E3A5F`

---

## H. PANDUAN KONVERSI KE PPTX

### Menggunakan generate_pptx_v3.mjs

File generator akan dibuat di: `bahan_paparan_walikota_v3/generate_pptx_v3.mjs`

Dependensi yang diperlukan:
```
pptxgenjs: ^3.12.0 (sudah terinstall di node_modules)
```

Perintah generate:
```powershell
cd "d:\Project\GAWE\Paparan Roadmap STP"
node bahan_paparan_walikota_v3/generate_pptx_v3.mjs
```

Output:
```
d:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V3.pptx
```

---

## I. CHECKLIST KUALITAS SEBELUM PAPARAN

- [ ] Semua angka finansial terverifikasi (cross-check dengan 00_AUDIT_REKONSILIASI_V3.md)
- [ ] Font Inter terinstall di sistem presenter
- [ ] Resolusi slide 1920×1080 atau minimal 1280×720
- [ ] Slide 09 & 10: angka dalam format Rp X.XXX.XXX.XXX (titik pemisah ribuan)
- [ ] Catatan narasi pembicara dicetak terpisah atau dibuka di mode Presenter View
- [ ] Test tampilan di proyektor (pastikan warna navy tidak terlalu gelap)
- [ ] Slide 12: pastikan 3 permohonan kebijakan terbaca jelas dengan hierarki visual
- [ ] Footer nomor slide tampil di semua slide

