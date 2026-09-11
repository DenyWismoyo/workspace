import pptxgen from 'pptxgenjs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { execSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const pptx = new pptxgen();

// Set true Widescreen 16:9 (13.333 x 7.5 inches)
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';

pptx.author = 'Solo Technopark & BRIDA Surakarta';
pptx.company = 'Pemerintah Kota Surakarta';
pptx.title = 'Roadmap Solo Technopark 2026-2030';

// Color Palette
const C_NAVY = '0A192F';
const C_BLUE = '1E3A8A';
const C_CYAN = '0284C7';
const C_LIGHT_CYAN = 'E0F2FE';
const C_AMBER = 'D97706';
const C_LIGHT_AMBER = 'FEF3C7';
const C_BG = 'F8FAFC';
const C_CARD = 'FFFFFF';
const C_TEXT = '1E293B';
const C_MUTED = '64748B';
const C_WHITE = 'FFFFFF';
const C_GREEN = '059669';
const C_LIGHT_GREEN = 'D1FAE5';

const coverImgPath = path.join(__dirname, 'assets', 'slide_01_cover.jpg');
const labImgPath = path.join(__dirname, 'assets', 'slide_04_05_innovation_lab.jpg');

// Helper: Add standard slide header
function addHeader(slide, title, category, slideNum) {
  slide.background = { color: C_BG };
  
  // Top thin accent bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.333, h: 0.1, fill: { color: C_CYAN }
  });

  // Category Tag
  slide.addText(category.toUpperCase(), {
    x: 0.8, y: 0.35, w: 9.0, h: 0.3,
    fontSize: 10.5, bold: true, color: C_CYAN, fontFace: 'Arial'
  });

  // Slide Title
  slide.addText(title, {
    x: 0.8, y: 0.65, w: 11.7, h: 0.5,
    fontSize: 21, bold: true, color: C_NAVY, fontFace: 'Arial'
  });

  // Footer
  slide.addText('Solo Technopark 2026–2030 | Bahan Paparan Walikota Surakarta', {
    x: 0.8, y: 6.95, w: 8.0, h: 0.3,
    fontSize: 9.5, color: C_MUTED, fontFace: 'Arial'
  });

  slide.addText(`Slide ${slideNum} / 10`, {
    x: 11.2, y: 6.95, w: 1.3, h: 0.3,
    fontSize: 9.5, bold: true, align: 'right', color: C_MUTED, fontFace: 'Arial'
  });
}

// ==========================================
// SLIDE 01: COVER
// ==========================================
const s1 = pptx.addSlide();
if (fs.existsSync(coverImgPath)) {
  s1.addImage({ path: coverImgPath, x: 0, y: 0, w: 13.333, h: 7.5 });
  // Dark overlay
  s1.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.333, h: 7.5,
    fill: { color: C_NAVY, transparency: 40 }
  });
} else {
  s1.background = { color: C_NAVY };
}

// Government Badge
s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 0.9, w: 4.2, h: 0.45,
  fill: { color: C_CYAN }, rectRadius: 0.08
});
s1.addText('PEMERINTAH KOTA SURAKARTA & BRIDA', {
  x: 1.0, y: 0.9, w: 4.2, h: 0.45,
  fontSize: 11, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

s1.addText('STRATEGI INOVATIF PENGEMBANGAN\nSOLO TECHNOPARK 2026–2030', {
  x: 1.0, y: 1.65, w: 11.3, h: 1.8,
  fontSize: 34, bold: true, color: C_WHITE, fontFace: 'Arial', lineSpacing: 40
});

s1.addText('Dari Pengelolaan Kawasan Fisik Menuju Mesin Hilirisasi Riset & Lokomotif Ekonomi Daerah', {
  x: 1.0, y: 3.65, w: 11.3, h: 0.6,
  fontSize: 16, color: 'E2E8F0', fontFace: 'Arial'
});

// 4 Highlight Cards on Cover
const coverCards = [
  { title: '🏛️ KELEMBAGAAN BLUD', desc: 'Delivery Unit BRIDA Surakarta' },
  { title: '🔬 HILIRISASI RISET', desc: 'Jembatan Invensi Menuju Pasar' },
  { title: '📈 TARGET Rp 23,4 M', desc: 'Lompatan Pendapatan 300% (2030)' },
  { title: '🤝 95+ MITRA GLOBAL', desc: 'Ekosistem AI, Cyber, 5G Terpasang' }
];

coverCards.forEach((c, idx) => {
  const cx = 1.0 + (idx * 2.88);
  s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: cx, y: 4.9, w: 2.68, h: 1.5,
    fill: { color: C_NAVY, transparency: 25 }, line: { color: '38BDF8', width: 1.5 }, rectRadius: 0.1
  });
  s1.addText(c.title, {
    x: cx + 0.15, y: 5.1, w: 2.38, h: 0.4,
    fontSize: 11, bold: true, color: '38BDF8', fontFace: 'Arial'
  });
  s1.addText(c.desc, {
    x: cx + 0.15, y: 5.6, w: 2.38, h: 0.65,
    fontSize: 10, color: 'CBD5E1', fontFace: 'Arial'
  });
});

// ==========================================
// SLIDE 02: KELEMBAGAAN & POLA HUBUNGAN BRIDA–STP
// ==========================================
const s2 = pptx.addSlide();
addHeader(s2, 'Kedudukan Kelembagaan & Harmonisasi Peran BRIDA–STP', 'Tata Kelola Pemerintahan', 2);

const s2Cards = [
  {
    role: 'BRIDA KOTA SURAKARTA',
    badge: 'ORCHESTRATOR & REGULATOR',
    badgeColor: C_BLUE,
    points: ['Perumusan kebijakan riset & inovasi daerah', 'Penyusunan kajian strategis ekonomi kota', 'Koordinasi lintas OPD & perguruan tinggi', 'Monitoring, evaluasi & akuntabilitas program'],
    x: 0.8
  },
  {
    role: 'UPTD KST SOLO TECHNOPARK',
    badge: 'DELIVERY UNIT PELAKSANA (BLUD)',
    badgeColor: C_CYAN,
    points: ['Eksekusi teknis operasional hilirisasi riset', 'Fasilitasi inkubasi bisnis & pendampingan startup', 'Penyelenggaraan diklat vokasi industri spesifik', 'Penerapan praktik bisnis sehat & fleksibilitas BLUD'],
    x: 4.8
  },
  {
    role: 'MASYARAKAT & DUNIA USAHA',
    badge: 'BENEFICIARIES & MITRA STRATEGIS',
    badgeColor: C_GREEN,
    points: ['Penyerapan tenaga kerja terampil pemuda Solo', 'Pengembangan produk dan omzet UMKM binaan', 'Investasi industri & pemanfaatan fasilitas riset', 'Peningkatan PAD dan perputaran ekonomi daerah'],
    x: 8.8
  }
];

s2Cards.forEach(c => {
  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: c.x, y: 1.35, w: 3.7, h: 4.4,
    fill: { color: C_CARD }, line: { color: 'E2E8F0', width: 1 }, rectRadius: 0.1
  });

  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: c.x + 0.2, y: 1.55, w: 3.3, h: 0.35,
    fill: { color: c.badgeColor }, rectRadius: 0.06
  });
  s2.addText(c.badge, {
    x: c.x + 0.2, y: 1.55, w: 3.3, h: 0.35,
    fontSize: 9.5, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });

  s2.addText(c.role, {
    x: c.x + 0.2, y: 2.0, w: 3.3, h: 0.45,
    fontSize: 13, bold: true, color: C_NAVY, align: 'center', fontFace: 'Arial'
  });

  const bulletItems = c.points.map(p => ({ text: p, options: { bullet: true, fontSize: 10.5, color: C_TEXT, lineSpacing: 18 } }));
  s2.addText(bulletItems, {
    x: c.x + 0.25, y: 2.55, w: 3.2, h: 3.0, fontFace: 'Arial'
  });
});

// Legal Box Footer
s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 5.95, w: 11.7, h: 0.65,
  fill: { color: C_LIGHT_CYAN }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.08
});
s2.addText('⚖️ Landasan Regulasi: Perwali No. 15/2022 & No. 38/2022 (Pola Tata Kelola BLUD STP) • Perpres No. 106/2017 (Kawasan Sains & Teknologi) • Permendagri No. 79/2018 (Fleksibilitas BLUD)', {
  x: 1.0, y: 5.95, w: 11.3, h: 0.65,
  fontSize: 10, bold: true, color: '0369A1', align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 03: 4 PILAR PORTOFOLIO LAYANAN
// ==========================================
const s3 = pptx.addSlide();
addHeader(s3, '4 Klaster Layanan Utama Sesuai Standar Nasional KST', 'Portofolio Layanan BLUD', 3);

const s3Pillars = [
  {
    title: '1. LAYANAN TEKNIS',
    subtitle: 'Technical Training & Vocational',
    color: '2563EB',
    items: ['Diklat Manufaktur Presisi & CNC', 'Underwater Wet Welding (Las Bawah Air)', 'Otomasi Industri & Mekatronika', 'Pelatihan Manajerial & Desain Teknis']
  },
  {
    title: '2. PENGEMBANGAN TEKNOLOGI',
    subtitle: 'Engineering & Prototyping',
    color: '0D9488',
    items: ['Jasa Pemesinan & Precision Parts', 'Pembuatan Prototipe Produk UKM/IKM', 'Pengujian Standarisasi Produk', 'R&D Bersama Mitra Industri']
  },
  {
    title: '3. INKUBASI BISNIS & TEKNOLOGI',
    subtitle: 'Startup & Commercialization',
    color: 'D97706',
    items: ['Pra-Inkubasi & Seleksi Inovator', 'Pendampingan Legalitas & HKI Paten', 'Fasilitasi Pendanaan & Investor Matching', 'Fasilitas Co-Working & Business Hub']
  },
  {
    title: '4. LAYANAN PENDUKUNG',
    subtitle: 'Community & Ecosystem Enabler',
    color: '7C3AED',
    items: ['Prakerin & Uji Sertifikasi Kompetensi SMK', 'Pusat Peragaan IPTEK (Solo Science Center)', 'Penyewaan Gedung Pertemuan & Exhibition', 'Penyelenggaraan Event Ekosistem Teknologi']
  }
];

s3Pillars.forEach((p, idx) => {
  const col = idx % 2;
  const row = Math.floor(idx / 2);
  const px = 0.8 + (col * 6.0);
  const py = 1.35 + (row * 2.45);

  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px, y: py, w: 5.7, h: 2.3,
    fill: { color: C_CARD }, line: { color: p.color, width: 1.5 }, rectRadius: 0.08
  });

  s3.addShape(pptx.shapes.RECTANGLE, {
    x: px, y: py, w: 5.7, h: 0.45,
    fill: { color: p.color }
  });

  s3.addText(`${p.title} — ${p.subtitle}`, {
    x: px + 0.2, y: py + 0.05, w: 5.3, h: 0.35,
    fontSize: 10.5, bold: true, color: C_WHITE, fontFace: 'Arial'
  });

  const bItems = p.items.map(it => ({ text: it, options: { bullet: true, fontSize: 10, color: C_TEXT, lineSpacing: 16 } }));
  s3.addText(bItems, {
    x: px + 0.25, y: py + 0.55, w: 5.2, h: 1.65, fontFace: 'Arial'
  });
});

s3.addText('💡 "BLUD menggeser paradigma STP dari sekadar menyewakan aset fisik menjadi penyedia solusi teknologi yang responsif pasar."', {
  x: 0.8, y: 6.35, w: 11.7, h: 0.4,
  fontSize: 10.5, italic: true, bold: true, color: C_CYAN, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 04: PIPELINE HILIRISASI RISET
// ==========================================
const s4 = pptx.addSlide();
addHeader(s4, 'Pipeline Hilirisasi Riset: Dari Laboratorium Menuju Pasar Komersial', 'Rantai Nilai Hilirisasi', 4);

const steps = [
  { step: '01', name: 'Riset & Inventor', desc: 'Kampus (UNS, UMS) & Inovator' },
  { step: '02', name: 'Kurasi & Seleksi', desc: 'Uji Kelayakan Kebutuhan Pasar' },
  { step: '03', name: 'Prototipe & Uji', desc: 'Workshop Rekayasa Presisi' },
  { step: '04', name: 'Inkubasi Bisnis', desc: 'Paten, Legalitas & Business Plan' },
  { step: '05', name: 'Pasar Komersial', desc: 'e-Katalog LKPP & Industri' }
];

steps.forEach((st, idx) => {
  const sx = 0.8 + (idx * 2.38);
  s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: sx, y: 1.35, w: 2.18, h: 1.7,
    fill: { color: idx === 4 ? C_BLUE : C_CARD },
    line: { color: idx === 4 ? C_BLUE : C_CYAN, width: 1.5 },
    rectRadius: 0.08
  });

  s4.addText(st.step, {
    x: sx + 0.1, y: 1.45, w: 1.98, h: 0.35,
    fontSize: 15, bold: true, color: idx === 4 ? '38BDF8' : C_CYAN, fontFace: 'Arial'
  });

  s4.addText(st.name, {
    x: sx + 0.1, y: 1.85, w: 1.98, h: 0.4,
    fontSize: 11, bold: true, color: idx === 4 ? C_WHITE : C_NAVY, fontFace: 'Arial'
  });

  s4.addText(st.desc, {
    x: sx + 0.1, y: 2.25, w: 1.98, h: 0.65,
    fontSize: 9.5, color: idx === 4 ? 'CBD5E1' : C_MUTED, fontFace: 'Arial'
  });
});

// Bottom 2 Columns
s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 3.25, w: 5.7, h: 2.8,
  fill: { color: C_CARD }, line: { color: 'E2E8F0', width: 1 }, rectRadius: 0.08
});
s4.addText('🚀 CONTOH PRODUK TEKNOLOGI TERAPAN NYATA', {
  x: 1.0, y: 3.4, w: 5.3, h: 0.35,
  fontSize: 11, bold: true, color: C_CYAN, fontFace: 'Arial'
});
const techPoints = [
  'Mesin Sortasi Kopi Otomatis berbasis Vision Sensor & AI',
  'Teknologi Pertanian Cerdas (Smart Agriculture IoT)',
  'Mesin Penetas Telur Presisi untuk Kelompok Peternak',
  'Simulator Virtual Reality (VR) Pelatihan Vokasi',
  'Kendaraan Listrik Otonom (Autonomous Vehicle Prototyping)'
];
s4.addText(techPoints.map(p => ({ text: p, options: { bullet: true, fontSize: 10, color: C_TEXT, lineSpacing: 16 } })), {
  x: 1.0, y: 3.8, w: 5.3, h: 2.1, fontFace: 'Arial'
});

s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.8, y: 3.25, w: 5.7, h: 2.8,
  fill: { color: C_CARD }, line: { color: 'E2E8F0', width: 1 }, rectRadius: 0.08
});
s4.addText('🎯 TARGET OUTPUT HILIRISASI KE PASAR', {
  x: 7.0, y: 3.4, w: 5.3, h: 0.35,
  fontSize: 11, bold: true, color: C_AMBER, fontFace: 'Arial'
});
const outputPoints = [
  'Prototipe TRL 7-9 yang siap diproduksi massal',
  'Pendaftaran Hak Cipta, Merek, dan Paten (HKI)',
  'Kelahiran Perusahaan Pemula Berbasis Teknologi (Spin-off)',
  'Penyerapan Produk Inovasi Lokal dalam e-Katalog LKPP Pemkot',
  'Peningkatan Nilai Tambah Ekonomi & Daya Saing Pengrajin/IKM'
];
s4.addText(outputPoints.map(p => ({ text: p, options: { bullet: true, fontSize: 10, color: C_TEXT, lineSpacing: 16 } })), {
  x: 7.0, y: 3.8, w: 5.3, h: 2.1, fontFace: 'Arial'
});

s4.addText('⭐ Kredo: "STP adalah jembatan penghubung antara meja laboratorium kampus dan etalase pasar industri."', {
  x: 0.8, y: 6.25, w: 11.7, h: 0.4,
  fontSize: 10.5, italic: true, bold: true, color: C_NAVY, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 05: EKOSISTEM KEMITRAAN
// ==========================================
const s5 = pptx.addSlide();
addHeader(s5, 'Kekuatan Ekosistem: 95+ Mitra Korporasi Kelas Dunia di Solo', 'Kemitraan Quadruplehelix', 5);

if (fs.existsSync(labImgPath)) {
  s5.addImage({ path: labImgPath, x: 0.8, y: 1.35, w: 5.4, h: 4.8 });
  s5.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 5.35, w: 5.4, h: 0.8, fill: { color: C_NAVY, transparency: 20 }
  });
  s5.addText('SOLO TECHNOPARK INNOVATION LABS\nAI • Robotics • 5G • Cyber Security • Gaming', {
    x: 0.9, y: 5.4, w: 5.2, h: 0.7, fontSize: 9.5, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });
} else {
  s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.35, w: 5.4, h: 4.8, fill: { color: C_NAVY }, rectRadius: 0.1
  });
  s5.addText('95+ MITRA BISNIS\nAKTIF & TERUJI', {
    x: 1.0, y: 3.0, w: 5.0, h: 1.5, fontSize: 24, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });
}

const hubs = [
  { title: '🛡️ National Cyber Security Hub', desc: 'Pusat riset dan talenta pertahanan siber nasional bersama BSSN & industri.' },
  { title: '🎮 Solo Gaming Hub', desc: 'Inkubasi developer game lokal dan studio animasi berkelas ekspor.' },
  { title: '🤖 AI Experience Center & 5G Lab', desc: 'Fasilitas riset terapan kecerdasan buatan & telko bersama Indosat & Huawei.' },
  { title: '🛒 GoTo UMKM Center', desc: 'Digitalisasi ribuan pedagang & UMKM Solo ke platform e-commerce nasional.' },
  { title: '🏢 Korporasi Multinasional Terpasang', desc: 'Shopee, Garena, Mandiri, J&T, Kemenperin 3-in-1, BP3MI, dsb.' }
];

hubs.forEach((h, idx) => {
  const hy = 1.35 + (idx * 0.95);
  s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 6.5, y: hy, w: 6.0, h: 0.85,
    fill: { color: C_CARD }, line: { color: 'CBD5E1', width: 1 }, rectRadius: 0.08
  });
  s5.addText(h.title, {
    x: 6.65, y: hy + 0.08, w: 5.7, h: 0.35,
    fontSize: 11, bold: true, color: C_NAVY, fontFace: 'Arial'
  });
  s5.addText(h.desc, {
    x: 6.65, y: hy + 0.42, w: 5.7, h: 0.38,
    fontSize: 9.5, color: C_MUTED, fontFace: 'Arial'
  });
});

s5.addText('⭐ "STP sukses menarik investasi fasilitas teknologi mutakhir tanpa membebani APBD Kota Surakarta."', {
  x: 0.8, y: 6.35, w: 11.7, h: 0.4,
  fontSize: 10, italic: true, bold: true, color: C_CYAN, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 06: GAP ANALYSIS & ISU KRITIS
// ==========================================
const s6 = pptx.addSlide();
addHeader(s6, 'Gap Analysis: 5 Tantangan Utama & Rencana Aksi Solutif', 'Evaluasi Kritis & Perbaikan', 6);

const tableRows = [
  [
    { text: 'No', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Kesenjangan (GAP)', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Identifikasi Akar Masalah', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Solusi & Arah Perbaikan', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } }
  ],
  [
    { text: '1', options: { fontSize: 9.5, bold: true } },
    { text: 'Varian Pelatihan Terbatas', options: { fontSize: 9.5, bold: true, color: C_NAVY } },
    { text: 'Keterbatasan instruktur ahli & peremajaan alat mesin canggih', options: { fontSize: 9, color: C_TEXT } },
    { text: 'Sharing-equipment dengan industri & pembukaan modul AI / EV', options: { fontSize: 9, bold: true, color: C_GREEN } }
  ],
  [
    { text: '2', options: { fontSize: 9.5, bold: true } },
    { text: 'R&D Belum Maksimal', options: { fontSize: 9.5, bold: true, color: C_NAVY } },
    { text: 'Riset kampus masih sebatas jurnal, belum membaca kebutuhan industri', options: { fontSize: 9, color: C_TEXT } },
    { text: 'Membentuk Clearing House riset berbasis tantangan riil industri', options: { fontSize: 9, bold: true, color: C_GREEN } }
  ],
  [
    { text: '3', options: { fontSize: 9.5, bold: true } },
    { text: 'Investasi Lahan Terbatas', options: { fontSize: 9.5, bold: true, color: C_NAVY } },
    { text: 'Lahan kosong belum dikemas spesifik untuk menarik investor', options: { fontSize: 9, color: C_TEXT } },
    { text: 'Penyusunan Masterplan Investasi & skema KSO/BOT yang atraktif', options: { fontSize: 9, bold: true, color: C_GREEN } }
  ],
  [
    { text: '4', options: { fontSize: 9.5, bold: true } },
    { text: 'Kemandirian Finansial Rendah', options: { fontSize: 9.5, bold: true, color: C_NAVY } },
    { text: 'Hardware/sarana belum dikelola sebagai portofolio bisnis komersial', options: { fontSize: 9, color: C_TEXT } },
    { text: 'Standardisasi tarif layanan & pembentukan Strategic Business Unit (SBU)', options: { fontSize: 9, bold: true, color: C_GREEN } }
  ],
  [
    { text: '5', options: { fontSize: 9.5, bold: true } },
    { text: 'Promosi & Jejaring Kurang', options: { fontSize: 9.5, bold: true, color: C_NAVY } },
    { text: 'Jejaring mitra belum diikat dalam payung regulasi bisnis jangka panjang', options: { fontSize: 9, color: C_TEXT } },
    { text: 'Penguatan tim Corporate Partnership & regulasi kontrak bisnis yang valid', options: { fontSize: 9, bold: true, color: C_GREEN } }
  ]
];

s6.addTable(tableRows, {
  x: 0.8, y: 1.35, w: 11.7, h: 4.6,
  colW: [0.6, 2.7, 4.2, 4.2],
  border: { pt: 1, color: 'CBD5E1' },
  fill: { color: C_CARD }
});

s6.addText('⚠️ Kejujuran manajerial membaca gap adalah kunci akselerasi STP menuju kawasan sains mandiri dan berdaya saing tinggi.', {
  x: 0.8, y: 6.25, w: 11.7, h: 0.4,
  fontSize: 10, italic: true, bold: true, color: C_AMBER, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 07: ROADMAP PENGEMBANGAN 2026-2030
// ==========================================
const s7 = pptx.addSlide();
addHeader(s7, 'Roadmap Pengembangan Solo Technopark 2026–2030 (3 Fase)', 'Rencana Kerja Strategis', 7);

const phases = [
  {
    phase: 'FASE I (2026–2027)',
    theme: 'KONSOLIDASI & EKOSISTEM KERJA',
    color: C_CYAN,
    x: 0.8,
    points: [
      'Peningkatan kualitas layanan diklat',
      'Perluasan bidang vokasi masa depan (EV, AI)',
      'Penguatan SDM instruktur & sertifikasi',
      'Integrasi serapan kerja lokal & luar negeri',
      'Penyusunan regulasi kemitraan komersial'
    ]
  },
  {
    phase: 'FASE II (2028–2029)',
    theme: 'KOMERSIALISASI & FORMASI SBU',
    color: C_BLUE,
    x: 4.8,
    points: [
      'Pembentukan Strategic Business Unit (SBU)',
      'Kemandirian unit produksi rekayasa presisi',
      'Strategi corporate marketing terintegrasi',
      'Peningkatan skala hilirisasi produk UMKM',
      'Peningkatan kontribusi pendapatan BLUD'
    ]
  },
  {
    phase: 'FASE III (2029–2030)',
    theme: 'GLOBAL POSITIONING & HOLDING',
    color: C_AMBER,
    x: 8.8,
    points: [
      'Reposisi Indonesia Digital Technopark',
      'Penyesuaian tata kelola Holding Company',
      'Kemandirian finansial penuh (Self-Financing)',
      'Ekspansi kemitraan rantai pasok global',
      'Solo sebagai rujukan technopark ASEAN'
    ]
  }
];

phases.forEach(ph => {
  s7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ph.x, y: 1.35, w: 3.7, h: 4.5,
    fill: { color: C_CARD }, line: { color: ph.color, width: 2 }, rectRadius: 0.08
  });

  s7.addShape(pptx.shapes.RECTANGLE, {
    x: ph.x, y: 1.35, w: 3.7, h: 0.8,
    fill: { color: ph.color }
  });

  s7.addText(ph.phase, {
    x: ph.x + 0.1, y: 1.45, w: 3.5, h: 0.35,
    fontSize: 11, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });

  s7.addText(ph.theme, {
    x: ph.x + 0.1, y: 1.8, w: 3.5, h: 0.35,
    fontSize: 9.5, bold: true, color: 'E2E8F0', align: 'center', fontFace: 'Arial'
  });

  const bItems = ph.points.map(pt => ({ text: pt, options: { bullet: true, fontSize: 10, color: C_TEXT, lineSpacing: 18 } }));
  s7.addText(bItems, {
    x: ph.x + 0.25, y: 2.35, w: 3.2, h: 3.4, fontFace: 'Arial'
  });
});

s7.addText('🎯 Target Akhir 2030: "Centre of Innovation and Industrial Connectivity & Entrepreneur Technopark with Self-Financing"', {
  x: 0.8, y: 6.15, w: 11.7, h: 0.45,
  fontSize: 10.5, bold: true, color: C_NAVY, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 08: TARGET PENDAPATAN BLUD
// ==========================================
const s8 = pptx.addSlide();
addHeader(s8, 'Re-Setting Target Pendapatan BLUD 2026–2030 (Lompatan 3x Lipat)', 'Kemandirian Finansial', 8);

const revStreams = [
  { name: 'Kerjasama Program', pct: '51%', color: C_BLUE, desc: 'OGSCI Migas, Kemenperin 3-in-1, BP3MI, Pemda e-Katalog' },
  { name: 'Pelatihan & Sertifikasi', pct: '23%', color: C_CYAN, desc: 'Diklat Mandiri, Las Underwater, Desain & AI' },
  { name: 'Sewa Fasilitas & Kawasan', pct: '18%', color: C_AMBER, desc: 'Co-Working, Virtual Office, Gedung Pameran & Event' },
  { name: 'Produksi Komersial', pct: '8%', color: C_GREEN, desc: 'Rekayasa Manufaktur, Sparepart & Prototipe IKM' }
];

revStreams.forEach((rs, idx) => {
  const rx = 0.8 + (idx * 2.98);
  s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: rx, y: 1.35, w: 2.75, h: 1.35,
    fill: { color: C_CARD }, line: { color: rs.color, width: 1.5 }, rectRadius: 0.08
  });

  s8.addText(rs.pct, {
    x: rx + 0.15, y: 1.45, w: 2.45, h: 0.45,
    fontSize: 22, bold: true, color: rs.color, fontFace: 'Arial'
  });

  s8.addText(rs.name, {
    x: rx + 0.15, y: 1.9, w: 2.45, h: 0.35,
    fontSize: 10.5, bold: true, color: C_NAVY, fontFace: 'Arial'
  });

  s8.addText(rs.desc, {
    x: rx + 0.15, y: 2.22, w: 2.45, h: 0.42,
    fontSize: 8.5, color: C_MUTED, fontFace: 'Arial'
  });
});

const finTable = [
  [
    { text: 'Tahun Anggaran', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Target Pendapatan BLUD', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Pertumbuhan Tahunan', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } },
    { text: 'Fokus Kontributor Pendapatan Kunci', options: { bold: true, fill: { color: C_NAVY }, color: C_WHITE, fontSize: 10 } }
  ],
  [
    { text: '2026', options: { bold: true, fontSize: 9.5 } },
    { text: 'Rp 7.503.392.698', options: { bold: true, color: C_CYAN, fontSize: 10 } },
    { text: 'Baseline', options: { fontSize: 9 } },
    { text: 'Optimalisasi kemitraan industri & diklat mandiri', options: { fontSize: 9 } }
  ],
  [
    { text: '2027', options: { bold: true, fontSize: 9.5 } },
    { text: 'Rp 12.532.185.031', options: { bold: true, color: C_CYAN, fontSize: 10 } },
    { text: '+67.0%', options: { bold: true, color: C_GREEN, fontSize: 9 } },
    { text: 'Ekspansi Program OGSCI & Pelatihan Vokasi Kemenperin', options: { fontSize: 9 } }
  ],
  [
    { text: '2028', options: { bold: true, fontSize: 9.5 } },
    { text: 'Rp 16.100.658.831', options: { bold: true, color: C_CYAN, fontSize: 10 } },
    { text: '+28.5%', options: { bold: true, color: C_GREEN, fontSize: 9 } },
    { text: 'Peluncuran SBU Manufaktur & Jasa Pengujian Industri', options: { fontSize: 9 } }
  ],
  [
    { text: '2029', options: { bold: true, fontSize: 9.5 } },
    { text: 'Rp 20.615.658.831', options: { bold: true, color: C_CYAN, fontSize: 10 } },
    { text: '+28.0%', options: { bold: true, color: C_GREEN, fontSize: 9 } },
    { text: 'Sertifikasi JFT & optimalisasi tenant komersial', options: { fontSize: 9 } }
  ],
  [
    { text: '2030', options: { bold: true, fontSize: 9.5, fill: { color: C_LIGHT_AMBER } } },
    { text: 'Rp 23.468.158.831', options: { bold: true, color: C_AMBER, fontSize: 11, fill: { color: C_LIGHT_AMBER } } },
    { text: '+13.8%', options: { bold: true, color: C_GREEN, fontSize: 9, fill: { color: C_LIGHT_AMBER } } },
    { text: 'Kemandirian Finansial Penuh (Holding Unit Bisnis)', options: { bold: true, fontSize: 9, fill: { color: C_LIGHT_AMBER } } }
  ]
];

s8.addTable(finTable, {
  x: 0.8, y: 2.9, w: 11.7, h: 3.1,
  colW: [1.8, 2.7, 2.0, 5.2],
  border: { pt: 1, color: 'CBD5E1' },
  fill: { color: C_CARD }
});

s8.addText('📈 Pertumbuhan Pendapatan > 212% dalam 5 tahun: Menuju kemandirian penuh tanpa beban subsidi operasional APBD.', {
  x: 0.8, y: 6.25, w: 11.7, h: 0.4,
  fontSize: 10, bold: true, color: C_GREEN, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 09: DAMPAK BAGI KOTA SOLO
// ==========================================
const s9 = pptx.addSlide();
addHeader(s9, 'Dampak Nyata bagi Masyarakat & Pertumbuhan Ekonomi Kota Surakarta', 'Multiplier Effect', 9);

const impacts = [
  {
    title: '👨‍🏭 SDM & Lapangan Kerja',
    desc: 'Lulusan diklat vokasi langsung terserap kerja di industri mitra dengan keahlian bersertifikat (welding, digital, mesin presisi).'
  },
  {
    title: '💼 UMKM & Startup Naik Kelas',
    desc: 'Inkubasi bisnis membantu pelaku usaha lokal mendapatkan akses teknologi, legalitas HKI, sertifikasi mutu, dan perluasan pasar.'
  },
  {
    title: '🔬 Inovasi Kampus ke Pasar',
    desc: 'Menghubungkan potensi riset dosen dan mahasiswa UNS, UMS, dsb. agar menghasilkan produk nyata yang laku terjual.'
  },
  {
    title: '💻 Akses Fasilitas Teknologi',
    desc: 'Masyarakat, pelajar SMK, dan komunitas dapat memanfaatkan laboratorium modern (AI, 5G, Cyber, Gaming) tanpa biaya mahal.'
  },
  {
    title: '🏙️ Peningkatan PAD & Investasi',
    desc: 'Aktivitas bisnis di STP memicu perputaran ekonomi lokal, menyerap belanja korporasi, dan mendongkrak penerimaan pajak daerah.'
  },
  {
    title: '🌟 Reputasi Kota Inovasi',
    desc: 'Mengukuhkan posisi Kota Surakarta bukan hanya sebagai kota budaya, tetapi juga episentrum inovasi digital dan teknologi nasional.'
  }
];

impacts.forEach((imp, idx) => {
  const col = idx % 3;
  const row = Math.floor(idx / 3);
  const ix = 0.8 + (col * 4.0);
  const iy = 1.35 + (row * 2.4);

  s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ix, y: iy, w: 3.7, h: 2.2,
    fill: { color: C_CARD }, line: { color: 'E2E8F0', width: 1 }, rectRadius: 0.08
  });

  s9.addShape(pptx.shapes.RECTANGLE, {
    x: ix, y: iy, w: 3.7, h: 0.42,
    fill: { color: C_NAVY }
  });

  s9.addText(imp.title, {
    x: ix + 0.15, y: iy + 0.06, w: 3.4, h: 0.35,
    fontSize: 11, bold: true, color: C_WHITE, fontFace: 'Arial'
  });

  s9.addText(imp.desc, {
    x: ix + 0.2, y: iy + 0.55, w: 3.3, h: 1.5,
    fontSize: 9.8, color: C_TEXT, lineSpacing: 16, fontFace: 'Arial'
  });
});

s9.addText('🎯 Ukuran Keberhasilan STP: Peserta Terserap Kerja • Tenant Tumbuh • Produk Masuk Pasar • Pajak & Retribusi Naik', {
  x: 0.8, y: 6.35, w: 11.7, h: 0.4,
  fontSize: 10.5, bold: true, color: C_CYAN, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SLIDE 10: DUKUNGAN WALIKOTA & NEXT ACTION
// ==========================================
const s10 = pptx.addSlide();
addHeader(s10, 'Dukungan Kebijakan Walikota & Rencana Aksi Konkret', 'Langkah Strategis Bersama', 10);

const asks = [
  {
    num: '1',
    title: 'Regulasi Prioritas Produk Inovasi (Perwali / Surat Edaran)',
    desc: 'Mendorong seluruh Organisasi Perangkat Daerah (OPD) dan BUMD Kota Surakarta memprioritaskan penyerapan produk inovasi teknologi binaan STP lewat e-Katalog Lokal.'
  },
  {
    num: '2',
    title: 'Payung Hukum Pemanfaatan Lahan Investasi',
    desc: 'Dukungan penetapan regulasi kerja sama aset kawasan (KSO/BOT) untuk menarik investor teknologi skala besar mendirikan Tech-Office dan fasilitas riset.'
  },
  {
    num: '3',
    title: 'Mandat Orkestrasi Ekosistem Hexahelix',
    desc: 'Arahan pimpinan daerah untuk menyatukan agenda riset terapan kampus dan politeknik di Solo Raya agar berpusat dan berhilir di Solo Technopark.'
  }
];

asks.forEach((a, idx) => {
  const ay = 1.35 + (idx * 1.5);
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: ay, w: 11.7, h: 1.35,
    fill: { color: C_CARD }, line: { color: C_CYAN, width: 1.5 }, rectRadius: 0.08
  });

  s10.addShape(pptx.shapes.OVAL, {
    x: 1.1, y: ay + 0.35, w: 0.65, h: 0.65,
    fill: { color: C_CYAN }
  });

  s10.addText(a.num, {
    x: 1.1, y: ay + 0.35, w: 0.65, h: 0.65,
    fontSize: 14, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });

  s10.addText(a.title, {
    x: 1.9, y: ay + 0.15, w: 10.3, h: 0.4,
    fontSize: 12, bold: true, color: C_NAVY, fontFace: 'Arial'
  });

  s10.addText(a.desc, {
    x: 1.9, y: ay + 0.55, w: 10.3, h: 0.7,
    fontSize: 10, color: C_TEXT, lineSpacing: 15, fontFace: 'Arial'
  });
});

// Closing banner
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 6.0, w: 11.7, h: 0.75,
  fill: { color: C_NAVY }, rectRadius: 0.08
});
s10.addText('Bersama Walikota Surakarta, Solo Technopark Siap Menjadi Lokomotif Kemandirian Inovasi Indonesia 🇮🇩', {
  x: 1.0, y: 6.0, w: 11.3, h: 0.75,
  fontSize: 12, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

// ==========================================
// SAVE FILE DIRECTLY TO OUTPUT_PAPARAN
// ==========================================
const outDir = path.join(__dirname, '..', 'output_paparan');
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

const outPptx = path.join(outDir, 'PAPARAN_ROADMAP_STP_WALIKOTA.pptx');
const outPdf = path.join(outDir, 'PAPARAN_ROADMAP_STP_WALIKOTA.pdf');

pptx.writeFile({ fileName: outPptx }).then(() => {
  console.log(`✅ File PowerPoint (Widescreen 13.33 x 7.5) berhasil disimpan di: ${outPptx}`);
  
  // Otomatis konversi ke PDF via PowerPoint asli
  try {
    const psScript = path.join(__dirname, 'export_pdf.ps1');
    console.log('⏳ Mengonversi ke PDF menggunakan PowerPoint engine...');
    execSync(`powershell -ExecutionPolicy Bypass -File "${psScript}" -pptxPath "${outPptx}" -pdfPath "${outPdf}"`, { stdio: 'inherit' });
    console.log(`✅ File PDF berhasil dibuat di: ${outPdf}`);
  } catch (pdfErr) {
    console.error('⚠️ Gagal konversi otomatis ke PDF:', pdfErr.message);
  }
}).catch(err => {
  console.error('❌ Gagal membuat PowerPoint:', err);
});
