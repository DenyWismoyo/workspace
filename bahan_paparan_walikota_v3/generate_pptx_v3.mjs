import pptxgen from 'pptxgenjs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const pptx = new pptxgen();

// =========================================================================
// PRESENTATION CONFIGURATION (16:9 Widescreen)
// =========================================================================
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';

pptx.author = 'UPTD KST Solo Technopark & BRIDA Kota Surakarta';
pptx.company = 'Pemerintah Kota Surakarta';
pptx.title = 'Solo Technopark: Dari Kawasan Sains & Teknologi Menuju Mesin Impact Daerah (V3)';
pptx.subject = 'Bahan Paparan Resmi Walikota Surakarta 2026–2030';

// =========================================================================
// COLOR PALETTE (Executive Deep Navy & Bento Styling)
// =========================================================================
const C_DARK_BG     = '070D18'; // Deep Navy canvas
const C_BENTO_BG    = '0A192F'; // Dark Blue Bento card fill
const C_BENTO_BG2   = '0D213F'; // Secondary Bento card fill
const C_BENTO_ALT   = '081528'; // Alternating table/card fill
const C_BORDER      = '1E3A5F'; // Steel blue border
const C_CYAN        = '0EA5E9'; // Electric Cyan accent
const C_CYAN_LIGHT  = '38BDF8'; // Light Cyan
const C_AMBER       = 'F59E0B'; // Gold Amber for metrics & callouts
const C_AMBER_LIGHT = 'FDE68A'; // Soft amber text
const C_EMERALD     = '10B981'; // Success Green
const C_EMERALD_BG  = '064E3B'; // Dark Green tint
const C_RED         = 'EF4444'; // Gap / Alert Red
const C_WHITE       = 'FFFFFF'; // Primary white text
const C_SILVER      = '94A3B8'; // Secondary silver/muted text
const C_MUTED       = '64748B'; // Dark muted text
const C_NAVY_BAR    = '0F2D4E'; // Header / title bar fill
const C_BLUE_ACCENT = '1D4ED8'; // Royal Blue accent

const FONT_BODY  = 'Arial'; // Standard cross-platform fallback
const FONT_HEAD  = 'Arial';

// Helper: Standard Slide Header & Footer for Slides 02-12
function addHeader(slide, title, category, slideNum, subtitle = '') {
  slide.background = { color: C_DARK_BG };
  
  // Top thin accent bar (Cyan & Amber)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10.0, h: 0.08, fill: { color: C_CYAN }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.0, y: 0, w: 3.333, h: 0.08, fill: { color: C_AMBER }
  });

  // Category Tag / Pill
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 0.28, w: 3.6, h: 0.26,
    fill: { color: C_BENTO_BG }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.05
  });
  slide.addText(category.toUpperCase(), {
    x: 0.8, y: 0.28, w: 3.6, h: 0.26,
    fontSize: 8.5, bold: true, color: C_CYAN_LIGHT, align: 'center', fontFace: FONT_HEAD
  });

  // Slide Main Title
  slide.addText(title, {
    x: 0.8, y: 0.58, w: 11.733, h: 0.42,
    fontSize: 18.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });

  // Slide Subtitle if provided
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.8, y: 1.0, w: 11.733, h: 0.32,
      fontSize: 10, italic: true, color: C_SILVER, fontFace: FONT_BODY
    });
  }

  // Footer Separator Line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 7.02, w: 11.733, h: 0.02, fill: { color: C_BORDER }
  });

  // Footer Left Branding
  slide.addText('UPTD KST SOLO TECHNOPARK  |  BRIDA KOTA SURAKARTA  •  Bahan Paparan Resmi Walikota', {
    x: 0.8, y: 7.08, w: 8.5, h: 0.3,
    fontSize: 8.5, color: C_SILVER, fontFace: FONT_BODY
  });

  // Footer Right Slide Number
  slide.addText(`Slide ${String(slideNum).padStart(2, '0')} / 12`, {
    x: 11.0, y: 7.08, w: 1.533, h: 0.3,
    fontSize: 9, bold: true, align: 'right', color: C_CYAN_LIGHT, fontFace: FONT_HEAD
  });
}

// Helper: Bento Card Container with Optional Left Accent Bar
function addBentoCard(slide, x, y, w, h, options = {}) {
  const fill = options.fill || C_BENTO_BG;
  const line = options.line || C_BORDER;
  const lineWidth = options.lineWidth || 1;
  const radius = options.radius !== undefined ? options.radius : 0.08;

  // Background Box
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: line, width: lineWidth },
    rectRadius: radius
  });

  // Optional Left Accent Bar
  if (options.accentColor) {
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 0.08, h,
      fill: { color: options.accentColor },
      line: { color: options.accentColor },
      rectRadius: 0.04
    });
  }
}

// =========================================================================
// SLIDE 01: COVER & FRAMING VISI BESAR
// =========================================================================
const s1 = pptx.addSlide();
s1.background = { color: C_DARK_BG };

// Top Elegant Accent Bars
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 9.5, h: 0.12, fill: { color: C_CYAN }
});
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 9.5, y: 0, w: 3.833, h: 0.12, fill: { color: C_AMBER }
});

// Badges at Top
s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.9, y: 0.65, w: 4.8, h: 0.38,
  fill: { color: C_BENTO_BG }, line: { color: C_CYAN, width: 1.2 }, rectRadius: 0.08
});
s1.addText('PEMERINTAH KOTA SURAKARTA  •  BRIDA', {
  x: 0.9, y: 0.65, w: 4.8, h: 0.38,
  fontSize: 10, bold: true, color: C_CYAN_LIGHT, align: 'center', fontFace: FONT_HEAD
});

s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.85, y: 0.65, w: 3.2, h: 0.38,
  fill: { color: C_BENTO_BG }, line: { color: C_AMBER, width: 1.2 }, rectRadius: 0.08
});
s1.addText('ROADMAP STRATEGIS 2026–2030', {
  x: 5.85, y: 0.65, w: 3.2, h: 0.38,
  fontSize: 10, bold: true, color: C_AMBER, align: 'center', fontFace: FONT_HEAD
});

// Title & Framing
s1.addText('SOLO TECHNOPARK: DARI KAWASAN SAINS\nMENUJU MESIN IMPACT DAERAH', {
  x: 0.9, y: 1.35, w: 11.533, h: 1.75,
  fontSize: 32, bold: true, color: C_WHITE, fontFace: FONT_HEAD, lineSpacing: 38
});

s1.addText('"Roadmap Transformasi UPTD KST Surakarta: Mewujudkan Kemandirian BLUD, Hilirisasi Riset Kampus, dan Kesejahteraan Nyata Warga Kota"', {
  x: 0.9, y: 3.15, w: 11.533, h: 0.6,
  fontSize: 13.5, italic: true, color: C_AMBER_LIGHT, fontFace: FONT_BODY
});

// Horizontal Divider
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.9, y: 3.88, w: 11.533, h: 0.02, fill: { color: C_BORDER }
});

// 3 Bento Cards Horizontal
const c1w = 3.65;
const coverCards = [
  {
    icon: '🎯',
    tag: 'PARADIGMA BARU',
    title: 'Bukan Sekadar Gedung Fisik',
    desc: 'Beralih dari pengelola lahan/sewa menjadi DELIVERY UNIT inovasi BRIDA yang menghasilkan IMPACT nyata bagi masyarakat dan industri.',
    accent: C_CYAN
  },
  {
    icon: '💡',
    tag: 'KEMANDIRIAN PENUH',
    title: 'BLUD Mandiri Tanpa APBD',
    desc: 'Lompatan pendapatan terukur dari Rp 7,50 M (2026) menuju Rp 23,47 M (2030) — tumbuh +212,8% dengan 0% subsidi kas daerah.',
    accent: C_AMBER
  },
  {
    icon: '🌐',
    tag: 'HEXAHELIX & DAMPAK',
    title: '95+ Mitra Korporasi Dunia',
    desc: 'Investasi swasta murni berkolaborasi menyerap 9.000 tenaga kerja dengan komitmen afirmasi minimal 65% warga KTP Surakarta.',
    accent: C_EMERALD
  }
];

coverCards.forEach((c, idx) => {
  const cx = 0.9 + idx * 3.94;
  addBentoCard(s1, cx, 4.15, c1w, 2.5, { fill: C_BENTO_BG, accentColor: c.accent });

  // Tag Badge
  s1.addText(`${c.icon}  ${c.tag}`, {
    x: cx + 0.25, y: 4.3, w: c1w - 0.4, h: 0.35,
    fontSize: 10, bold: true, color: c.accent, fontFace: FONT_HEAD
  });

  // Title
  s1.addText(c.title, {
    x: cx + 0.25, y: 4.65, w: c1w - 0.4, h: 0.4,
    fontSize: 12.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });

  // Desc
  s1.addText(c.desc, {
    x: cx + 0.25, y: 5.1, w: c1w - 0.4, h: 1.4,
    fontSize: 10.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 15
  });
});

// Footer Notice
s1.addText('Bahan Paparan Resmi Rancangan Teknokratik UPTD KST Solo Technopark  •  Disampaikan kepada Walikota Surakarta', {
  x: 0.9, y: 6.92, w: 11.533, h: 0.32,
  fontSize: 9, color: C_MUTED, align: 'center', fontFace: FONT_BODY
});

s1.addNotes('Bapak Walikota yang kami hormati, izinkan kami mempersembahkan arah baru Solo Technopark untuk periode 2026 hingga 2030. Kita tidak lagi memposisikan Solo Technopark hanya sebagai aset fisik atau deretan gedung sewa. Dengan mandat yang Bapak berikan, kami hadir dengan proposisi yang berbeda: Solo Technopark adalah mesin impact daerah—yang mengukur keberhasilannya bukan dari saldo kas semata, melainkan dari berapa banyak warga Solo yang mendapat pekerjaan, berapa banyak riset kampus yang jadi produk bernilai, dan seberapa mandiri keuangannya tanpa beban APBD.');

// =========================================================================
// SLIDE 02: KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)
// =========================================================================
const s2 = pptx.addSlide();
addHeader(s2, 'KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)', 'Tata Kelola Pemerintahan & Kelembagaan', 2,
  'Sinergi Kokoh: BRIDA sebagai Otak Kebijakan, STP sebagai Tangan Eksekutor, Masyarakat sebagai Penerima Manfaat'
);

// Top Chain of Impact Flow Bar
addBentoCard(s2, 0.8, 1.4, 11.733, 0.85, { fill: C_BENTO_BG2, line: C_CYAN });
s2.addText([
  { text: 'RANTAI DAMPAK:   ', options: { bold: true, color: C_AMBER, fontSize: 11 } },
  { text: 'BRIDA (Otak & Regulator Kebijakan) ', options: { bold: true, color: C_WHITE, fontSize: 11 } },
  { text: ' ➔  Regulasi & Arahan  ➔ ', options: { color: C_CYAN_LIGHT, fontSize: 10.5 } },
  { text: ' UPTD KST STP (Tangan Eksekutor BLUD) ', options: { bold: true, color: C_WHITE, fontSize: 11 } },
  { text: ' ➔  Layanan & Inovasi  ➔ ', options: { color: C_CYAN_LIGHT, fontSize: 10.5 } },
  { text: ' WARGA SOLO & INDUSTRI (Benefisiari)', options: { bold: true, color: C_EMERALD, fontSize: 11 } }
], {
  x: 1.0, y: 1.4, w: 11.333, h: 0.85,
  fontFace: FONT_HEAD, align: 'center', valign: 'middle'
});

// 3 Roles Bento Columns
const rCols = [
  {
    title: 'BRIDA KOTA SURAKARTA',
    role: 'Otak & Regulator Kebijakan',
    accent: C_CYAN,
    items: [
      'Merumuskan arah kebijakan riset & inovasi daerah',
      'Penyusunan kajian teknokratik & roadmap strategis',
      'Orkestrasi ekosistem Iptek antar-OPD & kampus',
      'Monitoring dan evaluasi dampak sosial-ekonomi'
    ]
  },
  {
    title: 'UPTD KST SOLO TECHNOPARK',
    role: 'Tangan Eksekutor & Delivery Unit',
    accent: C_AMBER,
    items: [
      'Penyelenggaraan diklat vokasi & sertifikasi industri',
      'Inkubasi bisnis, alih teknologi & komersialisasi',
      'Uji coba prototipe lab presisi & penetrant testing',
      'Pengelolaan bisnis BLUD profesional & kemitraan'
    ]
  },
  {
    title: 'MASYARAKAT & DUNIA USAHA',
    role: 'Penerima Manfaat Utama (Benefisiari)',
    accent: C_EMERALD,
    items: [
      'SDM lokal kompeten & tersertifikasi kerja industri',
      'UMKM & startup teknologi tumbuh mandiri omzetnya',
      'Produk inovasi kampus terserap pasar & e-Katalog',
      'Masuknya investasi swasta & serapan kerja baru'
    ]
  }
];

rCols.forEach((col, idx) => {
  const rx = 0.8 + idx * 3.95;
  addBentoCard(s2, rx, 2.4, 3.8, 3.3, { fill: C_BENTO_BG, accentColor: col.accent });

  // Column Header Box
  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: rx + 0.15, y: 2.55, w: 3.5, h: 0.65,
    fill: C_BENTO_BG2, line: { color: col.accent, width: 1 }, rectRadius: 0.06
  });
  s2.addText(col.title, {
    x: rx + 0.2, y: 2.6, w: 3.4, h: 0.3,
    fontSize: 10.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });
  s2.addText(col.role, {
    x: rx + 0.2, y: 2.9, w: 3.4, h: 0.25,
    fontSize: 8.5, bold: true, color: col.accent, fontFace: FONT_BODY, align: 'center'
  });

  // Bullets
  const bulletText = col.items.map(it => `•  ${it}`).join('\n\n');
  s2.addText(bulletText, {
    x: rx + 0.25, y: 3.35, w: 3.3, h: 2.2,
    fontSize: 9.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 14
  });
});

// Bottom Legal Basis Bar
addBentoCard(s2, 0.8, 5.85, 11.733, 0.95, { fill: C_BENTO_BG2, line: C_BORDER });
s2.addText('LANDASAN HUKUM KELEMBAGAAN KOTA SURAKARTA:', {
  x: 1.0, y: 5.95, w: 11.333, h: 0.25,
  fontSize: 9, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});
s2.addText('• Perwali Surakarta No. 15 Tahun 2022: Penetapan kedudukan UPTD KST Solo Technopark sebagai unit pelaksana teknis BRIDA.\n• Perwali Surakarta No. 38 Tahun 2022: Penetapan Pola Tata Kelola Badan Layanan Umum Daerah (BLUD) UPTD KST Solo Technopark.', {
  x: 1.0, y: 6.22, w: 11.333, h: 0.5,
  fontSize: 9.2, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 13
});

s2.addNotes('Posisi kelembagaan kita sangat jelas dan terstruktur, Bapak Walikota. BRIDA adalah arsitek kebijakan—merumuskan arah, memonitor dampak, dan mengorkestrasi ekosistem Iptek kota. Solo Technopark adalah insinyur pelaksana—menguji teknologi, melatih SDM, menginkubasi startup, dan menghubungkan riset kampus ke pasar industri. Dan pada ujung rantai ini, yang merasakan manfaat langsung adalah warga Solo—mendapat pekerjaan, UMKM mereka naik kelas, dan produk lokal bisa bersaing.');

// =========================================================================
// SLIDE 03: MANDAT BLUD & FONDASI REGULASI KOKOH
// =========================================================================
const s3 = pptx.addSlide();
addHeader(s3, 'FONDASI HUKUM YANG KOKOH: MANDAT & FLEKSIBILITAS BLUD', 'Landasan Regulasi & Tata Kelola', 3,
  'Tiga Pilar Regulasi yang Memberi Kewenangan, Fleksibilitas Bisnis Sehat, dan Arah Strategis KST'
);

const regs = [
  {
    law: 'PERMENDAGRI 79/2018',
    sub: 'Tata Kelola Keuangan BLUD',
    quote: '"Fleksibilitas BLUD adalah keleluasaan dalam pola pengelolaan keuangan dengan menerapkan praktik bisnis yang sehat untuk meningkatkan layanan kepada masyarakat tanpa mencari keuntungan semata."',
    accent: C_CYAN,
    points: [
      'Keleluasaan bermitra dengan swasta tanpa birokrasi kaku APBD',
      'Mengelola dan memutar pendapatan fungsional jasa secara mandiri',
      'Rekrutmen tenaga ahli & instruktur profesional sesuai kebutuhan pasar'
    ]
  },
  {
    law: 'PERPRES 106/2017',
    sub: 'Kawasan Sains & Teknologi (KST)',
    quote: '"KST berfungsi sebagai wahana kerja sama riset berkelanjutan antara Pemerintah, PT, dan Industri, wajib menyediakan 4 layanan: Teknis, Teknologi, Inkubasi, & Pendukung."',
    accent: C_AMBER,
    points: [
      'Mandat nasional fasilitasi hilirisasi & spin-off perusahaan riset',
      'Kewajiban menyediakan 4 pilar layanan iptek terintegrasi',
      'Landasan kolaborasi hexahelix multi-stakeholder formal'
    ]
  },
  {
    law: 'PERWALI 15 & 38/2022',
    sub: 'Struktur UPTD & Pola Tata Kelola',
    quote: '"Menetapkan struktur UPTD KST Solo Technopark dalam naungan BRIDA Kota Surakarta dengan pola tata kelola BLUD profesional berorientasi layanan publik."',
    accent: C_EMERALD,
    points: [
      'Payung hukum operasional tingkat kota yang sah dan mengikat',
      'Pembagian peran teknis yang presisi antara BRIDA dan STP',
      'Dasar legalitas pemanfaatan aset kawasan & kerjasama industri'
    ]
  }
];

regs.forEach((r, idx) => {
  const rx = 0.8 + idx * 3.95;
  addBentoCard(s3, rx, 1.45, 3.8, 4.4, { fill: C_BENTO_BG, accentColor: r.accent });

  // Badge Law
  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: rx + 0.2, y: 1.65, w: 3.4, h: 0.55,
    fill: C_BENTO_BG2, line: { color: r.accent, width: 1 }, rectRadius: 0.05
  });
  s3.addText(r.law, {
    x: rx + 0.2, y: 1.68, w: 3.4, h: 0.28,
    fontSize: 11, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });
  s3.addText(r.sub, {
    x: rx + 0.2, y: 1.95, w: 3.4, h: 0.22,
    fontSize: 8.5, bold: true, color: r.accent, fontFace: FONT_BODY, align: 'center'
  });

  // Quote Box
  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: rx + 0.2, y: 2.35, w: 3.4, h: 1.45,
    fill: C_DARK_BG, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
  });
  s3.addText(r.quote, {
    x: rx + 0.3, y: 2.42, w: 3.2, h: 1.3,
    fontSize: 8.8, italic: true, color: C_AMBER_LIGHT, fontFace: FONT_BODY, lineSpacing: 12
  });

  // Points
  const pText = r.points.map(p => `✅  ${p}`).join('\n\n');
  s3.addText(pText, {
    x: rx + 0.2, y: 3.95, w: 3.4, h: 1.75,
    fontSize: 9.2, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 13
  });
});

// Bottom Synthesis Banner
addBentoCard(s3, 0.8, 6.0, 11.733, 0.8, { fill: C_BENTO_BG2, line: C_CYAN });
s3.addText('KESIMPULAN YURIDIS: Solo Technopark telah memiliki legitimasi hukum lengkap dari tingkat nasional (Perpres & Permendagri) hingga regulasi kepala daerah (Perwali) untuk bermanuver cepat, profesional, dan akuntabel.', {
  x: 1.0, y: 6.0, w: 11.333, h: 0.8,
  fontSize: 10.5, bold: true, color: C_WHITE, align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s3.addNotes('Sebelum bicara angka dan target, Bapak Walikota, perlu kami tegaskan: fondasi hukum kita sangat kuat. Permendagri 79/2018 memberi kita fleksibilitas BLUD—artinya kita bisa bergerak lincah bermitra dengan Shopee, Garena, atau OGSCI tanpa terjebak rigiditas pengadaan APBD. Perpres 106/2017 memberi mandat nasional: kita harus inkubasi startup, fasilitasi spin-off perusahaan berbasis riset. Dan Perwali 15 serta 38/2022 dari Bapak sendiri telah mengokohkan posisi STP sebagai delivery unit profesional di bawah BRIDA.');

// =========================================================================
// SLIDE 04: EMPAT PILAR LAYANAN TERINTEGRASI BLUD STP
// =========================================================================
const s4 = pptx.addSlide();
addHeader(s4, 'EMPAT PILAR LAYANAN TERINTEGRASI: MEMENUHI MANDAT KST NASIONAL', 'Mandat Layanan Perpres 106/2017', 4,
  'Portofolio Layanan Terintegrasi yang Menjawab Kebutuhan Industri Spesialis & Masyarakat Kota Surakarta'
);

const pilars = [
  {
    code: 'PILAR 1',
    name: 'LAYANAN TEKNIS (DIKLAT VOKASI INDUSTRI)',
    icon: '🔧',
    accent: C_CYAN,
    items: [
      'Mekanik, Otomasi & Desain Manufaktur Presisi',
      'Welding Intensif (Pengelasan Konstruksi & Pipa)',
      'OGSCI: Oil, Gas, Coal, Shipbuilding & Infra Training',
      'Underwater Wet Welding (Keahlian Langka Nasional)',
      'Teknologi Kendaraan Listrik (EV Conversion Lab)',
      'Cyber Security & Artificial Intelligence Practitioner'
    ]
  },
  {
    code: 'PILAR 2',
    name: 'PENGEMBANGAN TEKNOLOGI (REKAYASA MANUFAKTUR)',
    icon: '⚙️',
    accent: C_AMBER,
    items: [
      'Jasa Layanan Rekayasa Industri & Permesinan Presisi',
      'Produksi Komponen Presisi (Precision Parts Industri)',
      'Pembuatan & Uji Prototipe IKM/UKM Kota Surakarta',
      'Non-Destructive Testing: Uji Penetrant Testing (NDT)',
      'Konsultasi Laboratorium Terpadu & Data Center',
      'Standarisasi Mutu & Sertifikasi TKDN / SNI Produk'
    ]
  },
  {
    code: 'PILAR 3',
    name: 'LAYANAN INKUBASI (HILIRISASI KE PASAR)',
    icon: '🚀',
    accent: C_EMERALD,
    items: [
      'Pra-Inkubasi & Inkubasi Bisnis Startup Berbasis Riset',
      'Fasilitasi HKI, Paten Produk, & Hak Cipta Inovator',
      'Hilirisasi Riset Perguruan Tinggi ke Pasar Industri',
      'National Cyber Security Hub & AI Experience Center',
      'Digital Technopark, Gaming Hub, & GoTo UMKM Center',
      'Co-Working Space, Virtual Office & Akses Permodalan'
    ]
  },
  {
    code: 'PILAR 4',
    name: 'LAYANAN PENDUKUNG (COMMUNITY & EDUCATION)',
    icon: '🎓',
    accent: C_CYAN_LIGHT,
    items: [
      'Prakerin Kerja Industri Siswa SMK Se-Solo Raya',
      'Tempat Uji Kompetensi (TUK Mandiri) & Sertifikasi JFT',
      'Solo Science Center (Pusat Edukasi IPTEK & Wisata Sains)',
      'Pemanfaatan Sarana Kawasan, Gedung Expo & MICE',
      'Kemitraan Komunitas Kreatif & Kunjungan Industri',
      'Pusat Kolaborasi Multipihak Pemkot Surakarta'
    ]
  }
];

pilars.forEach((p, idx) => {
  const col = idx % 2;
  const row = Math.floor(idx / 2);
  const px = 0.8 + col * 5.95;
  const py = 1.45 + row * 2.5;

  addBentoCard(s4, px, py, 5.75, 2.35, { fill: C_BENTO_BG, accentColor: p.accent });

  // Header Box
  s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.15, y: py + 0.12, w: 5.45, h: 0.42,
    fill: C_BENTO_BG2, line: { color: p.accent, width: 1 }, rectRadius: 0.05
  });
  s4.addText(`${p.icon}  ${p.code}: ${p.name}`, {
    x: px + 0.25, y: py + 0.12, w: 5.3, h: 0.42,
    fontSize: 9.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD, valign: 'middle'
  });

  // Items (2 columns inside card)
  const leftItems = p.items.slice(0, 3).map(it => `• ${it}`).join('\n\n');
  const rightItems = p.items.slice(3, 6).map(it => `• ${it}`).join('\n\n');

  s4.addText(leftItems, {
    x: px + 0.25, y: py + 0.6, w: 2.65, h: 1.65,
    fontSize: 8.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 13
  });
  s4.addText(rightItems, {
    x: px + 2.95, y: py + 0.6, w: 2.65, h: 1.65,
    fontSize: 8.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 13
  });
});

// Bottom Tagline
addBentoCard(s4, 0.8, 6.48, 11.733, 0.45, { fill: C_BENTO_BG2, line: C_BORDER });
s4.addText('"BLUD menggeser paradigma: Dari sekadar pengelolaan kawasan fisik menjadi penyedia solusi teknologi yang responsif terhadap kebutuhan industri dan masyarakat."', {
  x: 1.0, y: 6.48, w: 11.333, h: 0.45,
  fontSize: 9.5, italic: true, bold: true, color: C_AMBER, align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s4.addNotes('Empat pilar ini bukan sekadar daftar layanan, Bapak Walikota. Ini adalah ekosistem yang saling mengunci. Di pilar vokasi, kita melatih keahlian langka seperti underwater welding—keahlian yang satu pesertanya bisa langsung terserap industri migas dengan gaji di atas Rp 15 juta per bulan. Di inkubasi, kita tidak sekadar menampung startup, kita mendampinginya hingga produknya masuk pasar. Dan semuanya berjalan bersama, membentuk satu ekosistem yang mengubah potensi menjadi nilai ekonomi nyata bagi kota Surakarta.');

// =========================================================================
// SLIDE 05: PIPELINE HILIRISASI RISET: DARI LAB KE PASAR
// =========================================================================
const s5 = pptx.addSlide();
addHeader(s5, 'PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR NYATA', 'Mekanisme Hilirisasi Inovasi', 5,
  'Solo Technopark sebagai Jembatan: Menghubungkan Riset Kampus Menuju Produk Bernilai Ekonomi bagi Industri'
);

// 5 Pipeline Stages Horizontal Cards
const stages = [
  {
    num: '01',
    name: 'RISET & INVENSI',
    sub: 'Kampus & Litbang',
    desc: 'Invensi UNS, UMS, ISI, Poltek Solo Raya. Dosen, peneliti, dan mahasiswa berinovasi.',
    accent: C_CYAN
  },
  {
    num: '02',
    name: 'SELEKSI & KURASI',
    sub: 'STP & BRIDA',
    desc: 'Kurasi komersial & relevansi kebutuhan industri serta solusi tantangan kota Solo.',
    accent: C_CYAN_LIGHT
  },
  {
    num: '03',
    name: 'PROTOTYPING',
    sub: 'Lab Presisi STP',
    desc: 'Uji teknis, penetrant testing, validasi mutu produk, dan standarisasi SNI/TKDN.',
    accent: C_AMBER
  },
  {
    num: '04',
    name: 'INKUBASI BISNIS',
    sub: 'Inkubator STP',
    desc: 'Penyusunan business plan, proteksi HKI/Paten, akses permodalan & jejaring offtaker.',
    accent: C_EMERALD
  },
  {
    num: '05',
    name: 'KOMERSIALISASI',
    sub: 'Pasar & Industri',
    desc: 'Penetrasi e-Katalog Lokal OPD, captive market pemda, industri nasional & ekspor.',
    accent: C_AMBER
  }
];

stages.forEach((st, idx) => {
  const sx = 0.8 + idx * 2.37;
  addBentoCard(s5, sx, 1.4, 2.25, 2.3, { fill: C_BENTO_BG, accentColor: st.accent });

  // Number Badge
  s5.addText(st.num, {
    x: sx + 0.15, y: 1.5, w: 0.8, h: 0.35,
    fontSize: 16, bold: true, color: st.accent, fontFace: FONT_HEAD
  });
  s5.addText(st.sub, {
    x: sx + 0.9, y: 1.55, w: 1.25, h: 0.25,
    fontSize: 8, bold: true, color: C_SILVER, fontFace: FONT_HEAD, align: 'right'
  });

  s5.addText(st.name, {
    x: sx + 0.15, y: 1.9, w: 2.0, h: 0.45,
    fontSize: 10, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s5.addText(st.desc, {
    x: sx + 0.15, y: 2.4, w: 2.0, h: 1.15,
    fontSize: 8.5, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 11
  });

  // Connector Arrow (except last)
  if (idx < 4) {
    s5.addText('➔', {
      x: sx + 2.22, y: 2.3, w: 0.2, h: 0.4,
      fontSize: 12, bold: true, color: C_CYAN, fontFace: FONT_HEAD, align: 'center'
    });
  }
});

// Middle Section: 2 Bento Blocks (Portfolio & Active Hubs)
// Block Left: Table Portofolio Teknologi Riil
addBentoCard(s5, 0.8, 3.85, 6.8, 2.9, { fill: C_BENTO_BG, line: C_BORDER });
s5.addText('CONTOH TEKNOLOGI YANG DIHILIRKAN & DIINKUBASI DI STP (DATA RIIL)', {
  x: 1.0, y: 3.95, w: 6.4, h: 0.3,
  fontSize: 9.5, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

const protoTableData = [
  [
    { text: 'KLASTER TEKNOLOGI', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'CONTOH PRODUK INOVASI', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'STATUS TAHAPAN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_SILVER } },
    { text: 'Mesin Sortasi Biji Kopi Otomatis', options: { color: C_WHITE, bold: true } },
    { text: 'Prototype Validasi', options: { color: C_CYAN_LIGHT, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_SILVER } },
    { text: 'Teknologi Pertanian Presisi IoT & AI', options: { color: C_WHITE, bold: true } },
    { text: 'Uji Lapangan', options: { color: C_CYAN_LIGHT, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_SILVER } },
    { text: 'Mesin Tetas Telur Pintar Mandiri', options: { color: C_WHITE, bold: true } },
    { text: 'Siap e-Katalog', options: { color: C_EMERALD, align: 'center', bold: true } }
  ],
  [
    { text: 'Deep Tech Mobility', options: { color: C_SILVER } },
    { text: 'Autonomous Vehicle Elektrik Kampus', options: { color: C_WHITE, bold: true } },
    { text: 'R&D Bersama', options: { color: C_AMBER, align: 'center' } }
  ],
  [
    { text: 'GovTech & Pelatihan', options: { color: C_SILVER } },
    { text: 'VR Pelatihan Industri & Chatbot AI OPD', options: { color: C_WHITE, bold: true } },
    { text: 'Komersial Aktif', options: { color: C_EMERALD, align: 'center', bold: true } }
  ]
];

s5.addTable(protoTableData, {
  x: 1.0, y: 4.3, w: 6.4, h: 2.3,
  colW: [1.8, 3.2, 1.4],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG2 },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.35, 0.38, 0.38, 0.38, 0.38, 0.38]
});

// Block Right: Ekosistem Pendukung Aktif
addBentoCard(s5, 7.75, 3.85, 4.783, 2.9, { fill: C_BENTO_BG, line: C_CYAN });
s5.addText('EKOSISTEM FASILITAS PENDUKUNG AKTIF', {
  x: 7.95, y: 3.95, w: 4.383, h: 0.3,
  fontSize: 9.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const hubs = [
  { icon: '🛡️', name: 'National Cyber Security Hub', desc: 'Pusat ketahanan digital & talenta siber' },
  { icon: '🎮', name: 'Gaming Hub & Digital Studio', desc: 'Pengembangan game & konten kreatif' },
  { icon: '🤖', name: 'AI Experience Center', desc: 'Lab aplikasi kecerdasan buatan terapan' },
  { icon: '🛒', name: 'GoTo UMKM Center', desc: 'Akselerasi digitalisasi pedagang lokal' },
  { icon: '🔬', name: 'Laboratorium Rekayasa Presisi', desc: 'Fasilitas fabrikasi komponen & NDT uji' }
];

hubs.forEach((hb, idx) => {
  const hy = 4.32 + idx * 0.46;
  s5.addText(`${hb.icon}  ${hb.name}`, {
    x: 7.95, y: hy, w: 4.383, h: 0.22,
    fontSize: 9, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s5.addText(hb.desc, {
    x: 8.35, y: hy + 0.2, w: 3.983, h: 0.22,
    fontSize: 8, color: C_SILVER, fontFace: FONT_BODY
  });
});

s5.addNotes('Bapak Walikota, kami ingin menghapus fenomena riset lemari arsip—di mana penelitian kampus hebat tapi berakhir hanya sebagai tumpukan laporan jurnal. Pipeline hilirisasi kami memutus kebuntuan itu. Riset dosen UNS tentang mesin sortasi kopi? Kita ambil, kita uji prototipenya di lab presisi STP, kita bantu urus patennya, kita carikan investor atau offtaker industrinya. Inilah nilai sejati Solo Technopark: jembatan dari ilmu pengetahuan menuju perputaran uang dan lapangan kerja bagi warga Solo.');

// =========================================================================
// SLIDE 06: EKOSISTEM HEXAHELIX: 95+ MITRA KELAS DUNIA
// =========================================================================
const s6 = pptx.addSlide();
addHeader(s6, 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD', 'Jejaring Kemitraan Strategis', 6,
  'Kepercayaan Korporasi Global & Nasional yang Menanamkan Investasi Langsung di Jantung Kota Surakarta'
);

// Left Box: Hexahelix Pillars
addBentoCard(s6, 0.8, 1.45, 4.0, 5.35, { fill: C_BENTO_BG, accentColor: C_CYAN });
s6.addText('6 AKTOR KOLABORASI HEXAHELIX', {
  x: 1.05, y: 1.6, w: 3.5, h: 0.35,
  fontSize: 11, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

const helixActors = [
  { icon: '🎓', name: 'AKADEMISI', desc: 'UNS, UMS, ISI, Poltek — Pemasok riset, invensi, dan peneliti' },
  { icon: '🏭', name: 'INDUSTRI & BISNIS', desc: '95+ Korporasi — Offtaker, penyedia modal kerja, & transfer teknologi' },
  { icon: '🏛️', name: 'PEMERINTAH', desc: 'Pemkot Solo & BRIDA — Regulator, fasilitator kebijakan & pembeli awal' },
  { icon: '🤝', name: 'KOMUNITAS', desc: 'Asosiasi UMKM, SMK, Komunitas Kreatif & Startup lokal' },
  { icon: '🌐', name: 'MEDIA MASSA', desc: 'Publikasi nasional & penguatan reputasi branding kota Solo' },
  { icon: '🌍', name: 'INTERNASIONAL', desc: 'BP3MI, OGSCI, Mitra Global — Akses sertifikasi & pasar dunia' }
];

helixActors.forEach((ha, idx) => {
  const ay = 2.05 + idx * 0.76;
  s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 1.05, y: ay, w: 3.5, h: 0.68,
    fill: C_BENTO_BG2, line: { color: C_BORDER, width: 1 }, rectRadius: 0.05
  });
  s6.addText(`${ha.icon}  ${ha.name}`, {
    x: 1.15, y: ay + 0.06, w: 3.3, h: 0.25,
    fontSize: 9.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s6.addText(ha.desc, {
    x: 1.15, y: ay + 0.3, w: 3.3, h: 0.34,
    fontSize: 8.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 10
  });
});

// Right Top Box: Table Mitra Korporasi Kelas Dunia
addBentoCard(s6, 5.0, 1.45, 7.533, 3.8, { fill: C_BENTO_BG, line: C_BORDER });
s6.addText('MITRA INDUSTRI KELAS DUNIA YANG AKTIF BERINVESTASI DI STP', {
  x: 5.2, y: 1.6, w: 7.133, h: 0.3,
  fontSize: 10.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const partnerTableData = [
  [
    { text: 'SEKTOR INDUSTRI', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'MITRA STRATEGIS', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'KONTRIBUSI NYATA BAGI KAWASAN STP', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } }
  ],
  [
    { text: 'E-Commerce & Logistik', options: { color: C_SILVER } },
    { text: 'Shopee Indonesia', options: { color: C_WHITE, bold: true } },
    { text: 'Infrastruktur UMKM Hub, Co-location, sewa Rp 544 jt/th', options: { color: C_SILVER } }
  ],
  [
    { text: 'Gaming & Kreatif', options: { color: C_SILVER } },
    { text: 'Garena', options: { color: C_WHITE, bold: true } },
    { text: 'Gaming studio, beasiswa talenta digital, turnamen esports', options: { color: C_SILVER } }
  ],
  [
    { text: 'Perbankan & Fintech', options: { color: C_SILVER } },
    { text: 'Bank Mandiri', options: { color: C_WHITE, bold: true } },
    { text: 'Mandiri Innovation Hub, fasilitasi modal startup binaan', options: { color: C_SILVER } }
  ],
  [
    { text: 'Telekomunikasi', options: { color: C_SILVER } },
    { text: 'Indosat Ooredoo Hutchison', options: { color: C_WHITE, bold: true } },
    { text: 'Laboratorium 5G, konektivitas fiber optik kawasan', options: { color: C_SILVER } }
  ],
  [
    { text: 'Ride-Hailing & UMKM', options: { color: C_SILVER } },
    { text: 'GoTo (Gojek-Tokopedia)', options: { color: C_WHITE, bold: true } },
    { text: 'GoTo UMKM Center, pelatihan onboarding merchant lokal', options: { color: C_SILVER } }
  ],
  [
    { text: 'Migas & Industri Berat', options: { color: C_SILVER } },
    { text: 'OGSCI Consortium', options: { color: C_WHITE, bold: true } },
    { text: 'Pelatihan migas internasional: Target Rp 6,25 M (2030)', options: { color: C_AMBER, bold: true } }
  ]
];

s6.addTable(partnerTableData, {
  x: 5.2, y: 1.95, w: 7.133, h: 3.1,
  colW: [2.0, 2.1, 3.033],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG2 },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.35, 0.42, 0.42, 0.42, 0.42, 0.42, 0.45]
});

// Right Bottom Box: Zero APBD Burden Callout
addBentoCard(s6, 5.0, 5.4, 7.533, 1.4, { fill: C_BENTO_BG2, line: C_EMERALD });
s6.addText('💎  PRINSIP ZERO APBD BURDEN: SELURUH FASILITAS CANGGIH DIBANGUN OLEH INVESTASI SWASTA', {
  x: 5.2, y: 5.5, w: 7.133, h: 0.3,
  fontSize: 9.8, bold: true, color: C_EMERALD, fontFace: FONT_HEAD
});
s6.addText('Laboratorium AI, Studio Gaming, 5G Experience, hingga Cyber Security Hub tidak dibangun dengan membebani belanja modal APBD Kota Surakarta, melainkan murni melalui skema Public-Private Partnership (PPP). Pemkot menyediakan wahana dan arah kebijakan, korporasi global menanamkan investasi teknologi.', {
  x: 5.2, y: 5.82, w: 7.133, h: 0.85,
  fontSize: 9.2, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 13
});

s6.addNotes('Yang membuat Solo Technopark unik adalah ekosistem hexahelix-nya—melibatkan akademisi, industri, pemerintah, komunitas, media, hingga mitra internasional. Dan satu hal yang paling membanggakan: 95 lebih korporasi kelas dunia telah memercayakan fasilitas dan proyeknya di sini—Shopee, Garena, Mandiri, GoTo, Indosat—tanpa satu rupiah pun dari APBD Kota Surakarta. Ini adalah bukti nyata kepercayaan swasta global terhadap visi kota Surakarta.');

// =========================================================================
// SLIDE 07: GAP ANALYSIS JUJUR & 5 TEROBOSAN SOLUSI
// =========================================================================
const s7 = pptx.addSlide();
addHeader(s7, 'GAP ANALYSIS JUJUR & 5 TEROBOSAN SOLUSI 2026–2030', 'Diagnostik & Strategi Percepatan', 7,
  'Audit Transparan: Mengidentifikasi Titik Lemah Eksisting dan Menetapkan Solusi Terobosan Terukur'
);

const gapTableData = [
  [
    { text: 'NO', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'GAP EKSISTING', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'AKAR PERMASALAHAN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: '5 SOLUSI TEROBOSAN STRATEGIS 2026–2030', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } }
  ],
  [
    { text: '1', options: { bold: true, align: 'center', color: C_RED } },
    { text: 'Jenis Pelatihan Masih Terbatas', options: { bold: true, color: C_WHITE } },
    { text: 'Keterbatasan instruktur ahli internal & sarana diklat berbasis industri masa depan.', options: { color: C_SILVER } },
    { text: 'Gandeng korporasi global (OGSCI, Kemenperin) buka kelas spesialis: Welder Bawah Air, AI Engineer, Teknisi Baterai EV, & Drone Operator.', options: { color: C_WHITE } }
  ],
  [
    { text: '2', options: { bold: true, align: 'center', color: C_RED } },
    { text: 'R&D Belum Mengarah Hilirisasi', options: { bold: true, color: C_WHITE } },
    { text: 'Riset kampus berjalan terisolasi, kurang terpetakan, dan belum berorientasi pasar komersial.', options: { color: C_SILVER } },
    { text: 'Bentuk Technology Transfer Office (TTO) & kurasi riset kampus UNS/UMS berbasis kebutuhan industri nyata dan problem kota.', options: { color: C_WHITE } }
  ],
  [
    { text: '3', options: { bold: true, align: 'center', color: C_RED } },
    { text: 'Investasi Kawasan Terbatas', options: { bold: true, color: C_WHITE } },
    { text: 'Lahan kosong belum optimal dikerjasamakan; promosi investor teknologi belum sistematis.', options: { color: C_SILVER } },
    { text: 'Zonasi investasi teknologi terpadu & skema Long-Term Revenue Sharing (10–25 tahun) bagi investor swasta global.', options: { color: C_WHITE } }
  ],
  [
    { text: '4', options: { bold: true, align: 'center', color: C_RED } },
    { text: 'Kemandirian Finansial Rendah', options: { bold: true, color: C_WHITE } },
    { text: 'Pendapatan masih bertumpu pada sewa lahan; hardware lab belum menghasilkan revenue optimal.', options: { color: C_SILVER } },
    { text: 'Diversifikasi 6 revenue stream: Uji NDT Penetrant, TUK/JFT mandiri, Rekayasa Manufaktur, Co-Working, & Startup Binaan.', options: { color: C_WHITE } }
  ],
  [
    { text: '5', options: { bold: true, align: 'center', color: C_RED } },
    { text: 'Promosi & Penyerapan Rendah', options: { bold: true, color: C_WHITE } },
    { text: 'Networking kurang di-maintenance; belum ada regulasi penyerapan produk inovasi di pasar lokal.', options: { color: C_SILVER } },
    { text: 'Penerbitan Instruksi Walikota untuk e-Katalog Lokal Afirmatif bagi OPD Pemkot menyerap produk/layanan tenant STP.', options: { color: C_AMBER, bold: true } }
  ]
];

s7.addTable(gapTableData, {
  x: 0.8, y: 1.45, w: 11.733, h: 4.4,
  colW: [0.6, 2.5, 3.6, 5.033],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.38, 0.78, 0.78, 0.78, 0.78, 0.82]
});

// Bottom Priority Callout (Quick Win Inkwal)
addBentoCard(s7, 0.8, 6.0, 11.733, 0.85, { fill: C_BENTO_BG2, line: C_AMBER });
s7.addText('⭐ PRIORITAS #5: INSTRUKSI WALIKOTA e-KATALOG AFIRMATIF (QUICK WIN 30 HARI)', {
  x: 1.0, y: 6.08, w: 11.333, h: 0.28,
  fontSize: 10, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s7.addText('Target: Rp 2,25 Miliar serapan belanja OPD Pemkot di 2030. Memberikan pasar captive awal bagi 80+ produk tenant dan startup binaan STP tanpa perlu alokasi pagu belanja baru di APBD.', {
  x: 1.0, y: 6.36, w: 11.333, h: 0.42,
  fontSize: 9, color: C_WHITE, fontFace: FONT_BODY
});

s7.addNotes('Kami berdiri di sini dengan kejujuran penuh, Bapak Walikota. Lima gap ini adalah tantangan nyata yang kami identifikasi sendiri. Pendapatan kita masih terlalu bergantung sewa lahan dan satu program unggulan. Riset kampus belum terhubung ke pasar. Promosi produk inovasi kita masih lemah regulasinya. Namun kami tidak datang hanya dengan masalah—kami hadir dengan 5 solusi terobosan yang terukur dan ada satu yang sangat bergantung pada 1 lembar Instruksi Walikota: e-Katalog Lokal yang mengarahkan OPD Pemkot menyerap produk binaan STP.');

// =========================================================================
// SLIDE 08: ROADMAP 3 FASE TRANSFORMASI 2027–2030
// =========================================================================
const s8 = pptx.addSlide();
addHeader(s8, 'ROADMAP 3 FASE TRANSFORMASI: DARI CENTRE MENUJU HOLDING INOVASI DAERAH', 'Peta Jalan Strategis Kelembagaan', 8,
  'Milestone Terukur dan Bertahap Menuju Kemandirian Finansial Penuh Tanpa Subsidi Belanja APBD 2030'
);

// 3 Phase Horizontal Cards
const phases = [
  {
    phase: 'FASE 1 (2026–2027)',
    title: 'CENTRE OF INNOVATION & CONNECTIVITY',
    target: 'Rp 7,50 M ➔ Rp 12,53 M',
    accent: C_CYAN,
    points: [
      'Pondasi tata kelola kelembagaan BLUD akuntabel',
      'Akreditasi laboratorium industri & standarisasi uji',
      'Sertifikasi mutu diklat vokasi & kemitraan OGSCI',
      'Pembentukan Konsorsium Riset 5 Kampus Solo Raya',
      'Aktivasi TUK Mandiri & program SMK GoGlobal'
    ]
  },
  {
    phase: 'FASE 2 (2028–2029)',
    title: 'ENTREPRENEUR TECHNOPARK (SBU MANDIRI)',
    target: 'Rp 12,53 M ➔ Rp 20,62 M',
    accent: C_AMBER,
    points: [
      'Pembentukan Strategic Business Unit (SBU) profesional',
      'Hilirisasi 15–20 produk inovasi komersial ber-TKDN',
      'Strategi corporate marketing & penetrasi pasar nasional',
      'Ekspansi revenue OGSCI Migas mencapai Rp 5,5 Miliar',
      'Operasionalisasi Co-Working & Virtual Office penuh'
    ]
  },
  {
    phase: 'FASE 3 (2029–2030)',
    title: 'GLOBAL MARKET POSITIONING & HOLDING',
    target: 'Rp 20,62 M ➔ Rp 23,47 M',
    accent: C_EMERALD,
    points: [
      'Kemandirian finansial penuh: 0% subsidi belanja APBD',
      'Operasional holding company inovasi daerah pertama',
      'Pusat pengembangan Indonesia Digital Technopark',
      'Realisasi target pendapatan mandiri Rp 23,47 Miliar',
      'Ekspansi jejaring kemitraan dan penempatan global'
    ]
  }
];

phases.forEach((ph, idx) => {
  const px = 0.8 + idx * 3.95;
  addBentoCard(s8, px, 1.45, 3.8, 4.0, { fill: C_BENTO_BG, accentColor: ph.accent });

  // Phase Pill Header
  s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.15, y: 1.6, w: 3.5, h: 0.65,
    fill: C_BENTO_BG2, line: { color: ph.accent, width: 1 }, rectRadius: 0.05
  });
  s8.addText(ph.phase, {
    x: px + 0.2, y: 1.65, w: 3.4, h: 0.25,
    fontSize: 9.5, bold: true, color: ph.accent, fontFace: FONT_HEAD, align: 'center'
  });
  s8.addText(ph.title, {
    x: px + 0.2, y: 1.9, w: 3.4, h: 0.3,
    fontSize: 9.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });

  // Target Box
  s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.2, y: 2.35, w: 3.4, h: 0.4,
    fill: C_DARK_BG, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
  });
  s8.addText(`Target: ${ph.target}`, {
    x: px + 0.2, y: 2.35, w: 3.4, h: 0.4,
    fontSize: 10.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
  });

  // Bullets
  const pList = ph.points.map(pt => `•  ${pt}`).join('\n\n');
  s8.addText(pList, {
    x: px + 0.2, y: 2.85, w: 3.4, h: 2.45,
    fontSize: 9.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 13
  });
});

// Bottom Milestone Table Strip
addBentoCard(s8, 0.8, 5.6, 11.733, 1.25, { fill: C_BENTO_BG2, line: C_CYAN });
s8.addText('RINGKASAN TARGET TAHUNAN MENUJU KEMANDIRIAN 2030 (AUDIT FINANSIAL RIIL):', {
  x: 1.0, y: 5.7, w: 11.333, h: 0.25,
  fontSize: 9.5, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

const mstoneData = [
  [
    { text: 'TAHUN 2026 (Baseline)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'TAHUN 2027 (+67,0%)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'TAHUN 2028 (+28,5%)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'TAHUN 2029 (+28,0%)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'TAHUN 2030 (MANDIRI)', options: { bold: true, color: C_AMBER, fill: { color: C_NAVY_BAR }, align: 'center' } }
  ],
  [
    { text: 'Rp 7.503.392.698', options: { align: 'center', color: C_WHITE } },
    { text: 'Rp 12.532.185.031', options: { align: 'center', color: C_CYAN_LIGHT } },
    { text: 'Rp 16.100.658.831', options: { align: 'center', color: C_CYAN_LIGHT } },
    { text: 'Rp 20.615.658.831', options: { align: 'center', color: C_CYAN_LIGHT } },
    { text: 'Rp 23.468.158.831', options: { align: 'center', color: C_AMBER, bold: true } }
  ]
];

s8.addTable(mstoneData, {
  x: 1.0, y: 6.0, w: 11.333, h: 0.72,
  colW: [2.266, 2.266, 2.266, 2.266, 2.269],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG },
  fontSize: 9.2, fontFace: FONT_HEAD, valign: 'middle',
  rowH: [0.32, 0.38]
});

s8.addNotes('Roadmap ini kami rancang dalam tiga fase yang saling membangun. 2026–2027 adalah fase pondasi: tata kelola BLUD diperkuat, lab kita terakreditasi, dan kita mulai ekspansi kerjasama OGSCI dan SMK GoGlobal. 2028–2029 adalah fase akselerasi: SBU terbentuk dan mulai mandiri seperti entitas bisnis profesional. Puncaknya, 2030, Solo Technopark beroperasi sebagai holding company inovasi daerah yang sepenuhnya mandiri—mengharumkan nama Surakarta di panggung teknologi nasional bahkan global.');

// =========================================================================
// SLIDE 09: RE-SETTING TARGET FINANSIAL BLUD: +212,8%
// =========================================================================
const s9 = pptx.addSlide();
addHeader(s9, 'RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD +212,8%', 'Proyeksi Keuangan BLUD 2026–2030', 9,
  'Proyeksi Pendapatan Mandiri Riil dari Rp 7,50 Miliar (2026) menuju Rp 23,47 Miliar (2030) — Tanpa Subsidi APBD'
);

// Left Column: Yearly Progression Table
const finTableData = [
  [
    { text: 'TAHUN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'TARGET PENDAPATAN (RP)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'right' } },
    { text: 'PERTUMBUHAN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'KETERANGAN AKSELERATOR', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } }
  ],
  [
    { text: '2026', options: { bold: true, align: 'center', color: C_WHITE } },
    { text: 'Rp 7.503.392.698', options: { align: 'right', color: C_WHITE } },
    { text: 'Baseline', options: { align: 'center', bold: true, color: C_SILVER } },
    { text: 'Re-setting target; OGSCI Rp 1,79 M; Sewa Shopee Rp 544 jt', options: { color: C_SILVER } }
  ],
  [
    { text: '2027', options: { bold: true, align: 'center', color: C_WHITE } },
    { text: 'Rp 12.532.185.031', options: { align: 'right', color: C_CYAN_LIGHT } },
    { text: '+67,0%', options: { align: 'center', color: C_CYAN_LIGHT, bold: true } },
    { text: 'OGSCI Rp 2,5 M; TUK/JFT aktif; Startup binaan Rp 1,25 M', options: { color: C_SILVER } }
  ],
  [
    { text: '2028', options: { bold: true, align: 'center', color: C_WHITE } },
    { text: 'Rp 16.100.658.831', options: { align: 'right', color: C_CYAN_LIGHT } },
    { text: '+28,5%', options: { align: 'center', color: C_CYAN_LIGHT, bold: true } },
    { text: 'BEP Tercapai; SBU terbentuk; Rekayasa manufaktur Rp 1,11 M', options: { color: C_SILVER } }
  ],
  [
    { text: '2029', options: { bold: true, align: 'center', color: C_WHITE } },
    { text: 'Rp 20.615.658.831', options: { align: 'right', color: C_CYAN_LIGHT } },
    { text: '+28,0%', options: { align: 'center', color: C_CYAN_LIGHT, bold: true } },
    { text: 'OGSCI Rp 5,5 M; TUK/JFT Rp 2,27 M; Hilirisasi riset masif', options: { color: C_SILVER } }
  ],
  [
    { text: '2030', options: { bold: true, align: 'center', color: C_AMBER } },
    { text: 'Rp 23.468.158.831', options: { align: 'right', bold: true, color: C_AMBER } },
    { text: '0% APBD', options: { align: 'center', color: C_AMBER, bold: true } },
    { text: 'Kemandirian penuh; OGSCI Rp 6,25 M; Startup Rp 2,0 M; Holding', options: { color: C_AMBER_LIGHT } }
  ]
];

s9.addTable(finTableData, {
  x: 0.8, y: 1.45, w: 7.3, h: 2.7,
  colW: [0.9, 2.3, 1.4, 2.7],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.38, 0.44, 0.44, 0.44, 0.44, 0.5]
});

// Left Bottom: Strategic Financial Conclusion
addBentoCard(s9, 0.8, 4.3, 7.3, 2.55, { fill: C_BENTO_BG2, accentColor: C_CYAN });
s9.addText('LOMPATAN FINANSIAL: +212,8% (3,1× LIPAT DALAM 4 TAHUN)', {
  x: 1.05, y: 4.45, w: 6.9, h: 0.35,
  fontSize: 12, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s9.addText([
  { text: '• Mandiri Total Tanpa APBD di 2030: ', options: { bold: true, color: C_WHITE } },
  { text: 'Seluruh biaya operasional UPTD STP tertutup penuh oleh pendapatan layanan mandiri tanpa subsidi kas daerah.\n', options: { color: C_SILVER } },
  { text: '• Titik Impas (BEP) di Tahun 2028: ', options: { bold: true, color: C_WHITE } },
  { text: 'Tercapai di angka Rp 16,10 Miliar seiring beroperasinya Strategic Business Unit (SBU) manufaktur dan vokasi.\n', options: { color: C_SILVER } },
  { text: '• Kualitas Pendapatan Berkelanjutan: ', options: { bold: true, color: C_WHITE } },
  { text: '88% disumbang keahlian riset & vokasi industri (program migas, uji lab, diklat spesialis, dan startup).', options: { color: C_SILVER } }
], {
  x: 1.05, y: 4.85, w: 6.9, h: 1.85,
  fontSize: 9.2, fontFace: FONT_BODY, lineSpacing: 14
});

// Right Column: Composition & Top Catalysts
addBentoCard(s9, 8.3, 1.45, 4.233, 5.4, { fill: C_BENTO_BG, accentColor: C_AMBER });
s9.addText('KOMPOSISI REVENUE STREAM 2030', {
  x: 8.55, y: 1.6, w: 3.8, h: 0.3,
  fontSize: 10.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
});

const revStreams = [
  { rank: '🥇', name: 'Kerjasama Program', val: 'Rp 12,83 M', pct: '51%', color: C_CYAN },
  { rank: '🥈', name: 'Pelatihan & Sertifikasi', val: 'Rp 5,41 M', pct: '23%', color: C_CYAN_LIGHT },
  { rank: '🥉', name: 'Sewa Kawasan & Coworking', val: 'Rp 4,17 M', pct: '18%', color: C_EMERALD },
  { rank: '🏅', name: 'Produksi & Rekayasa', val: 'Rp 1,07 M', pct: '8%', color: C_SILVER }
];

revStreams.forEach((rs, idx) => {
  const rsy = 2.0 + idx * 0.55;
  s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 8.55, y: rsy, w: 3.8, h: 0.48,
    fill: C_BENTO_BG2, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
  });
  s9.addText(`${rs.rank} ${rs.name}`, {
    x: 8.65, y: rsy + 0.08, w: 2.1, h: 0.32,
    fontSize: 8.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s9.addText(`${rs.val} (${rs.pct})`, {
    x: 10.75, y: rsy + 0.08, w: 1.5, h: 0.32,
    fontSize: 8.8, bold: true, color: rs.color, fontFace: FONT_HEAD, align: 'right'
  });
});

// Top 5 Catalysts Box
s9.addText('TOP 5 KATALISATOR PENDAPATAN 2030:', {
  x: 8.55, y: 4.35, w: 3.8, h: 0.25,
  fontSize: 9.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const top5 = [
  { item: '#1 OGSCI Program Migas', nominal: 'Rp 6.250.000.000' },
  { item: '#2 Pemda e-Katalog Inovasi', nominal: 'Rp 2.250.000.000' },
  { item: '#3 Startup Binaan STP', nominal: 'Rp 2.000.000.000' },
  { item: '#4 Sertifikasi Tenaga JFT', nominal: 'Rp 1.375.000.000' },
  { item: '#5 Tempat Uji Kompetensi', nominal: 'Rp 1.350.000.000' }
];

top5.forEach((t, idx) => {
  const ty = 4.65 + idx * 0.42;
  s9.addText(t.item, {
    x: 8.55, y: ty, w: 2.1, h: 0.35,
    fontSize: 8.5, color: C_SILVER, fontFace: FONT_BODY
  });
  s9.addText(t.nominal, {
    x: 10.65, y: ty, w: 1.7, h: 0.35,
    fontSize: 8.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right'
  });
});

s9.addNotes('Inilah komitmen terukur kami, Bapak Walikota. Angka-angka ini bukan proyeksi asal-asalan—setiap rupiah memiliki sumber yang teridentifikasi. Pendorong terbesar adalah program OGSCI yang akan menyumbang Rp 6,25 Miliar di 2030—itu sendirian sudah hampir menyamai total pendapatan STP tahun 2026. Ditambah program Startup Binaan yang kami proyeksi menghasilkan Rp 2 Miliar, dan e-Katalog yang membutuhkan Instruksi Walikota untuk membuka pasar Rp 2,25 Miliar. Dengan diversifikasi ini, STP akan menjadi BLUD yang sepenuhnya mandiri pada 2030.');

// =========================================================================
// SLIDE 10: BREAKDOWN MESIN PENDAPATAN 2030 PER MATA ANGGARAN
// =========================================================================
const s10 = pptx.addSlide();
addHeader(s10, 'ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN', 'Transparansi Anggaran & Struktur Pendapatan', 10,
  'Transparansi Total Hasil Rekonsiliasi: Dari Mana Setiap Rupiah dari Target Rp 23,47 Miliar Berasal'
);

// 6 Clusters Cards in 3x2 Grid
const clusters = [
  {
    code: 'KLASTER A',
    title: 'KERJASAMA PROGRAM STRATEGIS',
    total: 'Rp 12.825.000.000 (51%)',
    accent: C_CYAN,
    lines: [
      'OGSCI (Oil, Gas, Coal, Ship & Infra): Rp 6.250.000.000',
      'Pemda Surakarta (e-Katalog Inovasi): Rp 2.250.000.000',
      'Startup Binaan STP (Komersialisasi): Rp 2.000.000.000',
      'BP3MI - SMK GoGlobal (Talenta Global): Rp 1.500.000.000',
      'LPK Lainnya & Kemenperin: Rp 825.000.000'
    ]
  },
  {
    code: 'KLASTER B',
    title: 'LAYANAN VOKASI & SERTIFIKASI',
    total: 'Rp 3.679.800.000 (16%)',
    accent: C_CYAN_LIGHT,
    lines: [
      'Sertifikasi JFT (Fungsional Tertentu): Rp 1.375.000.000',
      'TUK (Tempat Uji Kompetensi Mandiri): Rp 1.350.000.000',
      'Konsultasi Lab & Data Center: Rp 950.000.000',
      'Prakerin Siswa SMK Se-Solo Raya: Rp 4.800.000'
    ]
  },
  {
    code: 'KLASTER C',
    title: 'DIKLAT MANDIRI SPESIALIS',
    total: 'Rp 2.469.000.000 (10%)',
    accent: C_AMBER,
    lines: [
      'Underwater Wet Welding (Keahlian Langka): Rp 846.000.000',
      'Manufaktur Presisi & Welding Intensif: Rp 728.000.000',
      'Digital Tech (Cyber Security & AI): Rp 225.000.000',
      'Desain Produk, Manajerial & EV: Rp 555.000.000',
      'Kewirausahaan & Digitalisasi UMKM: Rp 115.000.000'
    ]
  },
  {
    code: 'KLASTER D',
    title: 'LAIN-LAIN PENDAPATAN BLUD',
    total: 'Rp 2.450.000.000 (10%)',
    accent: C_EMERALD,
    lines: [
      'Pemanfaatan Sarana Kawasan & Event/MICE: Rp 1.500.000.000',
      'Co-Working Space & Virtual Office: Rp 900.000.000',
      'Jasa Giro Rekening Kas BLUD: Rp 50.000.000'
    ]
  },
  {
    code: 'KLASTER E',
    title: 'REKAYASA MANUFAKTUR & PRODUKSI',
    total: 'Rp 1.465.000.000 (6%)',
    accent: C_WHITE,
    lines: [
      'Jasa Produksi Precision Parts: Rp 1.225.000.000',
      'Standarisasi Produk & Sertifikasi TKDN: Rp 135.000.000',
      'Produksi Pesanan Pemda (e-Katalog): Rp 105.000.000'
    ]
  },
  {
    code: 'KLASTER F',
    title: 'KEMITRAAN LAHAN EKSISTING',
    total: 'Rp 579.358.831 (3%)',
    accent: C_SILVER,
    lines: [
      'Shopee Indonesia (Sewa Lahan/Co-location): Rp 544.358.831',
      'Katulondi (Sewa Fasilitas Kawasan): Rp 35.000.000',
      'Catatan: Porsi sewa lahan ditekan ke 3% demi fokus layanan bernilai tambah tinggi.'
    ]
  }
];

clusters.forEach((cl, idx) => {
  const col = idx % 3;
  const row = Math.floor(idx / 3);
  const cx = 0.8 + col * 3.95;
  const cy = 1.45 + row * 2.38;

  addBentoCard(s10, cx, cy, 3.8, 2.25, { fill: C_BENTO_BG, accentColor: cl.accent });

  // Header Box
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: cx + 0.15, y: cy + 0.1, w: 3.5, h: 0.52,
    fill: C_BENTO_BG2, line: { color: cl.accent, width: 1 }, rectRadius: 0.04
  });
  s10.addText(`${cl.code}: ${cl.title}`, {
    x: cx + 0.2, y: cy + 0.12, w: 3.4, h: 0.22,
    fontSize: 8.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s10.addText(cl.total, {
    x: cx + 0.2, y: cy + 0.34, w: 3.4, h: 0.24,
    fontSize: 9.5, bold: true, color: cl.accent, fontFace: FONT_HEAD
  });

  // Lines
  const lText = cl.lines.map(l => `• ${l}`).join('\n');
  s10.addText(lText, {
    x: cx + 0.18, y: cy + 0.68, w: 3.45, h: 1.45,
    fontSize: 7.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 10
  });
});

// Grand Total Highlight Banner
addBentoCard(s10, 0.8, 6.3, 11.733, 0.62, { fill: C_BENTO_BG2, line: C_AMBER });
s10.addText([
  { text: 'TOTAL TARGET PENDAPATAN BLUD 2030:   ', options: { bold: true, color: C_WHITE, fontSize: 11 } },
  { text: 'Rp 23.468.158.831  ', options: { bold: true, color: C_AMBER, fontSize: 14 } },
  { text: '(Dua Puluh Tiga Miliar Empat Ratus Enam Puluh Delapan Juta Seratus Lima Puluh Delapan Ribu Rupiah)', options: { italic: true, color: C_SILVER, fontSize: 8.5 } }
], {
  x: 1.0, y: 6.3, w: 11.333, h: 0.62,
  fontFace: FONT_HEAD, align: 'center', valign: 'middle'
});

s10.addNotes('Bapak Walikota, slide ini adalah bukti transparansi kami. Setiap mata anggaran dari Rp 23,47 Miliar sudah teridentifikasi sumbernya. Ini bukan angka anggan-anggan—ini adalah komitmen yang sudah terverifikasi bersama mitra-mitra kami. Satu hal yang menarik perhatian: Startup Binaan yang kami proyeksi Rp 2 Miliar di 2030 adalah cerminan dari keberhasilan inkubasi—artinya startup yang kami bina sudah memberi kontribusi balik kepada STP. Ini adalah siklus ekosistem yang sehat.');

// =========================================================================
// SLIDE 11: 6 DIMENSI KEBERDAMPAKAN NYATA BAGI WARGA SOLO
// =========================================================================
const s11 = pptx.addSlide();
addHeader(s11, '6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA', 'Dampak Sosial-Ekonomi Daerah', 11,
  '"Indikator Keberhasilan Harus Mengukur IMPACT — Bukan Sekadar Aktivitas dan Pendapatan Kas" (Dokumen Strategi STP 2026)'
);

const impacts = [
  {
    num: '1',
    icon: '👷',
    title: 'SDM LEBIH KOMPETEN',
    desc: 'Pelatihan vokasi industri & uji kompetensi memperkuat daya saing kerja warga kota.',
    metric: '9.000 Peserta / Thn',
    submetric: 'Kuota afirmasi 65% KTP Surakarta',
    accent: C_CYAN
  },
  {
    num: '2',
    icon: '🌱',
    title: 'UMKM & STARTUP TUMBUH',
    desc: 'Inkubasi memberi pendampingan teknis, digitalisasi GoTo Hub, & akses pasar modal.',
    metric: '80+ Wirausaha Baru',
    submetric: 'Peningkatan omzet & legalitas',
    accent: C_AMBER
  },
  {
    num: '3',
    icon: '🔬',
    title: 'RISET DEKAT KE PASAR',
    desc: 'Hilirisasi menghubungkan inventor kampus dengan industri manufaktur dan OPD.',
    metric: '45+ Produk Hilirisasi',
    submetric: 'Paten HKI & Standar TKDN',
    accent: C_EMERALD
  },
  {
    num: '4',
    icon: '💻',
    title: 'AKSES TEKNOLOGI TERBUKA',
    desc: 'Fasilitas canggih AI, Cyber Security, 5G & Lab Presisi dibuka bagi pelajar SMK & publik.',
    metric: '10.000+ Kunjungan/Thn',
    submetric: 'Prakerin SMK se-Solo Raya',
    accent: C_CYAN_LIGHT
  },
  {
    num: '5',
    icon: '💼',
    title: 'PELUANG KERJA BARU',
    desc: 'Ekosistem menghasilkan SDM terserap industri dan lahirnya wirausaha mandiri.',
    metric: 'Rp 180 M Perputaran',
    submetric: 'Multiplier effect hotel & MICE',
    accent: C_AMBER
  },
  {
    num: '6',
    icon: '🤝',
    title: 'KOLABORASI INDUSTRI',
    desc: '95+ korporasi multinasional menanamkan fasilitas dan proyek langsung di Solo.',
    metric: '100% Swasta / PPP',
    submetric: '0% beban belanja modal APBD',
    accent: C_EMERALD
  }
];

impacts.forEach((im, idx) => {
  const col = idx % 3;
  const row = Math.floor(idx / 3);
  const ix = 0.8 + col * 3.95;
  const iy = 1.45 + row * 2.45;

  addBentoCard(s11, ix, iy, 3.8, 2.3, { fill: C_BENTO_BG, accentColor: im.accent });

  // Icon & Header
  s11.addText(`${im.icon}  ${im.num}. ${im.title}`, {
    x: ix + 0.2, y: iy + 0.15, w: 3.4, h: 0.35,
    fontSize: 10.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });

  // Description
  s11.addText(im.desc, {
    x: ix + 0.2, y: iy + 0.52, w: 3.4, h: 0.65,
    fontSize: 8.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 12
  });

  // Metric Callout Box inside Card
  s11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ix + 0.2, y: iy + 1.25, w: 3.4, h: 0.85,
    fill: C_BENTO_BG2, line: { color: im.accent, width: 1 }, rectRadius: 0.05
  });
  s11.addText(im.metric, {
    x: ix + 0.25, y: iy + 1.3, w: 3.3, h: 0.42,
    fontSize: 13, bold: true, color: im.accent, fontFace: FONT_HEAD, align: 'center'
  });
  s11.addText(im.submetric, {
    x: ix + 0.25, y: iy + 1.72, w: 3.3, h: 0.3,
    fontSize: 8.5, bold: true, color: C_SILVER, fontFace: FONT_BODY, align: 'center'
  });
});

// Bottom Synthesis Banner
addBentoCard(s11, 0.8, 6.45, 11.733, 0.48, { fill: C_BENTO_BG2, line: C_BORDER });
s11.addText('TOLAK UKUR KUNCI: Peserta terserap kerja (afirmasi KTP Solo) • Tenant tumbuh • Prototipe tervalidasi • HKI terdaftar • Produk masuk pasar • Investasi daerah bertambah', {
  x: 1.0, y: 6.45, w: 11.333, h: 0.48,
  fontSize: 9.2, bold: true, color: C_CYAN_LIGHT, align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s11.addNotes('Bapak Walikota, ukuran keberhasilan tertinggi kami bukan saldo rekening BLUD, melainkan enam dimensi dampak ini: seberapa banyak warga Solo mendapat pekerjaan, seberapa banyak UMKM naik kelas, seberapa banyak riset kampus menjadi produk nyata di pasar. Indikator-indikator ini yang setiap bulan kami pantau bersama BRIDA. Karena pada akhirnya, semua program yang kami jalankan harus bisa dijawab dengan satu pertanyaan sederhana: apakah warga Solo hidupnya lebih baik karena Solo Technopark ada?');

// =========================================================================
// SLIDE 12: 3 PERMOHONAN KEBIJAKAN & RENCANA AKSI 100 HARI
// =========================================================================
const s12 = pptx.addSlide();
addHeader(s12, '3 PERMOHONAN KEBIJAKAN KEPADA WALIKOTA & RENCANA AKSI 100 HARI', 'Dukungan Kepemimpinan & Komitmen Eksekusi', 12,
  'Kami Tidak Memohon Tambahan Pagu APBD — Kami Memohon 3 Payung Regulasi Kepemimpinan Walikota'
);

// 3 Leadership Asks Horizontal Cards
const asks = [
  {
    num: 'PERMOHONAN 1',
    title: 'INSTRUKSI WALIKOTA AFIRMASI e-KATALOG',
    accent: C_CYAN,
    desc: 'Mengarahkan seluruh OPD Pemkot memprioritaskan belanja produk teknologi & layanan tenant STP via e-Katalog Lokal.',
    target: 'Target: Rp 2,25 M / Thn di 2030 (80+ Produk Terserap)',
    status: 'Bisa diterbitkan dalam 30 hari pertama'
  },
  {
    num: 'PERMOHONAN 2',
    title: 'FLEKSIBILITAS KERJASAMA LAHAN BLUD',
    accent: C_AMBER,
    desc: 'Peraturan Walikota skema Long-Term Revenue Sharing (10–25 tahun) yang fleksibel bagi investor teknologi global di kawasan STP.',
    target: 'Target: Menarik 3–5 Investor Baru Tanpa Beban APBD',
    status: 'Harmonisasi bersama BPKAD & Bagian Hukum'
  },
  {
    num: 'PERMOHONAN 3',
    title: 'SURAT EDARAN KONSORSIUM RISET DAERAH',
    accent: C_EMERALD,
    desc: 'Mendorong perguruan tinggi (UNS, UMS, ISI, Poltek) mengarahkan minimal 30% riset terapan menyelesaikan problem kota di STP.',
    target: 'Target: 5 Kampus Mitra, 30 Klaster Riset Terapan',
    status: 'Penyusunan MoU bersama 5 Rektor Solo Raya'
  }
];

asks.forEach((a, idx) => {
  const ax = 0.8 + idx * 3.95;
  addBentoCard(s12, ax, 1.45, 3.8, 2.7, { fill: C_BENTO_BG, accentColor: a.accent });

  // Number Badge
  s12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ax + 0.15, y: 1.6, w: 3.5, h: 0.55,
    fill: C_BENTO_BG2, line: { color: a.accent, width: 1 }, rectRadius: 0.05
  });
  s12.addText(a.num, {
    x: ax + 0.2, y: 1.63, w: 3.4, h: 0.22,
    fontSize: 8.5, bold: true, color: a.accent, fontFace: FONT_HEAD, align: 'center'
  });
  s12.addText(a.title, {
    x: ax + 0.2, y: 1.88, w: 3.4, h: 0.25,
    fontSize: 9.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });

  // Description
  s12.addText(a.desc, {
    x: ax + 0.2, y: 2.25, w: 3.4, h: 0.9,
    fontSize: 8.8, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 12
  });

  // Target Box
  s12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ax + 0.2, y: 3.25, w: 3.4, h: 0.75,
    fill: C_DARK_BG, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
  });
  s12.addText(a.target, {
    x: ax + 0.25, y: 3.3, w: 3.3, h: 0.35,
    fontSize: 8.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s12.addText(`⚡ ${a.status}`, {
    x: ax + 0.25, y: 3.65, w: 3.3, h: 0.3,
    fontSize: 8.2, italic: true, color: a.accent, fontFace: FONT_BODY
  });
});

// Bottom 100-Day Action Table
addBentoCard(s12, 0.8, 4.3, 11.733, 2.65, { fill: C_BENTO_BG, line: C_CYAN });
s12.addText('RENCANA AKSI 100 HARI PERTAMA (QUICK WINS PASCA-ARAHAN WALIKOTA SURAKARTA):', {
  x: 1.0, y: 4.4, w: 11.333, h: 0.28,
  fontSize: 10, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const action100Data = [
  [
    { text: 'PERIODE', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'AGENDA STRATEGIS', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'OUTPUT KONKRET & TARGET KEBERHASILAN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } }
  ],
  [
    { text: 'Hari 1–10', options: { bold: true, align: 'center', color: C_CYAN_LIGHT } },
    { text: 'Penyusunan Draf Instruksi Walikota e-Katalog', options: { color: C_WHITE, bold: true } },
    { text: 'Draf Inkwal Afirmatif siap harmonisasi bersama Bagian Hukum Setda & BPKAD.', options: { color: C_SILVER } }
  ],
  [
    { text: 'Hari 11–30', options: { bold: true, align: 'center', color: C_CYAN_LIGHT } },
    { text: 'Harmonisasi & Penerbitan Instruksi Walikota', options: { color: C_WHITE, bold: true } },
    { text: 'Inkwal ditandatangani Walikota; Sosialisasi ke seluruh OPD prioritas Pemkot.', options: { color: C_EMERALD, bold: true } }
  ],
  [
    { text: 'Hari 31–50', options: { bold: true, align: 'center', color: C_AMBER } },
    { text: 'Kurasi & Seleksi 5 Produk Inovasi Unggulan', options: { color: C_WHITE, bold: true } },
    { text: 'Katalog 5 Produk Unggulan Hilirisasi STP (Smart Agri IoT, Alkes, Mesin Sortasi).', options: { color: C_SILVER } }
  ],
  [
    { text: 'Hari 51–60', options: { bold: true, align: 'center', color: C_AMBER } },
    { text: 'Showcase Inovasi di Depan Walikota & OPD', options: { color: C_WHITE, bold: true } },
    { text: 'Demonstrasi produk di hadapan Walikota & media; komitmen serapan perdana OPD.', options: { color: C_WHITE } }
  ],
  [
    { text: 'Hari 61–80', options: { bold: true, align: 'center', color: C_EMERALD } },
    { text: 'Inisiasi Konsorsium Riset 5 Perguruan Tinggi', options: { color: C_WHITE, bold: true } },
    { text: 'MoU Pra-Konsorsium bersama Rektor UNS, UMS, ISI, Poltek Solo Raya.', options: { color: C_SILVER } }
  ],
  [
    { text: 'Hari 81–100', options: { bold: true, align: 'center', color: C_EMERALD } },
    { text: 'Launching Resmi Konsorsium & Hilirisasi Inovasi', options: { color: C_WHITE, bold: true } },
    { text: 'Penandatanganan MoU Riset Terapan bersama Kadin/Apindo di STP; kick-off hilirisasi.', options: { color: C_EMERALD, bold: true } }
  ]
];

s12.addTable(action100Data, {
  x: 1.0, y: 4.75, w: 11.333, h: 2.05,
  colW: [1.4, 3.8, 6.133],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG2 },
  fontSize: 8.5, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.3, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28]
});

s12.addNotes('Menutup paparan ini, Bapak Walikota, manajemen Solo Technopark dan BRIDA tidak hadir meminta tambahan anggaran. Tiga permohonan kami semuanya adalah payung regulasi—bukan rupiah belanja. Instruksi Walikota untuk e-Katalog, Perwal kerjasama lahan, dan Surat Edaran konsorsium riset. Tiga lembar kebijakan ini yang akan menjadi kunci pembuka ekosistem senilai Rp 23 Miliar di 2030. Dengan kepemimpinan dan arahan Bapak, tim kami siap langsung tancap gas mengeksekusi 100 hari pertama. Bersama Bapak, Solo Technopark akan menjadi kebanggaan kota Surakarta dan teladan nasional pengelolaan kawasan inovasi yang mandiri, berdampak, dan berkelanjutan.');

// =========================================================================
// WRITE TO DISK
// =========================================================================
const outputPptxPathV3 = path.resolve(__dirname, 'PAPARAN_WALIKOTA_STP_V3.pptx');
const outputPptxPathOut = path.resolve(__dirname, '../output_paparan/PAPARAN_WALIKOTA_STP_V3.pptx');

pptx.writeFile({ fileName: outputPptxPathV3 })
  .then(() => {
    console.log(`[SUCCESS] PPTX Master V3 generated at: ${outputPptxPathV3}`);
    fs.copyFileSync(outputPptxPathV3, outputPptxPathOut);
    console.log(`[SUCCESS] Copied to output_paparan: ${outputPptxPathOut}`);
  })
  .catch(err => {
    console.error('[ERROR] Failed generating PPTX V3:', err);
    process.exit(1);
  });
