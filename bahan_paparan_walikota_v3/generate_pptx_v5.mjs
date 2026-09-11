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
pptx.title = 'Solo Technopark: Dari Kawasan Sains & Teknologi Menuju Mesin Impact Daerah (V5 - Native Charts & Diagrams)';
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
const C_PURPLE      = '8B5CF6'; // Purple accent
const C_PINK        = 'EC4899'; // Pink accent
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
    desc: 'Beralih dari pengelola sewa lahan menjadi DELIVERY UNIT inovasi BRIDA yang menghasilkan IMPACT nyata bagi masyarakat dan industri.',
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
// V5 Upgrade: True Integrated 3-Stage Value Chain Flow Diagram
// =========================================================================
const s2 = pptx.addSlide();
addHeader(s2, 'KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)', 'Tata Kelola Pemerintahan & Kelembagaan', 2,
  'Arsitektur Nilai: BRIDA sebagai Otak Kebijakan, STP sebagai Tangan Eksekutor, Warga & Industri sebagai Penerima Dampak'
);

// 3 Unified Value Chain Stages
const colW2 = 3.5;
const stages2 = [
  {
    x: 0.8,
    stage: 'TAHAP 1: REGULASI & ARAHAN (INPUT)',
    title: 'BRIDA KOTA SURAKARTA',
    role: 'Otak & Regulator Kebijakan Daerah',
    accent: C_CYAN,
    items: [
      'Merumuskan arah kebijakan riset & inovasi daerah',
      'Penyusunan kajian teknokratik & roadmap strategis',
      'Orkestrasi ekosistem Iptek antar-OPD & perguruan tinggi',
      'Monitoring & evaluasi akuntabilitas dampak sosial-ekonomi'
    ],
    kpi: 'Output: Regulasi, Roadmap, & Pengawasan Kebijakan'
  },
  {
    x: 4.9,
    stage: 'TAHAP 2: DELIVERY UNIT & OPERASI (PROSES)',
    title: 'UPTD KST SOLO TECHNOPARK',
    role: 'Tangan Eksekutor & Operator Bisnis BLUD',
    accent: C_AMBER,
    items: [
      'Penyelenggaraan diklat vokasi & sertifikasi industri',
      'Inkubasi bisnis, alih teknologi & komersialisasi riset',
      'Uji coba prototipe lab presisi & penetrant testing (NDT)',
      'Pengelolaan portofolio bisnis BLUD mandiri & kemitraan'
    ],
    kpi: 'Output: Layanan Teknis, Inkubasi, & Hilirisasi Produk'
  },
  {
    x: 9.0,
    stage: 'TAHAP 3: DAMPAK KESEJAHTERAAN (OUTCOME)',
    title: 'WARGA SURAKARTA & DUNIA USAHA',
    role: 'Penerima Manfaat Utama (Benefisiari)',
    accent: C_EMERALD,
    items: [
      'SDM lokal kompeten & tersertifikasi kerja industri global',
      'UMKM & startup teknologi tumbuh mandiri omzetnya',
      'Produk inovasi kampus terserap pasar & e-Katalog Pemda',
      'Masuknya investasi swasta & penyerapan 9.000 tenaga kerja'
    ],
    kpi: 'Impact: Lapangan Kerja, Daya Saing, & PAD Mandiri'
  }
];

// Draw 3 Stages
stages2.forEach((st) => {
  addBentoCard(s2, st.x, 1.45, colW2, 4.35, { fill: C_BENTO_BG, accentColor: st.accent });

  // Stage Pill
  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: st.x + 0.15, y: 1.6, w: colW2 - 0.3, h: 0.35,
    fill: C_BENTO_BG2, line: { color: st.accent, width: 1 }, rectRadius: 0.04
  });
  s2.addText(st.stage, {
    x: st.x + 0.15, y: 1.6, w: colW2 - 0.3, h: 0.35,
    fontSize: 8, bold: true, color: st.accent, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
  });

  // Title & Role
  s2.addText(st.title, {
    x: st.x + 0.15, y: 2.05, w: colW2 - 0.3, h: 0.45,
    fontSize: 12, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });
  s2.addText(st.role, {
    x: st.x + 0.15, y: 2.45, w: colW2 - 0.3, h: 0.28,
    fontSize: 9, bold: true, color: st.accent, fontFace: FONT_BODY, align: 'center'
  });

  // Divider line inside card
  s2.addShape(pptx.shapes.RECTANGLE, {
    x: st.x + 0.3, y: 2.8, w: colW2 - 0.6, h: 0.015, fill: { color: C_BORDER }
  });

  // Bullet items
  const bText = st.items.map(it => `•  ${it}`).join('\n\n');
  s2.addText(bText, {
    x: st.x + 0.22, y: 2.95, w: colW2 - 0.44, h: 2.1,
    fontSize: 9.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 13
  });

  // KPI / Mandate Box at bottom of card
  s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: st.x + 0.15, y: 5.15, w: colW2 - 0.3, h: 0.5,
    fill: C_DARK_BG, line: { color: st.accent, width: 1 }, rectRadius: 0.04
  });
  s2.addText(st.kpi, {
    x: st.x + 0.2, y: 5.15, w: colW2 - 0.4, h: 0.5,
    fontSize: 8.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
  });
});

// Flow Connector Arrows between Stages
// Arrow 1: BRIDA -> STP
s2.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 4.35, y: 3.3, w: 0.5, h: 0.45,
  fill: { color: C_CYAN }, line: { color: C_CYAN }
});
s2.addText('Regulasi &\nArahan', {
  x: 4.25, y: 2.7, w: 0.7, h: 0.55,
  fontSize: 7.5, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD, align: 'center'
});

// Arrow 2: STP -> Warga
s2.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 8.45, y: 3.3, w: 0.5, h: 0.45,
  fill: { color: C_AMBER }, line: { color: C_AMBER }
});
s2.addText('Layanan &\nInovasi', {
  x: 8.35, y: 2.7, w: 0.7, h: 0.55,
  fontSize: 7.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD, align: 'center'
});

// Bottom Legal Basis Bar
addBentoCard(s2, 0.8, 5.95, 11.733, 0.88, { fill: C_BENTO_BG2, line: C_BORDER });
s2.addText('LANDASAN HUKUM KELEMBAGAAN KOTA SURAKARTA:', {
  x: 1.0, y: 6.05, w: 11.333, h: 0.24,
  fontSize: 9.2, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});
s2.addText('• Perwali Surakarta No. 15 Tahun 2022: Penetapan kedudukan UPTD KST Solo Technopark sebagai unit pelaksana teknis operasional di bawah naungan BRIDA.\n• Perwali Surakarta No. 38 Tahun 2022: Penetapan Pola Tata Kelola Badan Layanan Umum Daerah (BLUD) UPTD KST Solo Technopark dengan fleksibilitas bisnis sehat.', {
  x: 1.0, y: 6.3, w: 11.333, h: 0.45,
  fontSize: 9, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 12
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
// V5 Upgrade: Seamless 5-Step Process Pipeline Flow (No Overlap)
// =========================================================================
const s5 = pptx.addSlide();
addHeader(s5, 'PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR NYATA', 'Mekanisme Hilirisasi Inovasi', 5,
  'Solo Technopark sebagai Jembatan: Menghubungkan Riset Kampus Menuju Produk Bernilai Ekonomi bagi Industri'
);

// 5 Connected Pipeline Stages
const stgW = 2.15;
const stages5 = [
  {
    num: '01',
    name: 'RISET & INVENSI',
    sub: 'Kampus & Peneliti',
    desc: 'Invensi UNS, UMS, ISI, Poltek Solo Raya. Riset dosen & mahasiswa.',
    accent: C_CYAN
  },
  {
    num: '02',
    name: 'SELEKSI & KURASI',
    sub: 'STP & BRIDA',
    desc: 'Kurasi komersial & relevansi kebutuhan industri serta problem kota Solo.',
    accent: C_CYAN_LIGHT
  },
  {
    num: '03',
    name: 'PROTOTYPING',
    sub: 'Lab Presisi STP',
    desc: 'Uji teknis, penetrant testing, validasi mutu produk & SNI/TKDN.',
    accent: C_AMBER
  },
  {
    num: '04',
    name: 'INKUBASI BISNIS',
    sub: 'Inkubator STP',
    desc: 'Penyusunan business model, proteksi HKI/Paten, akses modal & offtaker.',
    accent: C_EMERALD
  },
  {
    num: '05',
    name: 'KOMERSIALISASI',
    sub: 'Pasar & Industri',
    desc: 'Penetrasi e-Katalog Lokal OPD, captive market pemda & industri nasional.',
    accent: C_AMBER
  }
];

stages5.forEach((st, idx) => {
  const sx = 0.8 + idx * 2.41;
  addBentoCard(s5, sx, 1.45, stgW, 2.1, { fill: C_BENTO_BG, accentColor: st.accent });

  // Clean Header Box with Badge inside (No Overlapping circle!)
  s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: sx + 0.1, y: 1.55, w: stgW - 0.2, h: 0.45,
    fill: C_BENTO_BG2, line: { color: st.accent, width: 1 }, rectRadius: 0.04
  });
  s5.addText(`TAHAP ${st.num}`, {
    x: sx + 0.12, y: 1.58, w: 0.9, h: 0.38,
    fontSize: 9, bold: true, color: st.accent, fontFace: FONT_HEAD, valign: 'middle'
  });
  s5.addText(st.sub, {
    x: sx + 0.95, y: 1.58, w: stgW - 1.1, h: 0.38,
    fontSize: 7.5, bold: true, color: C_SILVER, fontFace: FONT_BODY, align: 'right', valign: 'middle'
  });

  // Stage Title
  s5.addText(st.name, {
    x: sx + 0.12, y: 2.08, w: stgW - 0.24, h: 0.4,
    fontSize: 10, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });

  // Stage Description
  s5.addText(st.desc, {
    x: sx + 0.12, y: 2.5, w: stgW - 0.24, h: 0.95,
    fontSize: 8.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 11
  });

  // Directional arrow between cards
  if (idx < 4) {
    s5.addShape(pptx.shapes.RIGHT_ARROW, {
      x: sx + stgW + 0.05, y: 2.35, w: 0.18, h: 0.28,
      fill: { color: st.accent }, line: { color: st.accent }
    });
  }
});

// Bottom Left: Table Portofolio Teknologi Riil (w: 6.8)
addBentoCard(s5, 0.8, 3.75, 6.8, 3.05, { fill: C_BENTO_BG, line: C_BORDER });
s5.addText('CONTOH TEKNOLOGI YANG DIHILIRKAN & DIINKUBASI DI STP (DATA RIIL)', {
  x: 1.0, y: 3.88, w: 6.4, h: 0.28,
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
  x: 1.0, y: 4.22, w: 6.4, h: 2.45,
  colW: [1.8, 3.2, 1.4],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG2 },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.38, 0.41, 0.41, 0.41, 0.41, 0.41]
});

// Bottom Right: Ekosistem Pendukung Aktif (w: 4.733)
addBentoCard(s5, 7.8, 3.75, 4.733, 3.05, { fill: C_BENTO_BG, line: C_CYAN });
s5.addText('EKOSISTEM FASILITAS PENDUKUNG AKTIF', {
  x: 8.0, y: 3.88, w: 4.333, h: 0.28,
  fontSize: 9.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const hubs = [
  { icon: '🛡️', name: 'National Cyber Security Hub', desc: 'Pusat ketahanan digital & talenta siber nasional' },
  { icon: '🎮', name: 'Gaming Hub & Digital Studio', desc: 'Pengembangan game & konten kreatif esport' },
  { icon: '🤖', name: 'AI Experience Center', desc: 'Laboratorium kecerdasan buatan terapan industri' },
  { icon: '🛒', name: 'GoTo UMKM Center', desc: 'Akselerasi digitalisasi & onboarding pedagang lokal' },
  { icon: '🔬', name: 'Laboratorium Rekayasa Presisi', desc: 'Fasilitas fabrikasi mesin & pengujian NDT industri' }
];

hubs.forEach((hb, idx) => {
  const hy = 4.25 + idx * 0.49;
  s5.addText(`${hb.icon}  ${hb.name}`, {
    x: 8.0, y: hy, w: 4.333, h: 0.22,
    fontSize: 9, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });
  s5.addText(hb.desc, {
    x: 8.4, y: hy + 0.2, w: 3.933, h: 0.24,
    fontSize: 8, color: C_SILVER, fontFace: FONT_BODY
  });
});

s5.addNotes('Bapak Walikota, kami ingin menghapus fenomena riset lemari arsip—di mana penelitian kampus hebat tapi berakhir hanya sebagai tumpukan laporan jurnal. Pipeline hilirisasi kami memutus kebuntuan itu. Riset dosen UNS tentang mesin sortasi kopi? Kita ambil, kita uji prototipenya di lab presisi STP, kita bantu urus patennya, kita carikan investor atau offtaker industrinya. Inilah nilai sejati Solo Technopark: jembatan dari ilmu pengetahuan menuju perputaran uang dan lapangan kerja bagi warga Solo.');

// =========================================================================
// SLIDE 06: EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD
// V5 Upgrade: Native PowerPoint Doughnut Chart for Hexahelix Composition
// =========================================================================
const s6 = pptx.addSlide();
addHeader(s6, 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD', 'Jejaring Kemitraan Strategis', 6,
  'Kepercayaan Korporasi Global & Nasional yang Menanamkan Investasi Langsung di Jantung Kota Surakarta'
);

// Left Container: Native Doughnut Chart Hexahelix (w: 4.6)
addBentoCard(s6, 0.8, 1.45, 4.6, 5.35, { fill: C_BENTO_BG, accentColor: C_CYAN });
s6.addText('KOMPOSISI 6 AKTOR HEXAHELIX (95+ MITRA)', {
  x: 1.0, y: 1.6, w: 4.2, h: 0.3,
  fontSize: 10.5, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

// Native Donut Chart
const hexaChartData = [{
  name: 'Aktor Hexahelix',
  labels: [
    'Industri & Bisnis (37%)',
    'Akademisi / Kampus (21%)',
    'Komunitas & Asosiasi (16%)',
    'Pemerintah / BUMN (13%)',
    'Media Massa (8%)',
    'Lembaga Global (5%)'
  ],
  values: [35, 20, 15, 12, 8, 5]
}];

s6.addChart(pptx.charts.DOUGHNUT, hexaChartData, {
  x: 1.5, y: 1.95, w: 3.2, h: 2.15,
  chartColors: ['0EA5E9', '38BDF8', '10B981', 'F59E0B', '8B5CF6', 'EC4899'],
  holeSize: 55,
  showLegend: false,
  showPercent: true,
  dataLabelColor: 'FFFFFF',
  dataLabelFontSize: 8
});

// 2-Column Legend Breakdown for Hexahelix Slices
const hexaLegend = [
  { name: 'Industri & Bisnis', pct: '37% (35)', col: C_CYAN },
  { name: 'Akademisi / Kampus', pct: '21% (20)', col: C_CYAN_LIGHT },
  { name: 'Komunitas & UMKM', pct: '16% (15)', col: C_EMERALD },
  { name: 'Pemerintah / BUMN', pct: '13% (12)', col: C_AMBER },
  { name: 'Media Massa', pct: '8% (8)', col: C_PURPLE },
  { name: 'Lembaga Global', pct: '5% (5)', col: C_PINK }
];

hexaLegend.forEach((hl, idx) => {
  const col = idx < 3 ? 0 : 1;
  const row = idx % 3;
  const hx = 1.0 + col * 2.15;
  const hy = 4.22 + row * 0.32;

  s6.addShape(pptx.shapes.OVAL, {
    x: hx, y: hy + 0.05, w: 0.12, h: 0.12,
    fill: { color: hl.col }, line: { color: hl.col }
  });
  s6.addText(`${hl.name}: `, {
    x: hx + 0.18, y: hy, w: 1.45, h: 0.25,
    fontSize: 7.8, color: C_SILVER, fontFace: FONT_BODY
  });
  s6.addText(hl.pct, {
    x: hx + 1.45, y: hy, w: 0.65, h: 0.25,
    fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right'
  });
});

// Bottom Synthesis inside Left Card
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 5.3, w: 4.2, h: 1.35,
  fill: C_BENTO_BG2, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
});
s6.addText('TOTAL: 95+ MITRA STRATEGIS AKTIF', {
  x: 1.1, y: 5.38, w: 4.0, h: 0.22,
  fontSize: 9.2, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s6.addText('• 100% Investasi Swasta & PPP (0% Beban APBD)\n• Akses fasilitas 5G, Cyber Security, AI, & Migas\n• Saluran penyerapan kerja langsung bagi warga Solo', {
  x: 1.1, y: 5.62, w: 4.0, h: 0.95,
  fontSize: 8.2, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 11
});

// Right Container: Table of World-Class Partners (w: 6.933)
addBentoCard(s6, 5.6, 1.45, 6.933, 5.35, { fill: C_BENTO_BG, line: C_BORDER });
s6.addText('MITRA INDUSTRI KELAS DUNIA YANG AKTIF BERINVESTASI DI STP', {
  x: 5.8, y: 1.6, w: 6.5, h: 0.3,
  fontSize: 10.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const partnersTableData = [
  [
    { text: 'SEKTOR INDUSTRI', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'MITRA STRATEGIS', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'KONTRIBUSI NYATA BAGI KAWASAN STP', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } }
  ],
  [
    { text: 'E-Commerce & Logistik', options: { color: C_SILVER } },
    { text: 'Shopee Indonesia', options: { color: C_WHITE, bold: true } },
    { text: 'Infrastruktur UMKM Hub, Co-location, sewa lahan Rp 544 jt/th', options: { color: C_SILVER } }
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
    { text: 'Laboratorium 5G, konektivitas fiber optik terintegrasi kawasan', options: { color: C_SILVER } }
  ],
  [
    { text: 'Ride-Hailing & UMKM', options: { color: C_SILVER } },
    { text: 'GoTo (Gojek-Tokopedia)', options: { color: C_WHITE, bold: true } },
    { text: 'GoTo UMKM Center, pelatihan onboarding merchant lokal', options: { color: C_SILVER } }
  ],
  [
    { text: 'Migas & Industri Berat', options: { color: C_SILVER } },
    { text: 'OGSCI Consortium', options: { color: C_AMBER, bold: true } },
    { text: 'Pelatihan migas internasional: Target Rp 6,25 M (2030)', options: { color: C_AMBER, bold: true } }
  ]
];

s6.addTable(partnersTableData, {
  x: 5.8, y: 2.0, w: 6.5, h: 3.2,
  colW: [1.8, 1.8, 2.9],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG2 },
  fontSize: 8.8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.38, 0.44, 0.44, 0.44, 0.44, 0.44, 0.48]
});

// Bottom Callout Box inside Right Container
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.8, y: 5.35, w: 6.5, h: 1.25,
  fill: '062E20', line: { color: C_EMERALD, width: 1 }, rectRadius: 0.04
});
s6.addText('💎  PRINSIP ZERO APBD BURDEN: SELURUH FASILITAS CANGGIH DIBANGUN OLEH SWASTA', {
  x: 6.0, y: 5.42, w: 6.1, h: 0.28,
  fontSize: 9.2, bold: true, color: C_EMERALD, fontFace: FONT_HEAD
});
s6.addText('Laboratorium AI, Studio Gaming, 5G Experience, hingga Cyber Security Hub tidak dibangun dengan membebani belanja modal APBD Kota Surakarta. Semuanya hadir melalui kemitraan murni Public-Private Partnership (PPP).', {
  x: 6.0, y: 5.72, w: 6.1, h: 0.78,
  fontSize: 8.5, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 12
});

s6.addNotes('Lihatlah daftar mitra ini, Bapak Walikota. Shopee, Garena, GoTo, Indosat, Bank Mandiri, hingga konsorsium migas internasional OGSCI. Mereka hadir di Solo bukan karena subsidi APBD—karena kita menerapkan prinsip Zero APBD Burden. Pemkot menyediakan wadah dan kepastian regulasi, pihak swasta menanamkan modal dan teknologinya. Hasilnya: fasilitas canggih berkelas internasional berdiri di jantung kota Surakarta, dan anak-anak muda Solo bisa belajar serta bekerja langsung di dalamnya.');

// =========================================================================
// SLIDE 07: DIAGNOSA & 5 GAP UTAMA BESERTA SOLUSI TEROBOSAN
// =========================================================================
const s7 = pptx.addSlide();
addHeader(s7, 'DIAGNOSA & 5 GAP UTAMA BESERTA SOLUSI TEROBOSANNYA', 'Refleksi Strategis & Solusi Konkret', 7,
  'Kajian Kritis Berdasarkan Dokumen Strategi STP 2026: Gap Riil, Risiko, dan Solusi Kebijakan Kepala Daerah'
);

const gapTableData = [
  [
    { text: 'NO', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR }, align: 'center' } },
    { text: 'GAP UTAMA EKSISTING', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'KONDISI FAKTUAL & AKAR MASALAH', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY_BAR } } },
    { text: 'SOLUSI TEROBOSAN 2026–2030', options: { bold: true, color: C_CYAN_LIGHT, fill: { color: C_NAVY_BAR } } }
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
// V5 Upgrade: Seamless 3-Phase Transformational Milestone Architecture
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
    target: 'Rp 7,50 M ➔ Rp 12,53 M (+67%)',
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
    target: 'Rp 12,53 M ➔ Rp 20,62 M (+65%)',
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
    target: 'Rp 20,62 M ➔ Rp 23,47 M (MANDIRI)',
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
  addBentoCard(s8, px, 1.45, 3.8, 3.9, { fill: C_BENTO_BG, accentColor: ph.accent });

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
    fontSize: 8.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'center'
  });

  // Target Box
  s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px + 0.2, y: 2.35, w: 3.4, h: 0.4,
    fill: C_DARK_BG, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
  });
  s8.addText(`Target: ${ph.target}`, {
    x: px + 0.2, y: 2.35, w: 3.4, h: 0.4,
    fontSize: 9.8, bold: true, color: C_AMBER, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
  });

  // Bullets
  const pList = ph.points.map(pt => `•  ${pt}`).join('\n\n');
  s8.addText(pList, {
    x: px + 0.2, y: 2.85, w: 3.4, h: 2.35,
    fontSize: 9.2, color: C_SILVER, fontFace: FONT_BODY, lineSpacing: 13
  });
});

// Bottom Milestone Table Strip
addBentoCard(s8, 0.8, 5.5, 11.733, 1.35, { fill: C_BENTO_BG2, line: C_CYAN });
s8.addText('RINGKASAN TARGET TAHUNAN MENUJU KEMANDIRIAN 2030 (AUDIT FINANSIAL RIIL):', {
  x: 1.0, y: 5.6, w: 11.333, h: 0.25,
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
  x: 1.0, y: 5.92, w: 11.333, h: 0.78,
  colW: [2.266, 2.266, 2.266, 2.266, 2.269],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_BENTO_BG },
  fontSize: 9.5, fontFace: FONT_HEAD, valign: 'middle',
  rowH: [0.35, 0.43]
});

s8.addNotes('Roadmap ini kami rancang dalam tiga fase yang saling membangun. 2026–2027 adalah fase pondasi: tata kelola BLUD diperkuat, lab kita terakreditasi, dan kita mulai ekspansi kerjasama OGSCI dan SMK GoGlobal. 2028–2029 adalah fase akselerasi: SBU terbentuk dan mulai mandiri seperti entitas bisnis profesional. Puncaknya, 2030, Solo Technopark beroperasi sebagai holding company inovasi daerah yang sepenuhnya mandiri—mengharumkan nama Surakarta di panggung teknologi nasional bahkan global.');

// =========================================================================
// SLIDE 09: RE-SETTING TARGET FINANSIAL BLUD: +212,8%
// V5 Upgrade: True Native PowerPoint Bar Chart & Doughnut Chart (Zero Overlap!)
// =========================================================================
const s9 = pptx.addSlide();
addHeader(s9, 'RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD +212,8%', 'Proyeksi Keuangan BLUD 2026–2030', 9,
  'Proyeksi Pendapatan Mandiri Riil dari Rp 7,50 Miliar (2026) menuju Rp 23,47 Miliar (2030) — Tanpa Subsidi APBD'
);

// Left Container: Native Bar Chart + Strategic Takeaways (w: 6.8)
addBentoCard(s9, 0.8, 1.45, 6.8, 5.35, { fill: C_BENTO_BG, accentColor: C_CYAN });
s9.addText('TREN PERTUMBUHAN PENDAPATAN BLUD 2026–2030 (dalam Rp Miliar)', {
  x: 1.0, y: 1.6, w: 6.4, h: 0.28,
  fontSize: 10, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

// Native Column Bar Chart
const financialChartData = [{
  name: 'Pendapatan BLUD (Rp Miliar)',
  labels: ['2026\n(Baseline)', '2027\n(+67%)', '2028\n(+28,5%)', '2029\n(+28%)', '2030\n(MANDIRI)'],
  values: [7.50, 12.53, 16.10, 20.62, 23.47]
}];

s9.addChart(pptx.charts.BAR, financialChartData, {
  x: 0.95, y: 1.95, w: 6.45, h: 3.1,
  barDir: 'col',
  chartColors: ['94A3B8', '38BDF8', '0EA5E9', '10B981', 'F59E0B'],
  chartColorsDiscardImageProps: true,
  showValue: true,
  dataLabelColor: 'FFFFFF',
  dataLabelFontSize: 9.5,
  dataLabelPosition: 'outEnd',
  valAxisLabelColor: '94A3B8',
  catAxisLabelColor: 'FFFFFF',
  catAxisFontSize: 8.5,
  valAxisMinVal: 0,
  valAxisMaxVal: 27,
  showLegend: false,
  valGridLine: { color: '1E3A5F', size: 1 },
  catGridLine: { style: 'none' }
});

// Bottom Box inside Left Card (NO OVERLAP! Starts at y: 5.15)
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 5.15, w: 6.4, h: 1.5,
  fill: C_BENTO_BG2, line: { color: C_AMBER, width: 1 }, rectRadius: 0.04
});
s9.addText('LOMPATAN FINANSIAL: +212,8% (3,1× LIPAT DALAM 4 TAHUN)', {
  x: 1.15, y: 5.25, w: 6.1, h: 0.25,
  fontSize: 10.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s9.addText('• Mandiri Total Tanpa APBD di 2030: Biaya operasional UPTD STP tertutup 100% dari jasa layanan mandiri.\n• Titik Impas (BEP) di Tahun 2028: Tercapai di angka Rp 16,10 Miliar seiring beroperasinya SBU manufaktur.\n• Kualitas Pendapatan Berkelanjutan: 88% pendapatan disumbang oleh riset, vokasi industri, & startup binaan.', {
  x: 1.15, y: 5.55, w: 6.1, h: 1.0,
  fontSize: 8.5, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 12
});

// Right Container: Native Doughnut Chart Revenue Stream + Top Catalysts (w: 4.733)
addBentoCard(s9, 7.8, 1.45, 4.733, 5.35, { fill: C_BENTO_BG, accentColor: C_AMBER });
s9.addText('STRUKTUR REVENUE STREAM 2030 (Rp 23,47 M)', {
  x: 8.0, y: 1.6, w: 4.333, h: 0.28,
  fontSize: 10, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

// Native Donut Chart for Revenue Stream
const revChartData = [{
  name: 'Revenue Stream 2030',
  labels: [
    'Kerjasama Program (54,6%)',
    'Pelatihan & Sertifikasi (15,7%)',
    'Diklat Mandiri (10,5%)',
    'Lain-lain BLUD / MICE (10,4%)',
    'Rekayasa Manufaktur (6,2%)',
    'Kemitraan Lahan (2,5%)'
  ],
  values: [54.6, 15.7, 10.5, 10.4, 6.2, 2.5]
}];

s9.addChart(pptx.charts.DOUGHNUT, revChartData, {
  x: 8.5, y: 1.95, w: 3.3, h: 2.05,
  chartColors: ['0EA5E9', '38BDF8', '10B981', 'F59E0B', '8B5CF6', '64748B'],
  holeSize: 55,
  showLegend: false,
  showPercent: true,
  dataLabelColor: 'FFFFFF',
  dataLabelFontSize: 7.5
});

// 2-Column Revenue Stream Legend Breakdown
const revLegend = [
  { name: 'Program Kemitraan', pct: '54,6% (Rp 12,83 M)', col: C_CYAN },
  { name: 'Pelatihan & Sertifikasi', pct: '15,7% (Rp 3,68 M)', col: C_CYAN_LIGHT },
  { name: 'Diklat Mandiri Spesialis', pct: '10,5% (Rp 2,47 M)', col: C_EMERALD },
  { name: 'Lain-lain BLUD / MICE', pct: '10,4% (Rp 2,45 M)', col: C_AMBER },
  { name: 'Rekayasa Manufaktur', pct: '6,2% (Rp 1,47 M)', col: C_PURPLE },
  { name: 'Kemitraan Lahan Eksisting', pct: '2,5% (Rp 0,58 M)', col: C_MUTED }
];

revLegend.forEach((rl, idx) => {
  const col = idx < 3 ? 0 : 1;
  const row = idx % 3;
  const rx = 8.0 + col * 2.25;
  const ry = 4.08 + row * 0.28;

  s9.addShape(pptx.shapes.OVAL, {
    x: rx, y: ry + 0.04, w: 0.1, h: 0.1,
    fill: { color: rl.col }, line: { color: rl.col }
  });
  s9.addText(`${rl.name}: `, {
    x: rx + 0.14, y: ry, w: 1.35, h: 0.22,
    fontSize: 7.2, color: C_SILVER, fontFace: FONT_BODY
  });
  s9.addText(rl.pct, {
    x: rx + 1.25, y: ry, w: 0.95, h: 0.22,
    fontSize: 7.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right'
  });
});

// Top 5 Catalysts Box at bottom of Right Container
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 8.0, y: 4.98, w: 4.333, h: 1.67,
  fill: C_BENTO_BG2, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
});
s9.addText('TOP 5 KATALISATOR PENDAPATAN 2030:', {
  x: 8.15, y: 5.06, w: 4.0, h: 0.22,
  fontSize: 8.8, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

const top5 = [
  { item: '#1 OGSCI Program Migas', nominal: 'Rp 6.250.000.000', col: C_AMBER },
  { item: '#2 Pemda e-Katalog Inovasi', nominal: 'Rp 2.250.000.000', col: C_WHITE },
  { item: '#3 Startup Binaan STP', nominal: 'Rp 2.000.000.000', col: C_EMERALD },
  { item: '#4 Sertifikasi Tenaga JFT', nominal: 'Rp 1.375.000.000', col: C_WHITE },
  { item: '#5 Tempat Uji Kompetensi', nominal: 'Rp 1.350.000.000', col: C_WHITE }
];

top5.forEach((t, idx) => {
  const ty = 5.32 + idx * 0.26;
  s9.addText(t.item, {
    x: 8.15, y: ty, w: 2.3, h: 0.24,
    fontSize: 7.8, color: C_SILVER, fontFace: FONT_BODY
  });
  s9.addText(t.nominal, {
    x: 10.5, y: ty, w: 1.7, h: 0.24,
    fontSize: 7.8, bold: true, color: t.col, fontFace: FONT_HEAD, align: 'right'
  });
});

s9.addNotes('Inilah komitmen terukur kami, Bapak Walikota. Angka-angka ini bukan proyeksi asal-asalan—setiap rupiah memiliki sumber yang teridentifikasi. Pendorong terbesar adalah program OGSCI yang akan menyumbang Rp 6,25 Miliar di 2030—itu sendirian sudah hampir menyamai total pendapatan STP tahun 2026. Ditambah program Startup Binaan yang kami proyeksi menghasilkan Rp 2 Miliar, dan e-Katalog yang membutuhkan Instruksi Walikota untuk membuka pasar Rp 2,25 Miliar. Dengan diversifikasi ini, STP akan menjadi BLUD yang sepenuhnya mandiri pada 2030.');

// =========================================================================
// =========================================================================
// SLIDE 10: BREAKDOWN MESIN PENDAPATAN 2030 PER MATA ANGGARAN
// V5.1 Upgrade: Asymmetric Executive Financial Bento & Visual Contribution Ribbon
// =========================================================================
const s10 = pptx.addSlide();
addHeader(s10, 'ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN', 'Transparansi Anggaran & Struktur Pendapatan', 10,
  'Transparansi Total Hasil Rekonsiliasi: Dari Mana Setiap Rupiah dari Target Rp 23,47 Miliar Berasal'
);

// 1. TOP VISUAL CONTRIBUTION RIBBON (Horizontal Macro Portfolio Split)
const ribbonSegments = [
  { name: 'Klaster A: Kemitraan Strategis (54,6%)', pct: 0.5465, col: C_CYAN, textCol: C_DARK_BG },
  { name: 'B: Vokasi (15,7%)', pct: 0.1568, col: C_CYAN_LIGHT, textCol: C_DARK_BG },
  { name: 'C: Diklat (10,5%)', pct: 0.1052, col: C_AMBER, textCol: C_DARK_BG },
  { name: 'D: MICE (10,4%)', pct: 0.1044, col: C_EMERALD, textCol: C_DARK_BG },
  { name: 'E: 6,2%', pct: 0.0624, col: C_PURPLE, textCol: C_WHITE },
  { name: 'F: 2,5%', pct: 0.0247, col: C_MUTED, textCol: C_WHITE }
];

let currRx = 0.8;
const totalRibbonW = 11.733;
const ribbonY = 1.40;
const ribbonH = 0.32;

ribbonSegments.forEach((seg) => {
  const segW = totalRibbonW * seg.pct;
  s10.addShape(pptx.shapes.RECTANGLE, {
    x: currRx, y: ribbonY, w: segW, h: ribbonH,
    fill: { color: seg.col }, line: { color: C_DARK_BG, width: 1 }
  });
  if (segW > 0.25) {
    const label = segW < 0.4 ? 'F' : seg.name;
    s10.addText(label, {
      x: currRx, y: ribbonY, w: segW, h: ribbonH,
      fontSize: segW > 2.0 ? 8.5 : (segW > 1.0 ? 7.5 : 6.5),
      bold: true, color: seg.textCol, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
    });
  }
  currRx += segW;
});

// 2. LEFT HERO PANEL: KLASTER A (THE 54.6% GROWTH ENGINE)
const heroX = 0.8;
const heroY = 1.82;
const heroW = 4.85;
const heroH = 4.38;

addBentoCard(s10, heroX, heroY, heroW, heroH, { fill: C_BENTO_BG, accentColor: C_CYAN });

// Hero Badge Header Container
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: heroX + 0.15, y: heroY + 0.12, w: heroW - 0.30, h: 0.74,
  fill: C_BENTO_BG2, line: { color: C_CYAN, width: 1.5 }, rectRadius: 0.05
});

// Left Header Text (Single multi-line text block)
s10.addText([
  { text: '🏆 ENGINE UTAMA (54,6% REVENUE)\n', options: { fontSize: 7.8, bold: true, color: C_CYAN_LIGHT } },
  { text: 'KLASTER A: KERJASAMA STRATEGIS\n', options: { fontSize: 8.8, bold: true, color: C_WHITE } },
  { text: '5 Program Kemitraan Unggulan 2030', options: { fontSize: 7.2, color: C_SILVER } }
], {
  x: heroX + 0.25, y: heroY + 0.14, w: 2.38, h: 0.70,
  fontFace: FONT_HEAD
});

// Right Header Nominal Box
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: heroX + 2.70, y: heroY + 0.18, w: 1.90, h: 0.62,
  fill: C_DARK_BG, line: { color: C_AMBER, width: 1 }, rectRadius: 0.04
});
s10.addText('Rp 12.825.000.000', {
  x: heroX + 2.70, y: heroY + 0.22, w: 1.90, h: 0.30,
  fontSize: 12.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
});
s10.addText('54,65% DARI TOTAL TARGET', {
  x: heroX + 2.70, y: heroY + 0.52, w: 1.90, h: 0.20,
  fontSize: 6.8, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD, align: 'center'
});

// Klaster A Line Items (Ledger Rows)
const heroItems = [
  { name: 'OGSCI Migas & Maritim (Konsorsium)', val: 'Rp 6.250.000.000', highlight: true, note: 'Anchor Revenue' },
  { name: 'Pemda Solo (e-Katalog Inovasi)', val: 'Rp 2.250.000.000', highlight: false, note: 'Instruksi Walikota' },
  { name: 'Startup Binaan STP (Komersialisasi)', val: 'Rp 2.000.000.000', highlight: false, note: 'Success Fee & Jasa' },
  { name: 'BP3MI - SMK GoGlobal (Talenta Luar Negeri)', val: 'Rp 1.500.000.000', highlight: false, note: 'Penempatan Global' },
  { name: 'Kemenperin & Mitra LPK Industri', val: 'Rp 825.000.000', highlight: false, note: 'Subsidi Pelatihan' }
];

heroItems.forEach((it, idx) => {
  const rowY = heroY + 0.95 + idx * 0.50;
  const rowH = 0.44;
  const isHi = it.highlight;

  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: heroX + 0.15, y: rowY, w: heroW - 0.30, h: rowH,
    fill: isHi ? '1E293B' : C_DARK_BG,
    line: { color: isHi ? C_AMBER : C_BORDER, width: isHi ? 1.2 : 0.8 },
    rectRadius: 0.04
  });

  // Indicator dot
  s10.addShape(pptx.shapes.OVAL, {
    x: heroX + 0.28, y: rowY + 0.18, w: 0.09, h: 0.09,
    fill: { color: isHi ? C_AMBER : C_CYAN },
    line: { color: isHi ? C_AMBER : C_CYAN }
  });

  // Name and Note
  s10.addText([
    { text: it.name, options: { bold: true, color: isHi ? C_WHITE : C_SILVER, fontSize: 8.5 } },
    { text: `  [${it.note}]`, options: { color: isHi ? C_AMBER : C_CYAN_LIGHT, fontSize: 7.2 } }
  ], {
    x: heroX + 0.42, y: rowY + 0.04, w: 2.70, h: rowH - 0.08,
    fontFace: FONT_BODY, valign: 'middle'
  });

  // Nominal
  s10.addText(it.val, {
    x: heroX + 3.10, y: rowY + 0.04, w: 1.45, h: rowH - 0.08,
    fontSize: 8.8, bold: true, color: isHi ? C_AMBER : C_WHITE,
    fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// Strategic Callout inside Hero Panel
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: heroX + 0.15, y: heroY + 3.52, w: heroW - 0.30, h: 0.74,
  fill: '0A2540', line: { color: '0284C7', width: 1 }, rectRadius: 0.04
});
s10.addText('💡 CATATAN STRATEGIS TRANSFORMASI:', {
  x: heroX + 0.25, y: heroY + 3.56, w: heroW - 0.50, h: 0.20,
  fontSize: 8.2, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});
s10.addText('Kontribusi 1 program OGSCI (Rp 6,25 M) saja sudah setara 83% total pendapatan STP tahun 2026. Ditambah e-Katalog dan Startup binaan, klaster ini menjamin kemandirian tanpa subsidi APBD.', {
  x: heroX + 0.25, y: heroY + 3.76, w: heroW - 0.50, h: 0.46,
  fontSize: 7.6, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 10
});

// 3. RIGHT AREA: 5 OTHER CLUSTERS (Row 1: B & C, Row 2: D, E, F)
const rightStartX = heroX + heroW + 0.18; // 5.83
const rightTotalW = 11.733 - heroW - 0.18; // 6.703

// ROW 1: KLASTER B & C (Talenta Vokasi & Keahlian Khusus)
const r1Y = heroY;
const r1H = 2.12;
const r1CardW = (rightTotalW - 0.15) / 2; // 3.276

// Card B: Layanan Vokasi & Sertifikasi
addBentoCard(s10, rightStartX, r1Y, r1CardW, r1H, { fill: C_BENTO_BG, accentColor: C_CYAN_LIGHT });
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: rightStartX + 0.12, y: r1Y + 0.10, w: r1CardW - 0.24, h: 0.44,
  fill: C_BENTO_BG2, line: { color: C_CYAN_LIGHT, width: 1 }, rectRadius: 0.04
});
s10.addText('KLASTER B: VOKASI & SERTIFIKASI', {
  x: rightStartX + 0.18, y: r1Y + 0.12, w: 1.70, h: 0.40,
  fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD, valign: 'middle'
});
s10.addText([
  { text: 'Rp 3.679.800.000\n', options: { bold: true, fontSize: 8.2, color: C_CYAN_LIGHT } },
  { text: '(15,7% Target 2030)', options: { fontSize: 6.8, color: C_SILVER } }
], {
  x: rightStartX + 1.85, y: r1Y + 0.12, w: 1.25, h: 0.40,
  fontFace: FONT_HEAD, align: 'right', valign: 'middle'
});

const bItems = [
  { name: 'Sertifikasi JFT Fungsional', val: 'Rp 1.375.000.000' },
  { name: 'TUK Mandiri Terakreditasi', val: 'Rp 1.350.000.000' },
  { name: 'Konsultasi Lab & Data Center', val: 'Rp 950.000.000' },
  { name: 'Prakerin Siswa SMK Solo Raya', val: 'Rp 4.800.000' }
];

bItems.forEach((bi, bidx) => {
  const by = r1Y + 0.60 + bidx * 0.36;
  s10.addShape(pptx.shapes.OVAL, {
    x: rightStartX + 0.16, y: by + 0.08, w: 0.06, h: 0.06,
    fill: { color: C_CYAN_LIGHT }, line: { color: C_CYAN_LIGHT }
  });
  s10.addText(bi.name, {
    x: rightStartX + 0.26, y: by, w: 1.85, h: 0.32,
    fontSize: 7.5, color: C_SILVER, fontFace: FONT_BODY, valign: 'middle'
  });
  s10.addText(bi.val, {
    x: rightStartX + 2.05, y: by, w: 1.08, h: 0.32,
    fontSize: 7.5, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// Card C: Diklat Mandiri Spesialis
const cardCX = rightStartX + r1CardW + 0.15;
addBentoCard(s10, cardCX, r1Y, r1CardW, r1H, { fill: C_BENTO_BG, accentColor: C_AMBER });
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: cardCX + 0.12, y: r1Y + 0.10, w: r1CardW - 0.24, h: 0.44,
  fill: C_BENTO_BG2, line: { color: C_AMBER, width: 1 }, rectRadius: 0.04
});
s10.addText('KLASTER C: DIKLAT SPESIALIS', {
  x: cardCX + 0.18, y: r1Y + 0.12, w: 1.70, h: 0.40,
  fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD, valign: 'middle'
});
s10.addText([
  { text: 'Rp 2.469.000.000\n', options: { bold: true, fontSize: 8.2, color: C_AMBER } },
  { text: '(10,5% Target 2030)', options: { fontSize: 6.8, color: C_SILVER } }
], {
  x: cardCX + 1.85, y: r1Y + 0.12, w: 1.25, h: 0.40,
  fontFace: FONT_HEAD, align: 'right', valign: 'middle'
});

const cItems = [
  { name: 'Underwater Wet Welding', val: 'Rp 846.000.000' },
  { name: 'Manufaktur & Welding Intensif', val: 'Rp 728.000.000' },
  { name: 'Desain, Manajerial & EV', val: 'Rp 555.000.000' },
  { name: 'Digital Tech (AI & Cyber)', val: 'Rp 225.000.000' },
  { name: 'Kewirausahaan & UMKM', val: 'Rp 115.000.000' }
];

cItems.forEach((ci, cidx) => {
  const cy = r1Y + 0.58 + cidx * 0.29;
  s10.addShape(pptx.shapes.OVAL, {
    x: cardCX + 0.16, y: cy + 0.08, w: 0.06, h: 0.06,
    fill: { color: C_AMBER }, line: { color: C_AMBER }
  });
  s10.addText(ci.name, {
    x: cardCX + 0.26, y: cy, w: 1.85, h: 0.26,
    fontSize: 7.2, color: C_SILVER, fontFace: FONT_BODY, valign: 'middle'
  });
  s10.addText(ci.val, {
    x: cardCX + 2.05, y: cy, w: 1.08, h: 0.26,
    fontSize: 7.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// ROW 2: KLASTER D, E, F (Asset Monetization, Precision Production, Land Leasing)
const r2Y = r1Y + r1H + 0.12; // 4.06
const r2H = 2.14;
const r2CardW = (rightTotalW - 0.24) / 3; // 2.154

// Card D: Lain-lain BLUD / MICE
const cardDX = rightStartX;
addBentoCard(s10, cardDX, r2Y, r2CardW, r2H, { fill: C_BENTO_BG, accentColor: C_EMERALD });
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: cardDX + 0.10, y: r2Y + 0.10, w: r2CardW - 0.20, h: 0.44,
  fill: C_BENTO_BG2, line: { color: C_EMERALD, width: 1 }, rectRadius: 0.04
});
s10.addText('KLASTER D: MICE & ASET', {
  x: cardDX + 0.14, y: r2Y + 0.12, w: r2CardW - 0.28, h: 0.18,
  fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
});
s10.addText('Rp 2.450.000.000 (10,4%)', {
  x: cardDX + 0.14, y: r2Y + 0.28, w: r2CardW - 0.28, h: 0.22,
  fontSize: 8.2, bold: true, color: C_EMERALD, fontFace: FONT_HEAD
});

const dItems = [
  { name: 'Sarana Kawasan / MICE', val: 'Rp 1.500 M' },
  { name: 'Co-Working & Virtual Off.', val: 'Rp 900 Jt' },
  { name: 'Jasa Giro Kas BLUD', val: 'Rp 50 Jt' }
];
dItems.forEach((di, didx) => {
  const dy = r2Y + 0.60 + didx * 0.38;
  s10.addText(di.name, {
    x: cardDX + 0.12, y: dy, w: 1.15, h: 0.32,
    fontSize: 7.2, color: C_SILVER, fontFace: FONT_BODY, valign: 'middle'
  });
  s10.addText(di.val, {
    x: cardDX + 1.20, y: dy, w: 0.84, h: 0.32,
    fontSize: 7.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// Card E: Rekayasa Manufaktur & Produksi
const cardEX = cardDX + r2CardW + 0.12;
addBentoCard(s10, cardEX, r2Y, r2CardW, r2H, { fill: C_BENTO_BG, accentColor: C_PURPLE });
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: cardEX + 0.10, y: r2Y + 0.10, w: r2CardW - 0.20, h: 0.44,
  fill: C_BENTO_BG2, line: { color: C_PURPLE, width: 1 }, rectRadius: 0.04
});
s10.addText('KLASTER E: MANUFAKTUR', {
  x: cardEX + 0.14, y: r2Y + 0.12, w: r2CardW - 0.28, h: 0.18,
  fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
});
s10.addText('Rp 1.465.000.000 (6,2%)', {
  x: cardEX + 0.14, y: r2Y + 0.28, w: r2CardW - 0.28, h: 0.22,
  fontSize: 8.2, bold: true, color: C_PURPLE, fontFace: FONT_HEAD
});

const eItems = [
  { name: 'Precision Parts Fabrikasi', val: 'Rp 1.225 M' },
  { name: 'Standarisasi Uji TKDN', val: 'Rp 135 Jt' },
  { name: 'Pesanan Pemda e-Katalog', val: 'Rp 105 Jt' }
];
eItems.forEach((ei, eidx) => {
  const ey = r2Y + 0.60 + eidx * 0.38;
  s10.addText(ei.name, {
    x: cardEX + 0.12, y: ey, w: 1.15, h: 0.32,
    fontSize: 7.2, color: C_SILVER, fontFace: FONT_BODY, valign: 'middle'
  });
  s10.addText(ei.val, {
    x: cardEX + 1.20, y: ey, w: 0.84, h: 0.32,
    fontSize: 7.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// Card F: Kemitraan Lahan Eksisting
const cardFX = cardEX + r2CardW + 0.12;
addBentoCard(s10, cardFX, r2Y, r2CardW, r2H, { fill: C_BENTO_BG, accentColor: C_MUTED });
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: cardFX + 0.10, y: r2Y + 0.10, w: r2CardW - 0.20, h: 0.44,
  fill: C_BENTO_BG2, line: { color: C_MUTED, width: 1 }, rectRadius: 0.04
});
s10.addText('KLASTER F: SEWA LAHAN', {
  x: cardFX + 0.14, y: r2Y + 0.12, w: r2CardW - 0.28, h: 0.18,
  fontSize: 7.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
});
s10.addText('Rp 579.358.831 (2,5%)', {
  x: cardFX + 0.14, y: r2Y + 0.28, w: r2CardW - 0.28, h: 0.22,
  fontSize: 8.2, bold: true, color: C_SILVER, fontFace: FONT_HEAD
});

const fItems = [
  { name: 'Shopee Co-Location', val: 'Rp 544 Jt' },
  { name: 'Katulondi Kawasan', val: 'Rp 35 Jt' }
];
fItems.forEach((fi, fidx) => {
  const fy = r2Y + 0.60 + fidx * 0.34;
  s10.addText(fi.name, {
    x: cardFX + 0.12, y: fy, w: 1.15, h: 0.28,
    fontSize: 7.2, color: C_SILVER, fontFace: FONT_BODY, valign: 'middle'
  });
  s10.addText(fi.val, {
    x: cardFX + 1.20, y: fy, w: 0.84, h: 0.28,
    fontSize: 7.2, bold: true, color: C_WHITE, fontFace: FONT_HEAD, align: 'right', valign: 'middle'
  });
});

// Policy Note inside Card F
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: cardFX + 0.10, y: r2Y + 1.34, w: r2CardW - 0.20, h: 0.70,
  fill: C_DARK_BG, line: { color: C_BORDER, width: 0.8 }, rectRadius: 0.03
});
s10.addText('📉 DE-ESKALASI SEWA:\nPorsi sewa ditekan ke 2,5% demi beralih ke inovasi bernilai tinggi.', {
  x: cardFX + 0.14, y: r2Y + 1.36, w: r2CardW - 0.28, h: 0.64,
  fontSize: 6.8, color: C_AMBER, fontFace: FONT_BODY, lineSpacing: 9
});

// 4. BOTTOM GRAND TOTAL RIBBON (y: 6.28, h: 0.58)
addBentoCard(s10, 0.8, 6.28, 11.733, 0.58, { fill: C_BENTO_BG2, line: C_AMBER });
s10.addText([
  { text: 'TOTAL TARGET PENDAPATAN BLUD 2030:   ', options: { bold: true, color: C_WHITE, fontSize: 10.5 } },
  { text: 'Rp 23.468.158.831  ', options: { bold: true, color: C_AMBER, fontSize: 13.5 } },
  { text: '(Dua Puluh Tiga Miliar Empat Ratus Enam Puluh Delapan Juta Rupiah) — Rekonsiliasi 100% Mandiri Bebas APBD', options: { italic: true, color: C_CYAN_LIGHT, fontSize: 8.2 } }
], {
  x: 1.0, y: 6.28, w: 11.333, h: 0.58,
  fontFace: FONT_HEAD, align: 'center', valign: 'middle'
});

s10.addNotes('Bapak Walikota, slide ini adalah bukti transparansi kami. Setiap mata anggaran dari Rp 23,47 Miliar sudah teridentifikasi sumbernya. Ini bukan angka anggan-anggan—ini adalah komitmen yang sudah terverifikasi bersama mitra-mitra kami. Perhatikan porsi Klaster A yang menyumbang 54,6%—OGSCI saja menghasilkan Rp 6,25 Miliar, dan e-Katalog Inovasi Rp 2,25 Miliar. Sementara itu, porsi sewa lahan pasif sengaja kami tekan tinggal 2,5%, membuktikan STP bukan sekadar penyewa lahan, melainkan pusat inovasi bernilai tambah tinggi.');

// =========================================================================
// SLIDE 11: 6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SOLO
// V5 Upgrade: Native Horizontal Bar Chart + High-Impact KPI Scorecards
// =========================================================================
const s11 = pptx.addSlide();
addHeader(s11, '6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA', 'Dampak Sosial-Ekonomi Daerah', 11,
  '"Indikator Keberhasilan Harus Mengukur IMPACT — Bukan Sekadar Aktivitas dan Pendapatan Kas" (Dokumen Strategi STP 2026)'
);

// Left Container: Native Horizontal Bar Chart (w: 5.3)
addBentoCard(s11, 0.8, 1.45, 5.3, 5.35, { fill: C_BENTO_BG, accentColor: C_CYAN });
s11.addText('SKOR CAPAIAN 6 DIMENSI DAMPAK 2030 (%)', {
  x: 1.0, y: 1.6, w: 4.9, h: 0.28,
  fontSize: 10, bold: true, color: C_CYAN_LIGHT, fontFace: FONT_HEAD
});

// Native Horizontal Bar Chart
const impactChartData = [{
  name: 'Capaian Target 2030 (%)',
  labels: [
    'Hilirisasi Riset Kampus',
    'Peluang Kerja Daerah',
    'UMKM & Startup Mandiri',
    'Akses Fasilitas Publik',
    'SDM Kompeten & Sertifikasi',
    'Kolaborasi Industri Murni'
  ],
  values: [70, 75, 80, 85, 90, 100]
}];

s11.addChart(pptx.charts.BAR, impactChartData, {
  x: 0.95, y: 1.95, w: 5.0, h: 3.1,
  barDir: 'bar',
  chartColors: ['0EA5E9', '38BDF8', '10B981', 'F59E0B', 'F97316', '10B981'],
  chartColorsDiscardImageProps: true,
  showValue: true,
  dataLabelColor: 'FFFFFF',
  dataLabelFontSize: 9,
  dataLabelPosition: 'outEnd',
  valAxisLabelColor: '94A3B8',
  catAxisLabelColor: 'FFFFFF',
  catAxisFontSize: 8,
  valAxisMinVal: 0,
  valAxisMaxVal: 115,
  showLegend: false,
  valGridLine: { color: '1E3A5F', size: 1 },
  catGridLine: { style: 'none' }
});

// Bottom Synthesis inside Left Card
s11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 5.15, w: 4.9, h: 1.5,
  fill: C_BENTO_BG2, line: { color: C_BORDER, width: 1 }, rectRadius: 0.04
});
s11.addText('FILOSOFI IMPACT MEASUREMENT KOTA:', {
  x: 1.15, y: 5.25, w: 4.6, h: 0.25,
  fontSize: 9.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s11.addText('Keberhasilan Solo Technopark tidak diukur dari megahnya gedung atau saldo kas BLUD semata, melainkan dari kesejahteraan nyata warga Surakarta: lapangan kerja yang terserap, UMKM yang naik kelas, serta keterhubungan riset kampus dengan kebutuhan industri.', {
  x: 1.15, y: 5.55, w: 4.6, h: 1.0,
  fontSize: 8.5, color: C_WHITE, fontFace: FONT_BODY, lineSpacing: 12
});

// Right Container: 6 Impact Scorecards in 3x2 Grid (w: 6.233)
const impactCards = [
  {
    icon: '👷',
    title: 'SDM LEBIH KOMPETEN',
    metric: '9.000 Peserta / Thn',
    sub: 'Kuota afirmasi 65% KTP Surakarta',
    accent: C_CYAN
  },
  {
    icon: '🌱',
    title: 'UMKM & STARTUP TUMBUH',
    metric: '80+ Wirausaha Baru',
    sub: 'Pendampingan & onboarding GoTo Hub',
    accent: C_AMBER
  },
  {
    icon: '🔬',
    title: 'RISET DEKAT KE PASAR',
    metric: '45+ Produk Hilirisasi',
    sub: 'Paten HKI, Uji NDT & Sertifikasi TKDN',
    accent: C_EMERALD
  },
  {
    icon: '💻',
    title: 'AKSES TEKNOLOGI TERBUKA',
    metric: '10.000+ Kunjungan/Thn',
    sub: 'Lab AI, Cyber & Prakerin SMK Solo Raya',
    accent: C_CYAN_LIGHT
  },
  {
    icon: '💼',
    title: 'PELUANG KERJA BARU',
    metric: 'Rp 180 M Multiplier',
    sub: 'Perputaran ekonomi hotel, resto & MICE',
    accent: C_AMBER
  },
  {
    icon: '🤝',
    title: 'KOLABORASI INDUSTRI',
    metric: '100% Investasi Swasta',
    sub: '95+ korporasi dunia tanpa beban APBD',
    accent: C_EMERALD
  }
];

impactCards.forEach((im, idx) => {
  const col = idx % 2;
  const row = Math.floor(idx / 2);
  const cx = 6.3 + col * 3.1;
  const cy = 1.45 + row * 1.8;

  addBentoCard(s11, cx, cy, 2.95, 1.68, { fill: C_BENTO_BG, accentColor: im.accent });

  // Icon & Title
  s11.addText(`${im.icon}  ${im.title}`, {
    x: cx + 0.15, y: cy + 0.12, w: 2.65, h: 0.28,
    fontSize: 8.8, bold: true, color: C_WHITE, fontFace: FONT_HEAD
  });

  // Large Metric Display
  s11.addText(im.metric, {
    x: cx + 0.15, y: cy + 0.45, w: 2.65, h: 0.6,
    fontSize: 14, bold: true, color: im.accent, fontFace: FONT_HEAD, align: 'center', valign: 'middle'
  });

  // Subtitle / Policy Affirmation
  s11.addText(im.sub, {
    x: cx + 0.15, y: cy + 1.12, w: 2.65, h: 0.45,
    fontSize: 8, color: C_SILVER, fontFace: FONT_BODY, align: 'center', lineSpacing: 10
  });
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
const outputPptxPathV5 = path.resolve(__dirname, 'PAPARAN_WALIKOTA_STP_V5.pptx');
const outputPptxPathOut = path.resolve(__dirname, '../output_paparan/PAPARAN_WALIKOTA_STP_V5.pptx');

pptx.writeFile({ fileName: outputPptxPathV5 })
  .then(() => {
    console.log(`[SUCCESS] PPTX Master V5 generated at: ${outputPptxPathV5}`);
    fs.copyFileSync(outputPptxPathV5, outputPptxPathOut);
    console.log(`[SUCCESS] Copied to output_paparan: ${outputPptxPathOut}`);
  })
  .catch(err => {
    console.error('[ERROR] Failed generating PPTX V5:', err);
    process.exit(1);
  });
