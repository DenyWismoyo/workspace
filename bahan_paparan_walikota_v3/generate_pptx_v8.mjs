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
pptx.title = 'Solo Technopark: Dari Kawasan Sains Menuju Mesin Impact Daerah (V8 - Swiss Enterprise White Edition)';
pptx.subject = 'Bahan Paparan Resmi Walikota Surakarta 2026–2030 (Edisi Eksekutif V8 Clean Enterprise)';

// =========================================================================
// COLOR PALETTE (V8: Clean Swiss Enterprise White)
// =========================================================================
const C_BG          = 'FFFFFF'; // Pure Snow White Canvas
const C_SURFACE     = 'F8FAFC'; // Crisp Off-White Panel Surface
const C_SURFACE_ALT = 'F1F5F9'; // Soft Slate Accent Surface
const C_RULE        = 'E2E8F0'; // 1px Hairline Structural Rule
const C_RULE_DARK   = 'CBD5E1'; // Defined Border Rule
const C_INK         = '0F172A'; // Primary Slate Charcoal Text (Headings)
const C_BODY        = '334155'; // Secondary Slate Gray (Body Text)
const C_MUTED       = '64748B'; // Muted Gray (Metadata & Footers)

// Enterprise Accents
const C_NAVY        = '0F2942'; // Deep Corporate Navy
const C_COBALT      = '1E40AF'; // Executive Royal Blue Accent
const C_BLUE        = '0284C7'; // Vivid Sky Blue Accent
const C_AMBER       = 'D97706'; // Warm Imperial Gold / Amber
const C_AMBER_LIGHT = 'FEF3C7'; // Soft Amber Tint
const C_EMERALD     = '059669'; // Forest Emerald Success Accent
const C_EMERALD_LIGHT = 'D1FAE5'; // Soft Emerald Tint
const C_RED         = 'DC2626'; // Alert Crimson
const C_RED_LIGHT   = 'FEE2E2'; // Soft Crimson Tint
const C_PURPLE      = '7C3AED'; // Deep Violet Accent

const FONT_HEAD     = 'Arial';
const FONT_BODY     = 'Arial';

// Helper: Standard Swiss Editorial Header & Footer for Slides 02-12
function addHeader(slide, title, category, slideNum, subtitle = '') {
  slide.background = { color: C_BG };

  // Top Minimalist Category Kicker
  slide.addText(`UPTD KST SOLO TECHNOPARK  //  ${category.toUpperCase()}`, {
    x: 0.8, y: 0.38, w: 9.5, h: 0.22,
    fontSize: 8.5, bold: true, color: C_COBALT, fontFace: FONT_HEAD
  });

  // Top Right Minimalist Slide Counter
  slide.addText(`${String(slideNum).padStart(2, '0')}  /  12`, {
    x: 10.5, y: 0.38, w: 2.033, h: 0.22,
    fontSize: 9, bold: true, color: C_MUTED, align: 'right', fontFace: FONT_HEAD
  });

  // Slide Main Headline
  slide.addText(title, {
    x: 0.8, y: 0.64, w: 11.733, h: 0.40,
    fontSize: 18, bold: true, color: C_INK, fontFace: FONT_HEAD
  });

  // Slide Subtitle
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.8, y: 1.04, w: 11.733, h: 0.26,
      fontSize: 9.5, italic: true, color: C_BODY, fontFace: FONT_BODY
    });
  }

  // Horizontal Architectural Hairline Rule
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.34, w: 11.733, h: 0.015,
    fill: { color: C_RULE }, line: { color: C_RULE, width: 0 }
  });

  // Footer Hairline Rule
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 7.02, w: 11.733, h: 0.015,
    fill: { color: C_RULE }, line: { color: C_RULE, width: 0 }
  });

  // Footer Left Attribution
  slide.addText('PEMERINTAH KOTA SURAKARTA  •  BADAN RISET DAN INOVASI DAERAH (BRIDA)  •  ROADMAP 2026–2030', {
    x: 0.8, y: 7.08, w: 8.5, h: 0.28,
    fontSize: 8, color: C_MUTED, fontFace: FONT_BODY
  });

  // Footer Right Edition Badge
  slide.addText('EDISI EKSEKUTIF ENTERPRISE (V8)', {
    x: 9.5, y: 7.08, w: 3.033, h: 0.28,
    fontSize: 8, bold: true, color: C_COBALT, align: 'right', fontFace: FONT_HEAD
  });
}

// Helper: Minimalist Architectural Panel Container
function addPanel(slide, x, y, w, h, options = {}) {
  const fill = options.fill || C_BG;
  const border = options.border || C_RULE_DARK;
  const borderWidth = options.borderWidth || 0.8;
  const topAccent = options.topAccent || null;

  // Base Box
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: border, width: borderWidth }
  });

  // Optional Top Accent Line (Clean corporate touch)
  if (topAccent) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w, h: 0.06,
      fill: { color: topAccent },
      line: { color: topAccent, width: 0 }
    });
  }
}

// =========================================================================
// SLIDE 01: COVER & FRAMING VISI BESAR (Swiss Minimalist Enterprise)
// =========================================================================
const s1 = pptx.addSlide();
s1.background = { color: C_BG };

// Top Corporate Identity Kicker
s1.addText('PEMERINTAH KOTA SURAKARTA  //  BADAN RISET DAN INOVASI DAERAH (BRIDA)', {
  x: 0.9, y: 0.65, w: 11.533, h: 0.25,
  fontSize: 9, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

// Main Headline Title
s1.addText('SOLO TECHNOPARK:\nDARI KAWASAN SAINS MENUJU MESIN IMPACT DAERAH', {
  x: 0.9, y: 1.05, w: 11.533, h: 1.45,
  fontSize: 30, bold: true, color: C_INK, fontFace: FONT_HEAD, lineSpacing: 34
});

// Subtitle
s1.addText('"Roadmap Transformasi UPTD KST Surakarta: Mewujudkan Kemandirian BLUD, Hilirisasi Riset Kampus, dan Kesejahteraan Nyata Warga Kota"', {
  x: 0.9, y: 2.58, w: 11.533, h: 0.40,
  fontSize: 11.5, italic: true, color: C_BODY, fontFace: FONT_BODY
});

// Horizontal Divider
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.9, y: 3.12, w: 11.533, h: 0.02,
  fill: { color: C_RULE_DARK }, line: { color: C_RULE_DARK, width: 0 }
});

// 4-Column Display Metric Stat Strip (Minimalist Architectural Cards)
const statStrip = [
  { val: '+212,8%', label: 'Lompatan Finansial Mandiri', sub: 'Target tumbuh 3,1x dalam 4 tahun', col: C_AMBER },
  { val: 'Rp 23,47 M', label: 'Target Pendapatan BLUD 2030', sub: '0% Subsidi Kas Daerah APBD', col: C_COBALT },
  { val: '9.000', label: 'Kapasitas SDM Terlatih / Thn', sub: 'Afirmasi min. 65% KTP Surakarta', col: C_EMERALD },
  { val: '95+', label: 'Mitra Industri Global & Nasional', sub: 'Prinsip Zero APBD Burden', col: C_PURPLE }
];

statStrip.forEach((st, idx) => {
  const sx = 0.9 + idx * 2.94;
  const sw = 2.76;
  
  addPanel(s1, sx, 3.30, sw, 1.25, { fill: C_SURFACE, border: C_RULE, topAccent: st.col });

  s1.addText(st.val, {
    x: sx + 0.15, y: 3.42, w: sw - 0.3, h: 0.45,
    fontSize: 22, bold: true, color: st.col, fontFace: FONT_HEAD
  });
  s1.addText(st.label, {
    x: sx + 0.15, y: 3.88, w: sw - 0.3, h: 0.30,
    fontSize: 8.8, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s1.addText(st.sub, {
    x: sx + 0.15, y: 4.18, w: sw - 0.3, h: 0.28,
    fontSize: 7.8, color: C_MUTED, fontFace: FONT_BODY
  });
});

// 3 Core Strategic Pillars (Clean Columnar Layout)
const coverPillars = [
  {
    num: '01',
    tag: 'PARADIGMA BARU',
    title: 'Bukan Sekadar Gedung Fisik',
    desc: 'Beralih dari pengelola sewa lahan menjadi DELIVERY UNIT inovasi BRIDA yang menghasilkan IMPACT nyata bagi masyarakat dan industri.',
    col: C_BLUE
  },
  {
    num: '02',
    tag: 'KEMANDIRIAN PENUH',
    title: 'BLUD Mandiri Tanpa APBD',
    desc: 'Lompatan pendapatan terukur dari Rp 7,50 M (2026) menuju Rp 23,47 M (2030) — tumbuh +212,8% dengan 0% subsidi kas daerah.',
    col: C_AMBER
  },
  {
    num: '03',
    tag: 'HEXAHELIX & DAMPAK',
    title: '95+ Mitra Korporasi Dunia',
    desc: 'Investasi swasta murni berkolaborasi menyerap 9.000 tenaga kerja dengan komitmen afirmasi minimal 65% warga KTP Surakarta.',
    col: C_EMERALD
  }
];

coverPillars.forEach((cp, idx) => {
  const cx = 0.9 + idx * 3.92;
  const cw = 3.68;
  addPanel(s1, cx, 4.75, cw, 1.95, { fill: C_BG, border: C_RULE_DARK, topAccent: cp.col });

  s1.addText(`${cp.num}  //  ${cp.tag}`, {
    x: cx + 0.18, y: 4.90, w: cw - 0.36, h: 0.24,
    fontSize: 8, bold: true, color: cp.col, fontFace: FONT_HEAD
  });
  s1.addText(cp.title, {
    x: cx + 0.18, y: 5.16, w: cw - 0.36, h: 0.30,
    fontSize: 11.5, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s1.addText(cp.desc, {
    x: cx + 0.18, y: 5.50, w: cw - 0.36, h: 1.05,
    fontSize: 8.8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 12
  });
});

// Footer Presenter Info
s1.addText('Disampaikan oleh: Kepala UPTD KST Solo Technopark  •  Didampingi: Kepala BRIDA Kota Surakarta  •  Surakarta, 2026', {
  x: 0.9, y: 6.95, w: 11.533, h: 0.28,
  fontSize: 8.5, color: C_MUTED, align: 'center', fontFace: FONT_BODY
});

s1.addNotes('Bapak Walikota yang kami hormati, izinkan kami mempersembahkan arah baru Solo Technopark untuk periode 2026 hingga 2030. Kita tidak lagi memposisikan Solo Technopark hanya sebagai aset fisik atau deretan gedung sewa. Dengan mandat yang Bapak berikan, kami hadir dengan proposisi yang berbeda: Solo Technopark adalah mesin impact daerah—yang mengukur keberhasilannya bukan dari saldo kas semata, melainkan dari berapa banyak warga Solo yang mendapat pekerjaan, berapa banyak riset kampus yang jadi produk bernilai, dan seberapa mandiri keuangannya tanpa beban APBD.');

// =========================================================================
// SLIDE 02: KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)
// =========================================================================
const s2 = pptx.addSlide();
addHeader(s2, 'KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)', 'Tata Kelola Pemerintahan & Kelembagaan', 2,
  'Arsitektur Nilai: BRIDA sebagai Otak Kebijakan, STP sebagai Tangan Eksekutor, Warga & Industri sebagai Penerima Dampak'
);

const colW2 = 3.65;
const stages2 = [
  {
    x: 0.8,
    num: '01 / INPUT',
    title: 'BRIDA KOTA SURAKARTA',
    role: 'Otak & Regulator Kebijakan Daerah',
    accent: C_COBALT,
    items: [
      'Merumuskan arah kebijakan riset & inovasi daerah',
      'Penyusunan kajian teknokratik & roadmap strategis',
      'Orkestrasi ekosistem Iptek antar-OPD & perguruan tinggi',
      'Monitoring & evaluasi akuntabilitas dampak sosial-ekonomi'
    ],
    kpi: 'OUTPUT: Regulasi, Roadmap, & Pengawasan Kebijakan'
  },
  {
    x: 4.84,
    num: '02 / PROSES',
    title: 'UPTD KST SOLO TECHNOPARK',
    role: 'Tangan Eksekutor & Operator Bisnis BLUD',
    accent: C_AMBER,
    items: [
      'Penyelenggaraan diklat vokasi & sertifikasi industri',
      'Inkubasi bisnis, alih teknologi & komersialisasi riset',
      'Uji coba prototipe lab presisi & penetrant testing (NDT)',
      'Pengelolaan portofolio bisnis BLUD mandiri & kemitraan'
    ],
    kpi: 'OUTPUT: Layanan Teknis, Inkubasi, & Hilirisasi Produk'
  },
  {
    x: 8.88,
    num: '03 / OUTCOME',
    title: 'WARGA SURAKARTA & INDUSTRI',
    role: 'Penerima Manfaat Utama (Benefisiari)',
    accent: C_EMERALD,
    items: [
      'SDM lokal kompeten & tersertifikasi kerja industri global',
      'UMKM & startup teknologi tumbuh mandiri omzetnya',
      'Produk inovasi kampus terserap pasar & e-Katalog Pemda',
      'Masuknya investasi swasta & penyerapan 9.000 tenaga kerja'
    ],
    kpi: 'IMPACT: Lapangan Kerja, Daya Saing, & PAD Mandiri'
  }
];

// Draw Transition Chevron Arrows between Columns
s2.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 4.52, y: 2.8, w: 0.26, h: 0.35,
  fill: { color: C_COBALT }, line: { color: C_COBALT, width: 0 }
});
s2.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 8.56, y: 2.8, w: 0.26, h: 0.35,
  fill: { color: C_AMBER }, line: { color: C_AMBER, width: 0 }
});

stages2.forEach(stg => {
  addPanel(s2, stg.x, 1.55, colW2, 4.25, { fill: C_BG, border: C_RULE_DARK, topAccent: stg.accent });

  // Stage Tag
  s2.addText(stg.num, {
    x: stg.x + 0.18, y: 1.70, w: colW2 - 0.36, h: 0.22,
    fontSize: 8.5, bold: true, color: stg.accent, fontFace: FONT_HEAD
  });

  // Stage Title
  s2.addText(stg.title, {
    x: stg.x + 0.18, y: 1.95, w: colW2 - 0.36, h: 0.32,
    fontSize: 11.5, bold: true, color: C_INK, fontFace: FONT_HEAD
  });

  // Role Subtitle
  s2.addText(stg.role, {
    x: stg.x + 0.18, y: 2.28, w: colW2 - 0.36, h: 0.26,
    fontSize: 9, bold: true, color: stg.accent, fontFace: FONT_BODY
  });

  // Hairline
  s2.addShape(pptx.shapes.RECTANGLE, {
    x: stg.x + 0.18, y: 2.60, w: colW2 - 0.36, h: 0.015,
    fill: { color: C_RULE }, line: { color: C_RULE, width: 0 }
  });

  // Bullet items
  stg.items.forEach((it, iidx) => {
    s2.addText(`•  ${it}`, {
      x: stg.x + 0.18, y: 2.75 + iidx * 0.58, w: colW2 - 0.36, h: 0.52,
      fontSize: 8.5, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
    });
  });

  // Bottom KPI Strip
  s2.addShape(pptx.shapes.RECTANGLE, {
    x: stg.x + 0.18, y: 5.25, w: colW2 - 0.36, h: 0.38,
    fill: { color: C_SURFACE_ALT }, line: { color: stg.accent, width: 0.8 }
  });
  s2.addText(stg.kpi, {
    x: stg.x + 0.18, y: 5.25, w: colW2 - 0.36, h: 0.38,
    fontSize: 8, bold: true, color: stg.accent, align: 'center', valign: 'middle', fontFace: FONT_HEAD
  });
});

// Bottom Legal Foundation Box
addPanel(s2, 0.8, 6.00, 11.733, 0.85, { fill: C_SURFACE, border: C_RULE_DARK });
s2.addText('LANDASAN HUKUM TATA KELOLA PEMKOT SURAKARTA:', {
  x: 1.0, y: 6.08, w: 11.333, h: 0.24,
  fontSize: 8.5, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});
s2.addText('• Perwali Surakarta No. 15 Tahun 2022: Kedudukan UPTD KST Solo Technopark sebagai unit pelaksana teknis operasional di bawah naungan BRIDA.\n• Perwali Surakarta No. 38 Tahun 2022: Pola Tata Kelola BLUD UPTD KST Solo Technopark yang memberikan fleksibilitas bisnis sehat & akuntabel.', {
  x: 1.0, y: 6.32, w: 11.333, h: 0.46,
  fontSize: 8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
});

s2.addNotes('Slide ini menjelaskan arsitektur kelembagaan: BRIDA adalah otak dan perumus kebijakan, sedangkan Solo Technopark adalah tangan eksekutor di lapangan. Dengan status UPTD berpola BLUD di bawah BRIDA, STP memiliki fleksibilitas operasional untuk bermitra dengan industri, namun tetap berada dalam koridor akuntabilitas dan prioritas pembangunan Kota Surakarta.');

// =========================================================================
// SLIDE 03: FONDASI HUKUM YANG KOKOH: MANDAT & FLEKSIBILITAS BLUD
// =========================================================================
const s3 = pptx.addSlide();
addHeader(s3, 'FONDASI HUKUM YANG KOKOH: MANDAT & FLEKSIBILITAS BLUD', 'Landasan Regulasi & Tata Kelola', 3,
  'Tiga Pilar Regulasi yang Memberi Kewenangan, Fleksibilitas Bisnis Sehat, dan Arah Strategis KST'
);

const colW3 = 3.65;
const legalPillars = [
  {
    x: 0.8,
    reg: 'PERMENDAGRI 79/2018',
    sub: 'Tata Kelola Keuangan BLUD',
    quote: '"Fleksibilitas BLUD adalah keleluasaan dalam pola pengelolaan keuangan dengan menerapkan praktik bisnis yang sehat untuk meningkatkan layanan kepada masyarakat tanpa mencari keuntungan semata."',
    checks: [
      'Keleluasaan bermitra dengan swasta tanpa birokrasi kaku APBD',
      'Mengelola dan memutar pendapatan fungsional jasa secara mandiri',
      'Rekrutmen tenaga ahli & instruktur profesional sesuai kebutuhan pasar'
    ],
    accent: C_COBALT
  },
  {
    x: 4.84,
    reg: 'PERPRES 106/2017',
    sub: 'Kawasan Sains & Teknologi (KST)',
    quote: '"KST berfungsi sebagai wahana kerja sama riset berkelanjutan antara Pemerintah, PT, dan Industri, wajib menyediakan 4 layanan: Teknis, Teknologi, Inkubasi, & Pendukung."',
    checks: [
      'Mandat nasional fasilitasi hilirisasi & spin-off perusahaan riset',
      'Kewajiban menyediakan 4 pilar layanan iptek terintegrasi',
      'Landasan kolaborasi hexahelix multi-stakeholder formal'
    ],
    accent: C_AMBER
  },
  {
    x: 8.88,
    reg: 'PERWALI 15 & 38/2022',
    sub: 'Struktur UPTD & Pola Tata Kelola',
    quote: '"Menetapkan struktur UPTD KST Solo Technopark dalam naungan BRIDA Kota Surakarta dengan pola tata kelola BLUD profesional berorientasi layanan publik."',
    checks: [
      'Payung hukum operasional tingkat kota yang sah dan mengikat',
      'Pembagian peran teknis yang presisi antara BRIDA dan STP',
      'Dasar legalitas pemanfaatan aset kawasan & kerjasama industri'
    ],
    accent: C_EMERALD
  }
];

legalPillars.forEach(lp => {
  addPanel(s3, lp.x, 1.55, colW3, 4.25, { fill: C_BG, border: C_RULE_DARK, topAccent: lp.accent });

  s3.addText(lp.reg, {
    x: lp.x + 0.18, y: 1.70, w: colW3 - 0.36, h: 0.28,
    fontSize: 11, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s3.addText(lp.sub, {
    x: lp.x + 0.18, y: 1.98, w: colW3 - 0.36, h: 0.22,
    fontSize: 8.5, bold: true, color: lp.accent, fontFace: FONT_HEAD
  });

  // Quote Box
  s3.addShape(pptx.shapes.RECTANGLE, {
    x: lp.x + 0.18, y: 2.26, w: colW3 - 0.36, h: 1.40,
    fill: { color: C_SURFACE }, line: { color: C_RULE, width: 0.8 }
  });
  s3.addText(lp.quote, {
    x: lp.x + 0.24, y: 2.30, w: colW3 - 0.48, h: 1.32,
    fontSize: 8, italic: true, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
  });

  // Checks
  lp.checks.forEach((ch, cidx) => {
    s3.addText(`✔  ${ch}`, {
      x: lp.x + 0.18, y: 3.78 + cidx * 0.62, w: colW3 - 0.36, h: 0.56,
      fontSize: 8.2, color: C_INK, fontFace: FONT_BODY, lineSpacing: 11
    });
  });
});

// Bottom Yuridis Conclusion Callout
addPanel(s3, 0.8, 6.00, 11.733, 0.85, { fill: C_NAVY, border: C_COBALT });
s3.addText('KESIMPULAN YURIDIS: Solo Technopark telah memiliki legitimasi hukum lengkap dari tingkat nasional (Perpres & Permendagri) hingga regulasi kepala daerah (Perwali) untuk bermanuver cepat, profesional, dan akuntabel.', {
  x: 1.0, y: 6.08, w: 11.333, h: 0.68,
  fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s3.addNotes('Bapak Walikota, kami ingin menegaskan bahwa transformasi ini berdiri di atas fondasi hukum yang sangat kokoh. Permendagri 79/2018 memberi mandat fleksibilitas BLUD agar kita bisa mengelola pendapatan fungsional langsung untuk peningkatan layanan tanpa hambatan birokrasi APBD. Perpres 106/2017 memberi mandat 4 layanan sains dan teknologi. Dan Perwali Kota Surakarta mengunci tata kelolanya secara akuntabel.');

// =========================================================================
// SLIDE 04: EMPAT PILAR LAYANAN TERINTEGRASI: MEMENUHI MANDAT KST
// =========================================================================
const s4 = pptx.addSlide();
addHeader(s4, 'EMPAT PILAR LAYANAN TERINTEGRASI: MEMENUHI MANDAT KST NASIONAL', 'Mandat Layanan Perpres 106/2017', 4,
  'Portofolio Layanan Terintegrasi yang Menjawab Kebutuhan Industri Spesialis & Masyarakat Kota Surakarta'
);

const qW = 5.72;
const qH = 2.40;
const fourPillars = [
  {
    x: 0.8, y: 1.55,
    tag: 'PILAR 1: LAYANAN TEKNIS (DIKLAT VOKASI INDUSTRI)',
    accent: C_BLUE,
    items: [
      '• Mekanik, Otomasi & Desain Manufaktur Presisi',
      '• Welding Intensif (Pengelasan Konstruksi & Pipa)',
      '• OGSCI: Oil, Gas, Coal, Shipbuilding & Infra Training',
      '• Underwater Wet Welding (Keahlian Langka Nasional)',
      '• Teknologi Kendaraan Listrik (EV Conversion Lab)',
      '• Cyber Security & Artificial Intelligence Practitioner'
    ]
  },
  {
    x: 6.81, y: 1.55,
    tag: 'PILAR 2: PENGEMBANGAN TEKNOLOGI (REKAYASA MANUFAKTUR)',
    accent: C_AMBER,
    items: [
      '• Jasa Layanan Rekayasa Industri & Permesinan Presisi',
      '• Produksi Komponen Presisi (Precision Parts Industri)',
      '• Pembuatan & Uji Prototipe IKM/UKM Kota Surakarta',
      '• Non-Destructive Testing: Uji Penetrant Testing (NDT)',
      '• Konsultasi Laboratorium Terpadu & Data Center',
      '• Standarisasi Mutu & Sertifikasi TKDN / SNI Produk'
    ]
  },
  {
    x: 0.8, y: 4.10,
    tag: 'PILAR 3: LAYANAN INKUBASI (HILIRISASI KE PASAR)',
    accent: C_EMERALD,
    items: [
      '• Pra-Inkubasi & Inkubasi Bisnis Startup Berbasis Riset',
      '• Fasilitasi HKI, Paten Produk, & Hak Cipta Inovator',
      '• Hilirisasi Riset Perguruan Tinggi ke Pasar Industri',
      '• National Cyber Security Hub & AI Experience Center',
      '• Digital Technopark, Gaming Hub, & GoTo UMKM Center',
      '• Co-Working Space, Virtual Office & Akses Permodalan'
    ]
  },
  {
    x: 6.81, y: 4.10,
    tag: 'PILAR 4: LAYANAN PENDUKUNG (COMMUNITY & EDUCATION)',
    accent: C_PURPLE,
    items: [
      '• Prakerin Kerja Industri Siswa SMK Se-Solo Raya',
      '• Tempat Uji Kompetensi (TUK Mandiri) & Sertifikasi JFT',
      '• Solo Science Center (Pusat Edukasi IPTEK & Wisata Sains)',
      '• Pemanfaatan Sarana Kawasan, Gedung Expo & MICE',
      '• Kemitraan Komunitas Kreatif & Kunjungan Industri',
      '• Pusat Kolaborasi Multipihak Pemkot Surakarta'
    ]
  }
];

fourPillars.forEach(fp => {
  addPanel(s4, fp.x, fp.y, qW, qH, { fill: C_BG, border: C_RULE_DARK, topAccent: fp.accent });

  s4.addText(fp.tag, {
    x: fp.x + 0.18, y: fp.y + 0.14, w: qW - 0.36, h: 0.28,
    fontSize: 9.2, bold: true, color: fp.accent, fontFace: FONT_HEAD
  });

  // 2 Columns inside quadrant
  const halfW = (qW - 0.40) / 2;
  const col1 = fp.items.slice(0, 3);
  const col2 = fp.items.slice(3, 6);

  col1.forEach((it, i) => {
    s4.addText(it, {
      x: fp.x + 0.18, y: fp.y + 0.48 + i * 0.58, w: halfW, h: 0.54,
      fontSize: 8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
    });
  });

  col2.forEach((it, i) => {
    s4.addText(it, {
      x: fp.x + 0.18 + halfW, y: fp.y + 0.48 + i * 0.58, w: halfW, h: 0.54,
      fontSize: 8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
    });
  });
});

// Bottom Callout Strip
addPanel(s4, 0.8, 6.60, 11.733, 0.32, { fill: C_SURFACE_ALT, border: C_RULE });
s4.addText('"BLUD menggeser paradigma: Dari sekadar pengelolaan kawasan fisik menjadi penyedia solusi teknologi yang responsif terhadap kebutuhan industri dan masyarakat."', {
  x: 0.8, y: 6.60, w: 11.733, h: 0.32,
  fontSize: 8.5, italic: true, bold: true, color: C_COBALT, align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s4.addNotes('Solo Technopark bukan sekadar tempat pelatihan dasar. Mandat Perpres 106/2017 mewajibkan empat pilar: Layanan Teknis, Pengembangan Teknologi, Inkubasi, dan Pendukung. Dari underwater welding hingga AI and cyber security, STP hadir menyediakan keahlian spesialis bernilai tinggi.');

// =========================================================================
// SLIDE 05: PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR
// =========================================================================
const s5 = pptx.addSlide();
addHeader(s5, 'PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR NYATA', 'Mekanisme Hilirisasi Inovasi', 5,
  'Solo Technopark sebagai Jembatan: Menghubungkan Riset Kampus Menuju Produk Bernilai Ekonomi bagi Industri'
);

// 5 Horizontal Stage Columns Connected by an Architectural Progress Rail
const stgW = 2.22;
const stages5 = [
  { num: '01', name: 'RISET & INVENSI', sub: 'Kampus & Peneliti', desc: 'Invensi UNS, UMS, ISI, Poltek Solo Raya. Riset dosen & mahasiswa.', accent: C_BLUE },
  { num: '02', name: 'SELEKSI & KURASI', sub: 'STP & BRIDA', desc: 'Kurasi komersial & relevansi kebutuhan industri serta problem kota Solo.', accent: C_COBALT },
  { num: '03', name: 'PROTOTYPING', sub: 'Lab Presisi STP', desc: 'Uji teknis, penetrant testing, validasi mutu produk & SNI/TKDN.', accent: C_AMBER },
  { num: '04', name: 'INKUBASI BISNIS', sub: 'Inkubator STP', desc: 'Penyusunan business model, proteksi HKI/Paten, akses modal & offtaker.', accent: C_EMERALD },
  { num: '05', name: 'KOMERSIALISASI', sub: 'Pasar & Industri', desc: 'Penetrasi e-Katalog Lokal OPD, captive market pemda & industri nasional.', accent: C_PURPLE }
];

stages5.forEach((st, idx) => {
  const sx = 0.8 + idx * 2.38;
  addPanel(s5, sx, 1.55, stgW, 2.05, { fill: C_BG, border: C_RULE_DARK, topAccent: st.accent });

  s5.addText(`TAHAP ${st.num}`, {
    x: sx + 0.12, y: 1.68, w: stgW - 0.24, h: 0.22,
    fontSize: 8.5, bold: true, color: st.accent, fontFace: FONT_HEAD
  });
  s5.addText(st.name, {
    x: sx + 0.12, y: 1.92, w: stgW - 0.24, h: 0.28,
    fontSize: 10, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s5.addText(st.sub, {
    x: sx + 0.12, y: 2.22, w: stgW - 0.24, h: 0.22,
    fontSize: 8, italic: true, color: C_MUTED, fontFace: FONT_BODY
  });
  s5.addText(st.desc, {
    x: sx + 0.12, y: 2.50, w: stgW - 0.24, h: 0.98,
    fontSize: 7.8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
  });

  // Chevron arrow between columns
  if (idx < 4) {
    s5.addShape(pptx.shapes.RIGHT_ARROW, {
      x: sx + stgW + 0.04, y: 2.45, w: 0.12, h: 0.20,
      fill: { color: C_MUTED }, line: { color: C_MUTED, width: 0 }
    });
  }
});

// Bottom Left: Table of Real Products
addPanel(s5, 0.8, 3.80, 6.7, 3.05, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s5.addText('CONTOH TEKNOLOGI YANG DIHILIRKAN & DIINKUBASI DI STP (DATA RIIL)', {
  x: 0.95, y: 3.92, w: 6.4, h: 0.24,
  fontSize: 8.8, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const protoTableData = [
  [
    { text: 'KLASTER TEKNOLOGI', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } },
    { text: 'CONTOH PRODUK INOVASI', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } },
    { text: 'STATUS TAHAPAN', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_BODY } },
    { text: 'Mesin Sortasi Biji Kopi Otomatis', options: { color: C_INK, bold: true } },
    { text: 'Prototype Validasi', options: { color: C_COBALT, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_BODY } },
    { text: 'Teknologi Pertanian Presisi IoT & AI', options: { color: C_INK, bold: true } },
    { text: 'Uji Lapangan', options: { color: C_COBALT, align: 'center' } }
  ],
  [
    { text: 'Smart Agriculture IoT', options: { color: C_BODY } },
    { text: 'Mesin Tetas Telur Pintar Mandiri', options: { color: C_INK, bold: true } },
    { text: 'Siap e-Katalog', options: { color: C_EMERALD, bold: true, align: 'center' } }
  ],
  [
    { text: 'Deep Tech Mobility', options: { color: C_BODY } },
    { text: 'Autonomous Vehicle Elektrik Kampus', options: { color: C_INK, bold: true } },
    { text: 'R&D Bersama', options: { color: C_AMBER, align: 'center' } }
  ],
  [
    { text: 'GovTech & Pelatihan', options: { color: C_BODY } },
    { text: 'VR Pelatihan Industri & Chatbot AI OPD', options: { color: C_INK, bold: true } },
    { text: 'Komersial Aktif', options: { color: C_EMERALD, bold: true, align: 'center' } }
  ]
];

s5.addTable(protoTableData, {
  x: 0.95, y: 4.22, w: 6.4, h: 2.45,
  colW: [1.8, 3.2, 1.4],
  border: { type: 'solid', pt: 0.8, color: C_RULE },
  fill: { color: C_BG },
  fontSize: 8, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.32, 0.38, 0.38, 0.38, 0.38, 0.38]
});

// Bottom Right: Active Facilities
addPanel(s5, 7.7, 3.80, 4.833, 3.05, { fill: C_SURFACE, border: C_RULE_DARK, topAccent: C_AMBER });
s5.addText('EKOSISTEM FASILITAS PENDUKUNG AKTIF', {
  x: 7.9, y: 3.92, w: 4.4, h: 0.24,
  fontSize: 8.8, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const facs = [
  { name: 'National Cyber Security Hub', desc: 'Pusat ketahanan digital & talenta siber nasional bersama BSSN.' },
  { name: 'Gaming Hub & Digital Studio', desc: 'Pengembangan game, animasi, & konten kreatif e-sport.' },
  { name: 'AI Experience Center', desc: 'Laboratorium kecerdasan buatan terapan industri & showcase solusi AI.' },
  { name: 'GoTo UMKM Center', desc: 'Akselerasi digitalisasi, kurasi produk, & onboarding pedagang lokal.' },
  { name: 'Laboratorium Rekayasa Presisi', desc: 'Fasilitas fabrikasi mesin, permesinan presisi & pengujian NDT industri.' }
];

facs.forEach((fc, fidx) => {
  s5.addText(`•  ${fc.name}`, {
    x: 7.9, y: 4.24 + fidx * 0.50, w: 4.4, h: 0.22,
    fontSize: 8.5, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s5.addText(fc.desc, {
    x: 8.1, y: 4.46 + fidx * 0.50, w: 4.2, h: 0.24,
    fontSize: 7.5, color: C_BODY, fontFace: FONT_BODY
  });
});

s5.addNotes('Di slide ini Bapak Walikota dapat melihat bahwa hilirisasi riset bukan lagi angan-angan. Pipeline 5 tahap dari Riset hingga Komersialisasi sudah berjalan. Contoh produk riilnya: mesin sortasi biji kopi, pertanian presisi IoT, mesin tetas telur yang siap masuk e-Katalog, hingga autonomous vehicle. Semuanya ditopang fasilitas aktif seperti Cyber Security Hub, AI Experience Center, dan GoTo UMKM Center.');

// =========================================================================
// SLIDE 06: EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD
// =========================================================================
const s6 = pptx.addSlide();
addHeader(s6, 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD', 'Jejaring Kemitraan Strategis', 6,
  'Kepercayaan Korporasi Global & Nasional yang Menanamkan Investasi Langsung di Jantung Kota Surakarta'
);

// Left Container: Native Doughnut Chart (w: 4.8)
addPanel(s6, 0.8, 1.55, 4.8, 5.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s6.addText('KOMPOSISI 6 AKTOR HEXAHELIX (95+ MITRA)', {
  x: 0.95, y: 1.70, w: 4.5, h: 0.26,
  fontSize: 9, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const hexaChartData = [{
  name: 'Komposisi Mitra',
  labels: ['Industri & Bisnis', 'Akademisi/Kampus', 'Komunitas & UMKM', 'Pemerintah/BUMN', 'Media Massa', 'Lembaga Global'],
  values: [35, 20, 15, 12, 8, 5]
}];

s6.addChart(pptx.charts.DOUGHNUT, hexaChartData, {
  x: 1.1, y: 2.05, w: 4.2, h: 2.4,
  chartColors: ['0284C7', '1E40AF', '059669', 'D97706', '7C3AED', 'DB2777'],
  chartColorsDiscardImageProps: true,
  holeSize: 60,
  showValue: true,
  dataLabelColor: '0F172A',
  dataLabelFontSize: 8.5,
  showLegend: false
});

// Hexahelix Legend Matrix (Clean 2-Column)
const hexaLegend = [
  { name: 'Industri & Bisnis:', val: '37% (35)', col: '0284C7' },
  { name: 'Pemerintah / BUMN:', val: '13% (12)', col: 'D97706' },
  { name: 'Akademisi / Kampus:', val: '21% (20)', col: '1E40AF' },
  { name: 'Media Massa:', val: '8% (8)', col: '7C3AED' },
  { name: 'Komunitas & UMKM:', val: '16% (15)', col: '059669' },
  { name: 'Lembaga Global:', val: '5% (5)', col: 'DB2777' }
];

hexaLegend.forEach((hl, hidx) => {
  const colX = hidx % 2 === 0 ? 0.95 : 3.35;
  const rowY = 4.55 + Math.floor(hidx / 2) * 0.32;
  s6.addShape(pptx.shapes.OVAL, {
    x: colX, y: rowY + 0.05, w: 0.12, h: 0.12,
    fill: { color: hl.col }, line: { color: hl.col, width: 0 }
  });
  s6.addText(`${hl.name}  ${hl.val}`, {
    x: colX + 0.18, y: rowY, w: 2.15, h: 0.24,
    fontSize: 7.8, color: C_INK, fontFace: FONT_BODY
  });
});

// Total Strategic Summary Callout
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 0.95, y: 5.65, w: 4.5, h: 1.05,
  fill: { color: C_SURFACE }, line: { color: C_RULE, width: 0.8 }
});
s6.addText('TOTAL: 95+ MITRA STRATEGIS AKTIF\n• 100% Investasi Swasta & PPP (0% Beban APBD)\n• Akses fasilitas 5G, Cyber Security, AI, & Migas\n• Saluran penyerapan kerja langsung bagi warga Solo', {
  x: 1.05, y: 5.72, w: 4.3, h: 0.90,
  fontSize: 8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 12
});

// Right Container: Corporate Partners Table (w: 6.7)
addPanel(s6, 5.8, 1.55, 6.733, 5.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_AMBER });
s6.addText('MITRA INDUSTRI KELAS DUNIA YANG AKTIF BERINVESTASI DI STP', {
  x: 6.0, y: 1.70, w: 6.3, h: 0.26,
  fontSize: 9, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const partnerTableData = [
  [
    { text: 'SEKTOR INDUSTRI', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } },
    { text: 'MITRA STRATEGIS', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } },
    { text: 'KONTRIBUSI NYATA BAGI KAWASAN STP', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } }
  ],
  [
    { text: 'E-Commerce & Logistik', options: { color: C_BODY } },
    { text: 'Shopee Indonesia', options: { color: C_INK, bold: true } },
    { text: 'Infrastruktur UMKM Hub, Co-location, sewa lahan Rp 544 jt/th', options: { color: C_BODY } }
  ],
  [
    { text: 'Gaming & Kreatif', options: { color: C_BODY } },
    { text: 'Garena', options: { color: C_INK, bold: true } },
    { text: 'Gaming studio, beasiswa talenta digital, turnamen esports', options: { color: C_BODY } }
  ],
  [
    { text: 'Perbankan & Fintech', options: { color: C_BODY } },
    { text: 'Bank Mandiri', options: { color: C_INK, bold: true } },
    { text: 'Mandiri Innovation Hub, fasilitasi modal startup binaan', options: { color: C_BODY } }
  ],
  [
    { text: 'Telekomunikasi', options: { color: C_BODY } },
    { text: 'Indosat Ooredoo Hutchison', options: { color: C_INK, bold: true } },
    { text: 'Laboratorium 5G, konektivitas fiber optik terintegrasi kawasan', options: { color: C_BODY } }
  ],
  [
    { text: 'Ride-Hailing & UMKM', options: { color: C_BODY } },
    { text: 'GoTo (Gojek-Tokopedia)', options: { color: C_INK, bold: true } },
    { text: 'GoTo UMKM Center, pelatihan onboarding merchant lokal', options: { color: C_BODY } }
  ],
  [
    { text: 'Migas & Industri Berat', options: { color: C_BODY } },
    { text: 'OGSCI Consortium', options: { color: C_AMBER, bold: true } },
    { text: 'Pelatihan migas internasional: Target Rp 6,25 M (2030)', options: { color: C_AMBER, bold: true } }
  ]
];

s6.addTable(partnerTableData, {
  x: 6.0, y: 2.05, w: 6.333, h: 3.10,
  colW: [1.7, 1.8, 2.833],
  border: { type: 'solid', pt: 0.8, color: C_RULE },
  fill: { color: C_BG },
  fontSize: 8.2, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.36, 0.44, 0.44, 0.44, 0.44, 0.44, 0.48]
});

// Bottom Callout Box inside Right Container (Zero APBD Principle)
s6.addShape(pptx.shapes.RECTANGLE, {
  x: 6.0, y: 5.30, w: 6.333, h: 1.35,
  fill: { color: C_EMERALD_LIGHT }, line: { color: C_EMERALD, width: 1.2 }
});
s6.addText('PRINSIP ZERO APBD BURDEN: SELURUH FASILITAS CANGGIH DIBANGUN OLEH SWASTA', {
  x: 6.15, y: 5.40, w: 6.0, h: 0.24,
  fontSize: 8.8, bold: true, color: '065F46', fontFace: FONT_HEAD
});
s6.addText('Laboratorium AI, Studio Gaming, 5G Experience, hingga Cyber Security Hub tidak dibangun dengan membebani belanja modal APBD Kota Surakarta. Semuanya hadir melalui kemitraan murni Public-Private Partnership (PPP).', {
  x: 6.15, y: 5.68, w: 6.0, h: 0.85,
  fontSize: 8.2, color: '047857', fontFace: FONT_BODY, lineSpacing: 11
});

s6.addNotes('Lihatlah daftar mitra ini, Bapak Walikota. Shopee, Garena, GoTo, Indosat, Bank Mandiri, hingga konsorsium migas internasional OGSCI. Mereka hadir di Solo bukan karena subsidi APBD—karena kita menerapkan prinsip Zero APBD Burden. Pemkot menyediakan wadah dan kepastian regulasi, pihak swasta menanamkan modal dan teknologinya. Hasilnya: fasilitas canggih berkelas internasional berdiri di jantung kota Surakarta, dan anak-anak muda Solo bisa belajar serta bekerja langsung di dalamnya.');

// =========================================================================
// SLIDE 07: DIAGNOSA 5 GAP UTAMA BESERTA SOLUSI TEROBOSANNYA
// =========================================================================
const s7 = pptx.addSlide();
addHeader(s7, 'DIAGNOSA & 5 GAP UTAMA BESERTA SOLUSI TEROBOSANNYA', 'Refleksi Strategis & Solusi Konkret', 7,
  'Kajian Kritis Berdasarkan Dokumen Strategi STP 2026: Gap Riil, Risiko, dan Solusi Kebijakan Kepala Daerah'
);

const gapItems = [
  {
    num: 'GAP 01',
    name: 'Jenis Pelatihan Masih Terbatas',
    issue: 'Keterbatasan instruktur ahli internal & sarana diklat berbasis industri masa depan.',
    solution: 'Gandeng Korporasi Global (OGSCI, Kemenperin) Buka Kelas Spesialis',
    detail: 'Welder Bawah Air, AI Engineer, Teknisi Baterai EV, & Drone Operator.',
    isPriority: false,
    accent: C_COBALT
  },
  {
    num: 'GAP 02',
    name: 'R&D Belum Mengarah Hilirisasi',
    issue: 'Riset kampus berjalan terisolasi, kurang terpetakan, dan belum berorientasi pasar.',
    solution: 'Bentuk Technology Transfer Office (TTO) & Kurasi Riset Terapan Kampus',
    detail: 'Konsorsium 5 Perguruan Tinggi Solo Raya (UNS, UMS, ISI, dll.) berbasis problem kota.',
    isPriority: false,
    accent: C_COBALT
  },
  {
    num: 'GAP 03',
    name: 'Investasi Kawasan Terbatas',
    issue: 'Lahan kosong belum optimal dikerjasamakan; promosi investor belum sistematis.',
    solution: 'Zonasi Investasi Teknologi Terpadu & Long-Term Revenue Sharing',
    detail: 'Skema kemitraan 10–25 tahun bagi korporasi teknologi global tanpa beban belanja APBD.',
    isPriority: false,
    accent: C_BLUE
  },
  {
    num: 'GAP 04',
    name: 'Kemandirian Finansial Rendah',
    issue: 'Pendapatan masih bertumpu sewa lahan; hardware lab belum optimal menghasilkan jasa.',
    solution: 'Diversifikasi 6 Revenue Stream Mandiri & Optimalisasi Aset Industri',
    detail: 'Uji NDT penetrant, TUK/JFT mandiri, rekayasa manufaktur, co-working, & startup binaan.',
    isPriority: false,
    accent: C_EMERALD
  },
  {
    num: 'GAP 05',
    name: 'Promosi & Penyerapan Rendah',
    issue: 'Networking kurang ter-maintenance; belum ada regulasi penyerapan inovasi lokal.',
    solution: 'Penerbitan Instruksi Walikota untuk e-Katalog Lokal Afirmatif',
    detail: 'Wajibkan OPD Pemkot menyerap produk/layanan tenant binaan STP (Target: Rp 2,25 Miliar).',
    isPriority: true,
    accent: C_AMBER
  }
];

gapItems.forEach((g, gidx) => {
  const gy = 1.48 + gidx * 0.86;
  const gh = 0.78;

  // 1. Problem Box (Left: w: 4.8)
  addPanel(s7, 0.8, gy, 4.80, gh, { fill: C_BG, border: g.isPriority ? C_AMBER : C_RULE_DARK });

  // Gap Badge
  s7.addShape(pptx.shapes.RECTANGLE, {
    x: 0.92, y: gy + 0.10, w: 0.75, h: 0.22,
    fill: { color: g.isPriority ? C_AMBER_LIGHT : C_RED_LIGHT },
    line: { color: g.isPriority ? C_AMBER : C_RED, width: 0.8 }
  });
  s7.addText(g.num, {
    x: 0.92, y: gy + 0.10, w: 0.75, h: 0.22,
    fontSize: 7.5, bold: true, color: g.isPriority ? C_AMBER : C_RED, align: 'center', fontFace: FONT_HEAD
  });

  // Gap Title
  s7.addText(g.name, {
    x: 1.75, y: gy + 0.08, w: 3.75, h: 0.24,
    fontSize: 8.8, bold: true, color: C_INK, fontFace: FONT_HEAD
  });

  // Gap Issue
  s7.addText(`Akar Masalah: ${g.issue}`, {
    x: 0.92, y: gy + 0.36, w: 4.58, h: 0.38,
    fontSize: 7.6, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 10
  });

  // 2. Connector Arrow
  s7.addShape(pptx.shapes.RIGHT_ARROW, {
    x: 5.68, y: gy + 0.24, w: 0.34, h: 0.26,
    fill: { color: g.isPriority ? C_AMBER : C_COBALT },
    line: { color: g.isPriority ? C_AMBER : C_COBALT, width: 0 }
  });

  // 3. Solution Box (Right: w: 6.433)
  addPanel(s7, 6.10, gy, 6.433, gh, { fill: g.isPriority ? C_AMBER_LIGHT : C_SURFACE, border: g.isPriority ? C_AMBER : C_RULE_DARK, topAccent: g.accent });

  // Solution Header
  s7.addText(`✔  ${g.solution}`, {
    x: 6.25, y: gy + 0.08, w: 6.15, h: 0.26,
    fontSize: 8.8, bold: true, color: g.isPriority ? C_AMBER : g.accent, fontFace: FONT_HEAD
  });

  // Solution Detail
  s7.addText(g.detail, {
    x: 6.25, y: gy + 0.36, w: 6.15, h: 0.38,
    fontSize: 7.8, color: C_INK, fontFace: FONT_BODY, lineSpacing: 10
  });
});

// Bottom Priority Callout (Quick Win Inkwal)
addPanel(s7, 0.8, 5.92, 11.733, 0.90, { fill: C_AMBER_LIGHT, border: C_AMBER, borderWidth: 1.2 });
s7.addText('PRIORITAS UTAMA #5: INSTRUKSI WALIKOTA e-KATALOG LOKAL AFIRMATIF (QUICK WIN 30 HARI)', {
  x: 1.0, y: 6.00, w: 11.333, h: 0.26,
  fontSize: 9.5, bold: true, color: '92400E', fontFace: FONT_HEAD
});
s7.addText('Target: Membuka captive market Rp 2,25 Miliar serapan belanja OPD Pemkot Surakarta di 2030 bagi 80+ produk tenant dan startup binaan STP tanpa perlu alokasi pagu belanja baru di APBD. Hanya butuh 1 lembar Instruksi Walikota.', {
  x: 1.0, y: 6.28, w: 11.333, h: 0.48,
  fontSize: 8.5, color: '78350F', fontFace: FONT_BODY, lineSpacing: 12
});

s7.addNotes('Bapak Walikota, dokumen strategi internal STP memetakan 5 gap riil. Dan yang paling kritikal adalah GAP nomor 5: produk inovasi lokal belum terserap karena belum ada regulasi afirmatif. Solusi terobosan nomor 5 ini adalah Quick Win: kami memohon 1 lembar Instruksi Walikota untuk mewajibkan OPD menyerap produk binaan STP via e-Katalog lokal. Ini langsung menggaransi pasar Rp 2,25 Miliar di 2030.');

// =========================================================================
// SLIDE 08: ROADMAP 3 FASE TRANSFORMASI (2026–2030)
// =========================================================================
const s8 = pptx.addSlide();
addHeader(s8, 'ROADMAP 3 FASE TRANSFORMASI: DARI CENTRE MENUJU HOLDING INOVASI DAERAH', 'Peta Jalan Strategis Kelembagaan', 8,
  'Milestone Terukur dan Bertahap Menuju Kemandirian Finansial Penuh Tanpa Subsidi Belanja APBD 2030'
);

const colW8 = 3.65;
const phases8 = [
  {
    x: 0.8,
    phase: 'FASE 1 (2026–2027)',
    name: 'CENTRE OF INNOVATION & CONNECTIVITY',
    target: 'Target: Rp 7,50 M ➔ Rp 12,53 M (+67%)',
    accent: C_BLUE,
    milestones: [
      'Pondasi tata kelola kelembagaan BLUD akuntabel',
      'Akreditasi laboratorium industri & standarisasi uji',
      'Sertifikasi mutu diklat vokasi & kemitraan OGSCI',
      'Pembentukan Konsorsium Riset 5 Kampus Solo Raya',
      'Aktivasi TUK Mandiri & program SMK GoGlobal'
    ]
  },
  {
    x: 4.84,
    phase: 'FASE 2 (2028–2029)',
    name: 'ENTREPRENEUR TECHNOPARK (SBU MANDIRI)',
    target: 'Target: Rp 12,53 M ➔ Rp 20,62 M (+65%)',
    accent: C_AMBER,
    milestones: [
      'Pembentukan Strategic Business Unit (SBU) profesional',
      'Hilirisasi 15–20 produk inovasi komersial ber-TKDN',
      'Strategi corporate marketing & penetrasi pasar nasional',
      'Ekspansi revenue OGSCI Migas mencapai Rp 5,5 Miliar',
      'Operasionalisasi Co-Working & Virtual Office penuh'
    ]
  },
  {
    x: 8.88,
    phase: 'FASE 3 (2029–2030)',
    name: 'GLOBAL MARKET POSITIONING & HOLDING',
    target: 'Target: Rp 20,62 M ➔ Rp 23,47 M (MANDIRI)',
    accent: C_EMERALD,
    milestones: [
      'Kemandirian finansial penuh: 0% subsidi belanja APBD',
      'Operasional holding company inovasi daerah pertama',
      'Pusat pengembangan Indonesia Digital Technopark',
      'Realisasi target pendapatan mandiri Rp 23,47 Miliar',
      'Ekspansi jejaring kemitraan dan penempatan global'
    ]
  }
];

// Chevron connectors between phases
s8.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 4.52, y: 2.8, w: 0.26, h: 0.35,
  fill: { color: C_BLUE }, line: { color: C_BLUE, width: 0 }
});
s8.addShape(pptx.shapes.RIGHT_ARROW, {
  x: 8.56, y: 2.8, w: 0.26, h: 0.35,
  fill: { color: C_AMBER }, line: { color: C_AMBER, width: 0 }
});

phases8.forEach(ph => {
  addPanel(s8, ph.x, 1.55, colW8, 4.05, { fill: C_BG, border: C_RULE_DARK, topAccent: ph.accent });

  s8.addText(ph.phase, {
    x: ph.x + 0.18, y: 1.70, w: colW8 - 0.36, h: 0.22,
    fontSize: 8.5, bold: true, color: ph.accent, fontFace: FONT_HEAD
  });
  s8.addText(ph.name, {
    x: ph.x + 0.18, y: 1.95, w: colW8 - 0.36, h: 0.32,
    fontSize: 10, bold: true, color: C_INK, fontFace: FONT_HEAD
  });

  // Target Box
  s8.addShape(pptx.shapes.RECTANGLE, {
    x: ph.x + 0.18, y: 2.32, w: colW8 - 0.36, h: 0.30,
    fill: { color: C_SURFACE_ALT }, line: { color: C_RULE, width: 0.8 }
  });
  s8.addText(ph.target, {
    x: ph.x + 0.18, y: 2.32, w: colW8 - 0.36, h: 0.30,
    fontSize: 8, bold: true, color: ph.accent, align: 'center', valign: 'middle', fontFace: FONT_HEAD
  });

  // Milestones
  ph.milestones.forEach((m, midx) => {
    s8.addText(`•  ${m}`, {
      x: ph.x + 0.18, y: 2.75 + midx * 0.52, w: colW8 - 0.36, h: 0.48,
      fontSize: 8.2, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
    });
  });
});

// Bottom Revenue Trajectory Summary Table
addPanel(s8, 0.8, 5.75, 11.733, 1.10, { fill: C_SURFACE, border: C_RULE_DARK });
s8.addText('RINGKASAN TARGET TAHUNAN MENUJU KEMANDIRIAN 2030 (AUDIT FINANSIAL RIIL):', {
  x: 1.0, y: 5.82, w: 11.333, h: 0.22,
  fontSize: 8.2, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const trajTableData = [
  [
    { text: 'TAHUN 2026 (Baseline)', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } },
    { text: 'TAHUN 2027 (+67,0%)', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } },
    { text: 'TAHUN 2028 (+28,5%)', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } },
    { text: 'TAHUN 2029 (+28,0%)', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } },
    { text: 'TAHUN 2030 (MANDIRI)', options: { bold: true, color: C_AMBER, fill: { color: C_AMBER_LIGHT }, align: 'center' } }
  ],
  [
    { text: 'Rp 7.503.392.698', options: { align: 'center', color: C_BODY } },
    { text: 'Rp 12.532.185.031', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Rp 16.100.658.831', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Rp 20.615.658.831', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Rp 23.468.158.831', options: { align: 'center', color: C_AMBER, bold: true } }
  ]
];

s8.addTable(trajTableData, {
  x: 1.0, y: 6.08, w: 11.333, h: 0.65,
  colW: [2.26, 2.26, 2.26, 2.26, 2.293],
  border: { type: 'solid', pt: 0.8, color: C_RULE },
  fill: { color: C_BG },
  fontSize: 8.5, fontFace: FONT_HEAD, valign: 'middle',
  rowH: [0.30, 0.32]
});

s8.addNotes('Roadmap ini kami bagi menjadi 3 fase terukur: Fase 1 (2026-2027) fokus pada akreditasi dan sertifikasi diklat. Fase 2 (2028-2029) ekspansi SBU mandiri dan hilirisasi produk TKDN. Dan puncaknya di Fase 3 (2030), Solo Technopark mencapai kemandirian total dengan pendapatan Rp 23,47 Miliar dan 0% subsidi kas daerah.');

// =========================================================================
// SLIDE 09: RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD
// =========================================================================
const s9 = pptx.addSlide();
addHeader(s9, 'RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD +212,8%', 'Proyeksi Keuangan BLUD 2026–2030', 9,
  'Proyeksi Pendapatan Mandiri Riil dari Rp 7,50 Miliar (2026) menuju Rp 23,47 Miliar (2030) — Tanpa Subsidi APBD'
);

// Left Container: Native Column Bar Chart
addPanel(s9, 0.8, 1.55, 5.75, 5.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s9.addText('TREN PERTUMBUHAN PENDAPATAN BLUD 2026–2030 (DALAM RP MILIAR)', {
  x: 0.95, y: 1.70, w: 5.4, h: 0.25,
  fontSize: 8.8, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const financialChartData = [{
  name: 'Pendapatan BLUD (Rp Miliar)',
  labels: ['2026\n(Baseline)', '2027\n(+67%)', '2028\n(+28,5%)', '2029\n(+28%)', '2030\n(MANDIRI)'],
  values: [7.50, 12.53, 16.10, 20.62, 23.47]
}];

s9.addChart(pptx.charts.BAR, financialChartData, {
  x: 0.95, y: 2.05, w: 5.45, h: 3.0,
  barDir: 'col',
  chartColors: ['94A3B8', '0284C7', '1E40AF', '059669', 'D97706'],
  chartColorsDiscardImageProps: true,
  showValue: true,
  dataLabelColor: '0F172A',
  dataLabelFontSize: 9,
  dataLabelPosition: 'outEnd',
  valAxisLabelColor: '64748B',
  catAxisLabelColor: '0F172A',
  catAxisFontSize: 8.5,
  valAxisMinVal: 0,
  valAxisMaxVal: 27,
  showLegend: false,
  valGridLine: { color: C_RULE, size: 1 },
  catGridLine: { style: 'none' }
});

// Bottom Callout Box inside Left Card
addPanel(s9, 0.95, 5.20, 5.45, 1.50, { fill: C_SURFACE, border: C_RULE });
s9.addText('LOMPATAN FINANSIAL: +212,8% (3,1× LIPAT DALAM 4 TAHUN)', {
  x: 1.10, y: 5.30, w: 5.15, h: 0.24,
  fontSize: 9, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s9.addText('• Mandiri Total Tanpa APBD di 2030: Biaya operasional UPTD STP tertutup 100% dari jasa layanan mandiri.\n• Titik Impas (BEP) di Tahun 2028: Tercapai di angka Rp 16,10 Miliar seiring beroperasinya SBU manufaktur.\n• Kualitas Pendapatan Berkelanjutan: 88% pendapatan disumbang oleh riset, vokasi industri, & startup binaan.', {
  x: 1.10, y: 5.58, w: 5.15, h: 1.0,
  fontSize: 8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
});

// Right Container: Native Doughnut Chart of Revenue Stream 2030
addPanel(s9, 6.78, 1.55, 5.753, 5.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_AMBER });
s9.addText('STRUKTUR REVENUE STREAM 2030 (RP 23,47 M)', {
  x: 6.95, y: 1.70, w: 5.4, h: 0.25,
  fontSize: 8.8, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});

const revenueChartData = [{
  name: 'Revenue 2030',
  labels: [
    'Klaster A: Kemitraan Strategis (54,6%)',
    'Klaster B: Vokasi & Sertifikasi (15,7%)',
    'Klaster C: Diklat Spesialis (10,5%)',
    'Klaster D: MICE & Sarana (10,4%)',
    'Klaster E: Rekayasa Manufaktur (6,2%)',
    'Klaster F: Sewa Lahan (2,5%)'
  ],
  values: [12.825, 3.680, 2.469, 2.450, 1.465, 0.579]
}];

s9.addChart(pptx.charts.DOUGHNUT, revenueChartData, {
  x: 7.0, y: 2.05, w: 5.3, h: 2.2,
  chartColors: ['0284C7', '1E40AF', 'D97706', '059669', '7C3AED', '94A3B8'],
  chartColorsDiscardImageProps: true,
  holeSize: 60,
  showValue: true,
  dataLabelColor: '0F172A',
  dataLabelFontSize: 8.2,
  showLegend: false
});

// Legend Matrix (2 Column)
const revLegend = [
  { name: 'Program Kemitraan:', val: '54,6% (Rp 12,83 M)', col: '0284C7' },
  { name: 'Lain-lain BLUD / MICE:', val: '10,4% (Rp 2,45 M)', col: '059669' },
  { name: 'Pelatihan & Sertifikasi:', val: '15,7% (Rp 3,68 M)', col: '1E40AF' },
  { name: 'Rekayasa Manufaktur:', val: '6,2% (Rp 1,47 M)', col: '7C3AED' },
  { name: 'Diklat Mandiri Spesialis:', val: '10,5% (Rp 2,47 M)', col: 'D97706' },
  { name: 'Kemitraan Lahan Eksisting:', val: '2,5% (Rp 0,58 M)', col: '94A3B8' }
];

revLegend.forEach((rl, ridx) => {
  const colX = ridx % 2 === 0 ? 6.95 : 9.75;
  const rowY = 4.35 + Math.floor(ridx / 2) * 0.32;
  s9.addShape(pptx.shapes.OVAL, {
    x: colX, y: rowY + 0.05, w: 0.12, h: 0.12,
    fill: { color: rl.col }, line: { color: rl.col, width: 0 }
  });
  s9.addText(`${rl.name}\n${rl.val}`, {
    x: colX + 0.18, y: rowY, w: 2.55, h: 0.30,
    fontSize: 7.2, color: C_INK, fontFace: FONT_BODY, lineSpacing: 9
  });
});

// Top 5 Catalysts Panel
addPanel(s9, 6.95, 5.40, 5.4, 1.30, { fill: C_SURFACE, border: C_RULE });
s9.addText('TOP 5 KATALISATOR PENDAPATAN 2030:', {
  x: 7.05, y: 5.46, w: 5.2, h: 0.20,
  fontSize: 8, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const top5 = [
  { item: '#1 OGSCI Program Migas', nominal: 'Rp 6.250.000.000', col: C_AMBER },
  { item: '#2 Pemda e-Katalog Inovasi', nominal: 'Rp 2.250.000.000', col: C_COBALT },
  { item: '#3 Startup Binaan STP', nominal: 'Rp 2.000.000.000', col: C_EMERALD },
  { item: '#4 Sertifikasi Tenaga JFT', nominal: 'Rp 1.375.000.000', col: C_INK },
  { item: '#5 Tempat Uji Kompetensi', nominal: 'Rp 1.350.000.000', col: C_INK }
];

top5.forEach((t, tidx) => {
  const ty = 5.68 + tidx * 0.19;
  s9.addText(t.item, {
    x: 7.05, y: ty, w: 3.2, h: 0.18,
    fontSize: 7.2, color: C_BODY, fontFace: FONT_BODY
  });
  s9.addText(t.nominal, {
    x: 10.25, y: ty, w: 2.0, h: 0.18,
    fontSize: 7.5, bold: true, color: t.col, align: 'right', fontFace: FONT_HEAD
  });
});

s9.addNotes('Bapak Walikota, angka Rp 23,47 Miliar ini bukan angka tebak-tebakan. Ini hasil audit matematis berbasis potensi riil. Lompatan +212,8% ini memiliki katalisator utama: kemitraan migas internasional OGSCI senilai Rp 6,25 Miliar dan e-Katalog Inovasi Pemda Rp 2,25 Miliar. Yang paling menggembirakan: porsi sewa lahan kita tekan hingga tinggal 2,5%, beralih ke pendapatan bernilai tambah tinggi.');

// =========================================================================
// SLIDE 10: ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN
// =========================================================================
const s10 = pptx.addSlide();
addHeader(s10, 'ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN', 'Transparansi Anggaran & Struktur Pendapatan', 10,
  'Transparansi Total Hasil Rekonsiliasi: Dari Mana Setiap Rupiah dari Target Rp 23,47 Miliar Berasal'
);

// Macro Split Top Ribbon
const clustersRibbon = [
  { name: 'Klaster A: Kemitraan Strategis (54,6%)', pct: 0.5465, col: '0284C7', textCol: 'FFFFFF' },
  { name: 'B: Vokasi (15,7%)', pct: 0.1568, col: '1E40AF', textCol: 'FFFFFF' },
  { name: 'C: Diklat (10,5%)', pct: 0.1052, col: 'D97706', textCol: 'FFFFFF' },
  { name: 'D: MICE (10,4%)', pct: 0.1044, col: '059669', textCol: 'FFFFFF' },
  { name: 'E: 6,2%', pct: 0.0624, col: '7C3AED', textCol: 'FFFFFF' },
  { name: 'F: 2,5%', pct: 0.0247, col: '94A3B8', textCol: 'FFFFFF' }
];

let curRibbonX = 0.8;
const totalRibbonW = 11.733;
clustersRibbon.forEach(cr => {
  const rw = totalRibbonW * cr.pct;
  s10.addShape(pptx.shapes.RECTANGLE, {
    x: curRibbonX, y: 1.48, w: rw - 0.02, h: 0.28,
    fill: { color: cr.col }, line: { color: cr.col, width: 0 }
  });
  s10.addText(cr.name, {
    x: curRibbonX, y: 1.48, w: rw - 0.02, h: 0.28,
    fontSize: 7.5, bold: true, color: cr.textCol, align: 'center', valign: 'middle', fontFace: FONT_HEAD
  });
  curRibbonX += rw;
});

// Left Hero Box: Klaster A (w: 4.8)
addPanel(s10, 0.8, 1.88, 4.80, 4.40, { fill: C_BG, border: C_COBALT, borderWidth: 1.5, topAccent: C_COBALT });

s10.addText('ENGINE UTAMA (54,6% REVENUE)\nKLASTER A: KERJASAMA STRATEGIS', {
  x: 0.98, y: 1.98, w: 2.50, h: 0.50,
  fontSize: 8.8, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

// Big Nominal Box Inside Hero
s10.addShape(pptx.shapes.RECTANGLE, {
  x: 3.50, y: 1.98, w: 1.95, h: 0.58,
  fill: { color: C_AMBER_LIGHT }, line: { color: C_AMBER, width: 1 }
});
s10.addText([
  { text: 'Rp 12.825.000.000\n', options: { fontSize: 10, bold: true, color: C_AMBER } },
  { text: '54,65% DARI TOTAL TARGET', options: { fontSize: 6.8, bold: true, color: '92400E' } }
], {
  x: 3.50, y: 1.98, w: 1.95, h: 0.58,
  align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

// 5 Key Programs of Klaster A
const klasterAItems = [
  { name: 'OGSCI Migas & Maritim (Konsorsium)', sub: '[Anchor Revenue Mandiri]', nominal: 'Rp 6.250.000.000', hi: true },
  { name: 'Pemda Solo (e-Katalog Inovasi)', sub: '[Instruksi Walikota]', nominal: 'Rp 2.250.000.000', hi: false },
  { name: 'Startup Binaan STP (Komersialisasi)', sub: '[Success Fee & Jasa]', nominal: 'Rp 2.000.000.000', hi: false },
  { name: 'BP3MI - SMK GoGlobal (Talenta Luar Negeri)', sub: '[Penempatan Global]', nominal: 'Rp 1.500.000.000', hi: false },
  { name: 'Kemenperin & Mitra LPK Industri', sub: '[Subsidi Pelatihan]', nominal: 'Rp 825.000.000', hi: false }
];

klasterAItems.forEach((it, idx) => {
  const iy = 2.68 + idx * 0.58;
  s10.addShape(pptx.shapes.RECTANGLE, {
    x: 0.98, y: iy, w: 4.44, h: 0.52,
    fill: { color: it.hi ? C_AMBER_LIGHT : C_SURFACE },
    line: { color: it.hi ? C_AMBER : C_RULE, width: 0.8 }
  });

  s10.addText([
    { text: `${it.name}\n`, options: { bold: true, color: C_INK, fontSize: 8.2 } },
    { text: it.sub, options: { color: it.hi ? C_AMBER : C_MUTED, fontSize: 7 } }
  ], {
    x: 1.10, y: iy + 0.04, w: 2.80, h: 0.44,
    fontFace: FONT_BODY, lineSpacing: 9
  });

  s10.addText(it.nominal, {
    x: 3.80, y: iy + 0.08, w: 1.50, h: 0.36,
    fontSize: 8.5, bold: true, color: it.hi ? C_AMBER : C_COBALT, align: 'right', fontFace: FONT_HEAD
  });
});

// Strategic Note at bottom of Klaster A
s10.addShape(pptx.shapes.RECTANGLE, {
  x: 0.98, y: 5.68, w: 4.44, h: 0.50,
  fill: { color: C_SURFACE_ALT }, line: { color: C_RULE, width: 0.8 }
});
s10.addText('CATATAN STRATEGIS: 1 program konsorsium OGSCI (Rp 6,25 M) setara 83% pendapatan STP tahun 2026. Ditambah e-Katalog Pemda, klaster ini menjamin kemandirian tanpa APBD.', {
  x: 1.06, y: 5.72, w: 4.28, h: 0.42,
  fontSize: 7.2, color: C_COBALT, fontFace: FONT_BODY, lineSpacing: 9
});

// Right Columns: Klaster B, C, D, E, F (w: 6.7)
// Top Right Row: Klaster B & C
addPanel(s10, 5.75, 1.88, 3.30, 2.15, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s10.addText('KLASTER B: VOKASI & SERTIFIKASI\nRp 3.679.800.000 (15,7%)', {
  x: 5.88, y: 1.96, w: 3.04, h: 0.38,
  fontSize: 8.2, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});
const bItems = [
  { name: 'Sertifikasi JFT Fungsional', val: 'Rp 1.375.000.000' },
  { name: 'TUK Mandiri Terakreditasi', val: 'Rp 1.350.000.000' },
  { name: 'Konsultasi Lab & Data Center', val: 'Rp 950.000.000' },
  { name: 'Prakerin SMK Solo Raya', val: 'Rp 4.800.000' }
];
bItems.forEach((bi, bidx) => {
  const by = 2.40 + bidx * 0.38;
  s10.addText(bi.name, { x: 5.88, y: by, w: 1.80, h: 0.34, fontSize: 7.5, color: C_BODY, fontFace: FONT_BODY });
  s10.addText(bi.val, { x: 7.68, y: by, w: 1.24, h: 0.34, fontSize: 7.5, bold: true, color: C_INK, align: 'right', fontFace: FONT_HEAD });
});

addPanel(s10, 9.20, 1.88, 3.333, 2.15, { fill: C_BG, border: C_RULE_DARK, topAccent: C_AMBER });
s10.addText('KLASTER C: DIKLAT SPESIALIS\nRp 2.469.000.000 (10,5%)', {
  x: 9.33, y: 1.96, w: 3.07, h: 0.38,
  fontSize: 8.2, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
const cItems = [
  { name: 'Underwater Wet Welding', val: 'Rp 846.000.000' },
  { name: 'Manufaktur & Welding Intensif', val: 'Rp 728.000.000' },
  { name: 'Desain, Manajerial & EV', val: 'Rp 555.000.000' },
  { name: 'Digital Tech (AI & Cyber)', val: 'Rp 225.000.000' },
  { name: 'Kewirausahaan & UMKM', val: 'Rp 115.000.000' }
];
cItems.forEach((ci, cidx) => {
  const cy = 2.38 + cidx * 0.31;
  s10.addText(ci.name, { x: 9.33, y: cy, w: 1.80, h: 0.28, fontSize: 7.2, color: C_BODY, fontFace: FONT_BODY });
  s10.addText(ci.val, { x: 11.13, y: cy, w: 1.27, h: 0.28, fontSize: 7.2, bold: true, color: C_INK, align: 'right', fontFace: FONT_HEAD });
});

// Bottom Right Row: Klaster D, E, F
addPanel(s10, 5.75, 4.15, 2.15, 2.13, { fill: C_BG, border: C_RULE_DARK, topAccent: C_EMERALD });
s10.addText('KLASTER D: MICE\nRp 2.450.000.000 (10,4%)', {
  x: 5.85, y: 4.22, w: 1.95, h: 0.34,
  fontSize: 7.8, bold: true, color: C_EMERALD, fontFace: FONT_HEAD
});
s10.addText('• Sarana/MICE: Rp 1,50 M\n• Co-Working: Rp 900 Jt\n• Jasa Giro Kas: Rp 50 Jt', {
  x: 5.85, y: 4.62, w: 1.95, h: 1.40,
  fontSize: 7.5, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 12
});

addPanel(s10, 8.05, 4.15, 2.15, 2.13, { fill: C_BG, border: C_RULE_DARK, topAccent: C_PURPLE });
s10.addText('KLASTER E: LAB & NDT\nRp 1.465.000.000 (6,2%)', {
  x: 8.15, y: 4.22, w: 1.95, h: 0.34,
  fontSize: 7.8, bold: true, color: C_PURPLE, fontFace: FONT_HEAD
});
s10.addText('• Fabrikasi Mesin: Rp 1,22 M\n• Uji Mutu SNI: Rp 135 Jt\n• Pesanan Pemda: Rp 105 Jt', {
  x: 8.15, y: 4.62, w: 1.95, h: 1.40,
  fontSize: 7.5, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 12
});

addPanel(s10, 10.35, 4.15, 2.183, 2.13, { fill: C_SURFACE, border: C_RULE_DARK, topAccent: '94A3B8' });
s10.addText('KLASTER F: SEWA LAHAN\nRp 579.358.831 (2,5%)', {
  x: 10.45, y: 4.22, w: 1.98, h: 0.34,
  fontSize: 7.8, bold: true, color: C_MUTED, fontFace: FONT_HEAD
});
s10.addText('• Shopee Hub: Rp 544 Jt\n• Katulondi: Rp 35 Jt\n\nDE-ESKALASI SEWA:\nDitekan ke 2,5% demi beralih ke inovasi.', {
  x: 10.45, y: 4.62, w: 1.98, h: 1.50,
  fontSize: 7.2, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
});

// Bottom Total Banner
addPanel(s10, 0.8, 6.38, 11.733, 0.58, { fill: C_NAVY, border: C_AMBER, borderWidth: 1.5 });
s10.addText([
  { text: 'TOTAL TARGET PENDAPATAN BLUD 2030:   ', options: { bold: true, color: 'FFFFFF', fontSize: 10 } },
  { text: 'Rp 23.468.158.831   ', options: { bold: true, color: '#F59E0B', fontSize: 11.5 } },
  { text: '(Dua Puluh Tiga Miliar Empat Ratus Enam Puluh Delapan Juta Rupiah) — 100% Mandiri Bebas APBD', options: { italic: true, color: '#93C5FD', fontSize: 8 } }
], {
  x: 1.0, y: 6.38, w: 11.333, h: 0.58,
  align: 'center', valign: 'middle', fontFace: FONT_HEAD
});

s10.addNotes('Di slide ini Bapak Walikota dapat melihat rincian per klaster pendapatan. Klaster A adalah engine utama dengan porsi 54,6% atau Rp 12,82 Miliar. Klaster B vokasi menyumbang Rp 3,68 M. Klaster C diklat Rp 2,47 M. Klaster D MICE Rp 2,45 M. Klaster E manufaktur Rp 1,47 M. Dan sewa lahan hanya tersisa 2,5%. Total keseluruhan adalah Rp 23.468.158.831 yang mengunci kemandirian penuh.');

// =========================================================================
// SLIDE 11: 6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA
// =========================================================================
const s11 = pptx.addSlide();
addHeader(s11, '6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA', 'Dampak Sosial-Ekonomi Daerah', 11,
  '"Indikator Keberhasilan Harus Mengukur IMPACT — Bukan Sekadar Aktivitas dan Pendapatan Kas" (Dokumen Strategi STP 2026)'
);

// Left Container: Horizontal Bar Chart of 6 Dimensions (w: 5.4)
addPanel(s11, 0.8, 1.55, 5.40, 5.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s11.addText('SKOR CAPAIAN 6 DIMENSI DAMPAK 2030 (%)', {
  x: 0.95, y: 1.70, w: 5.1, h: 0.25,
  fontSize: 8.8, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const impactChartData = [{
  name: 'Capaian Dampak (%)',
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
  x: 0.95, y: 2.05, w: 5.10, h: 3.10,
  barDir: 'bar',
  chartColors: ['0284C7', '1E40AF', '059669', 'D97706', 'EA580C', '059669'],
  chartColorsDiscardImageProps: true,
  showValue: true,
  dataLabelColor: '0F172A',
  dataLabelFontSize: 8.5,
  dataLabelPosition: 'outEnd',
  valAxisLabelColor: '64748B',
  catAxisLabelColor: '0F172A',
  catAxisFontSize: 8.2,
  valAxisMinVal: 0,
  valAxisMaxVal: 110,
  showLegend: false,
  catGridLine: { color: C_RULE, size: 1 },
  valGridLine: { style: 'none' }
});

// Bottom Impact Philosophy
addPanel(s11, 0.95, 5.25, 5.10, 1.45, { fill: C_SURFACE, border: C_RULE });
s11.addText('FILOSOFI IMPACT MEASUREMENT KOTA:', {
  x: 1.10, y: 5.35, w: 4.8, h: 0.22,
  fontSize: 8.5, bold: true, color: C_AMBER, fontFace: FONT_HEAD
});
s11.addText('Keberhasilan Solo Technopark tidak diukur dari megahnya gedung atau saldo kas BLUD semata, melainkan dari kesejahteraan nyata warga Surakarta: lapangan kerja yang terserap, UMKM yang naik kelas, serta keterhubungan riset kampus dengan kebutuhan industri.', {
  x: 1.10, y: 5.60, w: 4.8, h: 1.0,
  fontSize: 7.8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
});

// Right Container: 6 Impact Scorecards in 2x3 Grid (w: 6.1)
const impactKPIs = [
  { num: '9.000 Peserta/Thn', label: 'SDM LEBIH KOMPETEN', desc: 'Afirmasi kuota minimal 65% warga KTP Surakarta.', col: C_COBALT },
  { num: '80+ Wirausaha', label: 'UMKM & STARTUP TUMBUH', desc: 'Pendampingan inkubasi & digital onboarding GoTo Hub.', col: C_AMBER },
  { num: '45+ Produk', label: 'RISET DEKAT KE PASAR', desc: 'Paten HKI, Uji Presisi NDT & sertifikasi TKDN.', col: C_EMERALD },
  { num: '10.000+ Visit/Thn', label: 'AKSES TEKNOLOGI TERBUKA', desc: 'Kunjungan lab AI, Cyber, & magang siswa SMK Solo Raya.', col: C_BLUE },
  { num: 'Rp 180 M Multiplier', label: 'PELUANG KERJA BARU', desc: 'Estimasi perputaran ekonomi hotel, resto & MICE.', col: C_PURPLE },
  { num: '100% Swasta', label: 'KOLABORASI INDUSTRI MURNI', desc: '95+ korporasi dunia berinvestasi tanpa beban belanja APBD.', col: C_EMERALD }
];

impactKPIs.forEach((kpi, kidx) => {
  const kx = 6.42 + (kidx % 2) * 3.05;
  const ky = 1.55 + Math.floor(kidx / 2) * 1.76;
  const kw = 2.95;
  const kh = 1.66;

  addPanel(s11, kx, ky, kw, kh, { fill: C_BG, border: C_RULE_DARK, topAccent: kpi.col });

  s11.addText(kpi.label, {
    x: kx + 0.14, y: ky + 0.12, w: kw - 0.28, h: 0.22,
    fontSize: 8, bold: true, color: kpi.col, fontFace: FONT_HEAD
  });
  s11.addText(kpi.num, {
    x: kx + 0.14, y: ky + 0.38, w: kw - 0.28, h: 0.40,
    fontSize: 14.5, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s11.addText(kpi.desc, {
    x: kx + 0.14, y: ky + 0.82, w: kw - 0.28, h: 0.72,
    fontSize: 7.8, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 10
  });
});

s11.addNotes('Bapak Walikota, kami memegang teguh filosofi bahwa keberhasilan STP tidak diukur dari saldo kas BLUD semata. Inilah 6 dimensi dampaknya bagi warga Solo: 9.000 orang dilatih per tahun dengan afirmasi 65% KTP Solo, 80 wirausaha baru, 45 produk riset ke pasar, 10.000 kunjungan publik, dan dampak pengganda ekonomi hingga Rp 180 Miliar.');

// =========================================================================
// SLIDE 12: 3 PERMOHONAN KEBIJAKAN KEPADA WALIKOTA & RENCANA AKSI 100 HARI
// =========================================================================
const s12 = pptx.addSlide();
addHeader(s12, '3 PERMOHONAN KEBIJAKAN KEPADA WALIKOTA & RENCANA AKSI 100 HARI', 'Dukungan Kepemimpinan & Komitmen Eksekusi', 12,
  'Kami Tidak Memohon Tambahan Pagu APBD — Kami Memohon 3 Payung Regulasi Kepemimpinan Walikota'
);

const colW12 = 3.65;
const asks = [
  {
    x: 0.8,
    priority: 'PRIORITAS UTAMA // QUICK WIN 30 HARI',
    title: 'PERMOHONAN 1: INSTRUKSI WALIKOTA AFIRMASI e-KATALOG',
    desc: 'Mengarahkan seluruh OPD Pemkot memprioritaskan belanja produk teknologi & layanan tenant STP via e-Katalog Lokal.',
    target: 'Target: Rp 2,25 M / Thn di 2030 (80+ Produk Terserap)',
    status: '⚡ Bisa langsung diterbitkan dalam 30 hari pertama',
    accent: C_BLUE,
    tint: C_SURFACE
  },
  {
    x: 4.84,
    priority: 'INVESTASI JANGKA PANJANG (10–25 TAHUN)',
    title: 'PERMOHONAN 2: FLEKSIBILITAS KERJASAMA LAHAN BLUD',
    desc: 'Peraturan Walikota skema Long-Term Revenue Sharing yang fleksibel bagi investor teknologi global di kawasan STP.',
    target: 'Target: Menarik 3–5 Investor Baru Tanpa Beban APBD',
    status: '⚡ Harmonisasi bersama BPKAD & Bagian Hukum Setda',
    accent: C_AMBER,
    tint: C_AMBER_LIGHT
  },
  {
    x: 8.88,
    priority: 'KONSORSIUM RISET 5 KAMPUS SOLO RAYA',
    title: 'PERMOHONAN 3: SURAT EDARAN KONSORSIUM RISET DAERAH',
    desc: 'Mendorong perguruan tinggi (UNS, UMS, ISI, Poltek) mengarahkan minimal 30% riset terapan menyelesaikan problem kota di STP.',
    target: 'Target: 5 Kampus Mitra, 30 Klaster Riset Terapan',
    status: '⚡ Penyusunan MoU bersama 5 Rektor Solo Raya',
    accent: C_EMERALD,
    tint: C_EMERALD_LIGHT
  }
];

asks.forEach(p => {
  addPanel(s12, p.x, 1.55, colW12, 2.80, { fill: C_BG, border: C_RULE_DARK, topAccent: p.accent });

  s12.addText(p.priority, {
    x: p.x + 0.16, y: 1.68, w: colW12 - 0.32, h: 0.22,
    fontSize: 7.5, bold: true, color: p.accent, fontFace: FONT_HEAD
  });
  s12.addText(p.title, {
    x: p.x + 0.16, y: 1.94, w: colW12 - 0.32, h: 0.38,
    fontSize: 9.5, bold: true, color: C_INK, fontFace: FONT_HEAD
  });
  s12.addText(p.desc, {
    x: p.x + 0.16, y: 2.36, w: colW12 - 0.32, h: 0.72,
    fontSize: 8.2, color: C_BODY, fontFace: FONT_BODY, lineSpacing: 11
  });

  // Target Box inside Ask Card
  s12.addShape(pptx.shapes.RECTANGLE, {
    x: p.x + 0.16, y: 3.15, w: colW12 - 0.32, h: 0.44,
    fill: { color: C_SURFACE_ALT }, line: { color: C_RULE, width: 0.8 }
  });
  s12.addText(p.target, {
    x: p.x + 0.22, y: 3.18, w: colW12 - 0.44, h: 0.38,
    fontSize: 7.8, bold: true, color: p.accent, fontFace: FONT_HEAD, valign: 'middle'
  });

  s12.addText(p.status, {
    x: p.x + 0.16, y: 3.75, w: colW12 - 0.32, h: 0.45,
    fontSize: 7.5, italic: true, color: C_MUTED, fontFace: FONT_BODY
  });
});

// Bottom 100-Day Action Table
addPanel(s12, 0.8, 4.55, 11.733, 2.30, { fill: C_BG, border: C_RULE_DARK, topAccent: C_COBALT });
s12.addText('RENCANA AKSI 100 HARI PERTAMA (QUICK WINS PASCA-ARAHAN WALIKOTA SURAKARTA):', {
  x: 1.0, y: 4.65, w: 11.333, h: 0.22,
  fontSize: 8.5, bold: true, color: C_COBALT, fontFace: FONT_HEAD
});

const action100Data = [
  [
    { text: 'PERIODE', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT }, align: 'center' } },
    { text: 'AGENDA STRATEGIS', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } },
    { text: 'OUTPUT KONKRET & TARGET KEBERHASILAN', options: { bold: true, color: C_INK, fill: { color: C_SURFACE_ALT } } }
  ],
  [
    { text: 'Hari 1–10', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Penyusunan Draf Instruksi Walikota e-Katalog', options: { color: C_BODY } },
    { text: 'Draf Inkwal Afirmatif siap harmonisasi bersama Bagian Hukum Setda & BPKAD.', options: { color: C_BODY } }
  ],
  [
    { text: 'Hari 11–30', options: { align: 'center', color: C_AMBER, bold: true, fill: { color: C_AMBER_LIGHT } } },
    { text: 'Harmonisasi & Penerbitan Instruksi Walikota', options: { color: C_AMBER, bold: true, fill: { color: C_AMBER_LIGHT } } },
    { text: 'Inkwal ditandatangani Walikota; Sosialisasi ke seluruh OPD prioritas Pemkot.', options: { color: C_AMBER, bold: true, fill: { color: C_AMBER_LIGHT } } }
  ],
  [
    { text: 'Hari 31–50', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Kurasi & Seleksi 5 Produk Inovasi Unggulan', options: { color: C_BODY } },
    { text: 'Katalog 5 Produk Unggulan Hilirisasi STP (Smart Agri IoT, Alkes, Mesin Sortasi).', options: { color: C_BODY } }
  ],
  [
    { text: 'Hari 51–60', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Showcase Inovasi di Depan Walikota & OPD', options: { color: C_BODY } },
    { text: 'Demonstrasi produk di hadapan Walikota & media; komitmen serapan perdana OPD.', options: { color: C_BODY } }
  ],
  [
    { text: 'Hari 61–80', options: { align: 'center', color: C_COBALT, bold: true } },
    { text: 'Inisiasi Konsorsium Riset 5 Perguruan Tinggi', options: { color: C_BODY } },
    { text: 'MoU Pra-Konsorsium bersama Rektor UNS, UMS, ISI, Poltek Solo Raya.', options: { color: C_BODY } }
  ],
  [
    { text: 'Hari 81–100', options: { align: 'center', color: C_EMERALD, bold: true, fill: { color: C_EMERALD_LIGHT } } },
    { text: 'Launching Resmi Konsorsium & Hilirisasi Inovasi', options: { color: C_EMERALD, bold: true, fill: { color: C_EMERALD_LIGHT } } },
    { text: 'Penandatanganan MoU Riset Terapan bersama Kadin/Apindo di STP; kick-off hilirisasi.', options: { color: C_EMERALD, bold: true, fill: { color: C_EMERALD_LIGHT } } }
  ]
];

s12.addTable(action100Data, {
  x: 1.0, y: 4.92, w: 11.333, h: 1.85,
  colW: [1.4, 3.8, 6.133],
  border: { type: 'solid', pt: 0.8, color: C_RULE },
  fill: { color: C_BG },
  fontSize: 8.2, fontFace: FONT_BODY, valign: 'middle',
  rowH: [0.26, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25]
});

s12.addNotes('Menutup paparan ini, Bapak Walikota, manajemen Solo Technopark dan BRIDA tidak hadir meminta tambahan anggaran. Tiga permohonan kami semuanya adalah payung regulasi—bukan rupiah belanja. Instruksi Walikota untuk e-Katalog, Perwal kerjasama lahan, dan Surat Edaran konsorsium riset. Tiga lembar kebijakan ini yang akan menjadi kunci pembuka ekosistem senilai Rp 23 Miliar di 2030. Dengan kepemimpinan dan arahan Bapak, tim kami siap langsung tancap gas mengeksekusi 100 hari pertama. Bersama Bapak, Solo Technopark akan menjadi kebanggaan kota Surakarta dan teladan nasional pengelolaan kawasan inovasi yang mandiri, berdampak, dan berkelanjutan.');

// =========================================================================
// WRITE TO DISK (V8)
// =========================================================================
const outputPptxPathV8 = path.resolve(__dirname, 'PAPARAN_WALIKOTA_STP_V8.pptx');
const outputPptxPathOut = path.resolve(__dirname, '../output_paparan/PAPARAN_WALIKOTA_STP_V8.pptx');

pptx.writeFile({ fileName: outputPptxPathV8 })
  .then(() => {
    console.log(`[SUCCESS] PPTX Master V8 generated at: ${outputPptxPathV8}`);
    fs.copyFileSync(outputPptxPathV8, outputPptxPathOut);
    console.log(`[SUCCESS] Copied to output_paparan: ${outputPptxPathOut}`);
  })
  .catch(err => {
    console.error('[ERROR] Failed generating PPTX V8:', err);
    process.exit(1);
  });
