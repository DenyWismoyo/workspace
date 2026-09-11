import PptxGenJS from 'pptxgenjs';
import fs from 'fs';
import path from 'path';

// ==========================================
// PALET WARNA SWISS ENTERPRISE ULTRA (V9 - 18S POLISHED)
// ==========================================
const C = {
  bg: 'FFFFFF',          // Pure White
  bgSoft: 'F8FAFC',      // Slate 50
  bgTintBlue: 'EFF6FF',  // Blue 50
  bgTintGreen: 'F0FDF4', // Emerald 50
  bgTintAmber: 'FFFBEB', // Amber 50
  bgTintPurple: 'FAF5FF',// Purple 50
  bgTintRed: 'FEF2F2',   // Red 50
  
  border: 'E2E8F0',      // Hairline Slate 200
  borderDark: 'CBD5E1',  // Slate 300
  borderBlue: 'BFDBFE',  // Blue 200
  borderGreen: 'BBF7D0', // Green 200
  borderAmber: 'FDE68A', // Amber 200
  borderRed: 'FECACA',   // Red 200
  
  textMain: '0F172A',    // Ink Slate 900
  textMuted: '475569',   // Slate 600
  textDim: '64748B',     // Slate 500
  textLight: '94A3B8',   // Slate 400
  textWhite: 'FFFFFF',   // Pure White
  
  navy: '1E3A8A',        // Executive Navy
  blue: '1E40AF',        // Corporate Royal Blue
  sky: '0284C7',         // Tech Cyan / Sky
  emerald: '059669',     // Impact Emerald Green
  amber: 'D97706',       // Golden Amber
  orange: 'EA580C',      // Orange
  purple: '7C3AED',      // Sovereign Violet
  red: 'DC2626',         // Red Danger / Gap
  
  fontTitle: 'Arial',
  fontBody: 'Calibri'
};

// ==========================================
// KAMUS ICON VEKTOR SVG (LUCIDE / HEROICONS)
// ==========================================
const SVG_PATHS = {
  trendingUp: '<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline>',
  dollarSign: '<path d="M12 1v22M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>',
  users: '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path>',
  globe: '<circle cx="12" cy="12" r="10"></circle><line x1="2" y1="12" x2="22" y2="12"></line><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>',
  cpu: '<rect x="4" y="4" width="16" height="16" rx="2"></rect><rect x="9" y="9" width="6" height="6"></rect><path d="M9 1v3M15 1v3M9 20v3M15 20v3M20 9h3M20 14h3M1 9h3M1 14h3"></path>',
  shieldCheck: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path><path d="m9 12 2 2 4-4"></path>',
  award: '<circle cx="12" cy="8" r="7"></circle><polyline points="8.21 13.89 7 23 12 20 17 23 15.79 13.88"></polyline>',
  zap: '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon>',
  target: '<circle cx="12" cy="12" r="10"></circle><circle cx="12" cy="12" r="6"></circle><circle cx="12" cy="12" r="2"></circle>',
  checkCircle: '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline>',
  alertTriangle: '<polygon points="12 2 1 21 23 21 12 2"></polygon><line x1="12" y1="9" x2="12" y2="13"></line><circle cx="12" cy="17" r="1"></circle>',
  briefcase: '<rect x="2" y="7" width="20" height="14" rx="2" ry="2"></rect><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"></path>',
  building: '<rect x="4" y="2" width="16" height="20" rx="2" ry="2"></rect><line x1="9" y1="22" x2="9" y2="22.01"></line><line x1="15" y1="22" x2="15" y2="22.01"></line><line x1="9" y1="17" x2="9" y2="17.01"></line><line x1="15" y1="17" x2="15" y2="17.01"></line><line x1="9" y1="12" x2="9" y2="12.01"></line><line x1="15" y1="12" x2="15" y2="12.01"></line><line x1="9" y1="7" x2="9" y2="7.01"></line><line x1="15" y1="7" x2="15" y2="7.01"></line>',
  flask: '<path d="M10 2v7.31M14 9.3V1.99M8.5 2h7M14 9.3a6.5 6.5 0 1 1-4 0"></path>',
  fileText: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line>',
  layers: '<polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline>',
  wrench: '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path>',
  graduationCap: '<path d="M22 10v6M2 10l10-5 10 5-10 5z"></path><path d="M6 12v5c3 3 9 3 12 0v-5"></path>',
  shoppingCart: '<circle cx="8" cy="21" r="1"></circle><circle cx="19" cy="21" r="1"></circle><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"></path>',
  scale: '<line x1="12" y1="3" x2="12" y2="21"></line><path d="M3 7l9-4 9 4"></path><path d="M5 21h14"></path><path d="M5 10l-2 5h4z"></path><path d="M19 10l-2 5h4z"></path>',
  rocket: '<path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.84.7-2.13-.09-2.91a2.18 2.18 0 0 0-2.91-.09z"></path><path d="m12 15-3-3a22 22 0 0 1 2-3.95A12.88 12.88 0 0 1 22 2c0 2.72-.78 7.5-6 11a22.35 22.35 0 0 1-4 2z"></path>',
  sparkles: '<path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"></path>',
  mapPin: '<path d="M20 10c0 6-8 12-8 12s-8-6-8-12a8 8 0 0 1 16 0Z"></path><circle cx="12" cy="10" r="3"></circle>',
  anchor: '<circle cx="12" cy="5" r="3"></circle><line x1="12" y1="22" x2="12" y2="8"></line><path d="M5 12H2a10 10 0 0 0 20 0h-3"></path>',
  pieChart: '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"></path><path d="M22 12A10 10 0 0 0 12 2v10z"></path>',
  lock: '<rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect><path d="M7 11V7a5 5 0 0 1 10 0v4"></path>',
  presentation: '<path d="M2 3h20"></path><path d="M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"></path><path d="m7 21 5-5 5 5"></path>',
  sprout: '<path d="M7 20h10"></path><path d="M10 20c0-4.4 3.6-8 8-8"></path><path d="M14 20c0-3.3-2.7-6-6-6"></path><path d="M14 8a4 4 0 0 0-8 0c0 3 3 6 8 6"></path><path d="M17 4a4 4 0 0 1 4 4c0 3-3 6-8 6"></path>',
  microchip: '<rect x="5" y="5" width="14" height="14" rx="2"></rect><rect x="9" y="9" width="6" height="6"></rect><path d="M9 2v3M15 2v3M9 19v3M15 19v3M2 9h3M2 15h3M19 9h3M19 15h3"></path>',
  arrowRight: '<line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline>',
  chevronRight: '<polyline points="9 18 15 12 9 6"></polyline>'
};

function getSvgBase64(iconName, color = C.blue, strokeWidth = 2) {
  const inner = SVG_PATHS[iconName] || SVG_PATHS.checkCircle;
  const hexColor = '#' + color.replace('#', '');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96" fill="none" stroke="${hexColor}" stroke-width="${strokeWidth}" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;
  return Buffer.from(svg).toString('base64');
}

// Inisialisasi Presentasi 18 Slide
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';
pptx.author = 'Pemerintah Kota Surakarta - BRIDA & UPTD KST Solo Technopark';
pptx.title = 'Roadmap Transformasi Solo Technopark 2026-2030 (Swiss Enterprise Ultra 18S Polished)';

// ==========================================
// HELPER UI FUNCTIONS (18 SLIDES)
// ==========================================

function addHeader(slide, { kicker, title, subtitle, pageNum }) {
  // Kicker
  slide.addText(kicker.toUpperCase(), {
    x: 0.8, y: 0.36, w: 10.0, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.blue,
    charSpacing: 1.5
  });
  
  // Title
  slide.addText(title, {
    x: 0.8, y: 0.62, w: 10.5, h: 0.44,
    fontFace: C.fontTitle, fontSize: 18.5, bold: true, color: C.textMain
  });
  
  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 1.10, w: 10.5, h: 0.26,
    fontFace: C.fontBody, fontSize: 10, italic: true, color: C.textDim
  });
  
  // Page Counter (18 Total)
  if (pageNum) {
    slide.addText(`${pageNum.toString().padStart(2, '0')}  /  18`, {
      x: 11.5, y: 0.36, w: 1.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textDim, align: 'right'
    });
  }
  
  // Top hairline border
  slide.addShape(pptx.shapes.LINE, {
    x: 0.8, y: 1.45, w: 11.733, h: 0,
    line: { color: C.border, width: 1 }
  });
}

function addFooter(slide) {
  slide.addShape(pptx.shapes.LINE, {
    x: 0.8, y: 6.95, w: 11.733, h: 0,
    line: { color: C.border, width: 1 }
  });
  
  slide.addText('PEMERINTAH KOTA SURAKARTA  •  BADAN RISET DAN INOVASI DAERAH (BRIDA)  •  ROADMAP 2026–2030', {
    x: 0.8, y: 7.02, w: 7.5, h: 0.25,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.textDim, charSpacing: 1
  });
  
  slide.addText('EDISI ZEN ENTERPRISE ULTRA (18 SLIDES)', {
    x: 8.5, y: 7.02, w: 4.033, h: 0.25,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue, align: 'right', charSpacing: 1
  });
}

function addSvgIcon(slide, { iconName, x, y, size = 0.36, color = C.blue, bgCircle = true, bgColor = C.bgTintBlue }) {
  if (bgCircle) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y, w: size + 0.16, h: size + 0.16,
      fill: { color: bgColor },
      line: { color: C.border, width: 0.75 }
    });
    slide.addImage({
      data: `image/svg+xml;base64,${getSvgBase64(iconName, color)}`,
      x: x + 0.08, y: y + 0.08, w: size, h: size
    });
  } else {
    slide.addImage({
      data: `image/svg+xml;base64,${getSvgBase64(iconName, color)}`,
      x, y, w: size, h: size
    });
  }
}

console.log('Building 18-Slide Polished Executive Master Deck...');

// ----------------------------------------------------
// SLIDE 01: COVER MASTHEAD & 4 KEY METRICS (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  
  // Top Corporate Accent Bar
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.333, h: 0.2,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  
  // Hero Badge
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 0.65, w: 5.4, h: 0.36,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('ROADMAP STRATEGIS PEMERINTAH KOTA SURAKARTA 2026–2030', {
    x: 0.9, y: 0.70, w: 5.2, h: 0.26,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.blue, charSpacing: 1.2
  });
  
  // Main Title
  s.addText('TRANSFORMASI SOLO TECHNOPARK:\nKEMANDIRIAN FISKAL & PUSAT INOVASI', {
    x: 0.8, y: 1.15, w: 11.5, h: 1.45,
    fontFace: C.fontTitle, fontSize: 30, bold: true, color: C.textMain, lineSpacingMultiple: 1.15
  });
  
  // Subtitle
  s.addText('Rencana Strategis UPTD KST Solo Technopark Menuju Kemandirian Penuh BLUD 2029 Tanpa Subsidi APBD,\nHilirisasi Riset Terapan Konsorsium Perguruan Tinggi, dan Penciptaan Multiplier Effect Bagi Warga Surakarta', {
    x: 0.8, y: 2.70, w: 11.5, h: 0.75,
    fontFace: C.fontBody, fontSize: 13, color: C.textMuted, lineSpacingMultiple: 1.25
  });
  
  // Divider
  s.addShape(pptx.shapes.LINE, {
    x: 0.8, y: 3.65, w: 11.733, h: 0,
    line: { color: C.border, width: 1.5 }
  });
  
  // 4 KPI Summary Cards with Bottom Badges (Lebar 2.78" masing-masing)
  const kpis = [
    { num: 'Rp 23,47 M', label: 'TARGET PENDAPATAN 2030', sub: 'Lompatan +212,8% dari baseline 2026', tag: 'MANDIRI PENUH +212%', color: C.navy, icon: 'trendingUp', bg: C.bgTintBlue },
    { num: '2029', label: 'TARGET BREAK-EVEN POINT (BEP)', sub: '100% Mandiri operasional non-APBD', tag: 'NOL SUBSIDI APBD', color: C.emerald, icon: 'shieldCheck', bg: C.bgTintGreen },
    { num: '95+ Mitra', label: 'EKOSISTEM KEMITRAAN GLOBAL', sub: 'Shopee, Garena, Indosat, ACER, 5 Kampus', tag: 'SKEMA HEXAHELIX', color: C.blue, icon: 'globe', bg: C.bgTintBlue },
    { num: 'Rp 18,5 M', label: 'DAMPAK EKONOMI DAERAH', sub: 'Omzet tenant lokal & 1.500 naker/thn', tag: '1.500 NAKER / THN', color: C.purple, icon: 'award', bg: C.bgTintPurple }
  ];
  
  kpis.forEach((k, i) => {
    const cardX = 0.8 + (i * 2.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 3.90, w: 2.78, h: 2.45,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 3.90, w: 2.78, h: 0.08,
      fill: { color: k.color }, line: { color: k.color }
    });
    
    addSvgIcon(s, { iconName: k.icon, x: cardX + 0.2, y: 4.10, size: 0.36, color: k.color, bgColor: k.bg });
    
    s.addText(k.num, {
      x: cardX + 0.2, y: 4.60, w: 2.38, h: 0.45,
      fontFace: C.fontTitle, fontSize: 23, bold: true, color: k.color
    });
    s.addText(k.label, {
      x: cardX + 0.2, y: 5.08, w: 2.38, h: 0.28,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.textMain, charSpacing: 0.5
    });
    s.addText(k.sub, {
      x: cardX + 0.2, y: 5.38, w: 2.38, h: 0.45,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
    });
    
    // Bottom Pill Tag inside card
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.2, y: 5.92, w: 2.38, h: 0.28,
      fill: { color: k.bg }, line: { color: k.color, width: 0.75 }
    });
    s.addText(k.tag, {
      x: cardX + 0.2, y: 5.94, w: 2.38, h: 0.24,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: k.color, align: 'center', charSpacing: 0.8
    });
  });
  
  // Executive Presenter Metadata Strip
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.50, w: 11.733, h: 0.34,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 0.75 }
  });
  s.addText('PENYAJI: Kepala UPTD STP & BRIDA Kota Surakarta   •   AUDIENS: Walikota Surakarta   •   DOKUMEN: Paparan Eksekutif Master Plan 2026–2030', {
    x: 0.8, y: 6.53, w: 11.733, h: 0.28,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.textMuted, align: 'center', charSpacing: 0.8
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 02: KEDUDUKAN HUKUM KELEMBAGAAN UPTD DI BAWAH BRIDA (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'ARSITEKTUR TATA KELOLA PEMERINTAHAN // REGULASI KELEMBAGAAN',
    title: 'KEDUDUKAN HUKUM UPTD KST SOLO TECHNOPARK DI BAWAH BRIDA',
    subtitle: 'Landasan Struktural Berdasarkan Peraturan Walikota Surakarta No. 15 Tahun 2022',
    pageNum: 2
  });
  
  const pillars = [
    {
      pill: 'REGULATOR DAERAH',
      title: 'BRIDA KOTA SURAKARTA',
      tag: 'INDUK KEBIJAKAN & PENGARAH STRATEGIS',
      color: C.navy, bg: C.bgTintBlue, icon: 'building',
      desc: 'Lembaga teknis daerah yang memayungi arah riset, inovasi, dan hilirisasi teknologi sesuai visi Walikota.',
      points: [
        { b: 'Fungsi Regulasi:', t: 'Menetapkan pedoman prioritas riset terapan pemecah masalah perkotaan.' },
        { b: 'Pengarah Program:', t: 'Memastikan seluruh aset dan fasilitas STP selaras dengan RPJMD Surakarta.' },
        { b: 'Evaluasi Kinerja:', t: 'Menilai akuntabilitas ilmiah, dampak sosio-ekonomi, dan tata kelola inovasi.' },
        { b: 'Dasar Hukum:', t: 'Perda No. 10/2021 & Perwali Surakarta No. 14/2022 tentang Susunan Organisasi BRIDA.' }
      ]
    },
    {
      pill: 'OPERATOR KAWASAN',
      title: 'UPTD KST SOLO TECHNOPARK',
      tag: 'OPERATOR BISNIS & TEKNOLOGI BLUD',
      color: C.blue, bg: C.bgTintBlue, icon: 'cpu',
      desc: 'Unit Pelaksana Teknis Daerah yang mengoperasikan ekosistem inovasi, inkubasi, dan kemitraan industri.',
      points: [
        { b: 'Status BLUD:', t: 'Menerapkan Pola Pengelolaan Keuangan BLUD Penuh (Permendagri 79/2018).' },
        { b: 'Operator Kawasan:', t: 'Mengelola 9 gedung tematik, 5 laboratorium canggih, dan 13 hektar kawasan STP.' },
        { b: 'Kemitraan Industri:', t: 'Menjalin kontrak komersial business-to-business (B2B) secara lincah dan mandiri.' },
        { b: 'Akselerasi Startup:', t: 'Menginkubasi 20+ wirausaha baru berbasis teknologi setiap tahun.' }
      ]
    },
    {
      pill: 'SINERGI BLUD',
      title: 'SINERGI TATA KELOLA BLUD',
      tag: 'KEUNGGULAN OPERASIONAL & EFISIENSI',
      color: C.emerald, bg: C.bgTintGreen, icon: 'scale',
      desc: 'Harmonisasi kepatuhan birokrasi daerah dengan kecepatan eksekusi standar industri kelas dunia.',
      points: [
        { b: 'Rekrutmen Fleksibel:', t: 'Berhak merekrut tenaga profesional non-PNS dengan standar remunerasi pasar.' },
        { b: 'Pengadaan Cepat:', t: 'Mekanisme pengadaan barang/jasa langsung tanpa birokrasi lelang APBD berbelit.' },
        { b: 'Reinvestasi SiLPA:', t: 'Pendapatan jasa layanan dapat langsung digunakan kembali untuk belanja operasional.' },
        { b: 'Akuntabilitas Ganda:', t: 'Diaudit rutin oleh Kantor Akuntan Publik (KAP) independen, Inspektorat, dan BPK.' }
      ]
    }
  ];
  
  pillars.forEach((p, i) => {
    const cardX = 0.8 + (i * 3.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 0.08,
      fill: { color: p.color }, line: { color: p.color }
    });
    
    // Top Micro Pill
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 1.82, w: 1.7, h: 0.24,
      fill: { color: p.bg }, line: { color: p.color, width: 0.75 }
    });
    s.addText(p.pill, {
      x: cardX + 0.25, y: 1.84, w: 1.7, h: 0.20,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: p.color, align: 'center', charSpacing: 0.8
    });
    
    addSvgIcon(s, { iconName: p.icon, x: cardX + 0.25, y: 2.15, size: 0.36, color: p.color, bgColor: p.bg });
    
    s.addText(p.title, {
      x: cardX + 0.75, y: 2.15, w: 2.85, h: 0.28,
      fontFace: C.fontTitle, fontSize: 12.5, bold: true, color: C.textMain
    });
    s.addText(p.tag, {
      x: cardX + 0.75, y: 2.42, w: 2.85, h: 0.20,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: p.color, charSpacing: 0.5
    });
    
    s.addText(p.desc, {
      x: cardX + 0.25, y: 2.70, w: 3.28, h: 0.58,
      fontFace: C.fontBody, fontSize: 10, italic: true, color: C.textMuted, lineSpacingMultiple: 1.18
    });
    
    s.addShape(pptx.shapes.LINE, {
      x: cardX + 0.25, y: 3.32, w: 3.28, h: 0,
      line: { color: C.border, width: 1 }
    });
    
    p.points.forEach((pt, idx) => {
      const py = 3.42 + (idx * 0.64);
      s.addShape(pptx.shapes.OVAL, {
        x: cardX + 0.25, y: py + 0.06, w: 0.09, h: 0.09,
        fill: { color: p.color }, line: { color: p.color }
      });
      s.addText([
        { text: pt.b + ' ', options: { bold: true, color: C.textMain } },
        { text: pt.t, options: { color: C.textMuted } }
      ], {
        x: cardX + 0.42, y: py, w: 3.1, h: 0.60,
        fontFace: C.fontBody, fontSize: 10.5, lineSpacingMultiple: 1.18
      });
    });
  });
  
  // High-Contrast Solid Navy Bottom Banner
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('KOMITMEN TATA KELOLA: Kedudukan UPTD di bawah BRIDA menjamin keselarasan ilmiah dengan program prioritas Kota Surakarta, sementara status BLUD memberikan kelincahan komersial penuh untuk bermitra cepat dengan industri global.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 03: CHAIN OF IMPACT TRIPARTIT (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'LOGIKA PENGEMBANGAN DAERAH // CHAIN OF IMPACT TRIPARTIT',
    title: 'RANTAI DAMPAK SOLO TECHNOPARK UNTUK WARGA SURAKARTA',
    subtitle: 'Menjembatani Hasil Riset Laboratorium Menjadi Nilai Ekonomi Riil dan Lapangan Kerja Daerah',
    pageNum: 3
  });
  
  const stages = [
    {
      step: 'LEVEL 1 // DIRECT OUTPUT',
      title: 'LAYANAN LANGSUNG STP',
      color: C.navy, bg: C.bgTintBlue, icon: 'wrench',
      hero: 'Fasilitas & Layanan Riil',
      points: [
        'Pengoperasian 9 gedung, 5 lab canggih, dan 13 hektar kawasan.',
        '1.500 talenta muda Surakarta dilatih & disertifikasi industri per tahun.',
        '20 startup teknologi baru diinkubasi & difasilitasi akses pendanaan.',
        '15 inovasi riset kampus diuji coba dan dibuatkan prototipe layak pasar.',
        'Kolaborasi aktif dengan Shopee, Garena, Indosat, ACER, dan 5 Kampus.'
      ]
    },
    {
      step: 'LEVEL 2 // INTERMEDIARY OUTCOME',
      title: 'DAMPAK EKOSISTEM & INDUSTRI',
      color: C.blue, bg: C.bgTintBlue, icon: 'rocket',
      hero: 'Serapan Kerja & Investasi',
      points: [
        'Tingkat serapan kerja lulusan vokasi mencapai 85%+ dalam 6 bulan.',
        'Masuknya investasi korporasi teknologi global ke Solo (Zero APBD Burden).',
        'Peningkatan omzet tenant inovasi hingga mencapai Rp 18,5 Miliar di 2030.',
        'Peningkatan pendapatan mandiri BLUD menuju target Rp 23,47 Miliar.',
        'Terbentuknya jejaring rantai pasok industri manufaktur dan digital lokal.'
      ]
    },
    {
      step: 'LEVEL 3 // SYSTEMIC IMPACT',
      title: 'KESEJAHTERAAN WARGA KOTA',
      color: C.emerald, bg: C.bgTintGreen, icon: 'award',
      hero: 'Multi-Efek Kota Surakarta',
      points: [
        'Diversifikasi ekonomi Solo dari sektor pariwisata-tekstil ke ekonomi inovasi.',
        'Penurunan angka pengangguran usia produktif dan peningkatan upah layak.',
        'Peningkatan Pendapatan Asli Daerah (PAD) melalui retribusi dan pajak daerah.',
        'Solo diakui sebagai Hub Inovasi & Riset Terapan Terdepan di Jawa Tengah.',
        'Kemandirian fiskal Pemkot Surakarta: APBD dapat dialihkan untuk bansos.'
      ]
    }
  ];
  
  stages.forEach((st, i) => {
    const cardX = 0.8 + (i * 3.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 0.08,
      fill: { color: st.color }, line: { color: st.color }
    });
    
    // Level Badge Pill
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 1.82, w: 3.28, h: 0.26,
      fill: { color: st.bg }, line: { color: st.color, width: 0.75 }
    });
    s.addText(st.step, {
      x: cardX + 0.25, y: 1.84, w: 3.28, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: st.color, align: 'center', charSpacing: 1
    });
    
    addSvgIcon(s, { iconName: st.icon, x: cardX + 0.25, y: 2.18, size: 0.36, color: st.color, bgColor: st.bg });
    
    s.addText(st.title, {
      x: cardX + 0.75, y: 2.20, w: 2.85, h: 0.32,
      fontFace: C.fontTitle, fontSize: 12.5, bold: true, color: C.textMain
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 2.65, w: 3.28, h: 0.42,
      fill: { color: st.bg }, line: { color: st.color, width: 0.75 }
    });
    s.addText(st.hero, {
      x: cardX + 0.35, y: 2.71, w: 3.08, h: 0.3,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: st.color, align: 'center'
    });
    
    st.points.forEach((pt, idx) => {
      const py = 3.25 + (idx * 0.55);
      s.addShape(pptx.shapes.OVAL, {
        x: cardX + 0.25, y: py + 0.06, w: 0.09, h: 0.09,
        fill: { color: st.color }, line: { color: st.color }
      });
      s.addText(pt, {
        x: cardX + 0.42, y: py, w: 3.1, h: 0.50,
        fontFace: C.fontBody, fontSize: 10.5, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
    
    // Chevron Transition Flow Indicator between Column 1-2 and 2-3
    if (i < 2) {
      s.addText('➔', {
        x: cardX + 3.78 + 0.03, y: 3.5, w: 0.2, h: 0.4,
        fontFace: C.fontTitle, fontSize: 14, bold: true, color: C.blue, align: 'center'
      });
    }
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('PRINSIP KEBERDAMPAKAN: Setiap rupiah belanja operasional dan pemanfaatan aset Solo Technopark wajib bermuara pada kesejahteraan nyata, daya saing talenta muda, dan pembukaan lapangan kerja berkualitas bagi warga Surakarta.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 04: LANDASAN YURIDIS & 4 FLEKSIBILITAS BLUD (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'KEPASTIAN HUKUM // REGULASI & KEPATUHAN',
    title: 'LANDASAN YURIDIS KOKOH & 4 FLEKSIBILITAS PENGELOLAAN BLUD',
    subtitle: 'Menjamin Kepatuhan Tata Kelola Pemerintahan Sekaligus Kelincahan Kemitraan Komersial',
    pageNum: 4
  });
  
  // Left: Hierarchy of Regulation
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.65, w: 4.5, h: 4.4,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.65, w: 4.5, h: 0.08,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('HIRARKI REGULASI PENGUAT BLUD', {
    x: 1.05, y: 1.85, w: 4.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.navy, charSpacing: 1
  });
  
  const regs = [
    { title: 'Perpres No. 106 Tahun 2017', sub: 'Mandat Kawasan Sains & Teknologi Nasional sebagai pusat riset terapan dan hilirisasi inovasi daerah.' },
    { title: 'Permendagri No. 79 Tahun 2018', sub: 'Pedoman Teknis BLUD: Memberikan fleksibilitas penuh pengelolaan keuangan, pendapatan, belanja, dan tarif.' },
    { title: 'Peraturan Walikota Surakarta No. 15/2022', sub: 'Kedudukan, Susunan Organisasi, Tugas, Fungsi, dan Tata Kerja UPTD KST Solo Technopark pada BRIDA.' },
    { title: 'Peraturan Walikota Surakarta No. 38/2022', sub: 'Penetapan Pola Pengelolaan Keuangan BLUD Penuh pada UPTD KST Solo Technopark.' }
  ];
  
  regs.forEach((r, i) => {
    const ry = 2.22 + (i * 0.92);
    addSvgIcon(s, { iconName: 'fileText', x: 1.05, y: ry + 0.04, size: 0.28, color: C.blue, bgColor: C.bgTintBlue });
    s.addText(r.title, {
      x: 1.5, y: ry, w: 3.6, h: 0.25,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.textMain
    });
    s.addText(r.sub, {
      x: 1.5, y: ry + 0.26, w: 3.6, h: 0.60,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  // Right: 4 Structured Flexibility Cards (Balanced to exact 0.8" right margin)
  const flexs = [
    {
      title: 'FLEKSIBILITAS TARIF LAYANAN',
      sub: 'Penetapan Tarif Berdasarkan Nilai Pasar',
      color: C.blue, icon: 'dollarSign', bg: C.bgTintBlue,
      p1: 'Mekanisme Tarif: Ditetapkan langsung via SK Kepala UPTD/Perwali tanpa perlu pembahasan APBD tahunan.',
      p2: 'Keunggulan Bisnis: Tarif kompetitif mengikuti dinamika pasar untuk 45+ jenis uji lab & sewa fasilitas.'
    },
    {
      title: 'FLEKSIBILITAS BELANJA & SiLPA',
      sub: 'Reinvestasi Pendapatan Mandiri',
      color: C.emerald, icon: 'trendingUp', bg: C.bgTintGreen,
      p1: 'Rekening Otonom: Seluruh penerimaan disetor ke kas BLUD dan dapat langsung dibelanjakan mandiri.',
      p2: 'Efisiensi Operasional: Reinvestasi instan untuk pemeliharaan lab presisi dan bahan uji tanpa jeda birokrasi.'
    },
    {
      title: 'KERJASAMA LAHAN & ASET',
      sub: 'Kemitraan Jangka Panjang (10-25 Tahun)',
      color: C.amber, icon: 'building', bg: C.bgTintAmber,
      p1: 'Skema Fleksibel: Monetisasi lahan tidur melalui skema Bangun Guna Serah (BGS) & Long-Term Revenue Sharing.',
      p2: 'Zero APBD: Pembangunan gedung riset dan fasilitas baru 100% didanai mitra korporasi swasta global.'
    },
    {
      title: 'PENGADAAN & SDM PROFESIONAL',
      sub: 'Kelincahan Standar Industri Swasta',
      color: C.purple, icon: 'users', bg: C.bgTintPurple,
      p1: 'SDM Non-PNS: Hak merekrut tenaga ahli dan praktisi industri global dengan standar remunerasi pasar.',
      p2: 'Pengadaan Cepat: Mekanisme pengadaan barang/jasa teknologi langsung tanpa hambatan lelang birokratis.'
    }
  ];
  
  flexs.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const fx = 5.533 + (col * 3.60);
    const fy = 1.65 + (row * 2.25);
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: fx, y: fy, w: 3.4, h: 2.15,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: fx, y: fy, w: 3.4, h: 0.08,
      fill: { color: f.color }, line: { color: f.color }
    });
    
    addSvgIcon(s, { iconName: f.icon, x: fx + 0.2, y: fy + 0.18, size: 0.32, color: f.color, bgColor: f.bg });
    
    s.addText(f.title, {
      x: fx + 0.65, y: fy + 0.16, w: 2.6, h: 0.24,
      fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.textMain
    });
    s.addText(f.sub, {
      x: fx + 0.65, y: fy + 0.40, w: 2.6, h: 0.20,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: f.color
    });
    
    s.addShape(pptx.shapes.LINE, {
      x: fx + 0.2, y: fy + 0.68, w: 3.0, h: 0,
      line: { color: C.border, width: 0.75 }
    });
    
    // Structured 2 points per card
    s.addShape(pptx.shapes.OVAL, {
      x: fx + 0.2, y: fy + 0.78, w: 0.08, h: 0.08,
      fill: { color: f.color }, line: { color: f.color }
    });
    s.addText(f.p1, {
      x: fx + 0.35, y: fy + 0.74, w: 2.85, h: 0.58,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
    });
    
    s.addShape(pptx.shapes.OVAL, {
      x: fx + 0.2, y: fy + 1.40, w: 0.08, h: 0.08,
      fill: { color: f.color }, line: { color: f.color }
    });
    s.addText(f.p2, {
      x: fx + 0.35, y: fy + 1.36, w: 2.85, h: 0.58,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('KESIMPULAN YURIDIS: Payung hukum BLUD Solo Technopark telah sah dan berkekuatan hukum penuh. Tantangan utama saat ini bukan pada ketiadaan izin, melainkan pada eksekusi regulasi turunan pemanfaatan aset dan kemitraan lahan.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.navy, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 05: EMPAT PILAR LAYANAN TERPADU KST (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'FUNGSI STRATEGIS KAWASAN SAINS & TEKNOLOGI // PERPRES 106/2017',
    title: 'EMPAT PILAR LAYANAN TERPADU SOLO TECHNOPARK',
    subtitle: 'Mandat Nasional Pengembangan KST sebagai Mesin Hilirisasi dan Pertumbuhan Ekonomi Daerah',
    pageNum: 5
  });
  
  const pillars = [
    {
      title: 'INKUBASI BISNIS TEKNOLOGI',
      sub: 'Mencetak Wirausaha Berbasis Inovasi',
      icon: 'rocket', color: C.navy, bg: C.bgTintBlue,
      kpi: 'TARGET: 20+ STARTUP DIGITAL & MANUFAKTUR / THN',
      items: [
        'Akselerasi 20+ tenant startup digital & manufaktur per tahun.',
        'Mentoring intensif oleh praktisi industri global (Shopee, Garena).',
        'Fasilitasi legalitas PT/CV, HKI/Paten, dan sertifikasi SNI/Halal.',
        'Akses pendanaan investor melalui Demo Day & Business Matching.'
      ]
    },
    {
      title: 'PELATIHAN VOKASI & SERTIFIKASI',
      sub: 'Upskilling Talenta Masa Depan',
      icon: 'graduationCap', color: C.blue, bg: C.bgTintBlue,
      kpi: 'TARGET: 1.500 TALENTA TERSERTIFIKASI / THN',
      items: [
        'Diklat vokasi industri: Otomasi, Pengelasan Underwater, & AI.',
        'Sertifikasi standar BNSP & internasional (AWS, Cisco, TUV).',
        'Target 1.500 talenta muda tersertifikasi per tahun.',
        'Tingkat keterserapan kerja lulusan ke industri mencapai 85%+.'
      ]
    },
    {
      title: 'RISET TERAPAN & INOVASI',
      sub: 'Menjembatani Kampus ke Industri',
      icon: 'flask', color: C.emerald, bg: C.bgTintGreen,
      kpi: 'TARGET: 30+ KLASTER RISET INDUSTRI & KAMPUS',
      items: [
        'Konsorsium riset terapan bersama 5 kampus ternama Solo Raya.',
        'Fasilitasi uji laboratorium, reverse engineering, & CAD/CAM.',
        'Fasilitas Fablab, Precision Machining, dan Biotech Lab.',
        'Fokus riset terapan: Alkes, Smart Agri IoT, & Mesin Pengolah.'
      ]
    },
    {
      title: 'HILIRISASI & KOMERSIALISASI',
      sub: 'Membawa Prototipe ke Pasar Nyata',
      icon: 'shoppingCart', color: C.purple, bg: C.bgTintPurple,
      kpi: 'TARGET: 5+ PRODUK INOVASI TEMBUS e-KATALOG / THN',
      items: [
        'Pabrikasi skala terbatas (pilot production) produk inovasi.',
        'Onboarding produk inovasi ke etalase e-Katalog Lokal Pemkot.',
        'Kerjasama lisensi teknologi dan royalti paten bersama industri.',
        'Membuka captive market belanja inovasi seluruh OPD Surakarta.'
      ]
    }
  ];
  
  pillars.forEach((p, i) => {
    const cardX = 0.8 + (i * 2.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 2.78, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 2.78, h: 0.08,
      fill: { color: p.color }, line: { color: p.color }
    });
    
    addSvgIcon(s, { iconName: p.icon, x: cardX + 0.2, y: 1.88, size: 0.36, color: p.color, bgColor: p.bg });
    
    s.addText(p.title, {
      x: cardX + 0.2, y: 2.36, w: 2.38, h: 0.42,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    s.addText(p.sub, {
      x: cardX + 0.2, y: 2.80, w: 2.38, h: 0.25,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: p.color
    });
    
    s.addShape(pptx.shapes.LINE, {
      x: cardX + 0.2, y: 3.15, w: 2.38, h: 0,
      line: { color: C.border, width: 1 }
    });
    
    p.items.forEach((it, idx) => {
      const iy = 3.28 + (idx * 0.52);
      s.addShape(pptx.shapes.OVAL, {
        x: cardX + 0.2, y: iy + 0.06, w: 0.08, h: 0.08,
        fill: { color: p.color }, line: { color: p.color }
      });
      s.addText(it, {
        x: cardX + 0.36, y: iy, w: 2.22, h: 0.48,
        fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
    
    // Bottom Anchor KPI Box
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.15, y: 5.52, w: 2.48, h: 0.38,
      fill: { color: p.bg }, line: { color: p.color, width: 1 }
    });
    s.addText(p.kpi, {
      x: cardX + 0.18, y: 5.54, w: 2.42, h: 0.34,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: p.color, align: 'center', charSpacing: 0.5
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('KONSISTENSI 4 PILAR: Solo Technopark bukan sekadar gedung pertemuan atau kantor bersama, melainkan ekosistem terpadu yang memadukan talenta, riset kampus, fasilitas canggih, dan akses pasar industri.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 06: PIPELINE 5 TAHAPAN HILIRISASI RISET (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'ALUR KOMERSIALISASI INOVASI // DARI LABORATORIUM KE PASAR',
    title: 'PIPELINE 5 TAHAPAN HILIRISASI RISET SOLO TECHNOPARK',
    subtitle: 'Mengatasi Valley of Death Inovasi melalui Fasilitasi Terpadu Kesiapan Teknologi TRL 1 hingga TRL 9',
    pageNum: 6
  });
  
  const steps = [
    {
      num: '01', title: 'KURASI IDE', trl: 'TRL 1–3',
      color: C.navy, bg: C.bgTintBlue, icon: 'target',
      sub: 'Seleksi Riset Potensial',
      output: 'OUTPUT: DOKUMEN PATEN & HKI RESMI',
      points: [
        'Review proposal riset dosen & mahasiswa.',
        'Validasi kebaruan dan perlindungan HKI/Paten.',
        'Penyelarasan dengan kebutuhan industri lokal.'
      ]
    },
    {
      num: '02', title: 'VALIDASI LAB', trl: 'TRL 4–5',
      color: C.blue, bg: C.bgTintBlue, icon: 'flask',
      sub: 'Uji Fungsi Terbatas',
      output: 'OUTPUT: HASIL UJI LAB & SIMULASI CAD',
      points: [
        'Pengujian komponen di laboratorium canggih STP.',
        'Simulasi performa & optimasi rekayasa mesin.',
        'Verifikasi standar teknis dasar kelayakan.'
      ]
    },
    {
      num: '03', title: 'PROTOTYPING', trl: 'TRL 6–7',
      color: C.sky, bg: C.bgTintBlue, icon: 'layers',
      sub: 'Rekayasa Produk Nyata',
      output: 'OUTPUT: PURWARUPA FISIK TRL 7',
      points: [
        'Fabrikasi purwarupa skala industri di Fablab.',
        'Uji coba operasional pada lingkungan kerja nyata.',
        'Penyempurnaan desain ergonomi & biaya produksi.'
      ]
    },
    {
      num: '04', title: 'UJI PASAR', trl: 'TRL 8',
      color: C.amber, bg: C.bgTintAmber, icon: 'shoppingCart',
      sub: 'Sertifikasi & Pasar Terbatas',
      output: 'OUTPUT: SERTIFIKASI SNI & IZIN EDAR',
      points: [
        'Uji klinis/kelayakan resmi (BBPK, Kemenkes, SNI).',
        'Piloting produk pada 3–5 pengguna percontohan.',
        'Perhitungan kelayakan harga jual dan margin.'
      ]
    },
    {
      num: '05', title: 'KOMERSIALISASI', trl: 'TRL 9',
      color: C.emerald, bg: C.bgTintGreen, icon: 'award',
      sub: 'Produksi Massal & Lisensi',
      output: 'OUTPUT: LISTING e-KATALOG & KONTRAK',
      points: [
        'Penayangan produk pada e-Katalog Lokal Pemkot.',
        'Kerjasama lisensi royalti atau spin-off startup.',
        'Penetrasi pasar swasta dan ekspor regional.'
      ]
    }
  ];
  
  steps.forEach((st, i) => {
    const cardX = 0.8 + (i * 2.38);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 2.22, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 2.22, h: 0.08,
      fill: { color: st.color }, line: { color: st.color }
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.15, y: 1.85, w: 0.65, h: 0.32,
      fill: { color: st.bg }, line: { color: st.color, width: 1 }
    });
    s.addText(st.num, {
      x: cardX + 0.15, y: 1.88, w: 0.65, h: 0.26,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: st.color, align: 'center'
    });
    
    s.addText(st.trl, {
      x: cardX + 0.85, y: 1.9, w: 1.25, h: 0.25,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textDim, align: 'right'
    });
    
    s.addText(st.title, {
      x: cardX + 0.15, y: 2.28, w: 1.95, h: 0.30,
      fontFace: C.fontTitle, fontSize: 11.5, bold: true, color: C.textMain
    });
    s.addText(st.sub, {
      x: cardX + 0.15, y: 2.58, w: 1.95, h: 0.24,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: st.color
    });
    
    s.addShape(pptx.shapes.LINE, {
      x: cardX + 0.15, y: 2.92, w: 1.92, h: 0,
      line: { color: C.border, width: 1 }
    });
    
    st.points.forEach((pt, idx) => {
      const py = 3.08 + (idx * 0.72);
      s.addShape(pptx.shapes.OVAL, {
        x: cardX + 0.15, y: py + 0.06, w: 0.08, h: 0.08,
        fill: { color: st.color }, line: { color: st.color }
      });
      s.addText(pt, {
        x: cardX + 0.28, y: py, w: 1.82, h: 0.65,
        fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
    
    // Bottom Milestone Output Box
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.12, y: 5.48, w: 1.98, h: 0.42,
      fill: { color: st.bg }, line: { color: st.color, width: 1 }
    });
    s.addText(st.output, {
      x: cardX + 0.14, y: 5.50, w: 1.94, h: 0.38,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: st.color, align: 'center', charSpacing: 0.5
    });
    
    // Chevron Transition Flow Indicator between stages
    if (i < 4) {
      s.addText('➔', {
        x: cardX + 2.22 + 0.02, y: 3.5, w: 0.15, h: 0.4,
        fontFace: C.fontTitle, fontSize: 12, bold: true, color: C.blue, align: 'center'
      });
    }
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('MENJAWAB LEMBAH KEMATIAN INOVASI: Dengan pipeline terstruktur ini, hasil riset dosen dan mahasiswa Solo Raya tidak lagi mandek di perpustakaan kampus, melainkan secara bertahap dibimbing hingga terbit izin edar dan masuk ke pasar komersial.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.navy, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 07: SHOWCASE 5 PRODUK TERUJI & 5 FASILITAS CANGGIH (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'PORTOFOLIO INOVASI RIIL // BUKTI IMPLEMENTASI',
    title: 'SHOWCASE PRODUK HILIRISASI & FASILITAS CANGGIH KAWASAN',
    subtitle: 'Bukti Nyata Hasil Riset yang Telah Digunakan Industri dan Infrastruktur Laboratorium Pendukung',
    pageNum: 7
  });
  
  // Left: 5 Produk Inovasi Unggulan
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.65, w: 5.7, h: 4.4,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.65, w: 5.7, h: 0.08,
    fill: { color: C.blue }, line: { color: C.blue }
  });
  s.addText('5 PRODUK INOVASI UNGGULAN TERUJI', {
    x: 1.05, y: 1.85, w: 5.2, h: 0.25,
    fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.blue, charSpacing: 1
  });
  
  const prods = [
    { sector: 'AGRO-TECH', name: 'Smart Agri IoT & Monitoring Tanah', mitra: 'UNS & Kelompok Tani', trl: 'TRL 8', desc: 'Sensor kelembaban tanah & cuaca terintegrasi smartphone, hemat air 40%.' },
    { sector: 'ALKES / MEDIS', name: 'Kursi Roda Cerdas Kontrol Sinyal Otak', mitra: 'Poltekkes & RSUD', trl: 'TRL 7', desc: 'Alat bantu disabilitas motorik berat berbasis sensor EEG dan navigasi AI.' },
    { sector: 'IKM PANGAN', name: 'Mesin Sortasi Biji Kopi Presisi Tinggi', mitra: 'Fak. Teknik UMS & IKM', trl: 'TRL 9', desc: 'Mesin sortasi optik otomatis meningkatkan kapasitas sortir IKM kopi 300%.' },
    { sector: 'FARMASI HERBAL', name: 'Ekstrak Herbal Terstandar Antioksidan', mitra: 'Unisri & Industri Jamu', trl: 'TRL 8', desc: 'Formulasi jamu modern berizin BPOM siap ekspor ke kawasan Asia Tenggara.' },
    { sector: 'BIOMEDIS 3D', name: 'Implan Tulang Titan Kustomisasi 3D', mitra: 'Konsorsium Medis Solo', trl: 'TRL 7', desc: 'Implan medis presisi dibuat dengan 3D Metal Printer di Fablab STP.' }
  ];
  
  prods.forEach((pr, idx) => {
    const py = 2.2 + (idx * 0.74);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 1.05, y: py, w: 5.2, h: 0.66,
      fill: { color: C.bg }, line: { color: C.border, width: 0.75 }
    });
    
    // Sector badge
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 1.15, y: py + 0.08, w: 1.1, h: 0.22,
      fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 0.75 }
    });
    s.addText(pr.sector, {
      x: 1.15, y: py + 0.10, w: 1.1, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue, align: 'center'
    });
    
    s.addText(pr.name, {
      x: 2.32, y: py + 0.06, w: 2.7, h: 0.25,
      fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.textMain
    });
    s.addText(pr.trl, {
      x: 5.1, y: py + 0.06, w: 1.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9, bold: true, color: C.emerald, align: 'right'
    });
    s.addText([
      { text: pr.mitra + '  •  ', options: { bold: true, color: C.textMain } },
      { text: pr.desc, options: { color: C.textMuted } }
    ], {
      x: 1.15, y: py + 0.34, w: 4.95, h: 0.28,
      fontFace: C.fontBody, fontSize: 9
    });
  });
  
  // Right: 5 Fasilitas Laboratorium Canggih
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 6.833, y: 1.65, w: 5.7, h: 4.4,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 6.833, y: 1.65, w: 5.7, h: 0.08,
    fill: { color: C.emerald }, line: { color: C.emerald }
  });
  s.addText('5 FASILITAS LABORATORIUM CANGGIH KAWASAN', {
    x: 7.083, y: 1.85, w: 5.2, h: 0.25,
    fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.emerald, charSpacing: 1
  });
  
  const facs = [
    { name: 'Fablab & Studio Desain Digital', util: 'Utilisasi 85%', desc: '3D Metal/Resin Printer, Laser Cutting, dan Workstation CAD/CAM canggih.' },
    { name: 'Precision Machining & CNC Center', util: 'Utilisasi 90%', desc: 'Mesin CNC Milling 5-Axis dan Bubut Presisi untuk cetakan mold & suku cadang alkes.' },
    { name: 'Agro-Biotech & Food Testing Lab', util: 'Utilisasi 75%', desc: 'Laboratorium pengujian nutrisi makanan, sertifikasi halal, dan olahan pangan IKM.' },
    { name: 'Game & Digital Animation Studio', util: 'Utilisasi 80%', desc: 'Didukung Garena & Shopee: Motion capture studio, render farm, dan coding game.' },
    { name: 'Oil & Gas Skill Center (OGSCI)', util: 'Utilisasi 92%', desc: 'Pusat diklat & sertifikasi pengelasan bawah air (underwater welding) internasional.' }
  ];
  
  facs.forEach((fc, idx) => {
    const fy = 2.2 + (idx * 0.74);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 7.083, y: fy, w: 5.2, h: 0.66,
      fill: { color: C.bg }, line: { color: C.border, width: 0.75 }
    });
    s.addText(fc.name, {
      x: 7.233, y: fy + 0.06, w: 3.8, h: 0.25,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.textMain
    });
    s.addText(fc.util, {
      x: 11.133, y: fy + 0.06, w: 1.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9, bold: true, color: C.blue, align: 'right'
    });
    s.addText(fc.desc, {
      x: 7.233, y: fy + 0.32, w: 4.9, h: 0.28,
      fontFace: C.fontBody, fontSize: 9, color: C.textMuted
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintGreen }, line: { color: C.borderGreen, width: 1 }
  });
  s.addText('INFRASTRUKTUR KELAS DUNIA: Seluruh fasilitas laboratorium di atas telah beroperasi penuh dan siap melayani permintaan rekayasa industri nasional, pelatihan vokasi, serta pengujian mutu produk tenant.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.emerald, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 08: EKOSISTEM HEXAHELIX (POLISHED - NO BADGE LINE WRAP)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'KOLABORASI MULTI-STAKEHOLDER // JARINGAN MITRA',
    title: 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS GLOBAL & NASIONAL',
    subtitle: 'Membangun Ekosistem Berkelanjutan Murni Berbasis Kerjasama Tanpa Beban Belanja Modal APBD',
    pageNum: 8
  });
  
  const actors = [
    {
      actor: 'AKADEMISI (KAMPUS)',
      total: '18 Kampus Mitra',
      color: C.navy, bg: C.bgTintBlue, icon: 'graduationCap',
      members: 'UNS, UMS, ISI Surakarta, Unisri, Poltek Solo, ATMI, Undip, UGM, ITB.',
      role: 'Penyuplai riset terapan, dosen pakar, dan talenta mahasiswa tugas akhir.'
    },
    {
      actor: 'BISNIS & INDUSTRI',
      total: '45 Korporasi Mitra',
      color: C.blue, bg: C.bgTintBlue, icon: 'briefcase',
      members: 'Shopee, Garena, Indosat, ACER, Microsoft, Bukalapak, GoTo, Petrokimia.',
      role: 'Offtaker lulusan vokasi, penyedia kurikulum industri, dan investor tenant.'
    },
    {
      actor: 'KOMUNITAS INOVATOR',
      total: '12 Komunitas Aktif',
      color: C.emerald, bg: C.bgTintGreen, icon: 'users',
      members: 'Solo Creative City Network, Komunitas Robotik, Solo Developer, MakerSpace.',
      role: 'Aktivasi event kreatif, hackathon berkala, dan jejaring wirausaha pemula.'
    },
    {
      actor: 'PEMERINTAH & REGULATOR',
      total: '8 Instansi Terkait',
      color: C.amber, bg: C.bgTintAmber, icon: 'building',
      members: 'BRIN, Kemenperin, Kemenkop UKM, Kemendiktisaintek, Bappeda, Dinas Koperasi.',
      role: 'Penyedia regulasi afirmatif, sertifikasi paten, dan standardisasi produk.'
    },
    {
      actor: 'LEMBAGA KEUANGAN',
      total: '7 Bank & Investor',
      color: C.purple, bg: C.bgTintPurple, icon: 'dollarSign',
      members: 'Bank Indonesia Solo, OJK Solo Raya, Bank Jateng, BRI, Mandiri, Angel Investors.',
      role: 'Akses permodalan modal kerja, KUR inovasi, dan literasi transaksi digital.'
    },
    {
      actor: 'MEDIA & AGREGATOR',
      total: '5 Jaringan Media',
      color: C.red, bg: C.bgTintRed, icon: 'globe',
      members: 'Solopos, Radar Solo, Tribun Solo, RRI Surakarta, Sindikasi Media Nasional.',
      role: 'Amplifikasi branding Solo sebagai Kota Inovasi dan promosi produk tenant.'
    }
  ];
  
  actors.forEach((a, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const ax = 0.8 + (col * 3.98);
    const ay = 1.65 + (row * 2.25);
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: ax, y: ay, w: 3.78, h: 2.15,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: ax, y: ay, w: 3.78, h: 0.08,
      fill: { color: a.color }, line: { color: a.color }
    });
    
    addSvgIcon(s, { iconName: a.icon, x: ax + 0.2, y: ay + 0.18, size: 0.32, color: a.color, bgColor: a.bg });
    
    s.addText(a.actor, {
      x: ax + 0.62, y: ay + 0.16, w: 1.6, h: 0.25,
      fontFace: C.fontTitle, fontSize: 9.2, bold: true, color: C.textMain
    });
    
    // Widen badge container to 1.38" to completely eliminate awkward text wrapping
    s.addShape(pptx.shapes.RECTANGLE, {
      x: ax + 2.24, y: ay + 0.15, w: 1.38, h: 0.26,
      fill: { color: a.bg }, line: { color: a.color, width: 0.75 }
    });
    s.addText(a.total, {
      x: ax + 2.24, y: ay + 0.17, w: 1.38, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.2, bold: true, color: a.color, align: 'center'
    });
    
    s.addText(a.role, {
      x: ax + 0.2, y: ay + 0.55, w: 3.38, h: 0.38,
      fontFace: C.fontBody, fontSize: 9.5, italic: true, color: C.textDim, lineSpacingMultiple: 1.15
    });
    
    s.addShape(pptx.shapes.LINE, {
      x: ax + 0.2, y: ay + 1.0, w: 3.38, h: 0,
      line: { color: C.border, width: 1 }
    });
    
    s.addText('Mitra Tergabung:', {
      x: ax + 0.2, y: ay + 1.05, w: 3.38, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.textDim
    });
    s.addText(a.members, {
      x: ax + 0.2, y: ay + 1.25, w: 3.38, h: 0.78,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.2
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('KOLABORASI TANPA APBD: Lebih dari 95 mitra strategis di atas hadir membawa investasi, peralatan laboratorium, dan kurikulum industri secara mandiri tanpa membebani pagu APBD Kota Surakarta sepeser pun.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 09: DIAGNOSA 5 GAP STRUKTURAL (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'DIAGNOSA STRATEGIS DAERAH // PROBLEM STATEMENT',
    title: 'DIAGNOSA 5 GAP STRUKTURAL & BOTTLENECK EKSISTING STP',
    subtitle: 'Identifikasi Hambatan Utama yang Menahan Akselerasi Kemandirian dan Optimalisasi Dampak STP',
    pageNum: 9
  });
  
  const gaps = [
    {
      num: 'GAP 01', title: 'TATA KELOLA SDM & REMUNERASI PROFESIONAL',
      masalah: 'Standar penggajian birokrasi kaku menyulitkan retensi mentor teknologi & manajer bisnis kelas dunia.',
      dampak: 'Turn-over tenaga ahli tinggi; pengoperasian lab canggih (CNC/Biotech) sering terkendala ketiadaan operator ahli.'
    },
    {
      num: 'GAP 02', title: 'KEMANDIRIAN FINANSIAL & SUBSIDI OPERASIONAL',
      masalah: 'Rasio kemandirian finansial saat ini masih di bawah 45%, operasional dasar masih bergantung pada belanja rutin APBD.',
      dampak: 'Kapasitas ekspansi program terbatas pada besaran pagu anggaran daerah tahunan yang terus berfluktuasi.'
    },
    {
      num: 'GAP 03', title: 'UTILISASI ASET LAHAN TIDUR & REGULASI SEWA',
      masalah: 'Belum adanya regulasi fleksibilitas sewa lahan jangka panjang (10-25 tahun) yang memberikan kepastian hukum bagi investor.',
      dampak: 'Lahan strategis seluas ribuan meter persegi belum termonetisasi; investor swasta menunda pembangunan fasilitas fisik.'
    },
    {
      num: 'GAP 04', title: 'HILIRISASI RISET KAMPUS BELUM TERINTEGRASI PASAR',
      masalah: 'Lebih dari 80% riset dosen dan inovasi mahasiswa Solo Raya mandek pada laporan perpustakaan (TRL 3-4).',
      dampak: 'Kebutuhan teknologi industri manufaktur dan alkes di Solo Raya tetap dipasok oleh produk impor luar negeri.'
    },
    {
      num: 'GAP 05', title: 'MARKET ACCESS & SERAPAN PRODUK TENANT',
      masalah: 'Produk teknologi buatan tenant binaan STP belum diprioritaskan dalam sistem belanja pengadaan barang/jasa dinas Pemkot.',
      dampak: 'Startup inovasi kesulitan mendapatkan "First Paying Client" yang krusial untuk membuktikan rekam jejak pasar.'
    }
  ];
  
  gaps.forEach((g, idx) => {
    const gy = 1.65 + (idx * 0.88);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: gy, w: 11.733, h: 0.8,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: gy, w: 0.12, h: 0.8,
      fill: { color: C.red }, line: { color: C.red }
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 1.05, y: gy + 0.12, w: 0.9, h: 0.28,
      fill: { color: C.bgTintRed }, line: { color: C.borderRed, width: 1 }
    });
    s.addText(g.num, {
      x: 1.05, y: gy + 0.14, w: 0.9, h: 0.24,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.red, align: 'center'
    });
    
    s.addText(g.title, {
      x: 2.1, y: gy + 0.12, w: 9.8, h: 0.28,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    
    s.addText([
      { text: 'Akar Masalah: ', options: { bold: true, color: C.textMain } },
      { text: g.masalah + '   ', options: { color: C.textMuted } },
      { text: 'Dampak Nyata: ', options: { bold: true, color: C.red } },
      { text: g.dampak, options: { color: C.textMuted } }
    ], {
      x: 1.05, y: gy + 0.42, w: 11.2, h: 0.35,
      fontFace: C.fontBody, fontSize: 10, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintRed }, line: { color: C.borderRed, width: 1 }
  });
  s.addText('KESIMPULAN AUDIT MASALAH: Hambatan terbesar akselerasi Solo Technopark saat ini murni terletak pada sumbatan regulasi kepegawaian, skema sewa aset jangka panjang, dan ketiadaan perlindungan serapan pasar lokal.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.red, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 10: 5 SOLUSI TEROBOSAN STRATEGIS (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'STRATEGI TEROBOSAN // SOLUTION STATEMENT',
    title: 'LIMA SOLUSI TEROBOSAN STRATEGIS & QUICK WINS 30 HARI',
    subtitle: 'Jawaban Konkret untuk Mengatasi 5 GAP Struktural Menuju Kemandirian Penuh BLUD 2030',
    pageNum: 10
  });
  
  const sols = [
    {
      num: 'SOLUSI 01', gapRef: 'MENJAWAB GAP 1',
      title: 'REKRUTMEN MANAJEMEN PROFESIONAL NON-PNS DENGAN REMUNERASI BLUD',
      solusi: 'Penerbitan Keputusan Pemimpin BLUD tentang Struktur Organisasi Fleksibel dan Pola Insentif Kinerja berbasis standar korporasi.',
      quickWin: 'Rekrutmen terbuka 3 Manajer Inkubasi & Kemitraan Industri ex-korporasi teknologi dalam 30 hari pertama.'
    },
    {
      num: 'SOLUSI 02', gapRef: 'MENJAWAB GAP 2',
      title: 'DIVERSIFIKASI BISNIS KE 6 KLASTER PENDAPATAN KOMERSIAL',
      solusi: 'Restrukturisasi tarif layanan laboratorium, sertifikasi vokasi industri, dan optimalisasi sewa fasilitas MICE & Edukasi.',
      quickWin: 'Pemberlakuan e-Tarif Layanan STP baru yang kompetitif untuk 45 jenis uji laboratorium dan fabrikasi.'
    },
    {
      num: 'SOLUSI 03', gapRef: 'MENJAWAB GAP 3',
      title: 'PERATURAN WALIKOTA TENTANG FLEKSIBILITAS KERJASAMA LAHAN 10–25 TAHUN',
      solusi: 'Penyusunan Perwali yang memberikan kepastian hukum skema Bangun Guna Serah (BGS) dan Long-Term Revenue Sharing.',
      quickWin: 'Harmonisasi draf naskah Perwali bersama Bagian Hukum Setda & BPKAD Surakarta dalam 30 hari pertama.'
    },
    {
      num: 'SOLUSI 04', gapRef: 'MENJAWAB GAP 4',
      title: 'SURAT EDARAN WALIKOTA PEMBENTUKAN KONSORSIUM RISET 5 KAMPUS',
      solusi: 'Mewajibkan 30% anggaran riset terapan kampus Solo Raya (UNS, UMS, ISI, Unisri, Poltek) dialokasikan untuk pemecahan masalah industri di STP.',
      quickWin: 'Penandatanganan Nota Kesepahaman Bersama (MoU) 5 Rektor di hadapan Walikota Surakarta.'
    },
    {
      num: 'SOLUSI 05', gapRef: 'MENJAWAB GAP 5',
      title: 'INSTRUKSI WALIKOTA AFIRMASI e-KATALOG LOKAL BELANJA INOVASI',
      solusi: 'Menginstruksikan seluruh OPD Pemkot memprioritaskan belanja produk hasil riset tenant binaan STP melalui e-Katalog Lokal.',
      quickWin: 'Onboarding 5 produk inovasi unggulan (Smart Agri IoT, Kursi Roda AI, Alkes) ke etalase e-Katalog Lokal.'
    }
  ];
  
  sols.forEach((sl, idx) => {
    const sy = 1.65 + (idx * 0.88);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: sy, w: 11.733, h: 0.8,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: sy, w: 0.12, h: 0.8,
      fill: { color: C.emerald }, line: { color: C.emerald }
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 1.05, y: sy + 0.12, w: 1.0, h: 0.28,
      fill: { color: C.bgTintGreen }, line: { color: C.borderGreen, width: 1 }
    });
    s.addText(sl.num, {
      x: 1.05, y: sy + 0.14, w: 1.0, h: 0.24,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.emerald, align: 'center'
    });
    
    s.addText(sl.title, {
      x: 2.2, y: sy + 0.12, w: 9.8, h: 0.28,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.textMain
    });
    
    s.addText([
      { text: 'Terobosan Strategis: ', options: { bold: true, color: C.textMain } },
      { text: sl.solusi + '   ', options: { color: C.textMuted } },
      { text: 'Quick Win 30 Hari: ', options: { bold: true, color: C.emerald } },
      { text: sl.quickWin, options: { color: C.textMuted } }
    ], {
      x: 1.05, y: sy + 0.42, w: 11.2, h: 0.35,
      fontFace: C.fontBody, fontSize: 10, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintGreen }, line: { color: C.borderGreen, width: 1 }
  });
  s.addText('PRINSIP ZERO APBD BURDEN: Kelima solusi terobosan di atas tidak menuntut penambahan pagu belanja modal APBD, melainkan membuka keran kemandirian melalui kepastian regulasi kepemimpinan Walikota.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.emerald, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 11: ROADMAP 3 FASE TRANSFORMASI (POLISHED - NO ORPHAN "M")
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'PENTAHAPAN STRATEGIS // ROADMAP 5 TAHUN',
    title: 'TIGA FASE TRANSFORMASI SOLO TECHNOPARK (2026–2030)',
    subtitle: 'Tahapan Terukur Menuju Kemandirian Finansial Penuh dan Pengakuan sebagai Hub Inovasi Nasional',
    pageNum: 11
  });
  
  const phases = [
    {
      phase: 'FASE 1 (2026–2027)',
      theme: 'KONSOLIDASI & DEREGULASI',
      kicker: 'TARGET PENDAPATAN:',
      target: 'Rp 8,20 M → Rp 12,85 M',
      color: C.navy, bg: C.bgTintBlue, icon: 'shieldCheck',
      milestones: [
        'Penerbitan 3 Regulasi Kunci (Inkwal e-Katalog, Perwali Lahan, SE Konsorsium).',
        'Penetapan struktur organisasi baru & rekrutmen manajer profesional.',
        'Optimalisasi utilisasi 5 fasilitas laboratorium canggih hingga 85%.',
        'Revitalisasi infrastruktur digital & pemetaan aset lahan tidur kawasan.',
        'Rasio kemandirian finansial mencapai 50% dari total kebutuhan operasional.'
      ]
    },
    {
      phase: 'FASE 2 (2028–2029)',
      theme: 'AKSELERASI & BREAK-EVEN POINT',
      kicker: 'TARGET PENDAPATAN (BEP):',
      target: 'Rp 15,60 M → Rp 18,50 M',
      color: C.blue, bg: C.bgTintBlue, icon: 'zap',
      milestones: [
        'PENCAPAIAN BREAK-EVEN POINT (BEP) PENUH PADA AKHIR TAHUN 2029.',
        'Realisasi groundbreaking 2 gedung riset industri kemitraan lahan BLUD.',
        'Komersialisasi massal 10 produk riset unggulan konsorsium kampus.',
        'Ekspansi program diklat vokasi industri bersertifikasi internasional.',
        '100% biaya operasional rutin dibiayai mandiri tanpa subsidi APBD.'
      ]
    },
    {
      phase: 'FASE 3 (2030)',
      theme: 'MANDIRI PENUH & EKSPANSI',
      kicker: 'TARGET PENDAPATAN (SURPLUS):',
      target: 'Rp 23,47 Miliar / Tahun',
      color: C.emerald, bg: C.bgTintGreen, icon: 'award',
      milestones: [
        'Pencapaian surplus operasional bersih BLUD sebesar Rp 4,97 Miliar.',
        'Solo Technopark diakui sebagai Top-3 KST Terbaik di Tingkat Nasional.',
        'Ekspansi layanan sertifikasi & pengujian ke skala regional Jawa Tengah.',
        'Penyetoran dividen PAD ke kas daerah Pemkot Surakarta secara berkala.',
        'Kawasan mandiri energi dan pusat ekosistem ekonomi sirkular modern.'
      ]
    }
  ];
  
  phases.forEach((ph, i) => {
    const cardX = 0.8 + (i * 3.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 0.08,
      fill: { color: ph.color }, line: { color: ph.color }
    });
    
    s.addText(ph.phase, {
      x: cardX + 0.25, y: 1.85, w: 3.28, h: 0.24,
      fontFace: C.fontTitle, fontSize: 9, bold: true, color: ph.color, charSpacing: 1
    });
    
    addSvgIcon(s, { iconName: ph.icon, x: cardX + 0.25, y: 2.12, size: 0.36, color: ph.color, bgColor: ph.bg });
    
    s.addText(ph.theme, {
      x: cardX + 0.8, y: 2.15, w: 2.8, h: 0.32,
      fontFace: C.fontTitle, fontSize: 12.5, bold: true, color: C.textMain
    });
    
    // Stacked 2-Tier Metric Pill: Completely eliminates the orphan "M" wrapping!
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 2.62, w: 3.28, h: 0.50,
      fill: { color: ph.bg }, line: { color: ph.color, width: 0.75 }
    });
    s.addText(ph.kicker, {
      x: cardX + 0.35, y: 2.65, w: 3.08, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: ph.color, align: 'center', charSpacing: 0.8
    });
    s.addText(ph.target, {
      x: cardX + 0.35, y: 2.84, w: 3.08, h: 0.26,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain, align: 'center'
    });
    
    ph.milestones.forEach((m, idx) => {
      const my = 3.25 + (idx * 0.55);
      s.addShape(pptx.shapes.OVAL, {
        x: cardX + 0.25, y: my + 0.06, w: 0.09, h: 0.09,
        fill: { color: ph.color }, line: { color: ph.color }
      });
      s.addText(m, {
        x: cardX + 0.42, y: my, w: 3.1, h: 0.5,
        fontFace: C.fontBody, fontSize: 10.5, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('TRAJEKTORI PASTI: Roadmap 2026–2030 ini bukan angan-angan, melainkan rencana bertahap dengan milestone terukur yang mengarahkan Solo Technopark mencapai titik impas kemandirian pada tahun 2029.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 12: LINTASAN KEMANDIRIAN & TARGET BEP 2029 (POLISHED - NO BEP LINE COLLISION)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'TRAJEKTORI KEMANDIRIAN BLUD // TARGET FINANSIAL',
    title: 'RE-SETTING TARGET PENDAPATAN BLUD MENUJU BEP 2029',
    subtitle: 'Lonjakan Pendapatan dari Rp 8,20 M (2026) Menuju Rp 23,47 M (2030) dengan Titik Impas di 2029',
    pageNum: 12
  });
  
  const chartData = [
    {
      name: 'Target Pendapatan BLUD (Miliar Rp)',
      labels: ['2026 (Baseline)', '2027 (Konsolidasi)', '2028 (Akselerasi)', '2029 (BEP Mandiri)', '2030 (Surplus Penuh)'],
      values: [8.20, 12.85, 15.60, 18.50, 23.47]
    }
  ];
  
  // Set valAxisMaxVal: 28.0 so bars have comfortable headroom and do not collide with reference elements
  s.addChart(pptx.charts.BAR, chartData, {
    x: 0.8, y: 1.65, w: 6.8, h: 4.4,
    barDir: 'col',
    chartColors: ['1E40AF'],
    showValue: true,
    dataLabelColor: '0F172A',
    dataLabelFontFace: C.fontTitle,
    dataLabelFontSize: 11,
    dataLabelPosition: 'outEnd',
    dataLabelFormatCode: '#,##0.00 "M"',
    showLegend: false,
    catAxisLabelFontFace: C.fontTitle,
    catAxisLabelFontSize: 9.5,
    catAxisLabelColor: '475569',
    valAxisLabelFontFace: C.fontBody,
    valAxisLabelFontSize: 9,
    valAxisLabelColor: '64748B',
    valGridLine: { color: C.border, width: 0.5 },
    valAxisMaxVal: 28.0,
    valAxisMinVal: 0
  });
  
  // Elegant BEP Callout & Reference Line that does NOT intersect the 18.50 M data label
  s.addShape(pptx.shapes.LINE, {
    x: 1.3, y: 3.05, w: 3.6, h: 0,
    line: { color: C.emerald, width: 1.5, dashType: 'dash' }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 1.3, y: 2.76, w: 3.6, h: 0.26,
    fill: { color: C.bgTintGreen }, line: { color: C.borderGreen, width: 0.75 }
  });
  s.addText('AMBANG BEP MANDIRI: Rp 18,50 M (Dicapai 2029)', {
    x: 1.35, y: 2.78, w: 3.5, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.emerald, align: 'center'
  });
  
  const metrics = [
    {
      title: 'LOMPATAN FINANSIAL +212,8%',
      val: 'Rp 23,47 M',
      color: C.navy, bg: C.bgTintBlue,
      desc: 'Pertumbuhan pendapatan BLUD lebih dari 3 kali lipat dalam 5 tahun, dipicu oleh beroperasinya kemitraan sewa lahan dan hilirisasi produk riset.'
    },
    {
      title: 'TITIK IMPAS MANDIRI (BEP)',
      val: 'Tahun 2029',
      color: C.emerald, bg: C.bgTintGreen,
      desc: 'Pada tahun 2029, pendapatan mandiri sebesar Rp 18,50 Miliar telah menutup 100% biaya operasional rutin (gaji non-PNS, listrik lab, dan pemeliharaan).'
    },
    {
      title: 'EFISIENSI BELANJA APBD',
      val: 'Rp 10,2 M / Thn',
      color: C.purple, bg: C.bgTintPurple,
      desc: 'Pelepasan subsidi operasional secara bertahap menghemat belanja Pemkot hingga Rp 10,2 M/tahun yang dapat direalokasikan untuk program kesejahteraan warga.'
    }
  ];
  
  metrics.forEach((m, idx) => {
    const my = 1.65 + (idx * 1.48);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 7.933, y: my, w: 4.6, h: 1.35,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 7.933, y: my, w: 0.1, h: 1.35,
      fill: { color: m.color }, line: { color: m.color }
    });
    
    s.addText(m.title, {
      x: 8.15, y: my + 0.12, w: 3.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: m.color, charSpacing: 0.5
    });
    s.addText(m.val, {
      x: 8.15, y: my + 0.35, w: 4.2, h: 0.38,
      fontFace: C.fontTitle, fontSize: 18, bold: true, color: C.textMain
    });
    s.addText(m.desc, {
      x: 8.15, y: my + 0.75, w: 4.2, h: 0.55,
      fontFace: C.fontBody, fontSize: 10, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('AKUNTABILITAS FISKAL: Lintasan pertumbuhan ini disusun berdasarkan kontrak berjalan, re-setting tarif pengujian lab, dan proyeksi konservatif pemanfaatan aset lahan yang siap eksekusi.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.navy, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 13: PROPORSI 6 KLASTER & TOP 5 KATALISATOR (POLISHED - 1-LINE TITLES)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'STRUKTUR REVENUE STREAM // ANALISIS PORTOFOLIO 2030',
    title: 'KOMPOSISI 6 KLASTER PENDAPATAN & TOP 5 KATALISATOR',
    subtitle: 'Diversifikasi Sumber Penerimaan untuk Memastikan Stabilitas Keuangan Tanpa Ketergantungan Tunggal',
    pageNum: 13
  });
  
  const donutData = [
    {
      name: 'Proporsi Pendapatan 2030',
      labels: [
        'Klaster A: Kerjasama Lahan (54,6%)',
        'Klaster B: Diklat Vokasi (16,7%)',
        'Klaster C: Fabrikasi & Lab (10,4%)',
        'Klaster D: Inkubasi Startup (7,5%)',
        'Klaster E: Riset Industri (6,0%)',
        'Klaster F: MICE & Wisata Tekno (4,8%)'
      ],
      values: [54.6, 16.7, 10.4, 7.5, 6.0, 4.8]
    }
  ];
  
  s.addChart(pptx.charts.DOUGHNUT, donutData, {
    x: 0.8, y: 1.65, w: 5.6, h: 4.4,
    holeSize: 55,
    chartColors: ['1E3A8A', '1E40AF', '0284C7', 'D97706', '7C3AED', '059669'],
    showLabel: false,
    showValue: false,
    showLegend: true,
    legendPos: 'b',
    legendFontFace: C.fontBody,
    legendFontSize: 9,
    legendColor: '334155'
  });
  
  s.addText('TOTAL 2030\nRp 23,47 M', {
    x: 2.1, y: 3.1, w: 3.0, h: 0.8,
    fontFace: C.fontTitle, fontSize: 14, bold: true, color: C.navy, align: 'center'
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 6.733, y: 1.65, w: 5.8, h: 4.4,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 6.733, y: 1.65, w: 5.8, h: 0.08,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('TOP 5 KATALISATOR PENDAPATAN TERTINGGI 2030', {
    x: 6.983, y: 1.85, w: 5.3, h: 0.25,
    fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.navy, charSpacing: 1
  });
  
  const topDrivers = [
    { rank: '01', source: 'Sewa Lahan Kampus Satelit Industri (Zona Utara)', val: 'Rp 6,50 M / Thn', porsi: '27,7%', color: C.navy },
    { rank: '02', source: 'Sewa Gedung Komersial & Office IT (Zona Tengah)', val: 'Rp 3,20 M / Thn', porsi: '13,6%', color: C.blue },
    { rank: '03', source: 'Kelas Diklat Industri Vokasi & Migas (OGSCI)', val: 'Rp 2,80 M / Thn', porsi: '11,9%', color: C.sky },
    { rank: '04', source: 'Jasa Machining CNC Presisi & Fablab 3D', val: 'Rp 1,85 M / Thn', porsi: '7,9%', color: C.amber },
    { rank: '05', source: 'Sewa Coworking Space & Akselerasi Startup', val: 'Rp 1,40 M / Thn', porsi: '6,0%', color: C.purple }
  ];
  
  topDrivers.forEach((td, idx) => {
    const ty = 2.2 + (idx * 0.74);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 6.983, y: ty, w: 5.3, h: 0.66,
      fill: { color: C.bg }, line: { color: C.border, width: 0.75 }
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 7.08, y: ty + 0.12, w: 0.42, h: 0.42,
      fill: { color: C.bgTintBlue }, line: { color: td.color, width: 1 }
    });
    s.addText(td.rank, {
      x: 7.08, y: ty + 0.18, w: 0.42, h: 0.3,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: td.color, align: 'center'
    });
    
    // Title container widened to 3.65" at 9.5pt to fit perfectly on ONE line!
    s.addText(td.source, {
      x: 7.60, y: ty + 0.08, w: 3.65, h: 0.28,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textMain
    });
    s.addText(td.val, {
      x: 11.15, y: ty + 0.08, w: 1.05, h: 0.28,
      fontFace: C.fontTitle, fontSize: 10, bold: true, color: td.color, align: 'right'
    });
    s.addText(`Kontribusi Porsi: ${td.porsi} terhadap total pendapatan STP`, {
      x: 7.60, y: ty + 0.36, w: 4.5, h: 0.22,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('STABILITAS FINANSIAL: 5 katalisator terbesar di atas menyumbang 67,1% total penerimaan. Kombinasi sewa lahan jangka panjang (penerimaan pasti) dan jasa teknologi (pertumbuhan dinamis) menjamin arus kas BLUD sangat sehat.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 14: MESIN PENDAPATAN UTAMA: KLASTER A LAHAN & ASET (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'DEEP DIVE FINANSIAL // THE ANCHOR REVENUE ENGINE',
    title: 'ANATOMI KLASTER A: KERJASAMA PEMANFAATAN LAHAN & ASET',
    subtitle: 'Sumber Pendapatan Jangkar Senilai Rp 12.825.000.000 (54,6% dari Total Target Mandiri 2030)',
    pageNum: 14
  });
  
  // Left Hero Card
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.65, w: 4.6, h: 4.4,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  
  addSvgIcon(s, { iconName: 'anchor', x: 1.1, y: 1.95, size: 0.45, color: C.textWhite, bgCircle: false });
  
  s.addText('KLASTER A // ANCHOR ENGINE', {
    x: 1.7, y: 1.95, w: 3.4, h: 0.25,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.sky, charSpacing: 1
  });
  s.addText('PEMANFAATAN LAHAN & GEDUNG', {
    x: 1.1, y: 2.3, w: 4.0, h: 0.7,
    fontFace: C.fontTitle, fontSize: 18, bold: true, color: C.textWhite
  });
  
  s.addText('Rp 12.825.000.000', {
    x: 1.1, y: 3.1, w: 4.0, h: 0.5,
    fontFace: C.fontTitle, fontSize: 24, bold: true, color: C.textWhite
  });
  s.addText('Kontribusi: 54,6% dari Total Target Mandiri STP 2030', {
    x: 1.1, y: 3.65, w: 4.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 10, bold: true, color: C.sky
  });
  
  s.addShape(pptx.shapes.LINE, {
    x: 1.1, y: 4.0, w: 4.0, h: 0,
    line: { color: C.blue, width: 1 }
  });
  
  s.addText('Karakteristik Finansial:', {
    x: 1.1, y: 4.15, w: 4.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textWhite
  });
  s.addText([
    { text: '• Pendapatan Pasif Pasti: ', options: { bold: true } },
    { text: 'Kontrak sewa multi-tahun (10–25 thn) memberikan stabilitas penerimaan tanpa biaya variabel harian.\n' },
    { text: '• Zero APBD Burden: ', options: { bold: true } },
    { text: 'Pembangunan fisik gedung R&D dibiayai 100% oleh mitra industri swasta.\n' },
    { text: '• Kunci Keberhasilan: ', options: { bold: true } },
    { text: 'Membutuhkan payung hukum Perwali Fleksibilitas Kerjasama Lahan BLUD.' }
  ], {
    x: 1.1, y: 4.4, w: 4.0, h: 1.5,
    fontFace: C.fontBody, fontSize: 10, color: C.borderDark, lineSpacingMultiple: 1.2
  });
  
  // Right 4 Sub-Items with Micro Zone Badges
  const subA = [
    {
      zone: 'ZONA UTARA (1,2 HA)',
      title: 'Sewa Lahan Kampus Satelit Industri',
      nom: 'Rp 6.500.000.000 / Thn',
      porsi: '50,7% Klaster A',
      desc: 'Pemanfaatan lahan tidur 1,2 Ha untuk pendirian gedung pusat riset & pengembangan permanen oleh 2 korporasi teknologi global dengan sewa 20 tahun.'
    },
    {
      zone: 'ZONA TENGAH (15 TENANT)',
      title: 'Sewa Gedung Perkantoran IT & Retail Inovasi',
      nom: 'Rp 3.200.000.000 / Thn',
      porsi: '25,0% Klaster A',
      desc: 'Penyewaan ruang kantor siap pakai untuk 15 perusahaan software house, konsultan IT nasional, dan showroom produk teknologi tenant.'
    },
    {
      zone: 'KEMITRAAN MICE',
      title: 'Kerjasama Operasional MICE & Fasilitas Kawasan',
      nom: 'Rp 1.850.000.000 / Thn',
      porsi: '14,4% Klaster A',
      desc: 'Skema bagi hasil operasional gedung serbaguna, arena pertemuan, dan fasilitas penunjang kawasan bersama operator event management profesional.'
    },
    {
      zone: 'DIGITAL ADVERTISING',
      title: 'Ruang Terbuka & Digital Signage Advertising',
      nom: 'Rp 1.275.000.000 / Thn',
      porsi: '9,9% Klaster A',
      desc: 'Monetisasi videotron digital LED luar ruang di gerbang kawasan STP, penamaan gedung (naming rights), dan booth sponsor tahunan mitra korporasi.'
    }
  ];
  
  subA.forEach((sa, idx) => {
    const say = 1.65 + (idx * 1.1);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 5.65, y: say, w: 6.883, h: 0.98,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 5.65, y: say, w: 0.1, h: 0.98,
      fill: { color: C.blue }, line: { color: C.blue }
    });
    
    // Zone Badge
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 5.88, y: say + 0.10, w: 1.65, h: 0.22,
      fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 0.75 }
    });
    s.addText(sa.zone, {
      x: 5.88, y: say + 0.12, w: 1.65, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue, align: 'center'
    });
    
    s.addText(sa.title, {
      x: 7.60, y: say + 0.08, w: 3.1, h: 0.26,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.textMain
    });
    s.addText(sa.nom, {
      x: 10.5, y: say + 0.08, w: 1.9, h: 0.26,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.blue, align: 'right'
    });
    s.addText(`Porsi: ${sa.porsi}  •  ${sa.desc}`, {
      x: 5.88, y: say + 0.38, w: 6.45, h: 0.52,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.18
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('PONDASI FINANSIAL UTAMA: Penerbitan regulasi fleksibilitas kerjasama lahan jangka panjang adalah "kunci pembuka gerbang" pendapatan Rp 12,82 Miliar ini. Begitu regulasi terbit, pendapatan mandiri STP langsung melonjak drastis.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.navy, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 15: MESIN PENDAPATAN LAYANAN: KLASTER B-F (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'DEEP DIVE FINANSIAL // SERVICE & INNOVATION REVENUE',
    title: 'ANATOMI KLASTER B–F: LAYANAN INOVASI & JASA TEKNOLOGI',
    subtitle: 'Pendapatan Berbasis Jasa, Sertifikasi, dan Pengujian Senilai Rp 10.643.158.831 (45,4% dari Total Target)',
    pageNum: 15
  });
  
  const serviceClusters = [
    {
      code: 'KLASTER B', name: 'DIKLAT VOKASI & SERTIFIKASI KOMPETENSI',
      nom: 'Rp 3.920.000.000', porsi: '16,7%', color: C.navy, bg: C.bgTintBlue, icon: 'graduationCap',
      items: 'Diklat Industri Vokasi (Migas, Underwater Welding, Otomasi PLC), Sertifikasi BNSP & Internasional, In-House Training Korporasi.'
    },
    {
      code: 'KLASTER C', name: 'FABRIKASI PRESISI & PENGUJIAN LABORATORIUM',
      nom: 'Rp 2.450.000.000', porsi: '10,4%', color: C.blue, bg: C.bgTintBlue, icon: 'wrench',
      items: 'Jasa Pemesinan CNC 5-Axis presisi, Fablab 3D Printing Logam & Resin, Uji Mutu Bahan Pangan & Ekstrak Herbal, Kalibrasi Alat Industri.'
    },
    {
      code: 'KLASTER D', name: 'INKUBASI BISNIS & SEWA CO-WORKING SPACE',
      nom: 'Rp 1.750.000.000', porsi: '7,5%', color: C.emerald, bg: C.bgTintGreen, icon: 'rocket',
      items: 'Sewa ruang tenant startup, Dedicated Desk Co-working Space, Virtual Office UMKM, Manajemen Fee Akselerasi Bisnis Teknologi.'
    },
    {
      code: 'KLASTER E', name: 'KONSULTANSI TEKNOLOGI & RISET INDUSTRI',
      nom: 'Rp 1.400.000.000', porsi: '6,0%', color: C.amber, bg: C.bgTintAmber, icon: 'flask',
      items: 'Studi Kelayakan Teknis Industri, Reverse Engineering Mesin Manufaktur, Royalti Lisensi Paten Inovasi Kampus, Audit Energi & IoT.'
    },
    {
      code: 'KLASTER F', name: 'MICE & WISATA EDUKASI TEKNOLOGI DAERAH',
      nom: 'Rp 1.123.158.831', porsi: '4,8%', color: C.purple, bg: C.bgTintPurple, icon: 'globe',
      items: 'Tiket Wisata Edukasi Sains & Teknologi untuk Pelajar se-Jateng & DIY, Sewa Auditorium Seminar Nasional, Penyelenggaraan Expo Teknologi.'
    }
  ];
  
  serviceClusters.forEach((sc, idx) => {
    const scy = 1.65 + (idx * 0.88);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: scy, w: 11.733, h: 0.8,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: scy, w: 0.12, h: 0.8,
      fill: { color: sc.color }, line: { color: sc.color }
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 1.05, y: scy + 0.12, w: 1.1, h: 0.28,
      fill: { color: sc.bg }, line: { color: sc.color, width: 1 }
    });
    s.addText(sc.code, {
      x: 1.05, y: scy + 0.14, w: 1.1, h: 0.24,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: sc.color, align: 'center'
    });
    
    s.addText(sc.name, {
      x: 2.3, y: scy + 0.12, w: 6.8, h: 0.28,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    
    s.addText(`${sc.nom}  (${sc.porsi})`, {
      x: 9.2, y: scy + 0.12, w: 3.1, h: 0.28,
      fontFace: C.fontTitle, fontSize: 12, bold: true, color: sc.color, align: 'right'
    });
    
    s.addText([
      { text: 'Rincian Layanan Unggulan: ', options: { bold: true, color: C.textMain } },
      { text: sc.items, options: { color: C.textMuted } }
    ], {
      x: 1.05, y: scy + 0.44, w: 11.2, h: 0.32,
      fontFace: C.fontBody, fontSize: 10, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('TOTAL TARGET PENDAPATAN MANDIRI 2030 (Klaster A + Klaster B s.d. F): Rp 23.468.158.831 (100% Bebas Subsidi Belanja Operasional APBD Surakarta).', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 16: ENAM DIMENSI KEBERDAMPAKAN DAERAH (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'KONTRIBUSI NYATA // DAMPAK PEMBANGUNAN DAERAH',
    title: 'ENAM DIMENSI DAMPAK SOLO TECHNOPARK UNTUK KOTA SURAKARTA',
    subtitle: 'Parameter Kuantitatif dan Kualitatif Keberhasilan Pembangunan Ekosistem Inovasi 2026–2030',
    pageNum: 16
  });
  
  const dims = [
    {
      title: '1. DAMPAK EKONOMI',
      metric: 'Rp 18,5 Miliar',
      tag: '+300% MULTIPLIER',
      sub: 'Omzet Komulatif Tenant Inovasi 2030',
      color: C.navy, bg: C.bgTintBlue, icon: 'trendingUp',
      desc: 'Menciptakan perputaran ekonomi baru di Surakarta melalui komersialisasi produk inovasi teknologi dan pembentukan startup lokal.'
    },
    {
      title: '2. SERAPAN KERJA',
      metric: '1.500 Talenta / Tahun',
      tag: '85%+ LULUS TERSERAP',
      sub: 'Terserap Industri dengan Upah Layak',
      color: C.blue, bg: C.bgTintBlue, icon: 'users',
      desc: '85%+ lulusan diklat vokasi STP langsung terserap di industri teknologi nasional dan manufaktur presisi, menekan angka pengangguran terbuka.'
    },
    {
      title: '3. FISKAL & PAD',
      metric: 'Rp 10,2 Miliar / Tahun',
      tag: 'NOL SUBSIDI 2029',
      sub: 'Efisiensi Belanja APBD Pasca-BEP',
      color: C.emerald, bg: C.bgTintGreen, icon: 'dollarSign',
      desc: 'APBD tidak lagi terbebani subsidi operasional STP; kas daerah justru menerima kontribusi PAD melalui pajak dan retribusi kegiatan komersial.'
    },
    {
      title: '4. EKOSISTEM HKI',
      metric: '50+ Paten & Hak Cipta',
      tag: 'RUJUKAN INOVASI',
      sub: 'Didaftarkan & Dilisensikan ke Industri',
      color: C.amber, bg: C.bgTintAmber, icon: 'award',
      desc: 'Solo menjadi rujukan pendaftaran dan komersialisasi kekayaan intelektual (HKI) perguruan tinggi dan inovator lokal di Jawa Tengah.'
    },
    {
      title: '5. BRANDING KOTA',
      metric: 'Top-3 KST Nasional',
      tag: 'STANDAR GLOBAL',
      sub: 'Rujukan Nasional & Internasional',
      color: C.purple, bg: C.bgTintPurple, icon: 'globe',
      desc: 'Meningkatkan citra Kota Surakarta dari kota budaya-heritage menjadi kota modern yang memadukan budaya luhur dengan inovasi teknologi masa depan.'
    },
    {
      title: '6. INKLUSI WARGA',
      metric: '25.000 Pelajar / Tahun',
      tag: 'LITERASI GRATIS',
      sub: 'Penerima Manfaat Edukasi Sains',
      color: C.red, bg: C.bgTintRed, icon: 'sparkles',
      desc: 'Membuka akses gratis literasi digital, kunjungan laboratorium sains, dan pelatihan coding dasar bagi anak-anak keluarga kurang mampu di Surakarta.'
    }
  ];
  
  dims.forEach((d, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const dx = 0.8 + (col * 3.98);
    const dy = 1.65 + (row * 2.25);
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: dx, y: dy, w: 3.78, h: 2.15,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: dx, y: dy, w: 3.78, h: 0.08,
      fill: { color: d.color }, line: { color: d.color }
    });
    
    // Top Row: Icon + Short Title + Top-Right Tag Pill (Zero Collision!)
    addSvgIcon(s, { iconName: d.icon, x: dx + 0.2, y: dy + 0.16, size: 0.30, color: d.color, bgColor: d.bg });
    
    s.addText(d.title, {
      x: dx + 0.60, y: dy + 0.16, w: 1.70, h: 0.24,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textMain
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: dx + 2.36, y: dy + 0.16, w: 1.22, h: 0.24,
      fill: { color: d.bg }, line: { color: d.color, width: 0.75 }
    });
    s.addText(d.tag, {
      x: dx + 2.36, y: dy + 0.18, w: 1.22, h: 0.20,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: d.color, align: 'center'
    });
    
    // Metric Row: Full Width 3.38" - sits 100% on ONE single line without wrapping
    s.addText(d.metric, {
      x: dx + 0.2, y: dy + 0.52, w: 3.38, h: 0.36,
      fontFace: C.fontTitle, fontSize: 16.5, bold: true, color: d.color
    });
    
    // Subtitle Row
    s.addText(d.sub, {
      x: dx + 0.2, y: dy + 0.90, w: 3.38, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.textDim
    });
    
    // Hairline Divider
    s.addShape(pptx.shapes.LINE, {
      x: dx + 0.2, y: dy + 1.16, w: 3.38, h: 0,
      line: { color: C.border, width: 1 }
    });
    
    // Description Paragraph
    s.addText(d.desc, {
      x: dx + 0.2, y: dy + 1.24, w: 3.38, h: 0.78,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.18
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  s.addText('SOLO TECHNOPARK UNTUK WARGA: Sukses transformasi ini bukan diukur dari megahnya gedung, melainkan dari berapa banyak anak muda Solo yang mendapat kerja layak dan berapa banyak wirausaha baru yang lahir.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.navy, lineSpacingMultiple: 1.2
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 17: TIGA PERMOHONAN KEBIJAKAN NON-ANGGARAN (POLISHED)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'DUKUNGAN KEPEMIMPINAN // PAYUNG REGULASI',
    title: '3 PERMOHONAN KEBIJAKAN NON-ANGGARAN KEPADA WALIKOTA',
    subtitle: 'Kami Tidak Memohon Tambahan Pagu APBD — Kami Memohon 3 Payung Regulasi Kepemimpinan Walikota',
    pageNum: 17
  });
  
  const asks = [
    {
      num: 'PRIORITAS UTAMA // QUICK WIN 30 HARI',
      title: 'INSTRUKSI WALIKOTA AFIRMASI e-KATALOG LOKAL',
      jenis: 'INSTRUKSI WALIKOTA',
      color: C.blue, bg: C.bgTintBlue, icon: 'fileText',
      narasi: 'Mengarahkan seluruh OPD Pemkot Surakarta untuk memprioritaskan belanja produk teknologi, layanan digital, dan barang hasil riset tenant binaan STP melalui etalase e-Katalog Lokal.',
      impact: 'Membuka captive market pasti senilai Rp 2,25 Miliar/Tahun di 2030 bagi 80+ produk inovasi tenant tanpa menambah anggaran APBD.',
      mekanisme: 'Penerbitan 1 lembar Instruksi Walikota. Draf telah siap harmonisasi bersama Bagian Hukum Setda & BPKAD. Siap tuntas dalam 30 hari pertama.'
    },
    {
      num: 'INVESTASI JANGKA PANJANG (10–25 TAHUN)',
      title: 'REGULASI FLEKSIBILITAS KERJASAMA LAHAN BLUD',
      jenis: 'PERATURAN WALIKOTA',
      color: C.amber, bg: C.bgTintAmber, icon: 'scale',
      narasi: 'Penyusunan Peraturan Walikota yang memberikan kepastian hukum skema Long-Term Revenue Sharing dan sewa lahan jangka panjang (10–25 tahun) yang kompetitif bagi investor teknologi global di kawasan STP.',
      impact: 'Menarik 3–5 investor korporasi baru kelas dunia tanpa membebani belanja modal APBD daerah (Prinsip Zero APBD Burden).',
      mekanisme: 'Harmonisasi Perwali Pola Kerjasama Lahan bersama BPKAD, Bagian Hukum, dan Inspektorat Kota Surakarta.'
    },
    {
      num: 'SINERGI KAMPUS SOLO RAYA',
      title: 'SURAT EDARAN KONSORSIUM RISET DAERAH',
      jenis: 'SURAT EDARAN WALIKOTA',
      color: C.emerald, bg: C.bgTintGreen, icon: 'graduationCap',
      narasi: 'Mendorong 5 perguruan tinggi ternama (UNS, UMS, ISI, Unisri, Poltek) mengarahkan minimal 30% dana riset terapan dosen dan tugas akhir mahasiswa untuk memecahkan problem kota dan industri di STP.',
      impact: '5 Kampus mitra aktif, 30 klaster riset terapan berbasis problem kota, dan minimal 10 paten terhilirisasi ke industri per tahun.',
      mekanisme: 'Penerbitan Surat Edaran Walikota Surakarta dan penandatanganan kesepakatan bersama (MoU) 5 Rektor Kampus se-Solo Raya.'
    }
  ];
  
  asks.forEach((a, i) => {
    const cardX = 0.8 + (i * 3.98);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 4.4,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX, y: 1.65, w: 3.78, h: 0.08,
      fill: { color: a.color }, line: { color: a.color }
    });
    
    addSvgIcon(s, { iconName: a.icon, x: cardX + 0.25, y: 1.85, size: 0.38, color: a.color, bgColor: a.bg });
    
    s.addText(a.num, {
      x: cardX + 0.8, y: 1.85, w: 2.8, h: 0.24,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: a.color, charSpacing: 0.5
    });
    s.addText(a.title, {
      x: cardX + 0.8, y: 2.12, w: 2.8, h: 0.5,
      fontFace: C.fontTitle, fontSize: 12, bold: true, color: C.textMain
    });
    
    s.addText(a.narasi, {
      x: cardX + 0.25, y: 2.75, w: 3.28, h: 0.85,
      fontFace: C.fontBody, fontSize: 10.5, color: C.textMuted, lineSpacingMultiple: 1.2
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 3.7, w: 3.28, h: 1.05,
      fill: { color: C.bg }, line: { color: a.color, width: 1 }
    });
    s.addText('TARGET DAMPAK KONKRET:', {
      x: cardX + 0.4, y: 3.8, w: 2.98, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: a.color
    });
    s.addText(a.impact, {
      x: cardX + 0.4, y: 4.05, w: 2.98, h: 0.65,
      fontFace: C.fontBody, fontSize: 10, color: C.textMain, lineSpacingMultiple: 1.15
    });
    
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.25, y: 4.9, w: 3.28, h: 0.95,
      fill: { color: C.bgSoft }, line: { color: C.border, width: 0.75 }
    });
    
    // Header pill for regulation mechanism
    s.addShape(pptx.shapes.RECTANGLE, {
      x: cardX + 0.35, y: 4.96, w: 1.8, h: 0.22,
      fill: { color: a.bg }, line: { color: a.color, width: 0.75 }
    });
    s.addText(`MEKANISME: ${a.jenis}`, {
      x: cardX + 0.35, y: 4.98, w: 1.8, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7, bold: true, color: a.color, align: 'center'
    });
    
    s.addText(a.mekanisme, {
      x: cardX + 0.35, y: 5.24, w: 3.08, h: 0.58,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('KOMITMEN KAMI: Cukup dengan 3 payung regulasi Walikota di atas, UPTD KST Solo Technopark siap bergerak cepat, mandiri, dan menuntaskan target Rp 23,47 Miliar tanpa subsidi APBD.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(s);
}

// ----------------------------------------------------
// SLIDE 18: RENCANA AKSI 100 HARI PERTAMA (POLISHED - EXECUTIVE NAVY HEADER)
// ----------------------------------------------------
{
  const s = pptx.addSlide();
  addHeader(s, {
    kicker: 'KOMITMEN EKSEKUSI CEPAT // AKUNTABILITAS',
    title: 'RENCANA AKSI 100 HARI PERTAMA: QUICK WINS PASCA-ARAHAN WALIKOTA',
    subtitle: 'Milestone Terukur, Matriks Penanggung Jawab, dan Output Nyata dalam 100 Hari Pasca-Persetujuan Kebijakan',
    pageNum: 18
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.6, w: 11.733, h: 0.42,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  addSvgIcon(s, { iconName: 'sparkles', x: 0.95, y: 1.66, size: 0.28, color: C.blue, bgColor: C.bgTintBlue });
  s.addText('AGENDA EKSEKUTIF 100 HARI: DARI PENANDATANGANAN REGULASI HINGGA KONTRAK PERDANA KEMITRAAN INDUSTRI', {
    x: 1.35, y: 1.70, w: 11.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.blue, charSpacing: 0.5
  });
  
  // Executive Navy Table Header for maximum authority and contrast
  const rows = [
    [
      { text: 'PERIODE', options: { bold: true, fontSize: 9.5, color: C.textWhite, fill: C.navy } },
      { text: 'AGENDA STRATEGIS', options: { bold: true, fontSize: 9.5, color: C.textWhite, fill: C.navy } },
      { text: 'PENANGGUNG JAWAB (PIC)', options: { bold: true, fontSize: 9.5, color: C.textWhite, fill: C.navy, align: 'center' } },
      { text: 'OUTPUT KONKRET & TARGET KEBERHASILAN', options: { bold: true, fontSize: 9.5, color: C.textWhite, fill: C.navy } }
    ],
    [
      { text: 'Hari 1–10', options: { bold: true, fontSize: 10, color: C.textMain } },
      { text: 'Penyusunan Draf Regulasi & Harmonisasi Awal:\nFinalisasi draf Inkwal e-Katalog, draf SE Konsorsium Riset, dan reviu skema aset lahan BLUD.', options: { fontSize: 10, color: C.textMuted } },
      { text: 'Kepala UPTD STP &\nTim Hukum BRIDA', options: { fontSize: 10, color: C.textMain, align: 'center' } },
      { text: '3 Draf naskah regulasi siap harmonisasi Bagian Hukum Setda & BPKAD.', options: { fontSize: 10, color: C.textMain } }
    ],
    [
      { text: 'Hari 11–30', options: { bold: true, fontSize: 10, color: C.amber, fill: C.bgTintAmber } },
      { text: 'Harmonisasi & Penerbitan Instruksi Walikota (QUICK WIN UTAMA):\nPembahasan harmonisasi bersama Setda & BPKAD; Penandatanganan Inkwal e-Katalog Afirmatif.', options: { fontSize: 10, color: C.amber, bold: true, fill: C.bgTintAmber } },
      { text: 'Bagian Hukum, BPKAD,\nBRIDA, & STP', options: { fontSize: 10, color: C.textMain, align: 'center', fill: C.bgTintAmber } },
      { text: 'Inkwal resmi ditandatangani Walikota Surakarta; Sosialisasi perdana ke seluruh Kepala OPD Pemkot.', options: { fontSize: 10, color: C.amber, bold: true, fill: C.bgTintAmber } }
    ],
    [
      { text: 'Hari 31–50', options: { bold: true, fontSize: 10, color: C.textMain } },
      { text: 'Kurasi Produk Inovasi & Integrasi Sistem e-Katalog:\nKurasi 5 produk inovasi unggulan kampus/tenant untuk onboarding perdana ke etalase e-Katalog Lokal.', options: { fontSize: 10, color: C.textMuted } },
      { text: 'UPTD STP & Bagian\nPBJ Setda Solo', options: { fontSize: 10, color: C.textMain, align: 'center' } },
      { text: 'Katalog 5 Produk Unggulan Hilirisasi STP tayang aktif (Smart Agri IoT, Alkes, Mesin Sortasi).', options: { fontSize: 10, color: C.textMain } }
    ],
    [
      { text: 'Hari 51–60', options: { bold: true, fontSize: 10, color: C.textMain } },
      { text: 'Showcase Produk & Komitmen Serapan Perdana OPD:\nDemonstrasi produk inovasi di hadapan Walikota, Sekda, dan OPD terkait di Solo Technopark.', options: { fontSize: 10, color: C.textMuted } },
      { text: 'Kepala BRIDA &\nKepala UPTD STP', options: { fontSize: 10, color: C.textMain, align: 'center' } },
      { text: 'Demonstrasi produk di hadapan media & komitmen serapan perdana oleh 3 OPD percontohan.', options: { fontSize: 10, color: C.textMain } }
    ],
    [
      { text: 'Hari 61–80', options: { bold: true, fontSize: 10, color: C.textMain } },
      { text: 'Inisiasi Konsorsium Riset 5 Perguruan Tinggi:\nRapat koordinasi tindak lanjut Surat Edaran Walikota bersama pimpinan LPPM UNS, UMS, ISI, Poltek.', options: { fontSize: 10, color: C.textMuted } },
      { text: 'BRIDA Kota Surakarta\n& Konsorsium Kampus', options: { fontSize: 10, color: C.textMain, align: 'center' } },
      { text: 'MoU Pra-Konsorsium bersama 5 Rektor Solo Raya; Terpetakan 30 tema riset terapan prioritas.', options: { fontSize: 10, color: C.textMain } }
    ],
    [
      { text: 'Hari 81–100', options: { bold: true, fontSize: 10, color: C.emerald, fill: C.bgTintGreen } },
      { text: 'Launching Resmi Konsorsium & Kick-off Kemitraan:\nPenandatanganan MoU Riset Terapan bersama Kadin/Apindo & kontrak kemitraan industri mandiri.', options: { fontSize: 10, color: C.emerald, bold: true, fill: C.bgTintGreen } },
      { text: 'Walikota Surakarta,\nRektor, & Mitra Industri', options: { fontSize: 10, color: C.textMain, align: 'center', fill: C.bgTintGreen } },
      { text: 'Penandatanganan MoU Riset Terapan & Peresmian kelas diklat industri mandiri perdana.', options: { fontSize: 10, color: C.emerald, bold: true, fill: C.bgTintGreen } }
    ]
  ];
  
  s.addTable(rows, {
    x: 0.8, y: 2.15, w: 11.733,
    colW: [1.3, 4.4, 2.3, 3.733],
    border: { color: C.border, width: 0.75 },
    margin: [0.06, 0.10, 0.06, 0.10],
    lineSpacingMultiple: 1.15
  });
  
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.55,
    fill: { color: C.navy }, line: { color: C.navy }
  });
  s.addText('SISTEM MONITORING: Evaluasi progres dilaksanakan mingguan bersama BRIDA dan dilaporkan berkala setiap 30 hari langsung kepada Walikota Surakarta untuk memastikan seluruh quick wins tercapai 100%.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.42,
    fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(s);
}

// ==========================================
// SIMPAN PPTX FILE
// ==========================================
const outDir = path.resolve('d:/Project/GAWE/Paparan Roadmap STP/bahan_paparan_walikota_v3');
const pptxFile1 = path.join(outDir, 'PAPARAN_WALIKOTA_STP_V9_18S.pptx');
const distDir = path.resolve('d:/Project/GAWE/Paparan Roadmap STP/output_paparan');
const pptxFile2 = path.join(distDir, 'PAPARAN_WALIKOTA_STP_V9_18S.pptx');

console.log('Writing PowerPoint file to:', pptxFile1);
await pptx.writeFile({ fileName: pptxFile1 });

if (!fs.existsSync(distDir)) {
  fs.mkdirSync(distDir, { recursive: true });
}
fs.copyFileSync(pptxFile1, pptxFile2);
console.log('Copied PowerPoint to distribution folder:', pptxFile2);
console.log('PptxGenJS generation finished successfully!');
