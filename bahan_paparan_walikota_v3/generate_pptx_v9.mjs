import PptxGenJS from 'pptxgenjs';
import fs from 'fs';
import path from 'path';

// ==========================================
// PALET WARNA SWISS ENTERPRISE ULTRA (V9)
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
  presentation: '<path d="M2 3h20"></path><path d="M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"></path><path d="m7 21 5-5 5 5"></path>'
};

function getSvgBase64(iconName, color = C.blue, strokeWidth = 2) {
  const inner = SVG_PATHS[iconName] || SVG_PATHS.checkCircle;
  const hexColor = '#' + color.replace('#', '');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96" fill="none" stroke="${hexColor}" stroke-width="${strokeWidth}" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;
  return Buffer.from(svg).toString('base64');
}

// Inisialisasi Presentasi
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';
pptx.author = 'Pemerintah Kota Surakarta - BRIDA & UPTD KST Solo Technopark';
pptx.title = 'Roadmap Transformasi Solo Technopark 2026-2030 (Swiss Enterprise Ultra V9)';

// ==========================================
// HELPER UI FUNCTIONS (V9)
// ==========================================

// Header Standar V9
function addHeader(slide, { kicker, title, subtitle, pageNum }) {
  // Kicker
  slide.addText(kicker.toUpperCase(), {
    x: 0.8, y: 0.45, w: 10.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.blue,
    charSpacing: 1.5
  });
  
  // Title
  slide.addText(title, {
    x: 0.8, y: 0.68, w: 10.5, h: 0.48,
    fontFace: C.fontTitle, fontSize: 19, bold: true, color: C.textMain
  });
  
  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 1.15, w: 10.5, h: 0.28,
    fontFace: C.fontBody, fontSize: 10.5, italic: true, color: C.textDim
  });
  
  // Page Counter
  if (pageNum) {
    slide.addText(`${pageNum.toString().padStart(2, '0')}  /  12`, {
      x: 11.5, y: 0.52, w: 1.0, h: 0.25,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textDim, align: 'right'
    });
  }
  
  // Hairline Rule
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.48, w: 11.733, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
}

// Footer Standar V9
function addFooter(slide) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 7.02, w: 11.733, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
  
  slide.addText('PEMERINTAH KOTA SURAKARTA  •  BADAN RISET DAN INOVASI DAERAH (BRIDA)  •  ROADMAP 2026–2030', {
    x: 0.8, y: 7.1, w: 8.0, h: 0.22,
    fontFace: C.fontBody, fontSize: 7.5, bold: true, color: C.textDim, charSpacing: 0.8
  });
  
  slide.addText('EDISI EKSEKUTIF ENTERPRISE ULTRA (V9)', {
    x: 9.0, y: 7.1, w: 3.533, h: 0.22,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue, align: 'right', charSpacing: 0.8
  });
}

// Vector Icon Badge dengan Kotak Mikro Berlatar Lembut
function addIconBadge(slide, { iconName, x, y, size = 0.42, bgTint = C.bgTintBlue, strokeColor = C.blue, borderTint = C.borderBlue }) {
  // Container Box
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: size, h: size,
    fill: { color: bgTint },
    line: { color: borderTint, width: 1 }
  });
  
  // SVG Icon
  const pad = size * 0.18;
  slide.addImage({
    data: `image/svg+xml;base64,${getSvgBase64(iconName, strokeColor, 2)}`,
    x: x + pad, y: y + pad, w: size - 2 * pad, h: size - 2 * pad
  });
}

// Status Chip Mikro
function addStatusChip(slide, { text, type = 'blue', x, y, w = 1.35, h = 0.24 }) {
  let bg = C.bgTintBlue;
  let textCol = C.blue;
  let borderCol = C.borderBlue;
  
  if (type === 'green') { bg = C.bgTintGreen; textCol = C.emerald; borderCol = C.borderGreen; }
  else if (type === 'amber') { bg = C.bgTintAmber; textCol = C.amber; borderCol = C.borderAmber; }
  else if (type === 'red') { bg = C.bgTintRed; textCol = C.red; borderCol = C.borderRed; }
  else if (type === 'purple') { bg = C.bgTintPurple; textCol = C.purple; borderCol = 'E9D5FF'; }
  
  slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: bg },
    line: { color: borderCol, width: 0.8 }
  });
  
  slide.addText(text.toUpperCase(), {
    x, y: y + 0.01, w, h,
    fontFace: C.fontTitle, fontSize: 7, bold: true, color: textCol, align: 'center'
  });
}

// ==========================================
// SLIDE 01: COVER & METRIK KUNCI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  
  // Top Institutional Kicker
  slide.addText('PEMERINTAH KOTA SURAKARTA // BADAN RISET DAN INOVASI DAERAH (BRIDA)', {
    x: 0.8, y: 0.7, w: 10.0, h: 0.28,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.blue, charSpacing: 1.5
  });
  
  // Headline Utama
  slide.addText('SOLO TECHNOPARK:\nDARI KAWASAN SAINS MENUJU MESIN IMPACT DAERAH', {
    x: 0.8, y: 1.15, w: 11.5, h: 1.25,
    fontFace: C.fontTitle, fontSize: 26, bold: true, color: C.textMain, lineSpacingMultiple: 1.15
  });
  
  // Tagline Filosofis
  slide.addText('"Roadmap Transformasi UPTD KST Surakarta: Mewujudkan Kemandirian BLUD, Hilirisasi Riset Kampus, dan Kesejahteraan Nyata Warga Kota"', {
    x: 0.8, y: 2.5, w: 11.733, h: 0.35,
    fontFace: C.fontBody, fontSize: 11.5, italic: true, color: C.textDim
  });
  
  // Hairline Divider
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 2.95, w: 11.733, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
  
  // 4 METRIC DISPLAY CARDS (DENGAN VECTOR ICON BADGE)
  const metrics = [
    { num: '+212,8%', lbl: 'Lompatan Finansial Mandiri', sub: 'Target tumbuh 3,1x dalam 4 tahun', color: C.amber, bgTint: C.bgTintAmber, borderTint: C.borderAmber, icon: 'trendingUp' },
    { num: 'Rp 23,47 M', lbl: 'Target Pendapatan BLUD 2030', sub: '0% Subsidi Kas Daerah APBD', color: C.blue, bgTint: C.bgTintBlue, borderTint: C.borderBlue, icon: 'dollarSign' },
    { num: '9.000', lbl: 'Kapasitas SDM Terlatih / Thn', sub: 'Afirmasi min. 65% KTP Surakarta', color: C.emerald, bgTint: C.bgTintGreen, borderTint: C.borderGreen, icon: 'users' },
    { num: '95+', lbl: 'Mitra Industri Global & Nasional', sub: 'Prinsip Zero APBD Burden', color: C.purple, bgTint: C.bgTintPurple, borderTint: C.borderPurple, icon: 'globe' }
  ];
  
  const mY = 3.15, mW = 2.76, mH = 1.25, mGap = 0.23;
  metrics.forEach((m, i) => {
    const mX = 0.8 + i * (mW + mGap);
    
    // Base Panel
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: mX, y: mY, w: mW, h: mH,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    
    // Top Color Accent Strip
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: mX, y: mY, w: mW, h: 0.06,
      fill: { color: m.color }, line: { color: m.color, width: 0 }
    });
    
    // Icon Badge (Top Right)
    addIconBadge(slide, { iconName: m.icon, x: mX + mW - 0.52, y: mY + 0.12, size: 0.38, bgTint: m.bgTint, strokeColor: m.color, borderTint: m.borderTint });
    
    // Number
    slide.addText(m.num, {
      x: mX + 0.18, y: mY + 0.16, w: mW - 0.7, h: 0.42,
      fontFace: C.fontTitle, fontSize: 20, bold: true, color: m.color
    });
    
    // Label
    slide.addText(m.lbl, {
      x: mX + 0.18, y: mY + 0.62, w: mW - 0.36, h: 0.25,
      fontFace: C.fontBody, fontSize: 8.5, bold: true, color: C.textMain
    });
    
    // Sub
    slide.addText(m.sub, {
      x: mX + 0.18, y: mY + 0.88, w: mW - 0.36, h: 0.22,
      fontFace: C.fontBody, fontSize: 7.5, color: C.textDim
    });
  });
  
  // 3 PILAR STRATEGIS (DENGAN VECTOR ICON BADGE)
  const pillars = [
    { no: '01 // PARADIGMA BARU', title: 'Bukan Sekadar Gedung Fisik', desc: 'Beralih dari pengelola sewa lahan menjadi DELIVERY UNIT inovasi BRIDA yang menghasilkan IMPACT nyata bagi masyarakat dan industri.', color: C.sky, icon: 'cpu' },
    { no: '02 // KEMANDIRIAN PENUH', title: 'BLUD Mandiri Tanpa APBD', desc: 'Lompatan pendapatan terukur dari Rp 7,50 M (2026) menuju Rp 23,47 M (2030) — tumbuh +212,8% dengan 0% subsidi kas daerah.', color: C.amber, icon: 'shieldCheck' },
    { no: '03 // HEXAHELIX & DAMPAK', title: '95+ Mitra Korporasi Dunia', desc: 'Investasi swasta murni berkolaborasi menyerap 9.000 tenaga kerja dengan komitmen afirmasi minimal 65% warga KTP Surakarta.', color: C.emerald, icon: 'award' }
  ];
  
  const pY = 4.65, pW = 3.75, pH = 1.95, pGap = 0.24;
  pillars.forEach((p, i) => {
    const pX = 0.8 + i * (pW + pGap);
    
    // Panel
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX, y: pY, w: pW, h: pH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    // Top Accent
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX, y: pY, w: pW, h: 0.05,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: p.icon, x: pX + 0.2, y: pY + 0.18, size: 0.44, bgTint: C.bgSoft, strokeColor: p.color, borderTint: C.border });
    
    // Number Kicker
    slide.addText(p.no, {
      x: pX + 0.72, y: pY + 0.18, w: pW - 0.9, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: p.color, charSpacing: 1
    });
    
    // Title
    slide.addText(p.title, {
      x: pX + 0.72, y: pY + 0.38, w: pW - 0.9, h: 0.28,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    
    // Divider
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX + 0.2, y: pY + 0.72, w: pW - 0.4, h: 0.01,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    // Description
    slide.addText(p.desc, {
      x: pX + 0.2, y: pY + 0.82, w: pW - 0.4, h: 1.0,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted, lineSpacingMultiple: 1.2
    });
  });
  
  // Footer Presenter
  slide.addText('Disampaikan oleh: Kepala UPTD KST Solo Technopark  •  Didampingi: Kepala BRIDA Kota Surakarta  •  Surakarta, 2026', {
    x: 0.8, y: 6.95, w: 11.733, h: 0.25,
    fontFace: C.fontBody, fontSize: 8.5, color: C.textDim, align: 'center'
  });
}

// ==========================================
// SLIDE 02: KEDUDUKAN & RANTAI DAMPAK
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // TATA KELOLA PEMERINTAHAN & KELEMBAGAAN',
    title: 'KEDUDUKAN KELEMBAGAAN & RANTAI DAMPAK (CHAIN OF IMPACT)',
    subtitle: 'Arsitektur Nilai: BRIDA sebagai Otak Kebijakan, STP sebagai Tangan Eksekutor, Warga & Industri sebagai Penerima Dampak',
    pageNum: 2
  });
  
  const colY = 1.75, colW = 3.65, colH = 4.25;
  const cols = [
    {
      kicker: '01 / INPUT',
      title: 'BRIDA KOTA SURAKARTA',
      subtitle: 'Otak & Regulator Kebijakan Daerah',
      color: C.blue,
      icon: 'building',
      chip: 'OTORITAS KEBIJAKAN',
      points: [
        'Merumuskan arah kebijakan riset & inovasi daerah',
        'Penyusunan kajian teknokratik & roadmap strategis',
        'Orkestrasi ekosistem Iptek antar-OPD & perguruan tinggi',
        'Monitoring & evaluasi akuntabilitas dampak sosial-ekonomi'
      ],
      output: 'OUTPUT: Regulasi, Roadmap, & Pengawasan Kebijakan'
    },
    {
      kicker: '02 / PROSES',
      title: 'UPTD KST SOLO TECHNOPARK',
      subtitle: 'Tangan Eksekutor & Operator Bisnis BLUD',
      color: C.amber,
      icon: 'wrench',
      chip: 'OPERATOR BISNIS SEHAT',
      points: [
        'Penyelenggaraan diklat vokasi & sertifikasi industri',
        'Inkubasi bisnis, alih teknologi & komersialisasi riset',
        'Uji coba prototipe lab presisi & penetrant testing (NDT)',
        'Pengelolaan portofolio bisnis BLUD mandiri & kemitraan'
      ],
      output: 'OUTPUT: Layanan Teknis, Inkubasi, & Hilirisasi Produk'
    },
    {
      kicker: '03 / OUTCOME',
      title: 'WARGA SURAKARTA & INDUSTRI',
      subtitle: 'Penerima Manfaat Utama (Benefisiari)',
      color: C.emerald,
      icon: 'users',
      chip: 'BENEFISIARI DAMPAK',
      points: [
        'SDM lokal kompeten & tersertifikasi kerja industri global',
        'UMKM & startup teknologi tumbuh mandiri omzetnya',
        'Produk inovasi kampus terserap pasar & e-Katalog Pemda',
        'Masuknya investasi swasta & penyerapan 9.000 tenaga kerja'
      ],
      output: 'IMPACT: Lapangan Kerja, Daya Saing, & PAD Mandiri'
    }
  ];
  
  cols.forEach((c, i) => {
    const cX = 0.8 + i * 4.04;
    
    // Card Base
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: colY, w: colW, h: colH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    // Top Color Accent
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: colY, w: colW, h: 0.06,
      fill: { color: c.color }, line: { color: c.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: c.icon, x: cX + 0.25, y: colY + 0.2, size: 0.44, bgTint: C.bgSoft, strokeColor: c.color, borderTint: C.border });
    
    // Kicker
    slide.addText(c.kicker, {
      x: cX + 0.8, y: colY + 0.2, w: colW - 1.0, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: c.color, charSpacing: 1
    });
    
    // Title
    slide.addText(c.title, {
      x: cX + 0.8, y: colY + 0.38, w: colW - 1.0, h: 0.25,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    
    // Subtitle
    slide.addText(c.subtitle, {
      x: cX + 0.25, y: colY + 0.72, w: colW - 0.5, h: 0.22,
      fontFace: C.fontBody, fontSize: 8.5, bold: true, color: c.color
    });
    
    // Divider
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX + 0.25, y: colY + 0.98, w: colW - 0.5, h: 0.01,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    // Points
    c.points.forEach((p, pi) => {
      // Checkmark icon
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', c.color, 2)}`,
        x: cX + 0.25, y: colY + 1.15 + pi * 0.52, w: 0.16, h: 0.16
      });
      slide.addText(p, {
        x: cX + 0.48, y: colY + 1.12 + pi * 0.52, w: colW - 0.75, h: 0.46,
        fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted, lineSpacingMultiple: 1.1
      });
    });
    
    // Output Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX + 0.2, y: colY + 3.45, w: colW - 0.4, h: 0.42,
      fill: { color: C.bgSoft },
      line: { color: c.color, width: 1 }
    });
    slide.addText(c.output, {
      x: cX + 0.25, y: colY + 3.48, w: colW - 0.5, h: 0.36,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: c.color, align: 'center'
    });
    
    // Connector Arrow (between cols)
    if (i < 2) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: cX + colW + 0.11, y: colY + 1.8, w: 0.17, h: 0.3,
        fill: { color: cols[i+1].color }, line: { color: cols[i+1].color, width: 0 }
      });
    }
  });
  
  // Bottom Callout Box: Legal Basis
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.65,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  addIconBadge(slide, { iconName: 'scale', x: 0.95, y: 6.28, size: 0.46, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('LANDASAN HUKUM TATA KELOLA PEMKOT SURAKARTA:\n• Perwali Surakarta No. 15 Tahun 2022: Kedudukan UPTD KST Solo Technopark sebagai unit pelaksana teknis operasional di bawah naungan BRIDA.\n• Perwali Surakarta No. 38 Tahun 2022: Pola Tata Kelola BLUD UPTD KST Solo Technopark yang memberikan fleksibilitas bisnis sehat & akuntabel.', {
    x: 1.55, y: 6.24, w: 10.8, h: 0.58,
    fontFace: C.fontBody, fontSize: 8, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 03: FONDASI REGULASI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // LANDASAN REGULASI & TATA KELOLA',
    title: 'FONDASI HUKUM YANG KOKOH: MANDAT & FLEKSIBILITAS BLUD',
    subtitle: 'Tiga Pilar Regulasi yang Memberi Kewenangan, Fleksibilitas Bisnis Sehat, dan Arah Strategis KST',
    pageNum: 3
  });
  
  const regs = [
    {
      badge: 'PERMENDAGRI 79/2018',
      topic: 'Tata Kelola Keuangan BLUD',
      quote: '"Fleksibilitas BLUD adalah keleluasaan dalam pola pengelolaan keuangan dengan menerapkan praktik bisnis yang sehat untuk meningkatkan layanan kepada masyarakat tanpa mencari keuntungan semata."',
      color: C.blue,
      icon: 'fileText',
      checks: [
        'Keleluasaan bermitra dengan swasta tanpa birokrasi kaku APBD',
        'Mengelola dan memutar pendapatan fungsional jasa secara mandiri',
        'Rekrutmen tenaga ahli & instruktur profesional sesuai kebutuhan pasar'
      ]
    },
    {
      badge: 'PERPRES 106/2017',
      topic: 'Kawasan Sains & Teknologi (KST)',
      quote: '"KST berfungsi sebagai wahana kerja sama riset berkelanjutan antara Pemerintah, PT, dan Industri, wajib menyediakan 4 layanan: Teknis, Teknologi, Inkubasi, & Pendukung."',
      color: C.amber,
      icon: 'target',
      checks: [
        'Mandat nasional fasilitasi hilirisasi & spin-off perusahaan riset',
        'Kewajiban menyediakan 4 pilar layanan iptek terintegrasi',
        'Landasan kolaborasi hexahelix multi-stakeholder formal'
      ]
    },
    {
      badge: 'PERWALI 15 & 38/2022',
      topic: 'Struktur UPTD & Pola Tata Kelola',
      quote: '"Menetapkan struktur UPTD KST Solo Technopark dalam naungan BRIDA Kota Surakarta dengan pola tata kelola BLUD profesional berorientasi layanan publik."',
      color: C.emerald,
      icon: 'scale',
      checks: [
        'Payung hukum operasional tingkat kota yang sah dan mengikat',
        'Pembagian peran teknis yang presisi antara BRIDA dan STP',
        'Dasar legalitas pemanfaatan aset kawasan & kerjasama industri'
      ]
    }
  ];
  
  const rW = 3.65, rH = 4.25, rY = 1.75;
  regs.forEach((r, i) => {
    const rX = 0.8 + i * 4.04;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX, y: rY, w: rW, h: rH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX, y: rY, w: rW, h: 0.06,
      fill: { color: r.color }, line: { color: r.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: r.icon, x: rX + 0.25, y: rY + 0.2, size: 0.44, bgTint: C.bgSoft, strokeColor: r.color, borderTint: C.border });
    
    slide.addText(r.badge, {
      x: rX + 0.8, y: rY + 0.2, w: rW - 1.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 11, bold: true, color: C.textMain
    });
    slide.addText(r.topic, {
      x: rX + 0.8, y: rY + 0.42, w: rW - 1.0, h: 0.18,
      fontFace: C.fontBody, fontSize: 8.5, bold: true, color: r.color
    });
    
    // Quote Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX + 0.2, y: rY + 0.75, w: rW - 0.4, h: 1.45,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(r.quote, {
      x: rX + 0.35, y: rY + 0.82, w: rW - 0.7, h: 1.3,
      fontFace: C.fontBody, fontSize: 8.5, italic: true, color: C.textMuted, lineSpacingMultiple: 1.15
    });
    
    // Checks with vector icon
    r.checks.forEach((ch, ci) => {
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', r.color, 2)}`,
        x: rX + 0.25, y: rY + 2.38 + ci * 0.6, w: 0.18, h: 0.18
      });
      slide.addText(ch, {
        x: rX + 0.52, y: rY + 2.35 + ci * 0.6, w: rW - 0.75, h: 0.52,
        fontFace: C.fontBody, fontSize: 8.5, color: C.textMain, lineSpacingMultiple: 1.15
      });
    });
  });
  
  // Bottom Summary Ribbon
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.18, w: 11.733, h: 0.65,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('KESIMPULAN YURIDIS: Solo Technopark telah memiliki legitimasi hukum lengkap dari tingkat nasional (Perpres & Permendagri) hingga regulasi kepala daerah (Perwali) untuk bermanuver cepat, profesional, dan akuntabel.', {
    x: 1.0, y: 6.25, w: 11.333, h: 0.5,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textWhite, align: 'center', lineSpacingMultiple: 1.2
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 04: EMPAT PILAR LAYANAN TERINTEGRASI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // MANDAT LAYANAN PERPRES 106/2017',
    title: 'EMPAT PILAR LAYANAN TERINTEGRASI: MEMENUHI MANDAT KST NASIONAL',
    subtitle: 'Portofolio Layanan Terintegrasi yang Menjawab Kebutuhan Industri Spesialis & Masyarakat Kota Surakarta',
    pageNum: 4
  });
  
  const pillars = [
    {
      name: 'PILAR 1: LAYANAN TEKNIS (DIKLAT VOKASI INDUSTRI)',
      color: C.blue,
      icon: 'wrench',
      items: [
        { t: 'Mekanik, Otomasi & Desain Manufaktur Presisi', d: 'Pelatihan bersertifikat BNSP untuk operator CNC & permesinan modern.' },
        { t: 'Underwater Wet Welding (Keahlian Langka Nasional)', d: 'Pelatihan las bawah air spesialis industri maritim & migas berstandar global.' },
        { t: 'Welding Intensif (Pengelasan Konstruksi & Pipa)', d: 'Sertifikasi 3G/4G/6G pipa gas dan konstruksi berat untuk proyek strategis.' },
        { t: 'Teknologi Kendaraan Listrik (EV Conversion Lab)', d: 'Pelatihan konversi motor listrik & perakitan baterai EV bersama industri.' },
        { t: 'OGSCI: Oil, Gas, Coal, Shipbuilding & Infra Training', d: 'Pusat diklat konsorsium energi untuk penempatan kerja langsung.' },
        { t: 'Cyber Security & Artificial Intelligence Practitioner', d: 'Kelas intensif talenta digital bekerja sama dengan korporasi teknologi.' }
      ]
    },
    {
      name: 'PILAR 2: PENGEMBANGAN TEKNOLOGI (REKAYASA MANUFAKTUR)',
      color: C.amber,
      icon: 'cpu',
      items: [
        { t: 'Jasa Layanan Rekayasa Industri & Permesinan Presisi', d: 'Fabrikasi cetakan (*moulding*), jig & fixture untuk IKM Surakarta.' },
        { t: 'Non-Destructive Testing: Uji Penetrant Testing (NDT)', d: 'Pengujian kualitas las presisi tanpa merusak struktur material.' },
        { t: 'Produksi Komponen Presisi (Precision Parts Industri)', d: 'Sub-kontrak pembuatan spare part alat berat dan komponen otomotif.' },
        { t: 'Konsultasi Laboratorium Terpadu & Data Center', d: 'Fasilitas hosting data aman dan pendampingan teknis infrastruktur cloud.' },
        { t: 'Pembuatan & Uji Prototipe IKM/UKM Kota Surakarta', d: 'Membantu pelaku usaha lokal merealisasikan purwarupa produk teknologi.' },
        { t: 'Standarisasi Mutu & Sertifikasi TKDN / SNI Produk', d: 'Pendampingan pemenuhan sertifikasi kelayakan pasar & tingkat lokal.' }
      ]
    },
    {
      name: 'PILAR 3: LAYANAN INKUBASI (HILIRISASI KE PASAR)',
      color: C.emerald,
      icon: 'rocket',
      items: [
        { t: 'Pra-Inkubasi & Inkubasi Bisnis Startup Berbasis Riset', d: 'Akselerasi ide inovatif mahasiswa dan peneliti menjadi entitas bisnis.' },
        { t: 'National Cyber Security Hub & AI Experience Center', d: 'Showcase solusi teknologi masa depan dan pengujian ketahanan siber.' },
        { t: 'Fasilitasi HKI, Paten Produk, & Hak Cipta Inovator', d: 'Pendampingan pendaftaran paten secara cepat bekerja sama dengan Kemenkumham.' },
        { t: 'Digital Technopark, Gaming Hub, & GoTo UMKM Center', d: 'Fasilitas live commerce, game development studio, dan digital onboarding.' },
        { t: 'Hilirisasi Riset Perguruan Tinggi ke Pasar Industri', d: 'Penjembatan paten kampus agar dibeli atau dilisensikan ke industri manufaktur.' },
        { t: 'Co-Working Space, Virtual Office & Akses Permodalan', d: 'Ruang kerja bersama terintegrasi dengan akses angel investor & venture capital.' }
      ]
    },
    {
      name: 'PILAR 4: LAYANAN PENDUKUNG (COMMUNITY & EDUCATION)',
      color: C.purple,
      icon: 'graduationCap',
      items: [
        { t: 'Prakerin Kerja Industri Siswa SMK Se-Solo Raya', d: 'Penyelenggaraan magang kurikulum industri bagi ribuan siswa vokasi.' },
        { t: 'Pemanfaatan Sarana Kawasan, Gedung Expo & MICE', d: 'Penyewaan auditorium, ruang seminar, dan plaza pameran ekonomi kreatif.' },
        { t: 'Tempat Uji Kompetensi (TUK Mandiri) & Sertifikasi JFT', d: 'Uji sertifikasi keahlian teknis mandiri yang diakui kementerian terkait.' },
        { t: 'Kemitraan Komunitas Kreatif & Kunjungan Industri', d: 'Destinasi eduwisata teknologi dan pusat kumpul komunitas inovator muda.' },
        { t: 'Solo Science Center (Pusat Edukasi IPTEK & Wisata Sains)', d: 'Wahana peragaan sains interaktif untuk pelajar sekolah dasar hingga menengah.' },
        { t: 'Pusat Kolaborasi Multipihak Pemkot Surakarta', d: 'Ruang pertemuan strategis bagi OPD, akademisi, dan asosiasi pengusaha.' }
      ]
    }
  ];
  
  const qW = 5.75, qH = 2.15, qY1 = 1.75, qY2 = 4.05;
  const positions = [
    { x: 0.8, y: qY1 },
    { x: 6.78, y: qY1 },
    { x: 0.8, y: qY2 },
    { x: 6.78, y: qY2 }
  ];
  
  pillars.forEach((p, i) => {
    const pos = positions[i];
    
    // Panel
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pos.x, y: pos.y, w: qW, h: qH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    // Top Color Strip
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pos.x, y: pos.y, w: qW, h: 0.05,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: p.icon, x: pos.x + 0.2, y: pos.y + 0.14, size: 0.38, bgTint: C.bgSoft, strokeColor: p.color, borderTint: C.border });
    
    // Title
    slide.addText(p.name, {
      x: pos.x + 0.68, y: pos.y + 0.18, w: qW - 0.8, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: p.color
    });
    
    // 2-Column Items inside Quadrant
    p.items.forEach((it, idx) => {
      const col = idx % 2;
      const row = Math.floor(idx / 2);
      const itX = pos.x + 0.2 + col * 2.75;
      const itY = pos.y + 0.52 + row * 0.5;
      
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', p.color, 2)}`,
        x: itX, y: itY + 0.02, w: 0.13, h: 0.13
      });
      slide.addText(it.t, {
        x: itX + 0.18, y: itY, w: 2.5, h: 0.44,
        fontFace: C.fontBody, fontSize: 7.5, color: C.textMain, lineSpacingMultiple: 1.05
      });
    });
  });
  
  // Bottom Callout
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.32, w: 11.733, h: 0.5,
    fill: { color: C.bgTintBlue },
    line: { color: C.borderBlue, width: 1 }
  });
  slide.addText('"BLUD menggeser paradigma: Dari sekadar pengelolaan kawasan fisik menjadi penyedia solusi teknologi yang responsif terhadap kebutuhan industri dan masyarakat."', {
    x: 1.0, y: 6.38, w: 11.333, h: 0.38,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, italic: true, color: C.blue, align: 'center'
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 05: PIPELINE HILIRISASI RISET KAMPUS
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // MEKANISME HILIRISASI INOVASI',
    title: 'PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR NYATA',
    subtitle: 'Solo Technopark sebagai Jembatan: Menghubungkan Riset Kampus Menuju Produk Bernilai Ekonomi bagi Industri',
    pageNum: 5
  });
  
  // 5 TAHAP PIPELINE (CHEVRON & ICON BADGES)
  const stages = [
    { no: 'TAHAP 01', name: 'RISET & INVENSI', sub: 'Kampus & Peneliti', desc: 'Invensi UNS, UMS, ISI, Poltek Solo Raya. Riset dosen & mahasiswa.', color: C.sky, icon: 'flask' },
    { no: 'TAHAP 02', name: 'SELEKSI & KURASI', sub: 'STP & BRIDA', desc: 'Kurasi komersial & relevansi kebutuhan industri serta problem kota Solo.', color: C.blue, icon: 'target' },
    { no: 'TAHAP 03', name: 'PROTOTYPING', sub: 'Lab Presisi STP', desc: 'Uji teknis, penetrant testing, validasi mutu produk & SNI/TKDN.', color: C.amber, icon: 'layers' },
    { no: 'TAHAP 04', name: 'INKUBASI BISNIS', sub: 'Inkubator STP', desc: 'Penyusunan business model, proteksi HKI/Paten, akses modal & offtaker.', color: C.emerald, icon: 'briefcase' },
    { no: 'TAHAP 05', name: 'KOMERSIALISASI', sub: 'Pasar & Industri', desc: 'Penetrasi e-Katalog Lokal OPD, captive market pemda & industri nasional.', color: C.purple, icon: 'shoppingCart' }
  ];
  
  const stY = 1.75, stW = 2.18, stH = 2.45, stGap = 0.2;
  stages.forEach((st, i) => {
    const stX = 0.8 + i * (stW + stGap);
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: stX, y: stY, w: stW, h: stH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: stX, y: stY, w: stW, h: 0.06,
      fill: { color: st.color }, line: { color: st.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: st.icon, x: stX + 0.15, y: stY + 0.15, size: 0.38, bgTint: C.bgSoft, strokeColor: st.color, borderTint: C.border });
    
    slide.addText(st.no, {
      x: stX + 0.6, y: stY + 0.16, w: stW - 0.7, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: st.color, charSpacing: 1
    });
    
    slide.addText(st.name, {
      x: stX + 0.15, y: stY + 0.6, w: stW - 0.3, h: 0.25,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textMain
    });
    slide.addText(st.sub, {
      x: stX + 0.15, y: stY + 0.88, w: stW - 0.3, h: 0.18,
      fontFace: C.fontBody, fontSize: 8, italic: true, color: C.textDim
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: stX + 0.15, y: stY + 1.12, w: stW - 0.3, h: 0.01,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    slide.addText(st.desc, {
      x: stX + 0.15, y: stY + 1.22, w: stW - 0.3, h: 1.1,
      fontFace: C.fontBody, fontSize: 8, color: C.textMuted, lineSpacingMultiple: 1.15
    });
    
    if (i < 4) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: stX + stW + 0.06, y: stY + 1.1, w: 0.12, h: 0.22,
        fill: { color: C.borderDark }, line: { color: C.borderDark, width: 0 }
      });
    }
  });
  
  // BAGIAN BAWAH: TABEL PRODUK RIIL (DENGAN STATUS CHIPS)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 4.4, w: 6.9, h: 2.45,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 4.4, w: 6.9, h: 0.05,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  slide.addText('CONTOH TEKNOLOGI YANG DIHILIRKAN & DIINKUBASI DI STP (DATA RIIL)', {
    x: 1.0, y: 4.52, w: 6.5, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy
  });
  
  const pRows = [
    [
      { text: 'KLASTER TEKNOLOGI', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } },
      { text: 'CONTOH PRODUK INOVASI', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } },
      { text: 'STATUS TAHAPAN', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft }, align: 'center' } }
    ],
    ['Smart Agriculture IoT', 'Mesin Sortasi Biji Kopi Otomatis', 'Prototype Validasi'],
    ['Smart Agriculture IoT', 'Teknologi Pertanian Presisi IoT & AI', 'Uji Lapangan'],
    ['Smart Agriculture IoT', 'Mesin Tetas Telur Pintar Mandiri', 'Siap e-Katalog'],
    ['Deep Tech Mobility', 'Autonomous Vehicle Elektrik Kampus', 'R&D Bersama'],
    ['GovTech & Pelatihan', 'VR Pelatihan Industri & Chatbot AI OPD', 'Komersial Aktif']
  ];
  
  const statusStyles = [
    null,
    { text: 'PROTOTYPE VALIDASI', type: 'blue' },
    { text: 'UJI LAPANGAN', type: 'amber' },
    { text: 'SIAP e-KATALOG', type: 'green' },
    { text: 'R&D BERSAMA', type: 'purple' },
    { text: 'KOMERSIAL AKTIF', type: 'green' }
  ];
  
  const tableData = pRows.map((row, rIdx) => {
    if (rIdx === 0) return row;
    const st = statusStyles[rIdx];
    return [
      { text: row[0], options: { fontSize: 7.5, color: C.textDim } },
      { text: row[1], options: { fontSize: 7.5, bold: true, color: C.textMain } },
      { text: row[2], options: { fontSize: 7.5, bold: true, color: st.type === 'green' ? C.emerald : (st.type === 'amber' ? C.amber : C.blue), align: 'center' } }
    ];
  });
  
  slide.addTable(tableData, {
    x: 0.95, y: 4.8, w: 6.6,
    colW: [1.8, 3.2, 1.6],
    rowH: [0.26, 0.24, 0.24, 0.24, 0.24, 0.24],
    border: { pt: 0.5, color: C.border },
    autoPage: false
  });
  
  // Right Box: Ekosistem Fasilitas Pendukung
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 7.9, y: 4.4, w: 4.633, h: 2.45,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 7.9, y: 4.4, w: 4.633, h: 0.05,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  
  slide.addText('EKOSISTEM FASILITAS PENDUKUNG AKTIF', {
    x: 8.1, y: 4.52, w: 4.2, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber
  });
  
  const facil = [
    { title: 'National Cyber Security Hub', desc: 'Pusat ketahanan digital & talenta siber nasional bersama BSSN.', icon: 'shieldCheck' },
    { title: 'Gaming Hub & Digital Studio', desc: 'Pengembangan game, animasi, & konten kreatif e-sport.', icon: 'cpu' },
    { title: 'AI Experience Center', desc: 'Laboratorium kecerdasan buatan terapan industri & showcase solusi AI.', icon: 'zap' },
    { title: 'GoTo UMKM Center', desc: 'Akselerasi digitalisasi, kurasi produk, & onboarding pedagang lokal.', icon: 'shoppingCart' },
    { title: 'Laboratorium Rekayasa Presisi', desc: 'Fasilitas fabrikasi mesin, permesinan presisi & pengujian NDT industri.', icon: 'wrench' }
  ];
  
  facil.forEach((f, fi) => {
    const fY = 4.8 + fi * 0.38;
    slide.addImage({
      data: `image/svg+xml;base64,${getSvgBase64(f.icon, C.blue, 2)}`,
      x: 8.1, y: fY + 0.03, w: 0.16, h: 0.16
    });
    slide.addText(f.title, {
      x: 8.35, y: fY, w: 4.0, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.textMain
    });
    slide.addText(f.desc, {
      x: 8.35, y: fY + 0.17, w: 4.0, h: 0.18,
      fontFace: C.fontBody, fontSize: 7, color: C.textDim
    });
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 06: EKOSISTEM HEXAHELIX
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // JEJARING KEMITRAAN STRATEGIS',
    title: 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD',
    subtitle: 'Kepercayaan Korporasi Global & Nasional yang Menanamkan Investasi Langsung di Jantung Kota Surakarta',
    pageNum: 6
  });
  
  // Left Box: Donut Chart Hexahelix
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 4.7, h: 5.1,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 4.7, h: 0.06,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  // Header with vector icon
  addIconBadge(slide, { iconName: 'pieChart', x: 1.0, y: 1.9, size: 0.36, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('KOMPOSISI 6 AKTOR HEXAHELIX (95+ MITRA)', {
    x: 1.45, y: 1.95, w: 3.8, h: 0.25,
    fontFace: C.fontTitle, fontSize: 9, bold: true, color: C.navy
  });
  
  // Native Donut Chart
  const chartData = [
    {
      name: 'Aktor Hexahelix',
      labels: ['Industri & Bisnis', 'Akademisi / Kampus', 'Komunitas & UMKM', 'Pemerintah / BUMN', 'Media Massa', 'Lembaga Global'],
      values: [35, 20, 15, 12, 8, 5]
    }
  ];
  
  slide.addChart(pptx.charts.DOUGHNUT, chartData, {
    x: 1.0, y: 2.3, w: 4.3, h: 2.8,
    holeSize: 58,
    showLabel: false,
    showPercent: false,
    showValue: false,
    showLegend: false,
    chartColors: [C.sky, C.navy, C.emerald, C.amber, C.purple, 'F43F5E']
  });
  
  // Custom Clean Legend below Donut
  const legends = [
    { label: 'Industri & Bisnis: 37% (35)', color: C.sky },
    { label: 'Pemerintah / BUMN: 13% (12)', color: C.amber },
    { label: 'Akademisi / Kampus: 21% (20)', color: C.navy },
    { label: 'Media Massa: 8% (8)', color: C.purple },
    { label: 'Komunitas & UMKM: 16% (15)', color: C.emerald },
    { label: 'Lembaga Global: 5% (5)', color: 'F43F5E' }
  ];
  
  legends.forEach((lg, lIdx) => {
    const col = lIdx % 2;
    const row = Math.floor(lIdx / 2);
    const lX = 1.0 + col * 2.3;
    const lY = 5.2 + row * 0.28;
    
    slide.addShape(pptx.shapes.OVAL, {
      x: lX, y: lY + 0.03, w: 0.12, h: 0.12,
      fill: { color: lg.color }, line: { color: lg.color, width: 0 }
    });
    slide.addText(lg.label, {
      x: lX + 0.18, y: lY, w: 2.1, h: 0.22,
      fontFace: C.fontBody, fontSize: 7.5, color: C.textMain
    });
  });
  
  // Callout Box inside Left Panel
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 6.05, w: 4.3, h: 0.65,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  slide.addText('TOTAL: 95+ MITRA STRATEGIS AKTIF\n• 100% Investasi Swasta & PPP (0% Beban APBD)\n• Akses fasilitas 5G, Cyber Security, AI, & Migas\n• Saluran penyerapan kerja langsung bagi warga Solo', {
    x: 1.15, y: 6.1, w: 4.0, h: 0.55,
    fontFace: C.fontBody, fontSize: 7, color: C.textDim, lineSpacingMultiple: 1.15
  });
  
  // Right Box: Partner Table
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.7, y: 1.75, w: 6.833, h: 5.1,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.7, y: 1.75, w: 6.833, h: 0.06,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'award', x: 5.9, y: 1.9, size: 0.36, bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber });
  slide.addText('MITRA INDUSTRI KELAS DUNIA YANG AKTIF BERINVESTASI DI STP', {
    x: 6.35, y: 1.95, w: 6.0, h: 0.25,
    fontFace: C.fontTitle, fontSize: 9, bold: true, color: C.amber
  });
  
  const partners = [
    [
      { text: 'SEKTOR INDUSTRI', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } },
      { text: 'MITRA STRATEGIS', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } },
      { text: 'KONTRIBUSI NYATA BAGI KAWASAN STP', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } }
    ],
    ['E-Commerce & Logistik', 'Shopee Indonesia', 'Infrastruktur UMKM Hub, Co-location, sewa lahan Rp 544 jt/th'],
    ['Gaming & Kreatif', 'Garena', 'Gaming studio, beasiswa talenta digital, turnamen esports'],
    ['Perbankan & Fintech', 'Bank Mandiri', 'Mandiri Innovation Hub, fasilitasi modal startup binaan'],
    ['Telekomunikasi', 'Indosat Ooredoo Hutchison', 'Laboratorium 5G, konektivitas fiber optik terintegrasi kawasan'],
    ['Ride-Hailing & UMKM', 'GoTo (Gojek-Tokopedia)', 'GoTo UMKM Center, pelatihan onboarding merchant lokal'],
    ['Migas & Industri Berat', 'OGSCI Consortium', 'Pelatihan migas internasional: Target Rp 6,25 M (2030)']
  ];
  
  const tablePart = partners.map((r, i) => {
    if (i === 0) return r;
    const isOgsci = i === 6;
    return [
      { text: r[0], options: { fontSize: 7.5, color: C.textDim } },
      { text: r[1], options: { fontSize: 7.5, bold: true, color: isOgsci ? C.amber : C.textMain } },
      { text: r[2], options: { fontSize: 7.5, color: isOgsci ? C.amber : C.textMuted, bold: isOgsci } }
    ];
  });
  
  slide.addTable(tablePart, {
    x: 5.9, y: 2.35, w: 6.433,
    colW: [1.7, 1.8, 2.933],
    rowH: [0.28, 0.38, 0.38, 0.38, 0.38, 0.38, 0.38],
    border: { pt: 0.5, color: C.border },
    autoPage: false
  });
  
  // Right Bottom: Zero APBD Banner
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.9, y: 5.55, w: 6.433, h: 1.15,
    fill: { color: C.bgTintGreen },
    line: { color: C.borderGreen, width: 1 }
  });
  addIconBadge(slide, { iconName: 'shieldCheck', x: 6.1, y: 5.75, size: 0.55, bgTint: C.bg, strokeColor: C.emerald, borderTint: C.borderGreen });
  slide.addText('PRINSIP ZERO APBD BURDEN: SELURUH FASILITAS CANGGIH DIBANGUN OLEH SWASTA', {
    x: 6.8, y: 5.68, w: 5.4, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.emerald
  });
  slide.addText('Laboratorium AI, Studio Gaming, 5G Experience, hingga Cyber Security Hub tidak dibangun dengan membebani belanja modal APBD Kota Surakarta. Semuanya hadir melalui kemitraan murni Public-Private Partnership (PPP).', {
    x: 6.8, y: 5.94, w: 5.4, h: 0.65,
    fontFace: C.fontBody, fontSize: 7.5, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 07: DIAGNOSA 5 GAP & SOLUSI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // REFLEKSI STRATEGIS & SOLUSI KONKRET',
    title: 'DIAGNOSA & 5 GAP UTAMA BESERTA SOLUSI TEROBOSANNYA',
    subtitle: 'Kajian Kritis Berdasarkan Dokumen Strategi STP 2026: Gap Riil, Risiko, dan Solusi Kebijakan Kepala Daerah',
    pageNum: 7
  });
  
  const gaps = [
    {
      num: 'GAP 01',
      title: 'Jenis Pelatihan Masih Terbatas',
      cause: 'Keterbatasan instruktur ahli internal & sarana diklat berbasis industri masa depan.',
      solTitle: 'Gandeng Korporasi Global (OGSCI, Kemenperin) Buka Kelas Spesialis',
      solDesc: 'Welder Bawah Air, AI Engineer, Teknisi Baterai EV, & Drone Operator.',
      color: C.blue
    },
    {
      num: 'GAP 02',
      title: 'R&D Belum Mengarah Hilirisasi',
      cause: 'Riset kampus berjalan terisolasi, kurang terpetakan, dan belum berorientasi pasar.',
      solTitle: 'Bentuk Technology Transfer Office (TTO) & Kurasi Riset Terapan Kampus',
      solDesc: 'Konsorsium 5 Perguruan Tinggi Solo Raya (UNS, UMS, ISI, dll.) berbasis problem kota.',
      color: C.blue
    },
    {
      num: 'GAP 03',
      title: 'Investasi Kawasan Terbatas',
      cause: 'Lahan kosong belum optimal dikerjasamakan; promosi investor belum sistematis.',
      solTitle: 'Zonasi Investasi Teknologi Terpadu & Long-Term Revenue Sharing',
      solDesc: 'Skema kemitraan 10–25 tahun bagi korporasi teknologi global tanpa beban belanja APBD.',
      color: C.blue
    },
    {
      num: 'GAP 04',
      title: 'Kemandirian Finansial Rendah',
      cause: 'Pendapatan masih bertumpu sewa lahan; hardware lab belum optimal menghasilkan jasa.',
      solTitle: 'Diversifikasi 6 Revenue Stream Mandiri & Optimalisasi Aset Industri',
      solDesc: 'Uji NDT penetrant, TUK/JFT mandiri, rekayasa manufaktur, co-working, & startup binaan.',
      color: C.emerald
    },
    {
      num: 'GAP 05',
      title: 'Promosi & Penyerapan Rendah',
      cause: 'Networking kurang ter-maintenance; belum ada regulasi penyerapan inovasi lokal.',
      solTitle: 'Penerbitan Instruksi Walikota untuk e-Katalog Lokal Afirmatif',
      solDesc: 'Wajibkan OPD Pemkot menyerap produk/layanan tenant binaan STP (Target: Rp 2,25 Miliar).',
      color: C.amber,
      isPriority: true
    }
  ];
  
  const startY = 1.7, rowH = 0.85, rowGap = 0.12;
  gaps.forEach((g, idx) => {
    const rY = startY + idx * (rowH + rowGap);
    
    // Left Box: Problem
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y: rY, w: 4.8, h: rowH,
      fill: { color: g.isPriority ? C.bgTintAmber : C.bg },
      line: { color: g.isPriority ? C.amber : C.border, width: 1 }
    });
    
    // Icon Badge (Problem)
    addIconBadge(slide, { iconName: 'alertTriangle', x: 0.95, y: rY + 0.14, size: 0.32, bgTint: C.bgTintRed, strokeColor: C.red, borderTint: C.borderRed });
    
    slide.addText(g.num, {
      x: 1.35, y: rY + 0.16, w: 0.7, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.red
    });
    
    slide.addText(g.title, {
      x: 2.1, y: rY + 0.16, w: 3.4, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.textMain
    });
    
    slide.addText(`Akar Masalah: ${g.cause}`, {
      x: 0.95, y: rY + 0.46, w: 4.5, h: 0.32,
      fontFace: C.fontBody, fontSize: 7.5, color: C.textDim, lineSpacingMultiple: 1.1
    });
    
    // Arrow Connector
    slide.addShape(pptx.shapes.RIGHT_ARROW, {
      x: 5.75, y: rY + 0.28, w: 0.25, h: 0.24,
      fill: { color: g.isPriority ? C.amber : C.blue },
      line: { color: g.isPriority ? C.amber : C.blue, width: 0 }
    });
    
    // Right Box: Solution
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.15, y: rY, w: 6.383, h: rowH,
      fill: { color: g.isPriority ? C.bgTintAmber : C.bg },
      line: { color: g.isPriority ? C.amber : C.border, width: 1 }
    });
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.15, y: rY, w: 6.383, h: 0.04,
      fill: { color: g.isPriority ? C.amber : g.color }, line: { color: g.isPriority ? C.amber : g.color, width: 0 }
    });
    
    // Icon Badge (Solution)
    addIconBadge(slide, { iconName: g.isPriority ? 'zap' : 'checkCircle', x: 6.3, y: rY + 0.12, size: 0.32, bgTint: g.isPriority ? C.bgTintAmber : C.bgTintGreen, strokeColor: g.isPriority ? C.amber : C.emerald, borderTint: g.isPriority ? C.borderAmber : C.borderGreen });
    
    slide.addText(g.solTitle, {
      x: 6.7, y: rY + 0.14, w: 5.6, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: g.isPriority ? C.amber : g.color
    });
    
    slide.addText(g.solDesc, {
      x: 6.7, y: rY + 0.44, w: 5.6, h: 0.34,
      fontFace: C.fontBody, fontSize: 7.5, color: C.textMuted, lineSpacingMultiple: 1.1
    });
  });
  
  // Bottom Priority Callout
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.3, w: 11.733, h: 0.58,
    fill: { color: C.bgTintAmber },
    line: { color: C.amber, width: 1 }
  });
  addIconBadge(slide, { iconName: 'zap', x: 1.0, y: 6.38, size: 0.42, bgTint: C.bg, strokeColor: C.amber, borderTint: C.amber });
  slide.addText('PRIORITAS UTAMA #5: INSTRUKSI WALIKOTA e-KATALOG LOKAL AFIRMATIF (QUICK WIN 30 HARI)', {
    x: 1.55, y: 6.36, w: 10.8, h: 0.2,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.amber
  });
  slide.addText('Target: Membuka captive market Rp 2,25 Miliar serapan belanja OPD Pemkot Surakarta di 2030 bagi 80+ produk tenant dan startup binaan STP tanpa perlu alokasi pagu belanja baru di APBD. Hanya butuh 1 lembar Instruksi Walikota.', {
    x: 1.55, y: 6.58, w: 10.8, h: 0.26,
    fontFace: C.fontBody, fontSize: 7.5, color: C.textMuted
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 08: ROADMAP 3 FASE TRANSFORMASI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // PETA JALAN STRATEGIS KELEMBAGAAN',
    title: 'ROADMAP 3 FASE TRANSFORMASI: DARI CENTRE MENUJU HOLDING INOVASI DAERAH',
    subtitle: 'Milestone Terukur dan Bertahap Menuju Kemandirian Finansial Penuh Tanpa Subsidi Belanja APBD 2030',
    pageNum: 8
  });
  
  const phases = [
    {
      num: 'FASE 1 (2026–2027)',
      title: 'CENTRE OF INNOVATION & CONNECTIVITY',
      target: 'Target: Rp 7,50 M ➔ Rp 12,53 M (+67%)',
      color: C.blue,
      icon: 'sparkles',
      chip: 'PONDASI & AKREDITASI',
      items: [
        'Pondasi tata kelola kelembagaan BLUD akuntabel',
        'Akreditasi laboratorium industri & standarisasi uji',
        'Sertifikasi mutu diklat vokasi & kemitraan OGSCI',
        'Pembentukan Konsorsium Riset 5 Kampus Solo Raya',
        'Aktivasi TUK Mandiri & program SMK GoGlobal'
      ]
    },
    {
      num: 'FASE 2 (2028–2029)',
      title: 'ENTREPRENEUR TECHNOPARK (SBU MANDIRI)',
      target: 'Target: Rp 12,53 M ➔ Rp 20,62 M (+65%)',
      color: C.amber,
      icon: 'building',
      chip: 'KOMERSIALISASI SBU',
      items: [
        'Pembentukan Strategic Business Unit (SBU) profesional',
        'Hilirisasi 15–20 produk inovasi komersial ber-TKDN',
        'Strategi corporate marketing & penetrasi pasar nasional',
        'Ekspansi revenue OGSCI Migas mencapai Rp 5,5 Miliar',
        'Operasionalisasi Co-Working & Virtual Office penuh'
      ]
    },
    {
      num: 'FASE 3 (2029–2030)',
      title: 'GLOBAL MARKET POSITIONING & HOLDING',
      target: 'Target: Rp 20,62 M ➔ Rp 23,47 M (MANDIRI)',
      color: C.emerald,
      icon: 'award',
      chip: 'HOLDING & GLOBAL REACH',
      items: [
        'Kemandirian finansial penuh: 0% subsidi belanja APBD',
        'Operasional holding company inovasi daerah pertama',
        'Pusat pengembangan Indonesia Digital Technopark',
        'Realisasi target pendapatan mandiri Rp 23,47 Miliar',
        'Ekspansi jejaring kemitraan dan penempatan global'
      ]
    }
  ];
  
  const phW = 3.65, phH = 3.9, phY = 1.75;
  phases.forEach((p, i) => {
    const phX = 0.8 + i * 4.04;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX, y: phY, w: phW, h: phH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX, y: phY, w: phW, h: 0.06,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: p.icon, x: phX + 0.25, y: phY + 0.2, size: 0.44, bgTint: C.bgSoft, strokeColor: p.color, borderTint: C.border });
    
    slide.addText(p.num, {
      x: phX + 0.8, y: phY + 0.2, w: phW - 1.0, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: p.color, charSpacing: 1
    });
    
    slide.addText(p.title, {
      x: phX + 0.8, y: phY + 0.38, w: phW - 1.0, h: 0.4,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textMain
    });
    
    // Target Ribbon
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX + 0.2, y: phY + 0.85, w: phW - 0.4, h: 0.32,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(p.target, {
      x: phX + 0.2, y: phY + 0.88, w: phW - 0.4, h: 0.26,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: p.color, align: 'center'
    });
    
    // Checklist Items
    p.items.forEach((it, idx) => {
      const itY = phY + 1.35 + idx * 0.48;
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', p.color, 2)}`,
        x: phX + 0.25, y: itY + 0.02, w: 0.16, h: 0.16
      });
      slide.addText(it, {
        x: phX + 0.48, y: itY, w: phW - 0.75, h: 0.44,
        fontFace: C.fontBody, fontSize: 8, color: C.textMuted, lineSpacingMultiple: 1.1
      });
    });
    
    // Arrow between phases
    if (i < 2) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: phX + phW + 0.11, y: phY + 1.6, w: 0.17, h: 0.28,
        fill: { color: phases[i+1].color }, line: { color: phases[i+1].color, width: 0 }
      });
    }
  });
  
  // Bottom Table: Yearly Targets
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 5.82, w: 11.733, h: 1.0,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  
  slide.addText('RINGKASAN TARGET TAHUNAN MENUJU KEMANDIRIAN 2030 (AUDIT FINANSIAL RIIL):', {
    x: 1.0, y: 5.9, w: 11.333, h: 0.2,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.navy
  });
  
  const yrTable = [
    [
      { text: 'TAHUN 2026 (Baseline)', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft }, align: 'center' } },
      { text: 'TAHUN 2027 (+67,0%)', options: { bold: true, fontSize: 7.5, color: C.blue, fill: { color: C.bgSoft }, align: 'center' } },
      { text: 'TAHUN 2028 (+28,5%)', options: { bold: true, fontSize: 7.5, color: C.blue, fill: { color: C.bgSoft }, align: 'center' } },
      { text: 'TAHUN 2029 (+28,0%)', options: { bold: true, fontSize: 7.5, color: C.blue, fill: { color: C.bgSoft }, align: 'center' } },
      { text: 'TAHUN 2030 (MANDIRI)', options: { bold: true, fontSize: 7.5, color: C.amber, fill: { color: C.bgTintAmber }, align: 'center' } }
    ],
    [
      { text: 'Rp 7.503.392.698', options: { fontSize: 8, color: C.textDim, align: 'center' } },
      { text: 'Rp 12.532.185.031', options: { fontSize: 8, bold: true, color: C.blue, align: 'center' } },
      { text: 'Rp 16.100.658.831', options: { fontSize: 8, bold: true, color: C.blue, align: 'center' } },
      { text: 'Rp 20.615.658.831', options: { fontSize: 8, bold: true, color: C.blue, align: 'center' } },
      { text: 'Rp 23.468.158.831', options: { fontSize: 8.5, bold: true, color: C.amber, fill: { color: C.bgTintAmber }, align: 'center' } }
    ]
  ];
  
  slide.addTable(yrTable, {
    x: 0.95, y: 6.15, w: 11.433,
    colW: [2.28, 2.28, 2.28, 2.28, 2.313],
    rowH: [0.26, 0.28],
    border: { pt: 0.5, color: C.border },
    autoPage: false
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 09: RE-SETTING TARGET FINANSIAL
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // PROYEKSI KEUANGAN BLUD 2026–2030',
    title: 'RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD +212,8%',
    subtitle: 'Proyeksi Pendapatan Mandiri Riil dari Rp 7,50 Miliar (2026) menuju Rp 23,47 Miliar (2030) — Tanpa Subsidi APBD',
    pageNum: 9
  });
  
  // Left Box: Trend Column Chart
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 5.7, h: 5.1,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 5.7, h: 0.06,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'trendingUp', x: 1.0, y: 1.9, size: 0.36, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('TREN PERTUMBUHAN PENDAPATAN BLUD 2026–2030 (DALAM RP MILIAR)', {
    x: 1.45, y: 1.95, w: 4.8, h: 0.25,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy
  });
  
  const colChartData = [
    {
      name: 'Target Pendapatan (Rp M)',
      labels: ['2026\n(Baseline)', '2027\n(+67%)', '2028\n(+28,5%)', '2029\n(+28%)', '2030\n(MANDIRI)'],
      values: [7.50, 12.53, 16.10, 20.62, 23.47]
    }
  ];
  
  slide.addChart(pptx.charts.BAR, colChartData, {
    x: 1.0, y: 2.35, w: 5.3, h: 2.8,
    barDir: 'col',
    showLabel: true,
    showValue: true,
    showLegend: false,
    catAxisLabelColor: C.textMain,
    valAxisLabelColor: C.textDim,
    valGridLine: { color: C.border, width: 0.5 },
    chartColors: [C.sky, C.blue, C.navy, C.emerald, C.amber]
  });
  
  // Callout under chart
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 5.3, w: 5.3, h: 1.4,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  addIconBadge(slide, { iconName: 'award', x: 1.15, y: 5.42, size: 0.36, bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber });
  slide.addText('LOMPATAN FINANSIAL: +212,8% (3,1× LIPAT DALAM 4 TAHUN)', {
    x: 1.6, y: 5.45, w: 4.5, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.amber
  });
  slide.addText('• Mandiri Total Tanpa APBD di 2030: Biaya operasional UPTD STP tertutup 100% dari jasa layanan mandiri.\n• Titik Impas (BEP) di Tahun 2028: Tercapai di angka Rp 16,10 Miliar seiring beroperasinya SBU manufaktur.\n• Kualitas Pendapatan Berkelanjutan: 88% pendapatan disumbang oleh riset, vokasi industri, & startup binaan.', {
    x: 1.15, y: 5.75, w: 5.0, h: 0.9,
    fontFace: C.fontBody, fontSize: 7.5, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  // Right Box: Doughnut Revenue 2030
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.833, y: 1.75, w: 5.7, h: 5.1,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.833, y: 1.75, w: 5.7, h: 0.06,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'pieChart', x: 7.05, y: 1.9, size: 0.36, bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber });
  slide.addText('STRUKTUR REVENUE STREAM 2030 (RP 23,47 M)', {
    x: 7.5, y: 1.95, w: 4.8, h: 0.25,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber
  });
  
  const revDonut = [
    {
      name: 'Revenue 2030',
      labels: ['Program Kemitraan', 'Pelatihan & Sertifikasi', 'Diklat Mandiri Spesialis', 'Lain-lain BLUD / MICE', 'Rekayasa Manufaktur', 'Kemitraan Lahan Eksisting'],
      values: [12.825, 3.680, 2.469, 2.450, 1.465, 0.579]
    }
  ];
  
  slide.addChart(pptx.charts.DOUGHNUT, revDonut, {
    x: 7.05, y: 2.35, w: 5.3, h: 2.2,
    holeSize: 55,
    showLabel: false,
    showValue: false,
    showLegend: false,
    chartColors: [C.sky, C.navy, C.amber, C.emerald, C.purple, C.textDim]
  });
  
  // Donut Custom Legend
  const revLegs = [
    { l: 'Program Kemitraan: 54,6% (Rp 12,83 M)', c: C.sky },
    { l: 'Lain-lain BLUD / MICE: 10,4% (Rp 2,45 M)', c: C.emerald },
    { l: 'Pelatihan & Sertifikasi: 15,7% (Rp 3,68 M)', c: C.navy },
    { l: 'Rekayasa Manufaktur: 6,2% (Rp 1,47 M)', c: C.purple },
    { l: 'Diklat Mandiri Spesialis: 10,5% (Rp 2,47 M)', c: C.amber },
    { l: 'Kemitraan Lahan Eksisting: 2,5% (Rp 0,58 M)', c: C.textDim }
  ];
  
  revLegs.forEach((rl, rIdx) => {
    const col = rIdx % 2;
    const row = Math.floor(rIdx / 2);
    const rX = 7.05 + col * 2.75;
    const rY = 4.65 + row * 0.28;
    
    slide.addShape(pptx.shapes.OVAL, {
      x: rX, y: rY + 0.03, w: 0.12, h: 0.12,
      fill: { color: rl.c }, line: { color: rl.c, width: 0 }
    });
    slide.addText(rl.l, {
      x: rX + 0.18, y: rY, w: 2.5, h: 0.22,
      fontFace: C.fontBody, fontSize: 7, color: C.textMain
    });
  });
  
  // Right Bottom: Top Catalysts
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 7.05, y: 5.55, w: 5.3, h: 1.15,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  slide.addText('TOP 5 KATALISATOR PENDAPATAN 2030:', {
    x: 7.2, y: 5.62, w: 5.0, h: 0.18,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue
  });
  
  const catItems = [
    { n: '#1 OGSCI Program Migas', v: 'Rp 6.250.000.000', c: C.amber },
    { n: '#2 Pemda e-Katalog Inovasi', v: 'Rp 2.250.000.000', c: C.blue },
    { n: '#3 Startup Binaan STP', v: 'Rp 2.000.000.000', c: C.emerald },
    { n: '#4 Sertifikasi Tenaga JFT', v: 'Rp 1.375.000.000', c: C.textMain },
    { n: '#5 Tempat Uji Kompetensi', v: 'Rp 1.350.000.000', c: C.textMain }
  ];
  
  catItems.forEach((ct, ci) => {
    const cY = 5.82 + ci * 0.16;
    slide.addText(ct.n, {
      x: 7.2, y: cY, w: 3.2, h: 0.16,
      fontFace: C.fontBody, fontSize: 7, color: C.textMuted
    });
    slide.addText(ct.v, {
      x: 10.4, y: cY, w: 1.8, h: 0.16,
      fontFace: C.fontTitle, fontSize: 7, bold: true, color: ct.c, align: 'right'
    });
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 10: ANATOMY MESIN PENDAPATAN 2030
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // TRANSPARANSI ANGGARAN & STRUKTUR PENDAPATAN',
    title: 'ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN',
    subtitle: 'Transparansi Total Hasil Rekonsiliasi: Dari Mana Setiap Rupiah dari Target Rp 23,47 Miliar Berasal',
    pageNum: 10
  });
  
  // TOP CONTRIBUTION RIBBON
  const ribbonData = [
    { name: 'Klaster A: Kemitraan Strategis (54,6%)', w: 6.4, color: C.sky },
    { name: 'B: Vokasi (15,7%)', w: 1.84, color: C.navy },
    { name: 'C: Diklat (10,5%)', w: 1.23, color: C.amber },
    { name: 'D: MICE (10,4%)', w: 1.22, color: C.emerald },
    { name: 'E: 6,2%', w: 0.72, color: C.purple },
    { name: '2,5%', w: 0.323, color: C.textLight }
  ];
  
  let curX = 0.8;
  ribbonData.forEach((rb) => {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: curX, y: 1.65, w: rb.w, h: 0.28,
      fill: { color: rb.color }, line: { color: rb.color, width: 0 }
    });
    if (rb.w > 0.6) {
      slide.addText(rb.name, {
        x: curX, y: 1.68, w: rb.w, h: 0.22,
        fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.textWhite, align: 'center'
      });
    }
    curX += rb.w;
  });
  
  // LEFT HERO PANEL: KLASTER A (54,6% REVENUE)
  const leftX = 0.8, leftY = 2.05, leftW = 4.85, leftH = 4.3;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: leftX, y: leftY, w: leftW, h: leftH,
    fill: { color: C.bg },
    line: { color: C.blue, width: 1.5 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: leftX, y: leftY, w: leftW, h: 0.06,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'anchor', x: leftX + 0.2, y: leftY + 0.16, size: 0.44, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('ENGINE UTAMA (54,6% REVENUE)\nKLASTER A: KERJASAMA STRATEGIS', {
    x: leftX + 0.72, y: leftY + 0.18, w: 2.4, h: 0.45,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy, lineSpacingMultiple: 1.1
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: leftX + 2.95, y: leftY + 0.15, w: 1.75, h: 0.55,
    fill: { color: C.bgTintAmber },
    line: { color: C.amber, width: 1 }
  });
  slide.addText('Rp 12.825.000.000\n54,65% DARI TOTAL TARGET', {
    x: leftX + 2.95, y: leftY + 0.2, w: 1.75, h: 0.45,
    fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.amber, align: 'center', lineSpacingMultiple: 1.1
  });
  
  const progA = [
    { n: 'OGSCI Migas & Maritim (Konsorsium)', sub: '[Anchor Revenue Mandiri]', v: 'Rp 6.250.000.000', isBig: true },
    { n: 'Pemda Solo (e-Katalog Inovasi)', sub: '[Instruksi Walikota]', v: 'Rp 2.250.000.000' },
    { n: 'Startup Binaan STP (Komersialisasi)', sub: '[Success Fee & Jasa]', v: 'Rp 2.000.000.000' },
    { n: 'BP3MI - SMK GoGlobal (Talenta Luar Negeri)', sub: '[Penempatan Global]', v: 'Rp 1.500.000.000' },
    { n: 'Kemenperin & Mitra LPK Industri', sub: '[Subsidi Pelatihan]', v: 'Rp 825.000.000' }
  ];
  
  progA.forEach((pr, pi) => {
    const prY = leftY + 0.8 + pi * 0.56;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: leftX + 0.15, y: prY, w: leftW - 0.3, h: 0.48,
      fill: { color: pr.isBig ? C.bgTintAmber : C.bgSoft },
      line: { color: pr.isBig ? C.amber : C.border, width: pr.isBig ? 1.2 : 0.8 }
    });
    
    slide.addText(pr.n, {
      x: leftX + 0.28, y: prY + 0.05, w: 2.8, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8, bold: true, color: C.textMain
    });
    slide.addText(pr.sub, {
      x: leftX + 0.28, y: prY + 0.25, w: 2.8, h: 0.18,
      fontFace: C.fontBody, fontSize: 7, italic: true, color: pr.isBig ? C.amber : C.textDim
    });
    slide.addText(pr.v, {
      x: leftX + 3.1, y: prY + 0.12, w: 1.5, h: 0.25,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: pr.isBig ? C.amber : C.blue, align: 'right'
    });
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: leftX + 0.15, y: leftY + 3.65, w: leftW - 0.3, h: 0.52,
    fill: { color: C.bgTintBlue },
    line: { color: C.borderBlue, width: 0.8 }
  });
  slide.addText('CATATAN STRATEGIS: 1 program konsorsium OGSCI (Rp 6,25 M) setara 83% pendapatan STP tahun 2026. Ditambah e-Katalog Pemda, klaster ini menjamin kemandirian tanpa APBD.', {
    x: leftX + 0.25, y: leftY + 3.68, w: leftW - 0.5, h: 0.46,
    fontFace: C.fontBody, fontSize: 7, color: C.navy, lineSpacingMultiple: 1.15
  });
  
  // RIGHT PANELS: KLASTER B s.d. F
  const rightX = 5.85, rightW = 6.683;
  
  // Top Row: Klaster B & C
  const topBoxW = 3.24, topBoxH = 2.05;
  
  // Klaster B
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: rightX, y: leftY, w: topBoxW, h: topBoxH,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: rightX, y: leftY, w: topBoxW, h: 0.05,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  addIconBadge(slide, { iconName: 'graduationCap', x: rightX + 0.15, y: leftY + 0.12, size: 0.36, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('KLASTER B: VOKASI & SERTIFIKASI\nRp 3.679.800.000 (15,7%)', {
    x: rightX + 0.58, y: leftY + 0.12, w: topBoxW - 0.7, h: 0.34,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.navy
  });
  
  const bItems = [
    { n: 'Sertifikasi JFT Fungsional', v: 'Rp 1.375.000.000' },
    { n: 'TUK Mandiri Terakreditasi', v: 'Rp 1.350.000.000' },
    { n: 'Konsultasi Lab & Data Center', v: 'Rp 950.000.000' },
    { n: 'Prakerin SMK Solo Raya', v: 'Rp 4.800.000' }
  ];
  bItems.forEach((b, bi) => {
    const bY = leftY + 0.55 + bi * 0.35;
    slide.addText(b.n, { x: rightX + 0.18, y: bY, w: 1.8, h: 0.22, fontFace: C.fontBody, fontSize: 7.5, color: C.textMain });
    slide.addText(b.v, { x: rightX + 1.85, y: bY, w: 1.25, h: 0.22, fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.textMain, align: 'right' });
  });
  
  // Klaster C
  const cX = rightX + topBoxW + 0.2;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: cX, y: leftY, w: topBoxW, h: topBoxH,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: cX, y: leftY, w: topBoxW, h: 0.05,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  addIconBadge(slide, { iconName: 'wrench', x: cX + 0.15, y: leftY + 0.12, size: 0.36, bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber });
  slide.addText('KLASTER C: DIKLAT SPESIALIS\nRp 2.469.000.000 (10,5%)', {
    x: cX + 0.58, y: leftY + 0.12, w: topBoxW - 0.7, h: 0.34,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.amber
  });
  
  const cItems = [
    { n: 'Underwater Wet Welding', v: 'Rp 846.000.000' },
    { n: 'Manufaktur & Welding Intensif', v: 'Rp 728.000.000' },
    { n: 'Desain, Manajerial & EV', v: 'Rp 555.000.000' },
    { n: 'Digital Tech (AI & Cyber)', v: 'Rp 225.000.000' },
    { n: 'Kewirausahaan & UMKM', v: 'Rp 115.000.000' }
  ];
  cItems.forEach((c, ci) => {
    const cy = leftY + 0.52 + ci * 0.28;
    slide.addText(c.n, { x: cX + 0.18, y: cy, w: 1.8, h: 0.2, fontFace: C.fontBody, fontSize: 7, color: C.textMain });
    slide.addText(c.v, { x: cX + 1.85, y: cy, w: 1.25, h: 0.2, fontFace: C.fontTitle, fontSize: 7, bold: true, color: C.textMain, align: 'right' });
  });
  
  // Bottom Row: Klaster D, E, F
  const botBoxW = 2.12, botBoxH = 2.1, botY = leftY + topBoxH + 0.15;
  
  // Klaster D (MICE)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: rightX, y: botY, w: botBoxW, h: botBoxH,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: rightX, y: botY, w: botBoxW, h: 0.05,
    fill: { color: C.emerald }, line: { color: C.emerald, width: 0 }
  });
  addIconBadge(slide, { iconName: 'presentation', x: rightX + 0.12, y: botY + 0.1, size: 0.32, bgTint: C.bgTintGreen, strokeColor: C.emerald, borderTint: C.borderGreen });
  slide.addText('KLASTER D: MICE\nRp 2.450.000.000 (10,4%)', {
    x: rightX + 0.5, y: botY + 0.1, w: botBoxW - 0.55, h: 0.3,
    fontFace: C.fontTitle, fontSize: 7, bold: true, color: C.emerald
  });
  slide.addText('• Sarana/MICE: Rp 1,50 M\n• Co-Working: Rp 900 Jt\n• Jasa Giro Kas: Rp 50 Jt', {
    x: rightX + 0.15, y: botY + 0.6, w: botBoxW - 0.3, h: 1.3,
    fontFace: C.fontBody, fontSize: 7, color: C.textMuted, lineSpacingMultiple: 1.3
  });
  
  // Klaster E (Lab & NDT)
  const eX = rightX + botBoxW + 0.16;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: eX, y: botY, w: botBoxW, h: botBoxH,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: eX, y: botY, w: botBoxW, h: 0.05,
    fill: { color: C.purple }, line: { color: C.purple, width: 0 }
  });
  addIconBadge(slide, { iconName: 'flask', x: eX + 0.12, y: botY + 0.1, size: 0.32, bgTint: C.bgTintPurple, strokeColor: C.purple, borderTint: 'E9D5FF' });
  slide.addText('KLASTER E: LAB & NDT\nRp 1.465.000.000 (6,2%)', {
    x: eX + 0.5, y: botY + 0.1, w: botBoxW - 0.55, h: 0.3,
    fontFace: C.fontTitle, fontSize: 7, bold: true, color: C.purple
  });
  slide.addText('• Fabrikasi Mesin: Rp 1,22 M\n• Uji Mutu SNI: Rp 135 Jt\n• Pesanan Pemda: Rp 105 Jt', {
    x: eX + 0.15, y: botY + 0.6, w: botBoxW - 0.3, h: 1.3,
    fontFace: C.fontBody, fontSize: 7, color: C.textMuted, lineSpacingMultiple: 1.3
  });
  
  // Klaster F (Sewa Lahan)
  const fX = eX + botBoxW + 0.16;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: fX, y: botY, w: botBoxW, h: botBoxH,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: fX, y: botY, w: botBoxW, h: 0.05,
    fill: { color: C.textDim }, line: { color: C.textDim, width: 0 }
  });
  addIconBadge(slide, { iconName: 'mapPin', x: fX + 0.12, y: botY + 0.1, size: 0.32, bgTint: C.bgSoft, strokeColor: C.textDim, borderTint: C.border });
  slide.addText('KLASTER F: SEWA LAHAN\nRp 579.358.831 (2,5%)', {
    x: fX + 0.5, y: botY + 0.1, w: botBoxW - 0.55, h: 0.3,
    fontFace: C.fontTitle, fontSize: 7, bold: true, color: C.textDim
  });
  slide.addText('• Shopee Hub: Rp 544 Jt\n• Katulondi: Rp 35 Jt\n\nDE-ESKALASI SEWA:\nDitekan ke 2,5% demi beralih ke inovasi.', {
    x: fX + 0.15, y: botY + 0.5, w: botBoxW - 0.3, h: 1.4,
    fontFace: C.fontBody, fontSize: 6.8, color: C.textMuted, lineSpacingMultiple: 1.2
  });
  
  // BOTTOM TOTAL LEDGER BANNER
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.45, w: 11.733, h: 0.45,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  addIconBadge(slide, { iconName: 'lock', x: 0.95, y: 6.49, size: 0.36, bgTint: C.blue, strokeColor: C.textWhite, borderTint: C.blue });
  slide.addText('TOTAL TARGET PENDAPATAN BLUD 2030:  Rp 23.468.158.831  (Dua Puluh Tiga Miliar Empat Ratus Enam Puluh Delapan Juta Rupiah) — 100% Mandiri Bebas APBD', {
    x: 1.45, y: 6.53, w: 11.0, h: 0.28,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 11: 6 DIMENSI KEBERDAMPAKAN
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // DAMPAK SOSIAL-EKONOMI DAERAH',
    title: '6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA',
    subtitle: '"Indikator Keberhasilan Harus Mengukur IMPACT — Bukan Sekadar Aktivitas dan Pendapatan Kas" (Dokumen Strategi STP 2026)',
    pageNum: 11
  });
  
  // Left: Horizontal Bar Chart
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 5.3, h: 5.1,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.75, w: 5.3, h: 0.06,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'target', x: 1.0, y: 1.9, size: 0.36, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('SKOR CAPAIAN 6 DIMENSI DAMPAK 2030 (%)', {
    x: 1.45, y: 1.95, w: 4.4, h: 0.25,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy
  });
  
  const impChartData = [
    {
      name: 'Skor Dampak (%)',
      labels: ['Hilirisasi Riset Kampus', 'Peluang Kerja Daerah', 'UMKM & Startup Mandiri', 'Akses Fasilitas Publik', 'SDM Kompeten & Sertifikasi', 'Kolaborasi Industri Murni'],
      values: [70, 75, 80, 85, 90, 100]
    }
  ];
  
  slide.addChart(pptx.charts.BAR, impChartData, {
    x: 1.0, y: 2.35, w: 4.9, h: 3.1,
    barDir: 'bar',
    showLabel: true,
    showValue: true,
    showLegend: false,
    catAxisLabelColor: C.textMain,
    valAxisLabelColor: C.textDim,
    valGridLine: { color: C.border, width: 0.5 },
    chartColors: [C.blue, C.navy, C.emerald, C.amber, C.orange, C.emerald]
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 5.55, w: 4.9, h: 1.15,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  slide.addText('FILOSOFI IMPACT MEASUREMENT KOTA:', {
    x: 1.15, y: 5.62, w: 4.6, h: 0.18,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.amber
  });
  slide.addText('Keberhasilan Solo Technopark tidak diukur dari megahnya gedung atau saldo kas BLUD semata, melainkan dari kesejahteraan nyata warga Surakarta: lapangan kerja yang terserap, UMKM yang naik kelas, serta keterhubungan riset kampus dengan kebutuhan industri.', {
    x: 1.15, y: 5.82, w: 4.6, h: 0.8,
    fontFace: C.fontBody, fontSize: 7.2, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  // Right: 6 KPI SCORECARDS (WITH VECTOR ICON BADGES)
  const kpis = [
    {
      kicker: 'SDM LEBIH KOMPETEN',
      value: '9.000 Peserta/Thn',
      desc: 'Afirmasi kuota minimal 65% warga KTP Surakarta.',
      color: C.blue,
      icon: 'graduationCap'
    },
    {
      kicker: 'UMKM & STARTUP TUMBUH',
      value: '80+ Wirausaha',
      desc: 'Pendampingan inkubasi & digital onboarding GoTo Hub.',
      color: C.amber,
      icon: 'rocket'
    },
    {
      kicker: 'RISET DEKAT KE PASAR',
      value: '45+ Produk',
      desc: 'Paten HKI, Uji Presisi NDT & sertifikasi TKDN.',
      color: C.emerald,
      icon: 'cpu'
    },
    {
      kicker: 'AKSES TEKNOLOGI TERBUKA',
      value: '10.000+ Visit/Thn',
      desc: 'Kunjungan lab AI, Cyber, & magang siswa SMK Solo Raya.',
      color: C.sky,
      icon: 'users'
    },
    {
      kicker: 'PELUANG KERJA BARU',
      value: 'Rp 180 M Multiplier',
      desc: 'Estimasi perputaran ekonomi hotel, resto & MICE.',
      color: C.purple,
      icon: 'trendingUp'
    },
    {
      kicker: 'KOLABORASI INDUSTRI MURNI',
      value: '100% Swasta',
      desc: '95+ korporasi dunia berinvestasi tanpa beban belanja APBD.',
      color: C.emerald,
      icon: 'building'
    }
  ];
  
  const kw = 2.9, kh = 1.55;
  kpis.forEach((kp, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const kx = 6.433 + col * 3.1;
    const ky = 1.75 + row * 1.75;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: kx, y: ky, w: kw, h: kh,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: kx, y: ky, w: kw, h: 0.05,
      fill: { color: kp.color }, line: { color: kp.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: kp.icon, x: kx + kw - 0.52, y: ky + 0.14, size: 0.38, bgTint: C.bgSoft, strokeColor: kp.color, borderTint: C.border });
    
    slide.addText(kp.kicker, {
      x: kx + 0.18, y: ky + 0.14, w: kw - 0.8, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7, bold: true, color: kp.color, charSpacing: 1
    });
    
    slide.addText(kp.value, {
      x: kx + 0.18, y: ky + 0.38, w: kw - 0.36, h: 0.42,
      fontFace: C.fontTitle, fontSize: 16, bold: true, color: C.textMain
    });
    
    slide.addText(kp.desc, {
      x: kx + 0.18, y: ky + 0.9, w: kw - 0.36, h: 0.5,
      fontFace: C.fontBody, fontSize: 7.5, color: C.textDim, lineSpacingMultiple: 1.15
    });
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 12: 3 PERMOHONAN KEBIJAKAN & 100 HARI
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // DUKUNGAN KEPEMIMPINAN & KOMITMEN EKSEKUSI',
    title: '3 PERMOHONAN KEBIJAKAN KEPADA WALIKOTA & RENCANA AKSI 100 HARI',
    subtitle: 'Kami Tidak Memohon Tambahan Pagu APBD — Kami Memohon 3 Payung Regulasi Kepemimpinan Walikota',
    pageNum: 12
  });
  
  const asks = [
    {
      kicker: 'PRIORITAS UTAMA // QUICK WIN 30 HARI',
      title: 'PERMOHONAN 1: INSTRUKSI WALIKOTA AFIRMASI e-KATALOG',
      desc: 'Mengarahkan seluruh OPD Pemkot memprioritaskan belanja produk teknologi & layanan tenant STP via e-Katalog Lokal.',
      target: 'Target: Rp 2,25 M / Thn di 2030 (80+ Produk Terserap)',
      status: 'Bisa langsung diterbitkan dalam 30 hari pertama',
      color: C.blue,
      icon: 'fileText'
    },
    {
      kicker: 'INVESTASI JANGKA PANJANG (10–25 TAHUN)',
      title: 'PERMOHONAN 2: FLEKSIBILITAS KERJASAMA LAHAN BLUD',
      desc: 'Peraturan Walikota skema Long-Term Revenue Sharing yang fleksibel bagi investor teknologi global di kawasan STP.',
      target: 'Target: Menarik 3–5 Investor Baru Tanpa Beban APBD',
      status: 'Harmonisasi bersama BPKAD & Bagian Hukum Setda',
      color: C.amber,
      icon: 'scale'
    },
    {
      kicker: 'KONSORSIUM RISET 5 KAMPUS SOLO RAYA',
      title: 'PERMOHONAN 3: SURAT EDARAN KONSORSIUM RISET DAERAH',
      desc: 'Mendorong perguruan tinggi (UNS, UMS, ISI, Poltek) mengarahkan minimal 30% riset terapan menyelesaikan problem kota di STP.',
      target: 'Target: 5 Kampus Mitra, 30 Klaster Riset Terapan',
      status: 'Penyusunan MoU bersama 5 Rektor Solo Raya',
      color: C.emerald,
      icon: 'graduationCap'
    }
  ];
  
  const aW = 3.65, aH = 2.8, aY = 1.75;
  asks.forEach((a, i) => {
    const aX = 0.8 + i * 4.04;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX, y: aY, w: aW, h: aH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX, y: aY, w: aW, h: 0.06,
      fill: { color: a.color }, line: { color: a.color, width: 0 }
    });
    
    // Icon Badge
    addIconBadge(slide, { iconName: a.icon, x: aX + 0.2, y: aY + 0.16, size: 0.42, bgTint: C.bgSoft, strokeColor: a.color, borderTint: C.border });
    
    slide.addText(a.kicker, {
      x: aX + 0.72, y: aY + 0.16, w: aW - 0.9, h: 0.18,
      fontFace: C.fontTitle, fontSize: 7, bold: true, color: a.color, charSpacing: 1
    });
    
    slide.addText(a.title, {
      x: aX + 0.72, y: aY + 0.36, w: aW - 0.9, h: 0.4,
      fontFace: C.fontTitle, fontSize: 9, bold: true, color: C.textMain
    });
    
    slide.addText(a.desc, {
      x: aX + 0.2, y: aY + 0.85, w: aW - 0.4, h: 0.7,
      fontFace: C.fontBody, fontSize: 8, color: C.textMuted, lineSpacingMultiple: 1.15
    });
    
    // Target Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX + 0.2, y: aY + 1.62, w: aW - 0.4, h: 0.42,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(a.target, {
      x: aX + 0.25, y: aY + 1.66, w: aW - 0.5, h: 0.34,
      fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: a.color, align: 'center'
    });
    
    // Status text
    slide.addImage({
      data: `image/svg+xml;base64,${getSvgBase64('zap', a.color, 2)}`,
      x: aX + 0.2, y: aY + 2.22, w: 0.14, h: 0.14
    });
    slide.addText(a.status, {
      x: aX + 0.38, y: aY + 2.2, w: aW - 0.6, h: 0.4,
      fontFace: C.fontBody, fontSize: 7, italic: true, color: C.textDim
    });
  });
  
  // BOTTOM 100-DAY TIMELINE TABLE
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 4.75, w: 11.733, h: 2.05,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 4.75, w: 11.733, h: 0.05,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, { iconName: 'sparkles', x: 1.0, y: 4.88, size: 0.32, bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue });
  slide.addText('RENCANA AKSI 100 HARI PERTAMA (QUICK WINS PASCA-ARAHAN WALIKOTA SURAKARTA):', {
    x: 1.4, y: 4.92, w: 10.0, h: 0.22,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy
  });
  
  const planRows = [
    [
      { text: 'PERIODE', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft }, align: 'center' } },
      { text: 'AGENDA STRATEGIS', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } },
      { text: 'OUTPUT KONKRET & TARGET KEBERHASILAN', options: { bold: true, fontSize: 7.5, color: C.textMain, fill: { color: C.bgSoft } } }
    ],
    ['Hari 1–10', 'Penyusunan Draf Instruksi Walikota e-Katalog', 'Draf Inkwal Afirmatif siap harmonisasi bersama Bagian Hukum Setda & BPKAD.'],
    ['Hari 11–30', 'Harmonisasi & Penerbitan Instruksi Walikota', 'Inkwal ditandatangani Walikota; Sosialisasi ke seluruh OPD prioritas Pemkot.'],
    ['Hari 31–50', 'Kurasi & Seleksi 5 Produk Inovasi Unggulan', 'Katalog 5 Produk Unggulan Hilirisasi STP (Smart Agri IoT, Alkes, Mesin Sortasi).'],
    ['Hari 51–60', 'Showcase Inovasi di Depan Walikota & OPD', 'Demonstrasi produk di hadapan Walikota & media; komitmen serapan perdana OPD.'],
    ['Hari 61–80', 'Inisiasi Konsorsium Riset 5 Perguruan Tinggi', 'MoU Pra-Konsorsium bersama Rektor UNS, UMS, ISI, Poltek Solo Raya.'],
    ['Hari 81–100', 'Launching Resmi Konsorsium & Hilirisasi Inovasi', 'Penandatanganan MoU Riset Terapan bersama Kadin/Apindo di STP; kick-off hilirisasi.']
  ];
  
  const planTable = planRows.map((r, ri) => {
    if (ri === 0) return r;
    const isP1 = ri === 2; // Hari 11-30 penerbitan
    const isEnd = ri === 6; // Hari 81-100
    const bgCol = isP1 ? C.bgTintAmber : (isEnd ? C.bgTintGreen : (ri % 2 === 0 ? C.bgSoft : C.bg));
    return [
      { text: r[0], options: { fontSize: 7.5, bold: true, color: isP1 ? C.amber : (isEnd ? C.emerald : C.navy), fill: { color: bgCol }, align: 'center' } },
      { text: r[1], options: { fontSize: 7.5, bold: isP1 || isEnd, color: isP1 ? C.amber : C.textMain, fill: { color: bgCol } } },
      { text: r[2], options: { fontSize: 7.5, bold: isP1, color: isP1 ? C.amber : (isEnd ? C.emerald : C.textMuted), fill: { color: bgCol } } }
    ];
  });
  
  slide.addTable(planTable, {
    x: 0.95, y: 5.2, w: 11.433,
    colW: [1.5, 4.0, 5.933],
    rowH: [0.22, 0.22, 0.22, 0.22, 0.22, 0.22, 0.22],
    border: { pt: 0.5, color: C.border },
    autoPage: false
  });
  
  addFooter(slide);
}

// ==========================================
// SIMPAN PPTX V9
// ==========================================
const outDirLocal = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\bahan_paparan_walikota_v3';
const outDirPublic = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\output_paparan';
const outPptxLocal = path.join(outDirLocal, 'PAPARAN_WALIKOTA_STP_V9.pptx');
const outPptxPublic = path.join(outDirPublic, 'PAPARAN_WALIKOTA_STP_V9.pptx');

pptx.writeFile({ fileName: outPptxLocal }).then(() => {
  console.log('[SUCCESS] PPTX Master V9 generated at:', outPptxLocal);
  fs.copyFileSync(outPptxLocal, outPptxPublic);
  console.log('[SUCCESS] Copied to output_paparan:', outPptxPublic);
}).catch(err => {
  console.error('[ERROR] Failed to write PPTX V9:', err);
  process.exit(1);
});
