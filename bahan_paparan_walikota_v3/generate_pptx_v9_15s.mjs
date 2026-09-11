import PptxGenJS from 'pptxgenjs';
import fs from 'fs';
import path from 'path';

// ==========================================
// PALET WARNA SWISS ENTERPRISE ULTRA (V9 - 15S)
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
  microchip: '<rect x="5" y="5" width="14" height="14" rx="2"></rect><rect x="9" y="9" width="6" height="6"></rect><path d="M9 2v3M15 2v3M9 19v3M15 19v3M2 9h3M2 15h3M19 9h3M19 15h3"></path>'
};

function getSvgBase64(iconName, color = C.blue, strokeWidth = 2) {
  const inner = SVG_PATHS[iconName] || SVG_PATHS.checkCircle;
  const hexColor = '#' + color.replace('#', '');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="96" height="96" fill="none" stroke="${hexColor}" stroke-width="${strokeWidth}" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;
  return Buffer.from(svg).toString('base64');
}

// Inisialisasi Presentasi 15 Slide
const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';
pptx.author = 'Pemerintah Kota Surakarta - BRIDA & UPTD KST Solo Technopark';
pptx.title = 'Roadmap Transformasi Solo Technopark 2026-2030 (Swiss Enterprise Ultra 15S)';

// ==========================================
// HELPER UI FUNCTIONS (15 SLIDES)
// ==========================================

function addHeader(slide, { kicker, title, subtitle, pageNum }) {
  // Kicker
  slide.addText(kicker.toUpperCase(), {
    x: 0.8, y: 0.42, w: 10.0, h: 0.24,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.blue,
    charSpacing: 1.5
  });
  
  // Title
  slide.addText(title, {
    x: 0.8, y: 0.65, w: 10.5, h: 0.48,
    fontFace: C.fontTitle, fontSize: 19, bold: true, color: C.textMain
  });
  
  // Subtitle
  slide.addText(subtitle, {
    x: 0.8, y: 1.12, w: 10.5, h: 0.28,
    fontFace: C.fontBody, fontSize: 10.5, italic: true, color: C.textDim
  });
  
  // Page Counter
  if (pageNum) {
    slide.addText(`${pageNum.toString().padStart(2, '0')}  /  15`, {
      x: 11.5, y: 0.48, w: 1.0, h: 0.25,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textDim, align: 'right'
    });
  }
  
  // Hairline Rule
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.44, w: 11.733, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
}

function addFooter(slide) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 7.02, w: 11.733, h: 0.015,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
  
  slide.addText('PEMERINTAH KOTA SURAKARTA  •  BADAN RISET DAN INOVASI DAERAH (BRIDA)  •  ROADMAP 2026–2030', {
    x: 0.8, y: 7.1, w: 8.0, h: 0.22,
    fontFace: C.fontBody, fontSize: 7.5, bold: true, color: C.textDim, charSpacing: 0.8
  });
  
  slide.addText('EDISI EKSEKUTIF ENTERPRISE ULTRA (15 SLIDES)', {
    x: 8.5, y: 7.1, w: 4.033, h: 0.22,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.blue, align: 'right', charSpacing: 0.8
  });
}

function addIconBadge(slide, { iconName, x, y, size = 0.44, bgTint = C.bgTintBlue, strokeColor = C.blue, borderTint = C.borderBlue }) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w: size, h: size,
    fill: { color: bgTint },
    line: { color: borderTint, width: 1 }
  });
  
  const pad = size * 0.18;
  slide.addImage({
    data: `image/svg+xml;base64,${getSvgBase64(iconName, strokeColor, 2)}`,
    x: x + pad, y: y + pad, w: size - 2 * pad, h: size - 2 * pad
  });
}

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
    x: 0.8, y: 0.65, w: 10.0, h: 0.28,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.blue, charSpacing: 1.5
  });
  
  // Main Title
  slide.addText('SOLO TECHNOPARK:\nDARI KAWASAN SAINS MENUJU MESIN IMPACT DAERAH', {
    x: 0.8, y: 1.0, w: 11.5, h: 1.25,
    fontFace: C.fontTitle, fontSize: 27, bold: true, color: C.textMain, lineSpacingMultiple: 1.15
  });
  
  // Subtitle
  slide.addText('"Roadmap Transformasi UPTD KST Surakarta: Mewujudkan Kemandirian BLUD, Hilirisasi Riset Kampus, dan Kesejahteraan Nyata Warga Kota"', {
    x: 0.8, y: 2.35, w: 11.5, h: 0.45,
    fontFace: C.fontBody, fontSize: 12.0, italic: true, color: C.textMuted
  });
  
  // Divider Rule
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 2.92, w: 11.733, h: 0.02,
    fill: { color: C.border }, line: { color: C.border, width: 0 }
  });
  
  // 4 Display Metrics (More generous height & font size)
  const metrics = [
    { num: '+212,8%', label: 'Lompatan Finansial Mandiri', sub: 'Target tumbuh 3,1x dalam 4 tahun', color: C.amber, icon: 'trendingUp', bgTint: C.bgTintAmber, borderTint: C.borderAmber },
    { num: 'Rp 23,47 M', label: 'Target Pendapatan BLUD 2030', sub: '0% Subsidi Kas Daerah APBD', color: C.blue, icon: 'dollarSign', bgTint: C.bgTintBlue, borderTint: C.borderBlue },
    { num: '9.000', label: 'Kapasitas SDM Terlatih / Thn', sub: 'Afirmasi min. 65% KTP Surakarta', color: C.emerald, icon: 'users', bgTint: C.bgTintGreen, borderTint: C.borderGreen },
    { num: '95+', label: 'Mitra Industri Global & Nasional', sub: 'Prinsip Zero APBD Burden', color: C.purple, icon: 'globe', bgTint: C.bgTintPurple, borderTint: 'E9D5FF' }
  ];
  
  const cardW = 2.76;
  const cardGap = 0.23;
  metrics.forEach((m, i) => {
    const cX = 0.8 + i * (cardW + cardGap);
    const cY = 3.15;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: cY, w: cardW, h: 1.6,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    
    // Top Accent Line
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: cY, w: cardW, h: 0.045,
      fill: { color: m.color }, line: { color: m.color, width: 0 }
    });
    
    // Number
    slide.addText(m.num, {
      x: cX + 0.18, y: cY + 0.16, w: cardW - 0.75, h: 0.52,
      fontFace: C.fontTitle, fontSize: 23, bold: true, color: m.color
    });
    
    // Vector Icon
    addIconBadge(slide, {
      iconName: m.icon, x: cX + cardW - 0.58, y: cY + 0.18, size: 0.42,
      bgTint: m.bgTint, strokeColor: m.color, borderTint: m.borderTint
    });
    
    // Label & Sub
    slide.addText(m.label, {
      x: cX + 0.18, y: cY + 0.76, w: cardW - 0.36, h: 0.42,
      fontFace: C.fontBody, fontSize: 10.0, bold: true, color: C.textMain
    });
    slide.addText(m.sub, {
      x: cX + 0.18, y: cY + 1.18, w: cardW - 0.36, h: 0.32,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textDim
    });
  });
  
  // 3 Strategic Pillars (Spacious height & font 10.5pt)
  const pillars = [
    { num: '01 // PARADIGMA BARU', title: 'Bukan Sekadar Gedung Fisik', desc: 'Beralih dari pengelola sewa lahan menjadi DELIVERY UNIT inovasi BRIDA yang menghasilkan IMPACT nyata bagi masyarakat dan industri.', color: C.sky, icon: 'cpu' },
    { num: '02 // KEMANDIRIAN PENUH', title: 'BLUD Mandiri Tanpa APBD', desc: 'Lompatan pendapatan terukur dari Rp 7,50 M (2026) menuju Rp 23,47 M (2030) — tumbuh +212,8% dengan 0% subsidi kas daerah.', color: C.amber, icon: 'shieldCheck' },
    { num: '03 // HEXAHELIX & DAMPAK', title: '95+ Mitra Korporasi Dunia', desc: 'Investasi swasta murni berkolaborasi menyerap 9.000 tenaga kerja dengan komitmen afirmasi minimal 65% warga KTP Surakarta.', color: C.emerald, icon: 'award' }
  ];
  
  const pCardW = 3.75;
  const pGap = 0.24;
  pillars.forEach((p, i) => {
    const pX = 0.8 + i * (pCardW + pGap);
    const pY = 4.95;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX, y: pY, w: pCardW, h: 1.75,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX, y: pY, w: pCardW, h: 0.045,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: p.icon, x: pX + 0.2, y: pY + 0.2, size: 0.44,
      bgTint: C.bgTintBlue, strokeColor: p.color, borderTint: C.borderBlue
    });
    
    slide.addText(p.num, {
      x: pX + 0.75, y: pY + 0.22, w: pCardW - 0.9, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: p.color, charSpacing: 1
    });
    slide.addText(p.title, {
      x: pX + 0.75, y: pY + 0.42, w: pCardW - 0.9, h: 0.32,
      fontFace: C.fontTitle, fontSize: 12.0, bold: true, color: C.textMain
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: pX + 0.2, y: pY + 0.82, w: pCardW - 0.4, h: 0.015,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    slide.addText(p.desc, {
      x: pX + 0.2, y: pY + 0.95, w: pCardW - 0.4, h: 0.72,
      fontFace: C.fontBody, fontSize: 10.0, color: C.textMuted, lineSpacingMultiple: 1.2
    });
  });
  
  slide.addText('Disampaikan oleh: Kepala UPTD KST Solo Technopark  •  Didampingi: Kepala BRIDA Kota Surakarta  •  Surakarta, 2026', {
    x: 0.8, y: 6.9, w: 11.733, h: 0.3,
    fontFace: C.fontBody, fontSize: 9.0, color: C.textDim, align: 'center'
  });
}

// ==========================================
// SLIDE 02: KEDUDUKAN KELEMBAGAAN & CHAIN OF IMPACT
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
  
  const cols = [
    {
      step: '01 / INPUT',
      title: 'BRIDA KOTA SURAKARTA',
      role: 'Otak & Regulator Kebijakan Daerah',
      color: C.navy,
      icon: 'fileText',
      points: [
        'Merumuskan arah kebijakan riset & inovasi daerah',
        'Penyusunan kajian teknokratik & roadmap strategis',
        'Orkestrasi ekosistem iptek antar-OPD & perguruan tinggi',
        'Monitoring & evaluasi akuntabilitas dampak sosial-ekonomi'
      ],
      output: 'OUTPUT: Regulasi, Roadmap, & Pengawasan Kebijakan'
    },
    {
      step: '02 / PROSES',
      title: 'UPTD KST SOLO TECHNOPARK',
      role: 'Tangan Eksekutor & Operator Bisnis BLUD',
      color: C.amber,
      icon: 'wrench',
      points: [
        'Penyelenggaraan diklat vokasi & sertifikasi industri',
        'Inkubasi bisnis, alih teknologi & komersialisasi riset',
        'Uji coba prototipe lab presisi & penetrant testing (NDT)',
        'Pengelolaan portofolio bisnis BLUD mandiri & kemitraan'
      ],
      output: 'OUTPUT: Layanan Teknis, Inkubasi, & Hilirisasi Produk'
    },
    {
      step: '03 / OUTCOME',
      title: 'WARGA SURAKARTA & INDUSTRI',
      role: 'Penerima Manfaat Utama (Benefisiari)',
      color: C.emerald,
      icon: 'users',
      points: [
        'SDM lokal kompeten & tersertifikasi kerja industri global',
        'UMKM & startup teknologi tumbuh mandiri omzetnya',
        'Produk inovasi kampus terserap pasar & e-Katalog Pemda',
        'Masuknya investasi swasta & penyerapan 9.000 tenaga kerja'
      ],
      output: 'IMPACT: Lapangan Kerja, Daya Saing, & PAD Mandiri'
    }
  ];
  
  const colW = 3.65;
  const colGap = 0.39;
  cols.forEach((c, i) => {
    const cX = 0.8 + i * (colW + colGap);
    const cY = 1.68;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: cY, w: colW, h: 4.15,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX, y: cY, w: colW, h: 0.05,
      fill: { color: c.color }, line: { color: c.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: c.icon, x: cX + 0.25, y: cY + 0.2, size: 0.44,
      bgTint: C.bgSoft, strokeColor: c.color, borderTint: C.border
    });
    
    slide.addText(c.step, {
      x: cX + 0.8, y: cY + 0.22, w: colW - 0.95, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: c.color, charSpacing: 1
    });
    slide.addText(c.title, {
      x: cX + 0.8, y: cY + 0.42, w: colW - 0.95, h: 0.32,
      fontFace: C.fontTitle, fontSize: 12.0, bold: true, color: C.textMain
    });
    slide.addText(c.role, {
      x: cX + 0.25, y: cY + 0.78, w: colW - 0.5, h: 0.22,
      fontFace: C.fontBody, fontSize: 9.5, bold: true, color: c.color
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX + 0.25, y: cY + 1.05, w: colW - 0.5, h: 0.015,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    c.points.forEach((p, pi) => {
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', c.color, 2)}`,
        x: cX + 0.25, y: cY + 1.25 + pi * 0.56, w: 0.18, h: 0.18
      });
      slide.addText(p, {
        x: cX + 0.52, y: cY + 1.2 + pi * 0.56, w: colW - 0.75, h: 0.5,
        fontFace: C.fontBody, fontSize: 10.0, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: cX + 0.2, y: cY + 3.52, w: colW - 0.4, h: 0.45,
      fill: { color: C.bgSoft },
      line: { color: c.color, width: 1 }
    });
    slide.addText(c.output, {
      x: cX + 0.25, y: cY + 3.55, w: colW - 0.5, h: 0.38,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: c.color, align: 'center'
    });
    
    if (i < 2) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: cX + colW + 0.1, y: cY + 1.85, w: 0.19, h: 0.28,
        fill: { color: C.amber }, line: { color: C.amber, width: 0 }
      });
    }
  });
  
  // Legal Foundation Footer Box
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.05, w: 11.733, h: 0.78,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  addIconBadge(slide, {
    iconName: 'scale', x: 0.98, y: 6.16, size: 0.46,
    bgTint: C.bgTintBlue, strokeColor: C.navy, borderTint: C.borderBlue
  });
  slide.addText('LANDASAN HUKUM TATA KELOLA PEMKOT SURAKARTA:\n• Perwali Surakarta No. 15 Tahun 2022: Kedudukan UPTD KST Solo Technopark sebagai unit pelaksana teknis operasional di bawah naungan BRIDA.\n• Perwali Surakarta No. 38 Tahun 2022: Pola Tata Kelola BLUD UPTD KST Solo Technopark yang memberikan fleksibilitas bisnis sehat & akuntabel.', {
    x: 1.6, y: 6.12, w: 10.7, h: 0.65,
    fontFace: C.fontBody, fontSize: 9.0, color: C.textMain, lineSpacingMultiple: 1.18
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 03: FONDASI HUKUM & FLEKSIBILITAS BLUD
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
      reg: 'PERMENDAGRI 79/2018',
      topic: 'Tata Kelola Keuangan BLUD',
      color: C.blue,
      icon: 'fileText',
      quote: '"Fleksibilitas BLUD adalah keleluasaan dalam pola pengelolaan keuangan dengan menerapkan praktik bisnis yang sehat untuk meningkatkan layanan kepada masyarakat tanpa mencari keuntungan semata."',
      checks: [
        'Keleluasaan bermitra dengan swasta tanpa birokrasi kaku APBD',
        'Mengelola dan memutar pendapatan fungsional jasa secara mandiri',
        'Rekrutmen tenaga ahli & instruktur profesional sesuai kebutuhan pasar'
      ]
    },
    {
      reg: 'PERPRES 106/2017',
      topic: 'Kawasan Sains & Teknologi (KST)',
      color: C.amber,
      icon: 'target',
      quote: '"KST berfungsi sebagai wahana kerja sama riset berkelanjutan antara Pemerintah, PT, dan Industri, wajib menyediakan 4 layanan: Teknis, Teknologi, Inkubasi, & Pendukung."',
      checks: [
        'Mandat nasional fasilitasi hilirisasi & spin-off perusahaan riset',
        'Kewajiban menyediakan 4 pilar layanan iptek terintegrasi',
        'Landasan kolaborasi hexahelix multi-stakeholder formal'
      ]
    },
    {
      reg: 'PERWALI 15 & 38/2022',
      topic: 'Struktur UPTD & Pola Tata Kelola',
      color: C.emerald,
      icon: 'scale',
      quote: '"Menetapkan struktur UPTD KST Solo Technopark dalam naungan BRIDA Kota Surakarta dengan pola tata kelola BLUD profesional berorientasi layanan publik."',
      checks: [
        'Payung hukum operasional tingkat kota yang sah dan mengikat',
        'Pembagian peran teknis yang presisi antara BRIDA dan STP',
        'Dasar legalitas pemanfaatan aset kawasan & kerjasama industri'
      ]
    }
  ];
  
  const rW = 3.65;
  const rGap = 0.39;
  regs.forEach((r, i) => {
    const rX = 0.8 + i * (rW + rGap);
    const rY = 1.68;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX, y: rY, w: rW, h: 4.35,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX, y: rY, w: rW, h: 0.05,
      fill: { color: r.color }, line: { color: r.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: r.icon, x: rX + 0.25, y: rY + 0.2, size: 0.44,
      bgTint: C.bgSoft, strokeColor: r.color, borderTint: C.border
    });
    
    slide.addText(r.reg, {
      x: rX + 0.8, y: rY + 0.22, w: rW - 1.0, h: 0.22,
      fontFace: C.fontTitle, fontSize: 11.5, bold: true, color: C.textMain
    });
    slide.addText(r.topic, {
      x: rX + 0.8, y: rY + 0.44, w: rW - 1.0, h: 0.2,
      fontFace: C.fontBody, fontSize: 9.0, bold: true, color: r.color
    });
    
    // Quote Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: rX + 0.2, y: rY + 0.78, w: rW - 0.4, h: 1.5,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(r.quote, {
      x: rX + 0.35, y: rY + 0.86, w: rW - 0.7, h: 1.34,
      fontFace: C.fontBody, fontSize: 9.5, italic: true, color: C.textMuted, lineSpacingMultiple: 1.2
    });
    
    // Checks with vector icon
    r.checks.forEach((ch, ci) => {
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', r.color, 2)}`,
        x: rX + 0.25, y: rY + 2.45 + ci * 0.62, w: 0.18, h: 0.18
      });
      slide.addText(ch, {
        x: rX + 0.52, y: rY + 2.42 + ci * 0.62, w: rW - 0.75, h: 0.55,
        fontFace: C.fontBody, fontSize: 9.5, color: C.textMain, lineSpacingMultiple: 1.18
      });
    });
  });
  
  // Bottom Summary Ribbon
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.22, w: 11.733, h: 0.65,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('KESIMPULAN YURIDIS: Solo Technopark telah memiliki legitimasi hukum lengkap dari tingkat nasional (Perpres & Permendagri) hingga regulasi kepala daerah (Perwali) untuk bermanuver cepat, profesional, dan akuntabel.', {
    x: 1.0, y: 6.28, w: 11.333, h: 0.52,
    fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: C.textWhite, align: 'center', lineSpacingMultiple: 1.2
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
      code: 'PILAR 1', title: 'LAYANAN TEKNIS (DIKLAT VOKASI INDUSTRI)',
      color: C.blue, icon: 'wrench',
      items: [
        { name: 'Mekanik, Otomasi & Desain Manufaktur Presisi' },
        { name: 'Welding Intensif (Pengelasan Konstruksi & Pipa)' },
        { name: 'OGSCI: Oil, Gas, Coal, Shipbuilding & Infra Training' },
        { name: 'Underwater Wet Welding (Keahlian Langka Nasional)' },
        { name: 'Teknologi Kendaraan Listrik (EV Conversion Lab)' },
        { name: 'Cyber Security & Artificial Intelligence Practitioner' }
      ]
    },
    {
      code: 'PILAR 2', title: 'PENGEMBANGAN TEKNOLOGI (REKAYASA MANUFAKTUR)',
      color: C.amber, icon: 'cpu',
      items: [
        { name: 'Jasa Layanan Rekayasa Industri & Permesinan Presisi' },
        { name: 'Produksi Komponen Presisi (Precision Parts Industri)' },
        { name: 'Pembuatan & Uji Prototipe IKM/UKM Kota Surakarta' },
        { name: 'Non-Destructive Testing: Uji Penetrant Testing (NDT)' },
        { name: 'Konsultasi Laboratorium Terpadu & Data Center' },
        { name: 'Standarisasi Mutu & Sertifikasi TKDN / SNI Produk' }
      ]
    },
    {
      code: 'PILAR 3', title: 'LAYANAN INKUBASI (HILIRISASI KE PASAR)',
      color: C.emerald, icon: 'sprout',
      items: [
        { name: 'Pra-Inkubasi & Inkubasi Bisnis Startup Berbasis Riset' },
        { name: 'Fasilitasi HKI, Paten Produk, & Hak Cipta Inovator' },
        { name: 'Hilirisasi Riset Perguruan Tinggi ke Pasar Industri' },
        { name: 'National Cyber Security Hub & AI Experience Center' },
        { name: 'Digital Technopark, Gaming Hub, & GoTo UMKM Center' },
        { name: 'Co-Working Space, Virtual Office & Akses Permodalan' }
      ]
    },
    {
      code: 'PILAR 4', title: 'LAYANAN PENDUKUNG (COMMUNITY & EDUCATION)',
      color: C.purple, icon: 'graduationCap',
      items: [
        { name: 'Prakerin Kerja Industri Siswa SMK Se-Solo Raya' },
        { name: 'Tempat Uji Kompetensi (TUK Mandiri) & Sertifikasi JFT' },
        { name: 'Solo Science Center (Pusat Edukasi IPTEK & Wisata Sains)' },
        { name: 'Pemanfaatan Sarana Kawasan, Gedung Expo & MICE' },
        { name: 'Kemitraan Komunitas Kreatif & Kunjungan Industri' },
        { name: 'Pusat Kolaborasi Multipihak Pemkot Surakarta' }
      ]
    }
  ];
  
  const qW = 5.72;
  const qH = 2.45;
  const qGapX = 0.29;
  const qGapY = 0.22;
  
  pillars.forEach((p, i) => {
    const colIdx = i % 2;
    const rowIdx = Math.floor(i / 2);
    const qX = 0.8 + colIdx * (qW + qGapX);
    const qY = 1.68 + rowIdx * (qH + qGapY);
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: qX, y: qY, w: qW, h: qH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: qX, y: qY, w: qW, h: 0.045,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: p.icon, x: qX + 0.2, y: qY + 0.18, size: 0.42,
      bgTint: C.bgSoft, strokeColor: p.color, borderTint: C.border
    });
    
    slide.addText(p.title, {
      x: qX + 0.72, y: qY + 0.24, w: qW - 0.9, h: 0.3,
      fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: p.color
    });
    
    p.items.forEach((it, idx) => {
      const itCol = idx < 3 ? 0 : 1;
      const itRow = idx % 3;
      const itX = qX + 0.2 + itCol * 2.75;
      const itY = qY + 0.72 + itRow * 0.54;
      
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', p.color, 2)}`,
        x: itX, y: itY + 0.03, w: 0.16, h: 0.16
      });
      slide.addText(it.name, {
        x: itX + 0.24, y: itY, w: 2.45, h: 0.48,
        fontFace: C.fontBody, fontSize: 9.5, color: C.textMain, lineSpacingMultiple: 1.15
      });
    });
  });
  
  // Bottom Quote Ribbon
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.42, w: 11.733, h: 0.45,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 1 }
  });
  slide.addText('"BLUD menggeser paradigma: Dari sekadar pengelolaan kawasan fisik menjadi penyedia solusi teknologi yang responsif terhadap kebutuhan industri dan masyarakat."', {
    x: 1.0, y: 6.48, w: 11.333, h: 0.32,
    fontFace: C.fontTitle, fontSize: 10.0, italic: true, bold: true, color: C.blue, align: 'center'
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 05: PIPELINE ALUR HILIRISASI RISET KAMPUS (EXPANDED FOCUS)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // MEKANISME HILIRISASI INOVASI',
    title: 'PIPELINE HILIRISASI RISET: DARI LABORATORIUM MENUJU PASAR NYATA',
    subtitle: 'Arsitektur 5 Tahapan Terstruktur Menjembatani Riset Kampus Menjadi Produk Bernilai Ekonomi bagi Industri',
    pageNum: 5
  });
  
  const stages = [
    { num: 'TAHAP 01', name: 'RISET & INVENSI', actor: 'Kampus & Peneliti', color: C.blue, icon: 'flask', desc: 'Invensi riset dari UNS, UMS, ISI, Poltek Solo Raya. Berangkat dari riset dasar dosen & mahasiswa untuk kebutuhan daerah.' },
    { num: 'TAHAP 02', name: 'SELEKSI & KURASI', actor: 'STP & BRIDA', color: C.sky, icon: 'target', desc: 'Kurasi kelayakan komersial, kesiapan TRL 7–9, relevansi kebutuhan industri, serta pemecahan problem prioritas Pemkot.' },
    { num: 'TAHAP 03', name: 'PROTOTYPING', actor: 'Lab Presisi STP', color: C.amber, icon: 'layers', desc: 'Uji fungsi teknis di Lab Manufaktur Presisi, pengujian penetrant NDT, validasi mutu produk, serta sertifikasi SNI & TKDN.' },
    { num: 'TAHAP 04', name: 'INKUBASI BISNIS', actor: 'Inkubator STP', color: C.emerald, icon: 'building', desc: 'Penyusunan model bisnis teruji, valuasi pasar, perlindungan HKI/Paten, permodalan awal, dan fasilitasi investor industri.' },
    { num: 'TAHAP 05', name: 'KOMERSIALISASI', actor: 'Pasar & Industri', color: C.purple, icon: 'shoppingCart', desc: 'Penetrasi etalase e-Katalog Lokal Pemkot Surakarta, captive market belanja OPD, serta serapan jaringan korporasi nasional.' }
  ];
  
  const sW = 2.18;
  const sH = 3.35;
  const sGap = 0.208;
  
  stages.forEach((st, i) => {
    const sX = 0.8 + i * (sW + sGap);
    const sY = 1.68;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: sX, y: sY, w: sW, h: sH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: sX, y: sY, w: sW, h: 0.05,
      fill: { color: st.color }, line: { color: st.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: st.icon, x: sX + 0.2, y: sY + 0.2, size: 0.44,
      bgTint: C.bgSoft, strokeColor: st.color, borderTint: C.border
    });
    
    slide.addText(st.num, {
      x: sX + 0.72, y: sY + 0.25, w: sW - 0.8, h: 0.18,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: st.color, charSpacing: 1
    });
    slide.addText(st.name, {
      x: sX + 0.2, y: sY + 0.75, w: sW - 0.4, h: 0.45,
      fontFace: C.fontTitle, fontSize: 12.0, bold: true, color: C.textMain
    });
    slide.addText(st.actor, {
      x: sX + 0.2, y: sY + 1.25, w: sW - 0.4, h: 0.22,
      fontFace: C.fontBody, fontSize: 9.5, bold: true, color: st.color
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: sX + 0.2, y: sY + 1.55, w: sW - 0.4, h: 0.015,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    slide.addText(st.desc, {
      x: sX + 0.2, y: sY + 1.7, w: sW - 0.4, h: 1.5,
      fontFace: C.fontBody, fontSize: 10.0, color: C.textMuted, lineSpacingMultiple: 1.25
    });
    
    if (i < 4) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: sX + sW + 0.04, y: sY + 1.45, w: 0.13, h: 0.25,
        fill: { color: C.borderDark }, line: { color: C.borderDark, width: 0 }
      });
    }
  });
  
  // 3 Strategic Pillars of Hilirisasi at the Bottom
  const bottomPillars = [
    { title: 'KRITERIA KURASI KETAT', desc: 'Inovasi wajib mencapai TRL 7-9, memiliki nilai keekonomian jelas, bersertifikat paten/HKI, dan mampu menyelesaikan permasalahan kota/industri.', icon: 'target', color: C.blue },
    { title: 'PERAN STRATEGIS STP', desc: 'STP berperan sebagai akselerator: menyediakan fasilitas Lab Presisi, pengujian NDT penetrant testing, inkubasi bisnis, dan standardisasi TKDN/SNI.', icon: 'cpu', color: C.amber },
    { title: 'OFF-TAKER GUARANTEE', desc: 'Jaminan serapan pasar melalui payung Instruksi Walikota e-Katalog Lokal serta koneksi ke 95+ jaringan korporasi mitra skala nasional & global.', icon: 'shoppingCart', color: C.emerald }
  ];
  
  const bpW = 3.65;
  const bpGap = 0.39;
  bottomPillars.forEach((bp, bi) => {
    const bpX = 0.8 + bi * (bpW + bpGap);
    const bpY = 5.25;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: bpX, y: bpY, w: bpW, h: 1.55,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: bpX, y: bpY, w: bpW, h: 0.04,
      fill: { color: bp.color }, line: { color: bp.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: bp.icon, x: bpX + 0.2, y: bpY + 0.18, size: 0.42,
      bgTint: C.bg, strokeColor: bp.color, borderTint: C.border
    });
    
    slide.addText(bp.title, {
      x: bpX + 0.72, y: bpY + 0.24, w: bpW - 0.9, h: 0.28,
      fontFace: C.fontTitle, fontSize: 11.0, bold: true, color: bp.color
    });
    
    slide.addText(bp.desc, {
      x: bpX + 0.2, y: bpY + 0.68, w: bpW - 0.4, h: 0.75,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMain, lineSpacingMultiple: 1.2
    });
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 06: PORTOFOLIO INOVASI & FASILITAS CANGGIH (NEW EXPANDED)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // SHOWCASE INOVASI & FASILITAS CANGGIH',
    title: 'PORTOFOLIO PRODUK INOVASI TERUJI & INFRASTRUKTUR KAWASAN CANGGIH',
    subtitle: 'Bukti Nyata Hilirisasi Teknologi Siap Serap Pasar Didukung Fasilitas Kelas Dunia Tanpa Beban APBD',
    pageNum: 6
  });
  
  // Left: Table of Real Innovation Products (w: 6.8)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 6.9, h: 4.35,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 6.9, h: 0.045,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'microchip', x: 1.0, y: 1.84, size: 0.4,
    bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('CONTOH PRODUK INOVASI RIIL SIAP PASAR & e-KATALOG', {
    x: 1.5, y: 1.9, w: 6.0, h: 0.28,
    fontFace: C.fontTitle, fontSize: 11.0, bold: true, color: C.navy
  });
  
  const pRows = [
    ['KLASTER TEKNOLOGI', 'CONTOH PRODUK INOVASI', 'STATUS TAHAPAN', 'TARGET SERAPAN'],
    ['Smart Agriculture IoT', 'Mesin Sortasi Biji Kopi Otomatis', 'Validasi TRL 7', 'Kelompok Tani & IKM'],
    ['Smart Agriculture IoT', 'Teknologi Pertanian Presisi IoT & AI', 'Uji Lapangan', 'Dinas Pertanian'],
    ['Smart Agriculture IoT', 'Mesin Tetas Telur Pintar Mandiri', 'Siap e-Katalog', 'Peternak Solo Raya'],
    ['Deep Tech Mobility', 'Autonomous Vehicle Elektrik Kampus', 'R&D Bersama', 'Kawasan Wisata & Kampus'],
    ['GovTech & Pelatihan', 'VR Pelatihan Industri & Chatbot AI OPD', 'Komersial Aktif', 'OPD Pemkot & Industri']
  ];
  
  const tableData = pRows.map((row, rIdx) => {
    return row.map((cell, cIdx) => {
      const isHeader = rIdx === 0;
      let textCol = C.textMain;
      if (isHeader) textCol = C.textMain;
      else if (cIdx === 2) {
        if (cell.includes('Siap')) textCol = C.emerald;
        else if (cell.includes('Komersial')) textCol = C.blue;
        else if (cell.includes('Validasi')) textCol = C.amber;
        else textCol = C.sky;
      }
      
      return {
        text: cell,
        options: {
          fontFace: isHeader ? C.fontTitle : C.fontBody,
          fontSize: isHeader ? 8.5 : 9.5,
          bold: isHeader || cIdx === 1,
          color: textCol,
          fill: isHeader ? { color: 'F1F5F9' } : (rIdx % 2 === 1 ? { color: 'FFFFFF' } : { color: C.bgSoft }),
          align: cIdx >= 2 ? 'center' : 'left',
          valign: 'middle'
        }
      };
    });
  });
  
  slide.addTable(tableData, {
    x: 1.0, y: 2.35, w: 6.5, h: 3.45,
    colW: [1.6, 2.3, 1.3, 1.3],
    border: { pt: 0.5, color: C.border }
  });
  
  // Right: 5 Advanced Facilities Ecosystem (w: 4.5)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 8.0, y: 1.68, w: 4.533, h: 4.35,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 8.0, y: 1.68, w: 4.533, h: 0.045,
    fill: { color: C.emerald }, line: { color: C.emerald, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'building', x: 8.2, y: 1.84, size: 0.4,
    bgTint: C.bgTintGreen, strokeColor: C.emerald, borderTint: C.borderGreen
  });
  slide.addText('EKOSISTEM 5 FASILITAS PENDUKUNG CANGGIH', {
    x: 8.7, y: 1.9, w: 3.7, h: 0.28,
    fontFace: C.fontTitle, fontSize: 11.0, bold: true, color: C.emerald
  });
  
  const facil = [
    { title: 'National Cyber Security Hub', desc: 'Pusat ketahanan digital & talenta siber nasional bersama BSSN.', icon: 'shieldCheck', color: C.blue },
    { title: 'Gaming Hub & Digital Studio', desc: 'Pengembangan game, animasi, & konten kreatif e-sport.', icon: 'cpu', color: C.purple },
    { title: 'AI Experience Center', desc: 'Laboratorium kecerdasan buatan terapan industri & showcase solusi AI.', icon: 'zap', color: C.amber },
    { title: 'GoTo UMKM Center', desc: 'Akselerasi digitalisasi, kurasi produk, & onboarding pedagang lokal.', icon: 'shoppingCart', color: C.emerald },
    { title: 'Laboratorium Rekayasa Presisi & NDT', desc: 'Fasilitas fabrikasi mesin, permesinan presisi & pengujian NDT penetrant.', icon: 'wrench', color: C.sky }
  ];
  
  facil.forEach((f, fi) => {
    const fY = 2.38 + fi * 0.7;
    addIconBadge(slide, {
      iconName: f.icon, x: 8.2, y: fY, size: 0.38,
      bgTint: C.bgSoft, strokeColor: f.color, borderTint: C.border
    });
    slide.addText(f.title, {
      x: 8.7, y: fY, w: 3.6, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textMain
    });
    slide.addText(f.desc, {
      x: 8.7, y: fY + 0.22, w: 3.6, h: 0.42,
      fontFace: C.fontBody, fontSize: 9.0, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  // Bottom Callout Banner
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.2, w: 11.733, h: 0.65,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('NILAI STRATEGIS: Seluruh fasilitas canggih ini telah aktif dan dapat dimanfaatkan langsung oleh OPD Pemkot Surakarta, kampus riset, dan pelaku industri lokal tanpa perlu alokasi belanja modal baru di APBD.', {
    x: 1.0, y: 6.26, w: 11.333, h: 0.52,
    fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: C.textWhite, align: 'center', lineSpacingMultiple: 1.2
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 07: EKOSISTEM HEXAHELIX 95+ MITRA STRATEGIS
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // JEJARING KEMITRAAN STRATEGIS',
    title: 'EKOSISTEM HEXAHELIX: 95+ MITRA STRATEGIS TANPA BEBAN APBD',
    subtitle: 'Kepercayaan Korporasi Global & Nasional yang Menanamkan Investasi Langsung di Jantung Kota Surakarta',
    pageNum: 7
  });
  
  // Left: Donut Chart Hexahelix (w: 4.5)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 4.5, h: 5.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 4.5, h: 0.045,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'pieChart', x: 1.0, y: 1.84, size: 0.38,
    bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('KOMPOSISI 6 AKTOR HEXAHELIX (95+ MITRA)', {
    x: 1.5, y: 1.9, w: 3.6, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.navy
  });
  
  // Native Doughnut Chart (Without overlapping slice labels)
  slide.addChart(pptx.charts.DOUGHNUT, [
    {
      name: 'Proporsi Mitra',
      labels: ['Industri & Bisnis', 'Akademisi / Kampus', 'Komunitas & UMKM', 'Pemerintah / BUMN', 'Media Massa', 'Lembaga Global'],
      values: [35, 20, 15, 12, 8, 5]
    }
  ], {
    x: 0.95, y: 2.22, w: 4.2, h: 2.5,
    holeSize: 58,
    chartColors: [C.sky, C.navy, C.emerald, C.amber, C.purple, C.red],
    showLegend: false,
    showLabel: false,
    showValue: false
  });
  
  // Legend Cards
  const legends = [
    { name: 'Industri & Bisnis: 37% (35)', color: C.sky },
    { name: 'Pemerintah / BUMN: 13% (12)', color: C.amber },
    { name: 'Akademisi / Kampus: 21% (20)', color: C.navy },
    { name: 'Media Massa: 8% (8)', color: C.purple },
    { name: 'Komunitas & UMKM: 16% (15)', color: C.emerald },
    { name: 'Lembaga Global: 5% (5)', color: C.red }
  ];
  
  legends.forEach((lg, lIdx) => {
    const lCol = lIdx % 2;
    const lRow = Math.floor(lIdx / 2);
    const lgX = 1.0 + lCol * 2.15;
    const lgY = 4.8 + lRow * 0.36;
    
    slide.addShape(pptx.shapes.OVAL, {
      x: lgX, y: lgY + 0.04, w: 0.14, h: 0.14,
      fill: { color: lg.color }, line: { color: lg.color, width: 0 }
    });
    slide.addText(lg.name, {
      x: lgX + 0.22, y: lgY, w: 1.9, h: 0.25,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textMain
    });
  });
  
  // Callout Box inside left card
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 5.92, w: 4.1, h: 0.75,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  slide.addText('TOTAL: 95+ MITRA STRATEGIS AKTIF\n• 100% Investasi Swasta & PPP (0% Beban APBD)\n• Akses fasilitas 5G, Cyber Security, AI, & Migas\n• Saluran penyerapan kerja langsung bagi warga Solo', {
    x: 1.2, y: 5.96, w: 3.7, h: 0.68,
    fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  // Right: Table of World-Class Industrial Partners (w: 6.9)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 1.68, w: 6.933, h: 5.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.6, y: 1.68, w: 6.933, h: 0.045,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'award', x: 5.8, y: 1.84, size: 0.38,
    bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber
  });
  slide.addText('MITRA INDUSTRI KELAS DUNIA YANG AKTIF BERINVESTASI DI STP', {
    x: 6.3, y: 1.9, w: 6.0, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.amber
  });
  
  const partners = [
    ['SEKTOR INDUSTRI', 'MITRA STRATEGIS', 'KONTRIBUSI NYATA BAGI KAWASAN STP'],
    ['E-Commerce & Logistik', 'Shopee Indonesia', 'Infrastruktur UMKM Hub, Co-location, sewa lahan Rp 544 jt/th'],
    ['Gaming & Kreatif', 'Garena', 'Gaming studio, beasiswa talenta digital, turnamen esports'],
    ['Perbankan & Fintech', 'Bank Mandiri', 'Mandiri Innovation Hub, fasilitasi modal startup binaan'],
    ['Telekomunikasi', 'Indosat Ooredoo Hutchison', 'Laboratorium 5G, konektivitas fiber optik terintegrasi kawasan'],
    ['Ride-Hailing & UMKM', 'GoTo (Gojek-Tokopedia)', 'GoTo UMKM Center, pelatihan onboarding merchant lokal'],
    ['Migas & Industri Berat', 'OGSCI Consortium', 'Pelatihan migas internasional: Target Rp 6,25 M (2030)']
  ];
  
  const tablePart = partners.map((r, i) => {
    const isHead = i === 0;
    return r.map((c, ci) => {
      let tCol = C.textMain;
      if (i === 6) tCol = C.amber;
      return {
        text: c,
        options: {
          fontFace: isHead ? C.fontTitle : C.fontBody,
          fontSize: isHead ? 8.5 : 9.5,
          bold: isHead || ci === 1 || i === 6,
          color: tCol,
          fill: isHead ? { color: 'F1F5F9' } : (i % 2 === 1 ? { color: 'FFFFFF' } : { color: C.bgSoft }),
          valign: 'middle'
        }
      };
    });
  });
  
  slide.addTable(tablePart, {
    x: 5.8, y: 2.25, w: 6.533, h: 3.1,
    colW: [1.7, 1.8, 3.033],
    border: { pt: 0.5, color: C.border }
  });
  
  // Callout Principle Zero APBD
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.8, y: 5.5, w: 6.533, h: 1.15,
    fill: { color: C.bgTintGreen },
    line: { color: C.borderGreen, width: 1 }
  });
  addIconBadge(slide, {
    iconName: 'shieldCheck', x: 6.0, y: 5.75, size: 0.5,
    bgTint: C.bg, strokeColor: C.emerald, borderTint: C.borderGreen
  });
  slide.addText('PRINSIP ZERO APBD BURDEN: SELURUH FASILITAS CANGGIH DIBANGUN OLEH SWASTA', {
    x: 6.7, y: 5.62, w: 5.4, h: 0.28,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.emerald
  });
  slide.addText('Laboratorium AI, Studio Gaming, 5G Experience, hingga Cyber Security Hub tidak dibangun dengan membebani belanja modal APBD Kota Surakarta. Semuanya hadir melalui kemitraan murni Public-Private Partnership (PPP).', {
    x: 6.7, y: 5.92, w: 5.4, h: 0.65,
    fontFace: C.fontBody, fontSize: 9.0, color: C.textMain, lineSpacingMultiple: 1.18
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 08: DIAGNOSA 5 GAP & SOLUSI TEROBOSAN
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // REFLEKSI STRATEGIS & SOLUSI KONKRET',
    title: 'DIAGNOSA & 5 GAP UTAMA BESERTA SOLUSI TEROBOSANNYA',
    subtitle: 'Kajian Kritis Berdasarkan Dokumen Strategi STP 2026: Gap Riil, Risiko, dan Solusi Kebijakan Kepala Daerah',
    pageNum: 8
  });
  
  const gaps = [
    {
      num: 'GAP 01',
      problem: 'Jenis Pelatihan Masih Terbatas',
      probDetail: 'Akar Masalah: Keterbatasan instruktur ahli internal & sarana diklat berbasis industri masa depan.',
      solution: 'Gandeng Korporasi Global (OGSCI, Kemenperin) Buka Kelas Spesialis',
      solDetail: 'Welder Bawah Air, AI Engineer, Teknisi Baterai EV, & Drone Operator.',
      isPriority: false
    },
    {
      num: 'GAP 02',
      problem: 'R&D Belum Mengarah Hilirisasi',
      probDetail: 'Akar Masalah: Riset kampus berjalan terisolasi, kurang terpetakan, dan belum berorientasi pasar.',
      solution: 'Bentuk Technology Transfer Office (TTO) & Kurasi Riset Terapan Kampus',
      solDetail: 'Konsorsium 5 Perguruan Tinggi Solo Raya (UNS, UMS, ISI, dll.) berbasis problem kota.',
      isPriority: false
    },
    {
      num: 'GAP 03',
      problem: 'Investasi Kawasan Terbatas',
      probDetail: 'Akar Masalah: Lahan kosong belum optimal dikerjasamakan; promosi investor belum sistematis.',
      solution: 'Zonasi Investasi Teknologi Terpadu & Long-Term Revenue Sharing',
      solDetail: 'Skema kemitraan 10–25 tahun bagi korporasi teknologi global tanpa beban belanja APBD.',
      isPriority: false
    },
    {
      num: 'GAP 04',
      problem: 'Kemandirian Finansial Rendah',
      probDetail: 'Akar Masalah: Pendapatan masih bertumpu sewa lahan; hardware lab belum optimal menghasilkan jasa.',
      solution: 'Diversifikasi 6 Revenue Stream Mandiri & Optimalisasi Aset Industri',
      solDetail: 'Uji NDT penetrant, TUK/JFT mandiri, rekayasa manufaktur, co-working, & startup binaan.',
      isPriority: false
    },
    {
      num: 'GAP 05',
      problem: 'Promosi & Penyerapan Rendah',
      probDetail: 'Akar Masalah: Networking kurang ter-maintenance; belum ada regulasi penyerapan inovasi lokal.',
      solution: 'Penerbitan Instruksi Walikota untuk e-Katalog Lokal Afirmatif',
      solDetail: 'Wajibkan OPD Pemkot menyerap produk/layanan tenant binaan STP (Target: Rp 2,25 Miliar).',
      isPriority: true
    }
  ];
  
  const rowH = 0.85;
  const rowGap = 0.12;
  const startY = 1.68;
  
  gaps.forEach((g, idx) => {
    const y = startY + idx * (rowH + rowGap);
    
    // Left Box: Problem / GAP
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y, w: 4.8, h: rowH,
      fill: { color: g.isPriority ? C.bgTintAmber : C.bg },
      line: { color: g.isPriority ? C.orange : C.border, width: g.isPriority ? 1.5 : 1 }
    });
    
    addIconBadge(slide, {
      iconName: 'alertTriangle', x: 0.95, y: y + 0.18, size: 0.38,
      bgTint: C.bgTintRed, strokeColor: C.red, borderTint: C.borderRed
    });
    
    slide.addText(g.num, {
      x: 1.45, y: y + 0.16, w: 0.8, h: 0.2,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: C.red
    });
    slide.addText(g.problem, {
      x: 2.2, y: y + 0.16, w: 3.3, h: 0.22,
      fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: C.textMain
    });
    slide.addText(g.probDetail, {
      x: 1.0, y: y + 0.48, w: 4.4, h: 0.32,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textDim
    });
    
    // Middle Connector Arrow
    slide.addShape(pptx.shapes.RIGHT_ARROW, {
      x: 5.75, y: y + 0.28, w: 0.28, h: 0.25,
      fill: { color: g.isPriority ? C.orange : C.navy }, line: { color: g.isPriority ? C.orange : C.navy, width: 0 }
    });
    
    // Right Box: Solution
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.15, y, w: 6.383, h: rowH,
      fill: { color: g.isPriority ? C.bgTintAmber : C.bg },
      line: { color: g.isPriority ? C.orange : C.border, width: g.isPriority ? 1.5 : 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.15, y, w: 6.383, h: 0.035,
      fill: { color: g.isPriority ? C.orange : (idx === 3 ? C.emerald : C.navy) }, line: { color: g.isPriority ? C.orange : C.navy, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: g.isPriority ? 'zap' : 'checkCircle', x: 6.3, y: y + 0.18, size: 0.38,
      bgTint: g.isPriority ? C.bgTintAmber : C.bgTintGreen,
      strokeColor: g.isPriority ? C.orange : C.emerald,
      borderTint: g.isPriority ? C.borderAmber : C.borderGreen
    });
    
    slide.addText(g.solution, {
      x: 6.8, y: y + 0.16, w: 5.6, h: 0.22,
      fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: g.isPriority ? C.orange : C.navy
    });
    slide.addText(g.solDetail, {
      x: 6.8, y: y + 0.46, w: 5.6, h: 0.32,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted
    });
  });
  
  // Highlight Quick Win 30 Hari Banner at Bottom
  const bannerY = startY + 5 * (rowH + rowGap) + 0.05;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: bannerY, w: 11.733, h: 0.58,
    fill: { color: C.bgTintAmber },
    line: { color: C.orange, width: 1.5 }
  });
  addIconBadge(slide, {
    iconName: 'zap', x: 0.95, y: bannerY + 0.1, size: 0.38,
    bgTint: C.bg, strokeColor: C.orange, borderTint: C.borderAmber
  });
  slide.addText('PRIORITAS UTAMA #5: INSTRUKSI WALIKOTA e-KATALOG LOKAL AFIRMATIF (QUICK WIN 30 HARI)', {
    x: 1.45, y: bannerY + 0.08, w: 10.8, h: 0.22,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.orange
  });
  slide.addText('Target: Membuka captive market Rp 2,25 Miliar serapan belanja OPD Pemkot Surakarta di 2030 bagi 80+ produk tenant dan startup binaan STP tanpa perlu alokasi pagu belanja baru di APBD. Hanya butuh 1 lembar Instruksi Walikota.', {
    x: 1.45, y: bannerY + 0.3, w: 10.8, h: 0.24,
    fontFace: C.fontBody, fontSize: 8.5, color: C.textMain
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 09: ROADMAP 3 FASE TRANSFORMASI (2026-2030)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // PETA JALAN STRATEGIS KELEMBAGAAN',
    title: 'ROADMAP 3 FASE TRANSFORMASI: DARI CENTRE MENUJU HOLDING INOVASI',
    subtitle: 'Milestone Terukur dan Bertahap Menuju Kemandirian Finansial Penuh Tanpa Subsidi Belanja APBD 2030',
    pageNum: 9
  });
  
  const phases = [
    {
      phase: 'FASE 1 (2026–2027)',
      title: 'CENTRE OF INNOVATION & CONNECTIVITY',
      target: 'Target: Rp 7,50 M ➔ Rp 12,53 M (+67%)',
      color: C.navy, icon: 'sparkles',
      items: [
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
      target: 'Target: Rp 12,53 M ➔ Rp 20,62 M (+65%)',
      color: C.amber, icon: 'layers',
      items: [
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
      target: 'Target: Rp 20,62 M ➔ Rp 23,47 M (MANDIRI)',
      color: C.emerald, icon: 'award',
      items: [
        'Kemandirian finansial penuh: 0% subsidi belanja APBD',
        'Operasional holding company inovasi daerah pertama',
        'Pusat pengembangan Indonesia Digital Technopark',
        'Realisasi target pendapatan mandiri Rp 23,47 Miliar',
        'Ekspansi jejaring kemitraan dan penempatan global'
      ]
    }
  ];
  
  const phW = 3.65;
  const phGap = 0.39;
  phases.forEach((p, i) => {
    const phX = 0.8 + i * (phW + phGap);
    const phY = 1.68;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX, y: phY, w: phW, h: 3.95,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX, y: phY, w: phW, h: 0.05,
      fill: { color: p.color }, line: { color: p.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: p.icon, x: phX + 0.25, y: phY + 0.2, size: 0.44,
      bgTint: C.bgSoft, strokeColor: p.color, borderTint: C.border
    });
    
    slide.addText(p.phase, {
      x: phX + 0.8, y: phY + 0.22, w: phW - 0.95, h: 0.18,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: p.color, charSpacing: 1
    });
    slide.addText(p.title, {
      x: phX + 0.8, y: phY + 0.42, w: phW - 0.95, h: 0.38,
      fontFace: C.fontTitle, fontSize: 11.5, bold: true, color: C.textMain
    });
    
    // Target Badge
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: phX + 0.2, y: phY + 0.9, w: phW - 0.4, h: 0.35,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(p.target, {
      x: phX + 0.2, y: phY + 0.92, w: phW - 0.4, h: 0.3,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: p.color, align: 'center'
    });
    
    // Items
    p.items.forEach((it, idx) => {
      slide.addImage({
        data: `image/svg+xml;base64,${getSvgBase64('checkCircle', p.color, 2)}`,
        x: phX + 0.25, y: phY + 1.45 + idx * 0.48, w: 0.16, h: 0.16
      });
      slide.addText(it, {
        x: phX + 0.52, y: phY + 1.42 + idx * 0.48, w: phW - 0.75, h: 0.44,
        fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.15
      });
    });
    
    if (i < 2) {
      slide.addShape(pptx.shapes.RIGHT_ARROW, {
        x: phX + phW + 0.1, y: phY + 1.65, w: 0.19, h: 0.28,
        fill: { color: C.amber }, line: { color: C.amber, width: 0 }
      });
    }
  });
  
  // Bottom Table: Yearly Targets
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 5.75, w: 11.733, h: 1.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  
  slide.addText('RINGKASAN TARGET TAHUNAN MENUJU KEMANDIRIAN 2030 (AUDIT FINANSIAL RIIL):', {
    x: 1.0, y: 5.85, w: 11.3, h: 0.22,
    fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: C.navy
  });
  
  const trajTable = [
    [
      { text: 'TAHUN 2026 (Baseline)', options: { bold: true, fontFace: C.fontTitle, fontSize: 8.5, align: 'center', fill: { color: 'F1F5F9' } } },
      { text: 'TAHUN 2027 (+67,0%)', options: { bold: true, fontFace: C.fontTitle, fontSize: 8.5, align: 'center', fill: { color: 'F1F5F9' }, color: C.navy } },
      { text: 'TAHUN 2028 (+28,5%)', options: { bold: true, fontFace: C.fontTitle, fontSize: 8.5, align: 'center', fill: { color: 'F1F5F9' }, color: C.navy } },
      { text: 'TAHUN 2029 (+28,0%)', options: { bold: true, fontFace: C.fontTitle, fontSize: 8.5, align: 'center', fill: { color: 'F1F5F9' }, color: C.navy } },
      { text: 'TAHUN 2030 (MANDIRI)', options: { bold: true, fontFace: C.fontTitle, fontSize: 8.5, align: 'center', fill: { color: C.bgTintAmber }, color: C.amber } }
    ],
    [
      { text: 'Rp 7.503.392.698', options: { fontFace: C.fontBody, fontSize: 9.5, align: 'center' } },
      { text: 'Rp 12.532.185.031', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.navy, align: 'center' } },
      { text: 'Rp 16.100.658.831', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.navy, align: 'center' } },
      { text: 'Rp 20.615.658.831', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.navy, align: 'center' } },
      { text: 'Rp 23.468.158.831', options: { fontFace: C.fontBody, fontSize: 10.5, bold: true, color: C.amber, align: 'center', fill: { color: C.bgTintAmber } } }
    ]
  ];
  
  slide.addTable(trajTable, {
    x: 1.0, y: 6.12, w: 11.333, h: 0.65,
    colW: [2.26, 2.26, 2.26, 2.26, 2.29],
    border: { pt: 0.5, color: C.border }
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 10: RE-SETTING TARGET FINANSIAL (EXPANDED FOCUS COLUMN CHART)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // PROYEKSI KEUANGAN BLUD 2026–2030',
    title: 'RE-SETTING TARGET FINANSIAL: LOMPATAN KEMANDIRIAN BLUD +212,8%',
    subtitle: 'Proyeksi Pendapatan Mandiri Riil dari Rp 7,50 Miliar (2026) menuju Rp 23,47 Miliar (2030) — Tanpa Subsidi APBD',
    pageNum: 10
  });
  
  // Left: Large Column Chart (w: 6.8)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 6.9, h: 4.45,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 6.9, h: 0.045,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'trendingUp', x: 1.0, y: 1.84, size: 0.38,
    bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('TREN PERTUMBUHAN PENDAPATAN BLUD 2026–2030 (DALAM RP MILIAR)', {
    x: 1.5, y: 1.9, w: 6.0, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.navy
  });
  
  // Native Column Chart Full Scale
  slide.addChart(pptx.charts.BAR, [
    {
      name: 'Target Pendapatan (Rp M)',
      labels: ['2026\n(Baseline)', '2027\n(+67%)', '2028\n(+28,5%)', '2029\n(+28%)', '2030\n(MANDIRI)'],
      values: [7.50, 12.53, 16.10, 20.62, 23.47]
    }
  ], {
    x: 1.0, y: 2.25, w: 6.5, h: 3.65,
    barDir: 'col',
    chartColors: [C.sky, C.navy, C.navy, C.emerald, C.amber],
    showValue: true,
    dataLabelColor: C.textMain,
    dataLabelFontFace: C.fontTitle,
    dataLabelFontSize: 12.0,
    dataLabelFontBold: true,
    dataLabelPosition: 'outEnd',
    valGridLine: { color: C.border, width: 0.5 },
    catGridLine: { style: 'none' },
    valAxisMinVal: 0,
    valAxisMaxVal: 26,
    valAxisMajorUnit: 5
  });
  
  // Right: 3 Financial Analysis Cards (w: 4.5)
  const finCards = [
    {
      title: 'TITIK IMPAS OPERASIONAL (BEP 2028)',
      figure: 'Rp 16,10 Miliar',
      desc: 'Tercapai di tahun ke-3 seiring beroperasinya Strategic Business Unit (SBU) manufaktur presisi & serapan pasar produk hilirisasi.',
      icon: 'award', color: C.blue
    },
    {
      title: 'KEMANDIRIAN 100% BEBAS APBD (2030)',
      figure: 'Rp 23,47 Miliar',
      desc: 'Seluruh belanja operasional UPTD STP tertutup 100% dari jasa layanan BLUD mandiri tanpa membebani kas APBD Kota Surakarta.',
      icon: 'shieldCheck', color: C.emerald
    },
    {
      title: 'KUALITAS PENDAPATAN BERKELANJUTAN',
      figure: '88% Produktif',
      desc: '88% pendapatan dihasilkan dari riset vokasi industri & startup binaan, menggantikan paradigma lama yang bertumpu pada sewa lahan pasif.',
      icon: 'trendingUp', color: C.amber
    }
  ];
  
  const fcW = 4.533;
  const fcH = 1.35;
  finCards.forEach((fc, fci) => {
    const fcY = 1.68 + fci * (fcH + 0.2);
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 8.0, y: fcY, w: fcW, h: fcH,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 8.0, y: fcY, w: fcW, h: 0.045,
      fill: { color: fc.color }, line: { color: fc.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: fc.icon, x: 8.2, y: fcY + 0.18, size: 0.42,
      bgTint: C.bgSoft, strokeColor: fc.color, borderTint: C.border
    });
    
    slide.addText(fc.title, {
      x: 8.75, y: fcY + 0.16, w: fcW - 0.95, h: 0.2,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: fc.color
    });
    slide.addText(fc.figure, {
      x: 8.75, y: fcY + 0.36, w: fcW - 0.95, h: 0.32,
      fontFace: C.fontTitle, fontSize: 13.0, bold: true, color: C.textMain
    });
    slide.addText(fc.desc, {
      x: 8.2, y: fcY + 0.72, w: fcW - 0.4, h: 0.55,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMuted, lineSpacingMultiple: 1.18
    });
  });
  
  // Bottom Ribbon Full
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.25, w: 11.733, h: 0.62,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('KESIMPULAN EKSEKUTIF: Lompatan +212,8% bukan angka angan-angan, melainkan hasil rekonsiliasi berbasis kontrak riil bersama konsorsium industri, hilirisasi produk siap e-Katalog, serta optimalisasi aset BLUD yang akuntabel.', {
    x: 1.0, y: 6.3, w: 11.333, h: 0.5,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textWhite, align: 'center', lineSpacingMultiple: 1.2
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 11: STRUKTUR 6 KLASTER REVENUE & TOP KATALISATOR (NEW EXPANDED)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // STRUKTUR PENDAPATAN MANDIRI 2030',
    title: 'STRUKTUR 6 KLASTER REVENUE STREAM 2030 & TOP 5 KATALISATOR',
    subtitle: 'Diversifikasi Portofolio Pendapatan Sehat Menuju Target Rp 23,47 Miliar per Tahun',
    pageNum: 11
  });
  
  // Left: Native Doughnut Chart (w: 5.4)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 5.4, h: 5.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 5.4, h: 0.045,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'pieChart', x: 1.0, y: 1.84, size: 0.38,
    bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber
  });
  slide.addText('STRUKTUR REVENUE STREAM 2030 (RP 23,47 M)', {
    x: 1.5, y: 1.9, w: 4.5, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.amber
  });
  
  slide.addChart(pptx.charts.DOUGHNUT, [
    {
      name: 'Klaster Pendapatan',
      labels: ['Program Kemitraan', 'Pelatihan & Sertifikasi', 'Diklat Mandiri Spesialis', 'Lain-lain BLUD / MICE', 'Rekayasa Manufaktur', 'Kemitraan Lahan Eksisting'],
      values: [12.83, 3.68, 2.47, 2.45, 1.47, 0.58]
    }
  ], {
    x: 1.0, y: 2.22, w: 5.0, h: 2.7,
    holeSize: 58,
    chartColors: [C.sky, C.navy, C.orange, C.emerald, C.purple, C.textDim],
    showLegend: false,
    showLabel: false,
    showValue: false
  });
  
  // Legend Cards for Revenue Clusters
  const revLegs = [
    { name: 'Kemitraan Strategis: 54,6% (Rp 12,83 M)', color: C.sky },
    { name: 'Lain-lain BLUD / MICE: 10,4% (Rp 2,45 M)', color: C.emerald },
    { name: 'Pelatihan & Sertifikasi: 15,7% (Rp 3,68 M)', color: C.navy },
    { name: 'Rekayasa Manufaktur: 6,2% (Rp 1,47 M)', color: C.purple },
    { name: 'Diklat Mandiri Spesialis: 10,5% (Rp 2,47 M)', color: C.orange },
    { name: 'Kemitraan Lahan: 2,5% (Rp 0,58 M)', color: C.textDim }
  ];
  
  revLegs.forEach((rl, rIdx) => {
    const rlCol = rIdx % 2;
    const rlRow = Math.floor(rIdx / 2);
    const rlX = 1.0 + rlCol * 2.65;
    const rlY = 5.05 + rlRow * 0.45;
    
    slide.addShape(pptx.shapes.OVAL, {
      x: rlX, y: rlY + 0.04, w: 0.16, h: 0.16,
      fill: { color: rl.color }, line: { color: rl.color, width: 0 }
    });
    slide.addText(rl.name, {
      x: rlX + 0.24, y: rlY, w: 2.38, h: 0.4,
      fontFace: C.fontBody, fontSize: 9.0, color: C.textMain, lineSpacingMultiple: 1.15
    });
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 6.35, w: 5.0, h: 0.38,
    fill: { color: C.bgSoft }, line: { color: C.border, width: 1 }
  });
  slide.addText('STRUKTUR SEHAT: 88% disumbang riset & inovasi, sewa lahan pasif hanya 2,5%.', {
    x: 1.1, y: 6.38, w: 4.8, h: 0.3,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.textMuted, align: 'center'
  });
  
  // Right: Top 5 Catalysts Table (w: 6.0)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.5, y: 1.68, w: 6.033, h: 5.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.5, y: 1.68, w: 6.033, h: 0.045,
    fill: { color: C.emerald }, line: { color: C.emerald, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'award', x: 6.7, y: 1.84, size: 0.38,
    bgTint: C.bgTintGreen, strokeColor: C.emerald, borderTint: C.borderGreen
  });
  slide.addText('TOP 5 KATALISATOR PENDAPATAN MANDIRI 2030', {
    x: 7.2, y: 1.9, w: 5.1, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.5, bold: true, color: C.emerald
  });
  
  const catItems = [
    { rank: '#1', title: 'OGSCI Program Migas & Maritim', nominal: 'Rp 6.250.000.000', note: 'Anchor Revenue Mandiri: Pelatihan spesialis migas & sertifikasi internasional' },
    { rank: '#2', title: 'Pemda Solo (e-Katalog Inovasi)', nominal: 'Rp 2.250.000.000', note: 'Instruksi Walikota: Penyerapan produk teknologi & layanan tenant oleh OPD' },
    { rank: '#3', title: 'Startup Binaan STP (Komersialisasi)', nominal: 'Rp 2.000.000.000', note: 'Success Fee & Jasa Alih Teknologi dari hasil inkubasi hilirisasi riset' },
    { rank: '#4', title: 'Sertifikasi Tenaga JFT Fungsional', nominal: 'Rp 1.375.000.000', note: 'Program sertifikasi kompetensi ASN & instruktur teknis se-Jawa Tengah' },
    { rank: '#5', title: 'Tempat Uji Kompetensi (TUK) Mandiri', nominal: 'Rp 1.350.000.000', note: 'Uji sertifikasi BNSP bidang pengelasan, manufaktur, dan cyber security' }
  ];
  
  catItems.forEach((ct, ci) => {
    const ctY = 2.28 + ci * 0.82;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 6.7, y: ctY, w: 5.633, h: 0.72,
      fill: { color: ci === 0 ? C.bgTintAmber : (ci === 1 ? C.bgTintBlue : C.bgSoft) },
      line: { color: ci === 0 ? C.borderAmber : (ci === 1 ? C.borderBlue : C.border), width: 1 }
    });
    
    slide.addText(ct.rank, {
      x: 6.8, y: ctY + 0.12, w: 0.5, h: 0.45,
      fontFace: C.fontTitle, fontSize: 14.0, bold: true, color: ci === 0 ? C.amber : C.navy
    });
    
    slide.addText(ct.title, {
      x: 7.35, y: ctY + 0.1, w: 3.2, h: 0.24,
      fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: C.textMain
    });
    slide.addText(ct.note, {
      x: 7.35, y: ctY + 0.34, w: 3.2, h: 0.32,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textDim
    });
    
    slide.addText(ct.nominal, {
      x: 10.2, y: ctY + 0.18, w: 2.0, h: 0.35,
      fontFace: C.fontTitle, fontSize: 11.5, bold: true, color: ci === 0 ? C.amber : (ci === 1 ? C.blue : C.emerald), align: 'right'
    });
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 6.7, y: 6.38, w: 5.633, h: 0.35,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('TOTAL TOP 5 KATALISATOR: Rp 13,22 MILIAR (56,3% DARI TOTAL TARGET 2030)', {
    x: 6.8, y: 6.42, w: 5.433, h: 0.26,
    fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 12: ANATOMY MESIN PENDAPATAN 2030 (EXECUTIVE FINANCIAL LEDGER)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // TRANSPARANSI ANGGARAN & STRUKTUR PENDAPATAN',
    title: 'ANATOMY MESIN PENDAPATAN 2030: RINCIAN PER MATA ANGGARAN',
    subtitle: 'Transparansi Total Hasil Rekonsiliasi: Dari Mana Setiap Rupiah dari Target Rp 23,47 Miliar Berasal',
    pageNum: 12
  });
  
  // Top Proportional Ribbon
  const ribbonData = [
    { label: 'Klaster A: Kemitraan Strategis (54,6%)', pct: 0.546, color: C.blue },
    { label: 'B: Vokasi (15,7%)', pct: 0.157, color: C.navy },
    { label: 'C: Diklat (10,5%)', pct: 0.105, color: C.amber },
    { label: 'D: MICE (10,4%)', pct: 0.104, color: C.emerald },
    { label: 'E: 6,2%', pct: 0.062, color: C.purple },
    { label: 'F: 2,5%', pct: 0.026, color: C.textDim }
  ];
  
  let currRibX = 0.8;
  const totalRibW = 11.733;
  ribbonData.forEach((rb) => {
    const segW = totalRibW * rb.pct;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: currRibX, y: 1.62, w: segW, h: 0.28,
      fill: { color: rb.color }, line: { color: rb.color, width: 0 }
    });
    if (segW > 0.8) {
      slide.addText(rb.label, {
        x: currRibX, y: 1.65, w: segW, h: 0.22,
        fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.textWhite, align: 'center'
      });
    }
    currRibX += segW;
  });
  
  // Left: Hero Panel Klaster A (w: 4.8)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.98, w: 4.8, h: 4.45,
    fill: { color: C.bg },
    line: { color: C.blue, width: 1.5 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.98, w: 4.8, h: 0.05,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'anchor', x: 1.0, y: 2.15, size: 0.44,
    bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('ENGINE UTAMA (54,6% REVENUE)\nKLASTER A: KERJASAMA STRATEGIS', {
    x: 1.55, y: 2.15, w: 2.2, h: 0.5,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.navy
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 3.8, y: 2.12, w: 1.65, h: 0.56,
    fill: { color: C.bgTintAmber },
    line: { color: C.borderAmber, width: 1 }
  });
  slide.addText('Rp 12.825.000.000\n54,65% DARI TOTAL TARGET', {
    x: 3.8, y: 2.16, w: 1.65, h: 0.48,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber, align: 'center'
  });
  
  const progA = [
    { name: 'OGSCI Migas & Maritim (Konsorsium)', sub: '[Anchor Revenue Mandiri]', val: 'Rp 6.250.000.000', col: C.amber },
    { name: 'Pemda Solo (e-Katalog Inovasi)', sub: '[Instruksi Walikota]', val: 'Rp 2.250.000.000', col: C.blue },
    { name: 'Startup Binaan STP (Komersialisasi)', sub: '[Success Fee & Jasa]', val: 'Rp 2.000.000.000', col: C.blue },
    { name: 'BP3MI - SMK GoGlobal (Talenta Luar Negeri)', sub: '[Penempatan Global]', val: 'Rp 1.500.000.000', col: C.blue },
    { name: 'Kemenperin & Mitra LPK Industri', sub: '[Subsidi Pelatihan]', val: 'Rp 825.000.000', col: C.blue }
  ];
  
  progA.forEach((pr, pi) => {
    const prY = 2.8 + pi * 0.62;
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 1.0, y: prY, w: 4.4, h: 0.56,
      fill: { color: pi === 0 ? C.bgTintAmber : C.bgSoft },
      line: { color: pi === 0 ? C.borderAmber : C.border, width: 1 }
    });
    slide.addText(pr.name, {
      x: 1.15, y: prY + 0.08, w: 2.6, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: C.textMain
    });
    slide.addText(pr.sub, {
      x: 1.15, y: prY + 0.3, w: 2.6, h: 0.2,
      fontFace: C.fontBody, fontSize: 8.0, italic: true, color: C.textDim
    });
    slide.addText(pr.val, {
      x: 3.2, y: prY + 0.16, w: 2.1, h: 0.26,
      fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: pr.col, align: 'right'
    });
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 5.95, w: 4.4, h: 0.38,
    fill: { color: C.bgTintBlue }, line: { color: C.borderBlue, width: 0.8 }
  });
  slide.addText('CATATAN STRATEGIS: 1 program konsorsium OGSCI (Rp 6,25 M) setara 83% pendapatan STP tahun 2026. Ditambah e-Katalog Pemda, klaster ini menjamin kemandirian tanpa APBD.', {
    x: 1.05, y: 5.98, w: 4.3, h: 0.32,
    fontFace: C.fontBody, fontSize: 7.5, color: C.blue, lineSpacingMultiple: 1.1
  });
  
  // Right Top: Klaster B & Klaster C (w: 3.3 each)
  // Klaster B
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.8, y: 1.98, w: 3.25, h: 2.15,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.8, y: 1.98, w: 3.25, h: 0.04,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'graduationCap', x: 5.95, y: 2.1, size: 0.36,
    bgTint: C.bgTintBlue, strokeColor: C.navy, borderTint: C.borderBlue
  });
  slide.addText('KLASTER B: VOKASI & SERTIFIKASI\nRp 3.679.800.000 (15,7%)', {
    x: 6.4, y: 2.1, w: 2.55, h: 0.36,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy
  });
  
  const bItems = [
    ['Sertifikasi JFT Fungsional', 'Rp 1.375.000.000'],
    ['TUK Mandiri Terakreditasi', 'Rp 1.350.000.000'],
    ['Konsultasi Lab & Data Center', 'Rp 950.000.000'],
    ['Prakerin SMK Solo Raya', 'Rp 4.800.000']
  ];
  bItems.forEach((b, bi) => {
    const bY = 2.55 + bi * 0.38;
    slide.addText(b[0], {
      x: 6.0, y: bY, w: 1.9, h: 0.22,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textMain
    });
    slide.addText(b[1], {
      x: 7.7, y: bY, w: 1.25, h: 0.22,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.navy, align: 'right'
    });
  });
  
  // Klaster C
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.25, y: 1.98, w: 3.283, h: 2.15,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 9.25, y: 1.98, w: 3.283, h: 0.04,
    fill: { color: C.amber }, line: { color: C.amber, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'wrench', x: 9.4, y: 2.1, size: 0.36,
    bgTint: C.bgTintAmber, strokeColor: C.amber, borderTint: C.borderAmber
  });
  slide.addText('KLASTER C: DIKLAT SPESIALIS\nRp 2.469.000.000 (10,5%)', {
    x: 9.85, y: 2.1, w: 2.583, h: 0.36,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber
  });
  
  const cItems = [
    ['Underwater Wet Welding', 'Rp 846.000.000'],
    ['Manufaktur & Welding Intensif', 'Rp 728.000.000'],
    ['Desain, Manajerial & EV', 'Rp 555.000.000'],
    ['Digital Tech (AI & Cyber)', 'Rp 225.000.000'],
    ['Kewirausahaan & UMKM', 'Rp 115.000.000']
  ];
  cItems.forEach((c, ci) => {
    const cY = 2.52 + ci * 0.31;
    slide.addText(c[0], {
      x: 9.45, y: cY, w: 1.9, h: 0.2,
      fontFace: C.fontBody, fontSize: 8.0, color: C.textMain
    });
    slide.addText(c[1], {
      x: 11.2, y: cY, w: 1.25, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8.0, bold: true, color: C.amber, align: 'right'
    });
  });
  
  // Right Bottom: Klaster D, E, F (w: 2.16 each)
  // Klaster D
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.8, y: 4.25, w: 2.1, h: 2.18,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 5.8, y: 4.25, w: 2.1, h: 0.035,
    fill: { color: C.emerald }, line: { color: C.emerald, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'presentation', x: 5.92, y: 4.35, size: 0.32,
    bgTint: C.bgTintGreen, strokeColor: C.emerald, borderTint: C.borderGreen
  });
  slide.addText('KLASTER D: MICE\nRp 2.450.000.000 (10,4%)', {
    x: 6.32, y: 4.35, w: 1.5, h: 0.32,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.emerald
  });
  slide.addText('• Sarana/MICE: Rp 1,50 M\n• Co-Working: Rp 900 Jt\n• Jasa Giro Kas: Rp 50 Jt', {
    x: 5.92, y: 5.0, w: 1.9, h: 1.2,
    fontFace: C.fontBody, fontSize: 8.0, color: C.textMuted, lineSpacingMultiple: 1.2
  });
  
  // Klaster E
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 8.05, y: 4.25, w: 2.1, h: 2.18,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 8.05, y: 4.25, w: 2.1, h: 0.035,
    fill: { color: C.purple }, line: { color: C.purple, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'flask', x: 8.17, y: 4.35, size: 0.32,
    bgTint: C.bgTintPurple, strokeColor: C.purple, borderTint: 'E9D5FF'
  });
  slide.addText('KLASTER E: LAB & NDT\nRp 1.465.000.000 (6,2%)', {
    x: 8.57, y: 4.35, w: 1.5, h: 0.32,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.purple
  });
  slide.addText('• Fabrikasi Mesin: Rp 1,22 M\n• Uji Mutu SNI: Rp 135 Jt\n• Pesanan Pemda: Rp 105 Jt', {
    x: 8.17, y: 5.0, w: 1.9, h: 1.2,
    fontFace: C.fontBody, fontSize: 8.0, color: C.textMuted, lineSpacingMultiple: 1.2
  });
  
  // Klaster F
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.3, y: 4.25, w: 2.233, h: 2.18,
    fill: { color: C.bg }, line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 10.3, y: 4.25, w: 2.233, h: 0.035,
    fill: { color: C.textDim }, line: { color: C.textDim, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'mapPin', x: 10.42, y: 4.35, size: 0.32,
    bgTint: C.bgSoft, strokeColor: C.textDim, borderTint: C.border
  });
  slide.addText('KLASTER F: SEWA LAHAN\nRp 579.358.831 (2,5%)', {
    x: 10.82, y: 4.35, w: 1.6, h: 0.32,
    fontFace: C.fontTitle, fontSize: 7.5, bold: true, color: C.textDim
  });
  slide.addText('• Shopee Hub: Rp 544 Jt\n• Katulondi: Rp 35 Jt\n\nDE-ESKALASI SEWA:\nDitekan ke 2,5% demi beralih ke inovasi.', {
    x: 10.42, y: 4.95, w: 2.05, h: 1.35,
    fontFace: C.fontBody, fontSize: 8.0, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  // Bottom Full Lock Bar
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.55, w: 11.733, h: 0.45,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  addIconBadge(slide, {
    iconName: 'lock', x: 0.95, y: 6.6, size: 0.34,
    bgTint: C.blue, strokeColor: C.textWhite, borderTint: C.blue
  });
  slide.addText('TOTAL TARGET PENDAPATAN BLUD 2030: Rp 23.468.158.831 (Dua Puluh Tiga Miliar Empat Ratus Enam Puluh Delapan Juta Rupiah) — 100% Mandiri Bebas APBD', {
    x: 1.4, y: 6.62, w: 11.0, h: 0.32,
    fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.amber
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 13: 6 DIMENSI KEBERDAMPAKAN DAERAH
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // DAMPAK SOSIAL-EKONOMI DAERAH',
    title: '6 DIMENSI KEBERDAMPAKAN: SOLO TECHNOPARK UNTUK WARGA SURAKARTA',
    subtitle: '"Indikator Keberhasilan Harus Mengukur IMPACT — Bukan Sekadar Aktivitas dan Pendapatan Kas" (Dokumen Strategi STP 2026)',
    pageNum: 13
  });
  
  // Left: Horizontal Bar Chart of 6 Dimensions (w: 5.0)
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 5.0, h: 5.15,
    fill: { color: C.bg },
    line: { color: C.border, width: 1 }
  });
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.68, w: 5.0, h: 0.045,
    fill: { color: C.blue }, line: { color: C.blue, width: 0 }
  });
  
  addIconBadge(slide, {
    iconName: 'target', x: 1.0, y: 1.84, size: 0.38,
    bgTint: C.bgTintBlue, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('SKOR CAPAIAN 6 DIMENSI DAMPAK 2030 (%)', {
    x: 1.5, y: 1.9, w: 4.1, h: 0.26,
    fontFace: C.fontTitle, fontSize: 10.0, bold: true, color: C.navy
  });
  
  slide.addChart(pptx.charts.BAR, [
    {
      name: 'Skor Dampak (%)',
      labels: [
        'Hilirisasi Riset Kampus',
        'Peluang Kerja Daerah',
        'UMKM & Startup Mandiri',
        'Akses Fasilitas Publik',
        'SDM Kompeten & Sertifikasi',
        'Kolaborasi Industri Murni'
      ],
      values: [70, 75, 80, 85, 90, 100]
    }
  ], {
    x: 0.95, y: 2.25, w: 4.7, h: 3.55,
    barDir: 'bar',
    chartColors: [C.navy, C.navy, C.emerald, C.amber, C.orange, C.emerald],
    showValue: true,
    dataLabelColor: C.textMain,
    dataLabelFontFace: C.fontTitle,
    dataLabelFontSize: 11.0,
    dataLabelFontBold: true,
    valAxisMinVal: 0,
    valAxisMaxVal: 120,
    valAxisMajorUnit: 20
  });
  
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 1.0, y: 5.92, w: 4.6, h: 0.78,
    fill: { color: C.bgSoft },
    line: { color: C.border, width: 1 }
  });
  slide.addText('FILOSOFI IMPACT MEASUREMENT KOTA:\nKeberhasilan Solo Technopark tidak diukur dari megahnya gedung atau saldo kas BLUD semata, melainkan dari kesejahteraan nyata warga Surakarta: lapangan kerja yang terserap, UMKM yang naik kelas, serta keterhubungan riset kampus dengan kebutuhan industri.', {
    x: 1.15, y: 5.95, w: 4.3, h: 0.72,
    fontFace: C.fontBody, fontSize: 8.5, color: C.textMuted, lineSpacingMultiple: 1.15
  });
  
  // Right: 6 Scorecard KPI Cards
  const kpis = [
    { num: '9.000 Peserta/Thn', label: 'SDM LEBIH KOMPETEN', sub: 'Afirmasi kuota minimal 65% warga KTP Surakarta.', icon: 'graduationCap', color: C.blue },
    { num: '80+ Wirausaha', label: 'UMKM & STARTUP TUMBUH', sub: 'Pendampingan inkubasi & digital onboarding GoTo Hub.', icon: 'rocket', color: C.orange },
    { num: '45+ Produk', label: 'RISET DEKAT KE PASAR', sub: 'Paten HKI, Uji Presisi NDT & sertifikasi TKDN.', icon: 'microchip', color: C.emerald },
    { num: '10.000+ Visit/Thn', label: 'AKSES TEKNOLOGI TERBUKA', sub: 'Kunjungan lab AI, Cyber, & magang siswa SMK Solo Raya.', icon: 'users', color: C.sky },
    { num: 'Rp 180 M Multiplier', label: 'PELUANG KERJA BARU', sub: 'Estimasi perputaran ekonomi hotel, resto & MICE.', icon: 'trendingUp', color: C.purple },
    { num: '100% Swasta', label: 'KOLABORASI INDUSTRI MURNI', sub: '95+ korporasi dunia berinvestasi tanpa beban belanja APBD.', icon: 'layers', color: C.emerald }
  ];
  
  const kw = 3.15;
  const kh = 1.6;
  const kgX = 0.22;
  const kgY = 0.18;
  
  kpis.forEach((kp, idx) => {
    const cIdx = idx % 2;
    const rIdx = Math.floor(idx / 2);
    const kX = 6.1 + cIdx * (kw + kgX);
    const kY = 1.68 + rIdx * (kh + kgY);
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: kX, y: kY, w: kw, h: kh,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: kX, y: kY, w: kw, h: 0.045,
      fill: { color: kp.color }, line: { color: kp.color, width: 0 }
    });
    
    slide.addText(kp.label, {
      x: kX + 0.2, y: kY + 0.18, w: kw - 0.9, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: kp.color, charSpacing: 0.8
    });
    
    addIconBadge(slide, {
      iconName: kp.icon, x: kX + kw - 0.55, y: kY + 0.18, size: 0.38,
      bgTint: C.bgSoft, strokeColor: kp.color, borderTint: C.border
    });
    
    slide.addText(kp.num, {
      x: kX + 0.2, y: kY + 0.45, w: kw - 0.4, h: 0.45,
      fontFace: C.fontTitle, fontSize: 16.0, bold: true, color: C.textMain
    });
    
    slide.addText(kp.sub, {
      x: kX + 0.2, y: kY + 0.95, w: kw - 0.4, h: 0.52,
      fontFace: C.fontBody, fontSize: 8.5, color: C.textDim, lineSpacingMultiple: 1.15
    });
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 14: TIGA PERMOHONAN KEBIJAKAN KEPADA WALIKOTA (EXPANDED FOCUS)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // DUKUNGAN KEPEMIMPINAN & PAYUNG REGULASI',
    title: '3 PERMOHONAN KEBIJAKAN NON-ANGGARAN KEPADA WALIKOTA SURAKARTA',
    subtitle: 'Kami Tidak Memohon Tambahan Pagu APBD — Kami Memohon 3 Payung Regulasi Kepemimpinan Walikota',
    pageNum: 14
  });
  
  const asks = [
    {
      badge: 'PRIORITAS UTAMA // QUICK WIN 30 HARI',
      num: 'PERMOHONAN 1',
      title: 'INSTRUKSI WALIKOTA AFIRMASI e-KATALOG LOKAL',
      color: C.blue, icon: 'fileText',
      narrative: 'Mengarahkan seluruh OPD Pemkot Surakarta untuk memprioritaskan belanja produk teknologi, layanan digital, dan barang hasil riset tenant binaan STP melalui etalase e-Katalog Lokal.',
      targetTitle: 'TARGET DAMPAK KONKRET:',
      target: 'Membuka captive market pasti senilai Rp 2,25 Miliar/Tahun di 2030 bagi 80+ produk inovasi tenant tanpa menambah anggaran APBD.',
      mechanismTitle: 'MEKANISME REGULASI:',
      mechanism: 'Penerbitan 1 lembar Instruksi Walikota. Draf telah siap harmonisasi bersama Bagian Hukum Setda & BPKAD. Siap tuntas dalam 30 hari pertama.'
    },
    {
      badge: 'INVESTASI JANGKA PANJANG (10–25 TAHUN)',
      num: 'PERMOHONAN 2',
      title: 'REGULASI FLEKSIBILITAS KERJASAMA LAHAN BLUD',
      color: C.amber, icon: 'scale',
      narrative: 'Penyusunan Peraturan Walikota yang memberikan kepastian hukum skema Long-Term Revenue Sharing dan sewa lahan jangka panjang (10–25 tahun) yang kompetitif bagi investor teknologi global di kawasan STP.',
      targetTitle: 'TARGET DAMPAK KONKRET:',
      target: 'Menarik 3–5 investor korporasi baru kelas dunia tanpa membebani belanja modal APBD daerah (Prinsip Zero APBD Burden).',
      mechanismTitle: 'MEKANISME REGULASI:',
      mechanism: 'Harmonisasi Perwali Pola Kerjasama Lahan bersama BPKAD, Bagian Hukum, dan Inspektorat Kota Surakarta.'
    },
    {
      badge: 'SINERGI KAMPUS SOLO RAYA',
      num: 'PERMOHONAN 3',
      title: 'SURAT EDARAN KONSORSIUM RISET DAERAH',
      color: C.emerald, icon: 'graduationCap',
      narrative: 'Mendorong 5 perguruan tinggi ternama (UNS, UMS, ISI, Unisri, Poltek) mengarahkan minimal 30% dana riset terapan dosen dan tugas akhir mahasiswa untuk memecahkan problem kota dan industri di STP.',
      targetTitle: 'TARGET DAMPAK KONKRET:',
      target: '5 Kampus mitra aktif, 30 klaster riset terapan berbasis problem kota, dan minimal 10 paten terhilirisasi ke industri per tahun.',
      mechanismTitle: 'MEKANISME REGULASI:',
      mechanism: 'Penerbitan Surat Edaran Walikota Surakarta dan penandatanganan kesepakatan bersama (MoU) 5 Rektor Kampus se-Solo Raya.'
    }
  ];
  
  const aW = 3.65;
  const aGap = 0.39;
  asks.forEach((a, i) => {
    const aX = 0.8 + i * (aW + aGap);
    const aY = 1.68;
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX, y: aY, w: aW, h: 4.65,
      fill: { color: C.bg },
      line: { color: C.border, width: 1 }
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX, y: aY, w: aW, h: 0.05,
      fill: { color: a.color }, line: { color: a.color, width: 0 }
    });
    
    addIconBadge(slide, {
      iconName: a.icon, x: aX + 0.25, y: aY + 0.2, size: 0.44,
      bgTint: C.bgSoft, strokeColor: a.color, borderTint: C.border
    });
    
    slide.addText(a.badge, {
      x: aX + 0.8, y: aY + 0.22, w: aW - 0.95, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: a.color, charSpacing: 0.8
    });
    slide.addText(a.title, {
      x: aX + 0.8, y: aY + 0.44, w: aW - 0.95, h: 0.5,
      fontFace: C.fontTitle, fontSize: 11.5, bold: true, color: C.textMain
    });
    
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX + 0.25, y: aY + 1.02, w: aW - 0.5, h: 0.015,
      fill: { color: C.border }, line: { color: C.border, width: 0 }
    });
    
    slide.addText(a.narrative, {
      x: aX + 0.25, y: aY + 1.15, w: aW - 0.5, h: 1.1,
      fontFace: C.fontBody, fontSize: 10.0, color: C.textMuted, lineSpacingMultiple: 1.25
    });
    
    // Target Impact Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX + 0.2, y: aY + 2.35, w: aW - 0.4, h: 1.1,
      fill: { color: C.bgSoft },
      line: { color: a.color, width: 1 }
    });
    slide.addText(a.targetTitle, {
      x: aX + 0.35, y: aY + 2.42, w: aW - 0.7, h: 0.22,
      fontFace: C.fontTitle, fontSize: 9.0, bold: true, color: a.color
    });
    slide.addText(a.target, {
      x: aX + 0.35, y: aY + 2.66, w: aW - 0.7, h: 0.72,
      fontFace: C.fontBody, fontSize: 9.5, color: C.textMain, lineSpacingMultiple: 1.18
    });
    
    // Mechanism Box
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: aX + 0.2, y: aY + 3.55, w: aW - 0.4, h: 0.95,
      fill: { color: C.bgSoft },
      line: { color: C.border, width: 1 }
    });
    slide.addText(a.mechanismTitle, {
      x: aX + 0.35, y: aY + 3.62, w: aW - 0.7, h: 0.2,
      fontFace: C.fontTitle, fontSize: 8.5, bold: true, color: C.textDim
    });
    slide.addText(a.mechanism, {
      x: aX + 0.35, y: aY + 3.84, w: aW - 0.7, h: 0.6,
      fontFace: C.fontBody, fontSize: 9.0, color: C.textMuted, lineSpacingMultiple: 1.15
    });
  });
  
  // Bottom Full Width Summary Ribbon
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.45, w: 11.733, h: 0.45,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('KOMITMEN KAMI: Cukup dengan 3 payung regulasi Walikota di atas, UPTD KST Solo Technopark siap bergerak cepat, mandiri, dan menuntaskan target Rp 23,47 Miliar tanpa subsidi APBD.', {
    x: 1.0, y: 6.5, w: 11.333, h: 0.35,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(slide);
}

// ==========================================
// SLIDE 15: RENCANA AKSI 100 HARI PERTAMA (NEW EXPANDED FULL TABLE)
// ==========================================
{
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  addHeader(slide, {
    kicker: 'UPTD KST SOLO TECHNOPARK // KOMITMEN EKSEKUSI CEPAT & AKUNTABEL',
    title: 'RENCANA AKSI 100 HARI PERTAMA: QUICK WINS PASCA-ARAHAN WALIKOTA',
    subtitle: 'Milestone Terukur, Matriks Penanggung Jawab, dan Output Nyata dalam 100 Hari Pasca-Persetujuan Kebijakan',
    pageNum: 15
  });
  
  // Top Banner
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 1.62, w: 11.733, h: 0.42,
    fill: { color: C.bgTintBlue },
    line: { color: C.borderBlue, width: 1 }
  });
  addIconBadge(slide, {
    iconName: 'sparkles', x: 0.95, y: 1.66, size: 0.32,
    bgTint: C.bg, strokeColor: C.blue, borderTint: C.borderBlue
  });
  slide.addText('AGENDA EKSEKUTIF 100 HARI: DARI PENANDATANGANAN REGULASI HINGGA KONTRAK PERDANA KEMITRAAN INDUSTRI', {
    x: 1.4, y: 1.68, w: 11.0, h: 0.28,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.blue
  });
  
  const planRows = [
    [
      { text: 'PERIODE', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.0, fill: { color: 'F1F5F9' }, align: 'center', color: C.navy } },
      { text: 'AGENDA STRATEGIS', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.0, fill: { color: 'F1F5F9' }, color: C.navy } },
      { text: 'PENANGGUNG JAWAB (PIC)', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.0, fill: { color: 'F1F5F9' }, align: 'center', color: C.navy } },
      { text: 'OUTPUT KONKRET & TARGET KEBERHASILAN', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.0, fill: { color: 'F1F5F9' }, color: C.navy } }
    ],
    [
      { text: 'Hari 1–10', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center' } },
      { text: 'Penyusunan Draf Regulasi & Harmonisasi Awal:\nFinalisasi draf Inkwal e-Katalog, draf SE Konsorsium Riset, dan reviu skema aset lahan BLUD.', options: { fontFace: C.fontBody, fontSize: 9.5 } },
      { text: 'Kepala UPTD STP &\nTim Hukum BRIDA', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center' } },
      { text: '3 Draf naskah regulasi siap harmonisasi Bagian Hukum Setda & BPKAD.', options: { fontFace: C.fontBody, fontSize: 9.5, color: C.textMain } }
    ],
    [
      { text: 'Hari 11–30', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center', color: C.amber, fill: { color: C.bgTintAmber } } },
      { text: 'Harmonisasi & Penerbitan Instruksi Walikota (QUICK WIN UTAMA):\nPembahasan harmonisasi bersama Setda & BPKAD; Penandatanganan Inkwal e-Katalog Afirmatif.', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.amber, fill: { color: C.bgTintAmber } } },
      { text: 'Bagian Hukum, BPKAD,\nBRIDA, & STP', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center', fill: { color: C.bgTintAmber } } },
      { text: 'Inkwal resmi ditandatangani Walikota Surakarta; Sosialisasi perdana ke seluruh Kepala OPD Pemkot.', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.amber, fill: { color: C.bgTintAmber } } }
    ],
    [
      { text: 'Hari 31–50', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center' } },
      { text: 'Kurasi Produk Inovasi & Integrasi Sistem e-Katalog:\nKurasi 5 produk inovasi unggulan kampus/tenant untuk onboarding perdana ke etalase e-Katalog Lokal.', options: { fontFace: C.fontBody, fontSize: 9.5 } },
      { text: 'UPTD STP & Bagian\nPBJ Setda Solo', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center' } },
      { text: 'Katalog 5 Produk Unggulan Hilirisasi STP tayang aktif (Smart Agri IoT, Alkes, Mesin Sortasi).', options: { fontFace: C.fontBody, fontSize: 9.5, color: C.textMain } }
    ],
    [
      { text: 'Hari 51–60', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center' } },
      { text: 'Showcase Produk & Komitmen Serapan Perdana OPD:\nDemonstrasi produk inovasi di hadapan Walikota, Sekda, dan OPD terkait di Solo Technopark.', options: { fontFace: C.fontBody, fontSize: 9.5 } },
      { text: 'Kepala BRIDA &\nKepala UPTD STP', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center' } },
      { text: 'Demonstrasi produk di hadapan media & komitmen serapan perdana oleh 3 OPD percontohan.', options: { fontFace: C.fontBody, fontSize: 9.5, color: C.textMain } }
    ],
    [
      { text: 'Hari 61–80', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center' } },
      { text: 'Inisiasi Konsorsium Riset 5 Perguruan Tinggi:\nRapat koordinasi tindak lanjut Surat Edaran Walikota bersama pimpinan LPPM UNS, UMS, ISI, Poltek.', options: { fontFace: C.fontBody, fontSize: 9.5 } },
      { text: 'BRIDA Kota Surakarta\n& Konsorsium Kampus', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center' } },
      { text: 'MoU Pra-Konsorsium bersama 5 Rektor Solo Raya; Terpetakan 30 tema riset terapan prioritas.', options: { fontFace: C.fontBody, fontSize: 9.5, color: C.textMain } }
    ],
    [
      { text: 'Hari 81–100', options: { bold: true, fontFace: C.fontTitle, fontSize: 9.5, align: 'center', color: C.emerald, fill: { color: C.bgTintGreen } } },
      { text: 'Launching Resmi Konsorsium & Kick-off Kemitraan Industri:\nPenandatanganan MoU Riset Terapan bersama Kadin/Apindo & kontrak kemitraan kelas OGSCI Migas.', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.emerald, fill: { color: C.bgTintGreen } } },
      { text: 'Walikota Surakarta,\nRektor, & Mitra Industri', options: { fontFace: C.fontBody, fontSize: 9.0, align: 'center', fill: { color: C.bgTintGreen } } },
      { text: 'Penandatanganan MoU Riset Terapan & Peresmian kelas diklat industri mandiri perdana.', options: { fontFace: C.fontBody, fontSize: 9.5, bold: true, color: C.emerald, fill: { color: C.bgTintGreen } } }
    ]
  ];
  
  slide.addTable(planRows, {
    x: 0.8, y: 2.12, w: 11.733, h: 4.15,
    colW: [1.3, 4.4, 2.0, 4.033],
    border: { pt: 0.5, color: C.border }
  });
  
  // Bottom Monitoring Ribbon
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0.8, y: 6.38, w: 11.733, h: 0.48,
    fill: { color: C.navy }, line: { color: C.navy, width: 0 }
  });
  slide.addText('SISTEM MONITORING: Evaluasi progres dilaksanakan mingguan bersama BRIDA dan dilaporkan berkala setiap 30 hari langsung kepada Walikota Surakarta untuk memastikan seluruh quick wins tercapai 100%.', {
    x: 1.0, y: 6.44, w: 11.333, h: 0.35,
    fontFace: C.fontTitle, fontSize: 9.5, bold: true, color: C.textWhite, align: 'center'
  });
  
  addFooter(slide);
}

// ==========================================
// SIMPAN PPTX V9 (15 SLIDES)
// ==========================================
const outDirLocal = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\bahan_paparan_walikota_v3';
const outDirPublic = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\output_paparan';
const outPptxLocal = path.join(outDirLocal, 'PAPARAN_WALIKOTA_STP_V9_15S.pptx');
const outPptxPublic = path.join(outDirPublic, 'PAPARAN_WALIKOTA_STP_V9_15S.pptx');

pptx.writeFile({ fileName: outPptxLocal }).then(() => {
  console.log('[SUCCESS] PPTX Master 15S generated at:', outPptxLocal);
  fs.copyFileSync(outPptxLocal, outPptxPublic);
  console.log('[SUCCESS] Copied to output_paparan:', outPptxPublic);
}).catch(err => {
  console.error('[ERROR] Failed generating PPTX 15S:', err);
  process.exit(1);
});
