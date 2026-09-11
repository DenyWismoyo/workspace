import pptxgen from 'pptxgenjs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const pptx = new pptxgen();

// Set True Widescreen 16:9 (13.333 x 7.5 inches)
pptx.defineLayout({ name: 'LAYOUT_16x9_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'LAYOUT_16x9_WIDE';

pptx.author = 'UPTD KST Solo Technopark & BRIDA Kota Surakarta';
pptx.company = 'Pemerintah Kota Surakarta';
pptx.title = 'Strategi Inovatif Pengembangan Solo Technopark 2026-2030';
pptx.subject = 'Bahan Paparan Resmi Walikota Surakarta';

// High-Contrast Executive Color Palette
const C_NAVY        = '0A192F'; // #0a192f Primary Dark Navy
const C_BLUE        = '1E3A8A'; // #1e3a8a Deep Royal Blue
const C_CYAN        = '0284C7'; // #0284c7 Vivid Accent Cyan
const C_LIGHT_CYAN  = 'E0F2FE'; // #e0f2fe Light Cyan Fill
const C_AMBER       = 'D97706'; // #d97706 Golden Accent
const C_LIGHT_AMBER = 'FEF3C7'; // #fef3c7 Light Amber Fill
const C_EMERALD     = '059669'; // #059669 Emerald Green
const C_LIGHT_GREEN = 'D1FAE5'; // #d1fae5 Light Green Fill
const C_BG          = 'F8FAFC'; // #f8fafc Clean Executive Background
const C_CARD        = 'FFFFFF'; // #ffffff Card Fill
const C_TEXT        = '1E293B'; // #1e293b Slate Text
const C_MUTED       = '64748B'; // #64748b Muted Text
const C_WHITE       = 'FFFFFF'; // #ffffff Pure White
const C_BORDER      = 'E2E8F0'; // #e2e8f0 Card Border
const C_DARK_BG     = '070D18'; // #070d18 Hero Dark

// Catatan: Seluruh gambar dihilangkan sesuai instruksi (fokus murni tipografi, data, dan layout elegan)

// Helper: Standard Slide Header & Footer
function addHeader(slide, title, category, slideNum) {
  slide.background = { color: C_BG };
  
  // Top thin accent line
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.333, h: 0.1, fill: { color: C_CYAN }
  });

  // Category Tag
  slide.addText(category.toUpperCase(), {
    x: 0.8, y: 0.32, w: 9.0, h: 0.28,
    fontSize: 10, bold: true, color: C_CYAN, fontFace: 'Arial'
  });

  // Slide Main Title
  slide.addText(title, {
    x: 0.8, y: 0.60, w: 11.7, h: 0.45,
    fontSize: 20, bold: true, color: C_NAVY, fontFace: 'Arial'
  });

  // Footer Branding
  slide.addText('Solo Technopark 2026–2030 | Bahan Paparan Resmi Walikota Surakarta', {
    x: 0.8, y: 6.98, w: 8.5, h: 0.3,
    fontSize: 9, color: C_MUTED, fontFace: 'Arial'
  });

  // Slide Number Indicator
  slide.addText(`Slide ${slideNum} / 10`, {
    x: 11.2, y: 6.98, w: 1.333, h: 0.3,
    fontSize: 9.5, bold: true, align: 'right', color: C_MUTED, fontFace: 'Arial'
  });
}

// =========================================================================
// SLIDE 01: COVER & VISI BESAR TRANSFORMASI
// =========================================================================
const s1 = pptx.addSlide();
s1.background = { color: C_DARK_BG };

// Top Elegant Accent Bars
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 13.333, h: 0.12, fill: { color: C_CYAN }
});
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.12, w: 13.333, h: 0.04, fill: { color: C_AMBER }
});

// Government Badge
s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.9, y: 0.7, w: 5.2, h: 0.42,
  fill: { color: C_BLUE }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.08
});
s1.addText('PEMERINTAH KOTA SURAKARTA & BRIDA', {
  x: 0.9, y: 0.7, w: 5.2, h: 0.42,
  fontSize: 11, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

s1.addText('STRATEGI INOVATIF PENGEMBANGAN\nSOLO TECHNOPARK 2026–2030', {
  x: 0.9, y: 1.45, w: 11.5, h: 1.7,
  fontSize: 34, bold: true, color: C_WHITE, fontFace: 'Arial', lineSpacing: 40
});

s1.addText('"Dari Pengelolaan Kawasan Fisik Menuju Mesin Hilirisasi Riset & Lokomotif Ekonomi Daerah"', {
  x: 0.9, y: 3.25, w: 11.5, h: 0.55,
  fontSize: 15, italic: true, color: C_LIGHT_AMBER, fontFace: 'Arial'
});

// Sleek Cyan Divider Line
s1.addShape(pptx.shapes.RECTANGLE, {
  x: 0.9, y: 3.95, w: 11.533, h: 0.03, fill: { color: C_CYAN }
});

// 3 Bento Highlights on Cover
const c1w = 3.6;
const coverCards = [
  {
    title: 'MESIN HILIRISASI RISET',
    desc: 'Menghubungkan invensi kampus (UNS, UMS, dll.) menuju sertifikasi dan kebutuhan komersial industri.'
  },
  {
    title: 'KEMANDIRIAN BLUD (0% APBD)',
    desc: 'Re-setting target pendapatan melonjak dari Rp 7,50 M (2026) menuju Rp 23,47 M (2030) mandiri penuh.'
  },
  {
    title: 'DAMPAK WARGA KTP SOLO',
    desc: 'Penyerapan 9.000 tenaga kerja baru dengan komitmen kuota afirmasi minimal 65% KTP Kota Surakarta.'
  }
];

coverCards.forEach((c, idx) => {
  const cx = 0.9 + idx * 3.9;
  s1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: cx, y: 4.25, w: c1w, h: 2.3,
    fill: { color: '0F1E36' },
    line: { color: C_CYAN, width: 1.5 }, rectRadius: 0.1
  });
  s1.addText(c.title, {
    x: cx + 0.25, y: 4.45, w: c1w - 0.5, h: 0.45,
    fontSize: 12, bold: true, color: C_AMBER, fontFace: 'Arial'
  });
  s1.addText(c.desc, {
    x: cx + 0.25, y: 4.95, w: c1w - 0.5, h: 1.45,
    fontSize: 10.5, color: 'E2E8F0', fontFace: 'Arial', lineSpacing: 16
  });
});

// Presenter Context Footer
s1.addText('Bahan Paparan Resmi Rancangan Teknokratik Roadmap STP 2026–2030  •  Disampaikan kepada: Walikota Surakarta', {
  x: 0.9, y: 6.85, w: 11.533, h: 0.35,
  fontSize: 9.5, color: '94A3B8', align: 'center', fontFace: 'Arial'
});

s1.addNotes('Bapak Walikota yang kami hormati, perkenankan kami menyampaikan arah baru dan Roadmap Solo Technopark periode 2026 hingga 2030. Solo Technopark hari ini tidak lagi boleh dipandang hanya sebagai fasilitas fisik atau gedung pelatihan semata. Kami memosisikan STP sebagai motor penggerak inovasi daerah dan jembatan hilirisasi ekonomi yang mengubah potensi riset perguruan tinggi dan kemitraan industri menjadi lapangan kerja nyata, pendapatan daerah, serta peningkatan kesejahteraan warga Surakarta.');


// =========================================================================
// SLIDE 02: KEDUDUKAN KELEMBAGAAN & POLA HUBUNGAN SINERGIS BRIDA–STP
// =========================================================================
const s2 = pptx.addSlide();
addHeader(s2, 'KEDUDUKAN KELEMBAGAAN & SINERGI STRATEGIS BRIDA–STP', 'Tata Kelola Pemerintahan & Kelembagaan', 2);

// Banner Subtitle
s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.15, w: 11.733, h: 0.55,
  fill: { color: C_LIGHT_CYAN }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.08
});
s2.addText('BRIDA merumuskan & mengorkestrasi  ➔  Solo Technopark mengimplementasikan, menguji, menghilirkan & melayani', {
  x: 0.9, y: 1.15, w: 11.533, h: 0.55,
  fontSize: 12.5, bold: true, color: C_BLUE, align: 'center', fontFace: 'Arial'
});

// Left Card: BRIDA
s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.9, w: 5.65, h: 2.7,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1 }, rectRadius: 0.1
});
s2.addShape(pptx.shapes.RECTANGLE, {
  x: 0.8, y: 1.9, w: 5.65, h: 0.45, fill: { color: C_BLUE }
});
s2.addText('BRIDA SURAKARTA (Otak & Regulator)', {
  x: 1.0, y: 1.9, w: 5.25, h: 0.45,
  fontSize: 12, bold: true, color: C_WHITE, fontFace: 'Arial'
});
s2.addText([
  { text: '• Mandat Kebijakan & Riset: ', options: { bold: true, color: C_NAVY } },
  { text: 'Merumuskan arah kebijakan Iptek, roadmap inovasi kota, dan kajian strategis pembangunan daerah.\n', options: { color: C_TEXT } },
  { text: '• Orkestrasi Ekosistem: ', options: { bold: true, color: C_NAVY } },
  { text: 'Menghubungkan Perguruan Tinggi, litbang, dan OPD Pemkot Surakarta.\n', options: { color: C_TEXT } },
  { text: '• Monitoring & Evaluasi: ', options: { bold: true, color: C_NAVY } },
  { text: 'Menilai capaian dampak sosial-ekonomi inovasi bagi warga Kota Surakarta.', options: { color: C_TEXT } }
], {
  x: 1.05, y: 2.45, w: 5.15, h: 2.0,
  fontSize: 10.5, fontFace: 'Arial', lineSpacing: 16
});

// Right Card: UPTD KST Solo Technopark
s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.85, y: 1.9, w: 5.65, h: 2.7,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1 }, rectRadius: 0.1
});
s2.addShape(pptx.shapes.RECTANGLE, {
  x: 6.85, y: 1.9, w: 5.65, h: 0.45, fill: { color: C_CYAN }
});
s2.addText('UPTD KST SOLO TECHNOPARK (Tangan Eksekutor)', {
  x: 7.05, y: 1.9, w: 5.25, h: 0.45,
  fontSize: 12, bold: true, color: C_WHITE, fontFace: 'Arial'
});
s2.addText([
  { text: '• Teknis Operasional: ', options: { bold: true, color: C_NAVY } },
  { text: 'Penyelenggaraan diklat vokasi industri dan jasa rekayasa manufaktur presisi.\n', options: { color: C_TEXT } },
  { text: '• Fasilitasi Riset & Uji Coba: ', options: { bold: true, color: C_NAVY } },
  { text: 'Validasi prototype kampus hingga siap uji standar SNI/TKDN.\n', options: { color: C_TEXT } },
  { text: '• Inkubasi & Komersialisasi: ', options: { bold: true, color: C_NAVY } },
  { text: 'Mendampingi tenant bisnis, fasilitasi paten HKI, dan penetrasi pasar komersial.', options: { color: C_TEXT } }
], {
  x: 7.1, y: 2.45, w: 5.15, h: 2.0,
  fontSize: 10.5, fontFace: 'Arial', lineSpacing: 16
});

// Bottom Regulatory Box
s2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 4.8, w: 11.733, h: 1.9,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1 }, rectRadius: 0.1
});
s2.addText('LANDASAN REGULASI & POLA TATA KELOLA BLUD YANG KOKOH', {
  x: 1.0, y: 4.95, w: 11.3, h: 0.35,
  fontSize: 11.5, bold: true, color: C_NAVY, fontFace: 'Arial'
});

const regCards = [
  {
    tag: 'PERMENDAGRI 79/2018 (BLUD)',
    text: 'Pasal 1 angka 2: Fleksibilitas BLUD adalah keleluasaan pengelolaan keuangan menerapkan praktik bisnis sehat untuk meningkatkan layanan publik tanpa mencari keuntungan semata.'
  },
  {
    tag: 'PERPRES 106/2017 (KST)',
    text: 'Pasal 4 & 5: KST berfungsi sebagai wahana litbang berkelanjutan Pemda, Perguruan Tinggi, dan Industri melalui layanan teknis, alih teknologi, inkubasi, dan pendukung.'
  },
  {
    tag: 'PERWALI 15 & 38 TAHUN 2022',
    text: 'Pola Tata Kelola UPTD KST STP: Menugaskan pelaksanaan teknis operasional penunjang litbang dan penumbuhan perusahaan pemula berbasis inovasi (spin-off).'
  }
];

regCards.forEach((rc, idx) => {
  const rx = 1.0 + idx * 3.8;
  s2.addShape(pptx.shapes.RECTANGLE, {
    x: rx, y: 5.35, w: 3.6, h: 1.2,
    fill: { color: C_BG }, line: { color: C_BORDER, width: 1 }
  });
  s2.addText(rc.tag, {
    x: rx + 0.15, y: 5.42, w: 3.3, h: 0.25,
    fontSize: 9.5, bold: true, color: C_BLUE, fontFace: 'Arial'
  });
  s2.addText(rc.text, {
    x: rx + 0.15, y: 5.68, w: 3.3, h: 0.8,
    fontSize: 8.5, color: C_TEXT, fontFace: 'Arial', lineSpacing: 11
  });
});

s2.addNotes('Dasar kelembagaan kita sudah sangat kokoh, Bapak Walikota. Berdasarkan Perwali 15 dan 38 Tahun 2022, pembagian peran dibuat sangat tegas: BRIDA menjadi otak yang merumuskan dan mengorkestrasi arah kebijakan riset kota, sedangkan UPTD Solo Technopark bertindak sebagai tangan eksekutor atau delivery unit. Dengan status pola kelola BLUD sesuai Permendagri 79/2018, STP memiliki fleksibilitas bisnis yang sehat untuk melayani industri secara profesional tanpa birokrasi berbelit, namun tetap berorientasi penuh pada kemanfaatan publik.');


// =========================================================================
// SLIDE 03: 4 PILAR PORTOFOLIO LAYANAN TERINTEGRASI BLUD STP
// =========================================================================
const s3 = pptx.addSlide();
addHeader(s3, 'EMPAT PILAR PORTOFOLIO LAYANAN TERINTEGRASI BLUD STP', 'Mandat Layanan Perpres 106/2017', 3);

// Subtitle Quote Banner
s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.15, w: 11.733, h: 0.45,
  fill: { color: C_LIGHT_AMBER }, line: { color: C_AMBER, width: 1 }, rectRadius: 0.06
});
s3.addText('"BLUD menggeser paradigma: Dari sekadar menyewakan kawasan fisik menjadi penyedia solusi teknologi yang responsif terhadap kebutuhan industri dan masyarakat."', {
  x: 0.9, y: 1.15, w: 11.533, h: 0.45,
  fontSize: 11, bold: true, color: C_AMBER, align: 'center', fontFace: 'Arial'
});

const pilarCards = [
  {
    num: 'PILAR 1',
    title: 'LAYANAN TEKNIS (VOKASI SPESIALIS)',
    color: C_BLUE,
    items: [
      'Diklat Mekanik, Otomasi & Mekatronika Industri',
      'Desain Manufaktur & Pengelasan (Welding 3G/6G)',
      'OGSCI: Oil & Gas Skills Training Center',
      'Underwater Wet Welding (Keahlian Langka Nasional)'
    ]
  },
  {
    num: 'PILAR 2',
    title: 'PENGEMBANGAN TEKNOLOGI',
    color: C_CYAN,
    items: [
      'Jasa Layanan Rekayasa Industri & Permesinan',
      'Produksi Komponen Presisi (Precision Parts)',
      'Pembuatan & Uji Prototipe bagi IKM/UKM Solo',
      'Riset Terapan Bersama Industri & Laboratorium'
    ]
  },
  {
    num: 'PILAR 3',
    title: 'INKUBASI BISNIS & TEKNOLOGI',
    color: C_EMERALD,
    items: [
      'Program Pra-Inkubasi & Seleksi Startup Digital',
      'Pendampingan HKI, Hak Paten, & Izin Edar',
      'Hilirisasi Riset Invensi Perguruan Tinggi',
      'Fasilitas Co-Working Space & Akses Modal Ventura'
    ]
  },
  {
    num: 'PILAR 4',
    title: 'LAYANAN PENDUKUNG (SUPPORTING)',
    color: C_NAVY,
    items: [
      'Prakerin & Uji Kompetensi Siswa SMK se-Solo Raya',
      'Solo Science Center: Wahana Edukasi Iptek Publik',
      'Optimalisasi Gedung Pertemuan, MICE, & Expo',
      'Fasilitas Pusat Olahraga & Sarana Publik Kawasan'
    ]
  }
];

pilarCards.forEach((p, idx) => {
  const col = idx % 2;
  const row = Math.floor(idx / 2);
  const px = 0.8 + col * 5.95;
  const py = 1.75 + row * 2.5;

  s3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px, y: py, w: 5.75, h: 2.35,
    fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
  });
  
  // Header tag
  s3.addShape(pptx.shapes.RECTANGLE, {
    x: px, y: py, w: 5.75, h: 0.45, fill: { color: p.color }
  });
  s3.addText(`${p.num}: ${p.title}`, {
    x: px + 0.2, y: py, w: 5.35, h: 0.45,
    fontSize: 11, bold: true, color: C_WHITE, fontFace: 'Arial'
  });

  const bulletList = p.items.map(it => `• ${it}`).join('\n');
  s3.addText(bulletList, {
    x: px + 0.25, y: py + 0.55, w: 5.25, h: 1.65,
    fontSize: 10, color: C_TEXT, fontFace: 'Arial', lineSpacing: 15
  });
});

s3.addNotes('Sebagai BLUD, STP beroperasi di atas 4 pilar utama layanan sesuai standar nasional KST. Kami tidak hanya melatih operator mesin dasar, tetapi memiliki kompetensi unggulan langka seperti pengelasan bawah air (underwater welding) dan manufaktur presisi. Di sisi lain, kami memfasilitasi hilirisasi riset perguruan tinggi dan inkubasi usaha rintisan. Empat pilar ini saling mengunci: anak muda dilatih, teknologinya dikembangkan, startupnya diinkubasi, dan fasilitasnya dioptimalkan.');


// =========================================================================
// SLIDE 04: PERAN UTAMA: PIPELINE HILIRISASI RISET KE PASAR
// =========================================================================
const s4 = pptx.addSlide();
addHeader(s4, 'PERAN KUNCI: PIPELINE HILIRISASI RISET MENUJU PASAR', 'Rantai Nilai Hilirisasi (Value Chain)', 4);

// 5-Step Pipeline Flow
const steps = [
  { num: '01', title: 'Riset / Inventor', sub: 'UNS, UMS, ISI, Poltek' },
  { num: '02', title: 'Seleksi & Kurasi', sub: 'Validasi Kesiapan STP' },
  { num: '03', title: 'Prototype Lab', sub: 'Uji Teknis & Standarisasi' },
  { num: '04', title: 'Inkubasi Bisnis', sub: 'Business Plan & HKI' },
  { num: '05', title: 'Pasar Industri', sub: 'Komersial & e-Katalog' }
];

steps.forEach((st, idx) => {
  const sx = 0.8 + idx * 2.38;
  s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: sx, y: 1.25, w: 2.2, h: 1.1,
    fill: { color: idx === 4 ? C_EMERALD : (idx === 2 ? C_AMBER : C_BLUE) },
    line: { color: C_WHITE, width: 1.5 }, rectRadius: 0.08
  });
  s4.addText(`TAHAP ${st.num}`, {
    x: sx, y: 1.35, w: 2.2, h: 0.25,
    fontSize: 9, bold: true, color: C_LIGHT_AMBER, align: 'center', fontFace: 'Arial'
  });
  s4.addText(st.title, {
    x: sx, y: 1.6, w: 2.2, h: 0.35,
    fontSize: 11, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });
  s4.addText(st.sub, {
    x: sx, y: 1.95, w: 2.2, h: 0.3,
    fontSize: 8.5, color: C_LIGHT_CYAN, align: 'center', fontFace: 'Arial'
  });

  if (idx < 4) {
    s4.addText('➔', {
      x: sx + 2.15, y: 1.55, w: 0.3, h: 0.4,
      fontSize: 14, bold: true, color: C_CYAN, align: 'center', fontFace: 'Arial'
    });
  }
});

// 3 Supporting Bento Cards
const pipeCards = [
  {
    title: 'CONTOH TEKNOLOGI RIIL',
    color: C_BLUE,
    items: [
      'Smart Agriculture: Mesin Sortasi Kopi Otomatis',
      'IoT Pertanian: Pemantau Lahan & Tetas Telur Pintar',
      'Autonomous Electric Vehicle & Sensorik Kota',
      'Virtual Reality (VR) Pelatihan Industri & Chatbot AI'
    ]
  },
  {
    title: 'SENTRA EKOSISTEM AKTIF',
    color: C_CYAN,
    items: [
      'National Cyber Security Hub (Ketahanan Siber)',
      'Digital Technopark (Software & Cloud AI)',
      'Gaming Hub & GoTo UMKM Center',
      'AI Experience Center & Robotic Testing Sandbox'
    ]
  },
  {
    title: 'OUTPUT STRATEGIS DITARGETKAN',
    color: C_EMERALD,
    items: [
      'Prototipe TRL 7-9 Siap Produksi Massal',
      'Pendaftaran Hak Paten / HKI bagi Inventor',
      'Kelahiran Tenant / Startup Mandiri (Spin-off)',
      'Produk Bersertifikat TKDN Masuk e-Katalog Lokal'
    ]
  }
];

pipeCards.forEach((pc, idx) => {
  const px = 0.8 + idx * 3.95;
  s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px, y: 2.6, w: 3.8, h: 3.2,
    fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
  });
  s4.addShape(pptx.shapes.RECTANGLE, {
    x: px, y: 2.6, w: 3.8, h: 0.4, fill: { color: pc.color }
  });
  s4.addText(pc.title, {
    x: px + 0.15, y: 2.6, w: 3.5, h: 0.4,
    fontSize: 10.5, bold: true, color: C_WHITE, fontFace: 'Arial'
  });

  const bText = pc.items.map(it => `• ${it}`).join('\n\n');
  s4.addText(bText, {
    x: px + 0.2, y: 3.15, w: 3.4, h: 2.5,
    fontSize: 9.8, color: C_TEXT, fontFace: 'Arial', lineSpacing: 14
  });
});

// Bottom Big Quote
s4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 6.0, w: 11.733, h: 0.75,
  fill: { color: C_NAVY }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.08
});
s4.addText('"Solo Technopark tidak berhenti pada hasil riset; STP menjadi jembatan nyata dari ilmu pengetahuan menuju nilai ekonomi daerah dan manfaat sosial bagi warga Surakarta."', {
  x: 1.0, y: 6.0, w: 11.333, h: 0.75,
  fontSize: 11.5, italic: true, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

s4.addNotes('Bapak Walikota, persoalan terbesar riset di perguruan tinggi adalah fenomena riset berhenti di lemari arsip atau publikasi jurnal semata. Di Solo Technopark, kami memutus mata rantai kebuntuan itu. Kami mengambil prototipe dosen dan mahasiswa UNS, UMS, maupun politeknik, kami sempurnakan di lab permesinan presisi kami, kami uji sertifikasinya, lalu kami sambungkan ke 95 mitra industri untuk diproduksi massal.');


// =========================================================================
// SLIDE 05: EKOSISTEM KOLABORASI GLOBAL (95+ MITRA BISNIS)
// =========================================================================
const s5 = pptx.addSlide();
addHeader(s5, 'EKOSISTEM KOLABORASI GLOBAL: 95+ MITRA KORPORASI AKTIF', 'Public-Private Partnership (Non-APBD)', 5);

// Left Content Column (5.7 in)
s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 1.2, w: 5.7, h: 4.6,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
});
s5.addShape(pptx.shapes.RECTANGLE, {
  x: 0.8, y: 1.2, w: 5.7, h: 0.45, fill: { color: C_BLUE }
});
s5.addText('PORTOFOLIO FASILITAS TECH HUB DUNIA', {
  x: 1.0, y: 1.2, w: 5.3, h: 0.45,
  fontSize: 11.5, bold: true, color: C_WHITE, fontFace: 'Arial'
});

s5.addText([
  { text: '• Shopee Digital Tech & Export Hub: ', options: { bold: true, color: C_NAVY } },
  { text: 'Pusat inkubasi talenta digital, logistik e-commerce terintegrasi, dan kurasi UMKM ekspor Solo ke pasar ASEAN.\n\n', options: { color: C_TEXT } },
  { text: '• Garena Gaming & Creative Hub: ', options: { bold: true, color: C_NAVY } },
  { text: 'Pengembangan software engineering, animasi digital 3D, game development, dan inkubasi talenta kreatif.\n\n', options: { color: C_TEXT } },
  { text: '• Bank Mandiri Innovation Hub: ', options: { bold: true, color: C_NAVY } },
  { text: 'Sandbox fintech, kurasi startup binaan, dan fasilitas akses pembiayaan wirausaha baru berbasis teknologi.\n\n', options: { color: C_TEXT } },
  { text: '• ACER, Indosat & Telkomsel 5G Lab: ', options: { bold: true, color: C_NAVY } },
  { text: 'Fasilitas konektivitas 5G super cepat, lab pengujian hardware, cyber security, dan digitalisasi logistik pintar.\n\n', options: { color: C_TEXT } },
  { text: '• CyberTrend AI & Data Science Lab: ', options: { bold: true, color: C_NAVY } },
  { text: 'Laboratorium riset kecerdasan buatan, visualisasi data analytics terapan, dan pusat komputasi awan.', options: { color: C_TEXT } }
], {
  x: 1.0, y: 1.75, w: 5.3, h: 3.9,
  fontSize: 9.8, fontFace: 'Arial', lineSpacing: 13
});

// Right Content Column (5.7 in)
s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.833, y: 1.2, w: 5.7, h: 4.6,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
});
s5.addShape(pptx.shapes.RECTANGLE, {
  x: 6.833, y: 1.2, w: 5.7, h: 0.45, fill: { color: C_CYAN }
});
s5.addText('3 PRINSIP STRATEGIS KEMITRAAN (NON-APBD)', {
  x: 7.033, y: 1.2, w: 5.3, h: 0.45,
  fontSize: 11.5, bold: true, color: C_WHITE, fontFace: 'Arial'
});

// 3 Value Proposition Cards on Right
const partnershipPrinciples = [
  {
    title: '1. Efisiensi Penuh APBD (0 Rupiah Kas Daerah)',
    desc: 'Seluruh infrastruktur lab dan teknologi mutakhir dibangun 100% atas komitmen investasi korporasi mitra tanpa membebani keuangan daerah.'
  },
  {
    title: '2. Transfer Teknologi & Kurikulum Global',
    desc: 'Talenta lokal Surakarta dibimbing langsung oleh master engineer industri dunia dengan kurikulum dan sertifikasi berstandar internasional.'
  },
  {
    title: '3. Kepastian Penyerapan Pasar (Captive Market)',
    desc: 'Menjamin lulusan diklat vokasi dan produk riset hilirisasi kampus langsung terserap ke dalam ekosistem rantai pasok korporasi mitra.'
  }
];

partnershipPrinciples.forEach((pp, pIdx) => {
  const py = 1.75 + pIdx * 0.85;
  s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 7.033, y: py, w: 5.3, h: 0.75,
    fill: { color: C_BG }, line: { color: C_BORDER, width: 1 }, rectRadius: 0.06
  });
  s5.addText(pp.title, {
    x: 7.183, y: py + 0.06, w: 5.0, h: 0.22,
    fontSize: 9.8, bold: true, color: C_BLUE, fontFace: 'Arial'
  });
  s5.addText(pp.desc, {
    x: 7.183, y: py + 0.28, w: 5.0, h: 0.42,
    fontSize: 8.5, color: C_TEXT, fontFace: 'Arial', lineSpacing: 11
  });
});

// Right Stat Box below principles
s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 7.033, y: 4.45, w: 5.3, h: 1.2,
  fill: { color: C_NAVY }, line: { color: C_CYAN, width: 1.5 }, rectRadius: 0.08
});
s5.addText('95+ KORPORASI & PROYEK STRATEGIS AKTIF', {
  x: 7.233, y: 4.55, w: 4.9, h: 0.35,
  fontSize: 12.5, bold: true, color: C_AMBER, fontFace: 'Arial'
});
s5.addText('Solo Technopark terbukti memiliki daya tarik investasi industri yang sangat kuat dan menjadi rujukan nasional kemitraan triple-helix (Pemerintah - Kampus - Industri).', {
  x: 7.233, y: 4.95, w: 4.9, h: 0.6,
  fontSize: 9.5, color: C_WHITE, fontFace: 'Arial', lineSpacing: 13
});

// Bottom Summary Banner
s5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 6.0, w: 11.733, h: 0.75,
  fill: { color: C_LIGHT_GREEN }, line: { color: C_EMERALD, width: 1 }, rectRadius: 0.08
});
s5.addText('Kemitraan strategis ini membuktikan bahwa Solo Technopark memiliki daya tarik investasi tinggi dan menjadi rujukan kolaborasi industri-pemerintah di tingkat nasional.', {
  x: 1.0, y: 6.0, w: 11.333, h: 0.75,
  fontSize: 11, bold: true, color: C_EMERALD, align: 'center', fontFace: 'Arial'
});

s5.addNotes('Satu hal yang membanggakan dari Solo Technopark adalah: ekosistem ini dipercaya oleh 95 lebih korporasi kelas dunia. Shopee, Garena, Mandiri, hingga ACER telah menanamkan fasilitas dan proyek riil di sini. Dan poin paling penting bagi Bapak Walikota: pembangunan fasilitas mutakhir ini tidak menyedot anggaran APBD Kota Surakarta satu rupiah pun. Ini adalah murni kemitraan strategis investasi swasta yang dihadirkan untuk memajukan warga Solo.');


// =========================================================================
// SLIDE 06: ANALISIS GAP STRATEGIS & LIMA SOLUSI TEROBOSAN
// =========================================================================
const s6 = pptx.addSlide();
addHeader(s6, 'ANALISIS GAP STRATEGIS & LIMA TEROBOSAN SOLUSI 2026–2030', 'Hasil Audit Masalah Riil & Formula Solusi', 6);

// Table Data from Master Audit
const gapTableData = [
  [
    { text: 'NO', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY }, align: 'center' } },
    { text: 'IDENTIFIKASI GAP EKSISTING', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY } } },
    { text: 'AKAR MASALAH (AUDIT RIIL)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY } } },
    { text: 'SOLUSI TEROBOSAN ROADMAP 2026–2030', options: { bold: true, color: C_WHITE, fill: { color: C_CYAN } } }
  ],
  [
    { text: '1', options: { bold: true, align: 'center' } },
    { text: 'Jenis Pelatihan Terbatas' },
    { text: 'Keterbatasan SDM instruktur ahli dan kelengkapan peralatan diklat.' },
    { text: 'Kolaborasi kurikulum korporasi (AI, Cyber Security, EV, Welder Bawah Air).' }
  ],
  [
    { text: '2', options: { bold: true, align: 'center' } },
    { text: 'Kegiatan R&D Belum Maksimal' },
    { text: 'Riset tidak tampak di pasar; belum mengarah ke hilirisasi dan komersialisasi.' },
    { text: 'Bentuk Technology Transfer Office (TTO) & kurasi prototype siap industri.' }
  ],
  [
    { text: '3', options: { bold: true, align: 'center' } },
    { text: 'Investasi Lahan Terbatas' },
    { text: 'Lahan kosong kawasan belum optimal menarik investor teknologi.' },
    { text: 'Zonasi investasi terpadu & skema Long-Term Revenue Sharing bagi investor.' }
  ],
  [
    { text: '4', options: { bold: true, align: 'center' } },
    { text: 'Kemandirian Finansial Rendah' },
    { text: 'Hardware infrastruktur belum menghasilkan revenue penopang operasional KST.' },
    { text: 'Diversifikasi ke jasa uji lab penetrant, sertifikasi TUK/JFT, & royalti paten.' }
  ],
  [
    { text: '5', options: { bold: true, align: 'center' } },
    { text: 'Promosi & Penyerapan Lemah' },
    { text: 'Networking kurang di-maintenance; belum ada regulasi penyerapan pasar lokal.' },
    { text: 'Instruksi Walikota untuk e-Katalog Lokal Afirmatif produk inovasi STP.' }
  ]
];

s6.addTable(gapTableData, {
  x: 0.8, y: 1.2, w: 11.733, h: 4.6,
  colW: [0.6, 2.7, 3.8, 4.633],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_CARD },
  fontSize: 9.8,
  fontFace: 'Arial',
  valign: 'middle',
  rowH: [0.4, 0.75, 0.8, 0.8, 0.85, 0.8]
});

// Bottom Highlight
s6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 6.0, w: 11.733, h: 0.75,
  fill: { color: C_LIGHT_AMBER }, line: { color: C_AMBER, width: 1 }, rectRadius: 0.08
});
s6.addText('Kunci Keberhasilan: Beralih dari penyewaan aset fisik statis menjadi penyedia layanan bernilai tambah tinggi dan kepastian pasar melalui regulasi belanja daerah.', {
  x: 1.0, y: 6.0, w: 11.333, h: 0.75,
  fontSize: 11, bold: true, color: C_AMBER, align: 'center', fontFace: 'Arial'
});

s6.addNotes('Kami melakukan audit transparan terhadap tantangan STP saat ini, Bapak Walikota. Masalah utama selama ini adalah pendapatan BLUD kita masih terlalu bertumpu pada sewa lahan fisik, sementara riset kampus belum terhubung ke pasar. Solusi terobosan kami jelas: kita beralih ke hilirisasi produk paten, membuka laboratorium uji berbayar bagi korporasi luar kota, dan memohon arahan regulasi Bapak agar produk binaan STP diserap oleh dinas-dinas Pemkot lewat e-Katalog Lokal.');


// =========================================================================
// SLIDE 07: ROADMAP PENTAHAPAN 2026–2030
// =========================================================================
const s7 = pptx.addSlide();
addHeader(s7, 'ROADMAP PENTAHAPAN PENGEMBANGAN SOLO TECHNOPARK 2026–2030', 'Peta Jalan Strategis Kelembagaan', 7);

const phases = [
  {
    phase: 'FASE 1 (2026–2027)',
    title: 'CENTRE OF INNOVATION & CONNECTIVITY',
    color: C_BLUE,
    points: [
      'Penguatan kelembagaan BLUD & legalitas aset kawasan.',
      'Sertifikasi mutu diklat vokasi & akreditasi lab industri.',
      'Perluasan jejaring industri domestik dan multinasional.',
      'Pembentukan Konsorsium Riset 5 Kampus Solo Raya.'
    ]
  },
  {
    phase: 'FASE 2 (2028–2029)',
    title: 'ENTREPRENEUR TECHNOPARK & SBU MANDIRI',
    color: C_CYAN,
    points: [
      'Pembentukan Strategic Business Unit (SBU) yang mandiri.',
      'Penerapan Strategi Corporate Marketing terintegrasi.',
      'Hilirisasi 20 produk inovasi komersial ber-TKDN.',
      'Peningkatan pendapatan jasa industri hingga Rp 16,1 Miliar.'
    ]
  },
  {
    phase: 'FASE 3 (2029–2030)',
    title: 'GLOBAL MARKET POSITIONING & HOLDING',
    color: C_EMERALD,
    points: [
      'Kemandirian finansial penuh: 0% subsidi belanja APBD.',
      'Pola tata kelola manajemen holding company inovasi.',
      'Pusat pengembangan Indonesia Digital Technopark.',
      'Realisasi target pendapatan mandiri Rp 23,47 Miliar.'
    ]
  }
];

phases.forEach((ph, idx) => {
  const px = 0.8 + idx * 3.95;
  s7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: px, y: 1.25, w: 3.8, h: 4.1,
    fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
  });
  s7.addShape(pptx.shapes.RECTANGLE, {
    x: px, y: 1.25, w: 3.8, h: 0.65, fill: { color: ph.color }
  });
  s7.addText(ph.phase, {
    x: px + 0.15, y: 1.3, w: 3.5, h: 0.25,
    fontSize: 9.5, bold: true, color: C_LIGHT_AMBER, align: 'center', fontFace: 'Arial'
  });
  s7.addText(ph.title, {
    x: px + 0.15, y: 1.55, w: 3.5, h: 0.35,
    fontSize: 10, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });

  const pList = ph.points.map(pt => `• ${pt}`).join('\n\n');
  s7.addText(pList, {
    x: px + 0.2, y: 2.05, w: 3.4, h: 3.1,
    fontSize: 10, color: C_TEXT, fontFace: 'Arial', lineSpacing: 14
  });
});

// Bottom Revenue Stream Composition Bar (Audited Page 4)
s7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 5.6, w: 11.733, h: 1.15,
  fill: { color: C_NAVY }, line: { color: C_CYAN, width: 1.2 }, rectRadius: 0.08
});
s7.addText('PROYEKSI SUMBER PENDAPATAN BLUD 2026–2030 (TARGET REVENUE STREAM):', {
  x: 1.0, y: 5.7, w: 11.333, h: 0.25,
  fontSize: 10, bold: true, color: C_LIGHT_AMBER, fontFace: 'Arial'
});
s7.addText('Kerjasama Program Strategis: 51%  |  Pelatihan & Sertifikasi: 23%  |  Sewa Fasilitas & Coworking: 18%  |  Produksi & Produk Komersial: 8%', {
  x: 1.0, y: 6.0, w: 11.333, h: 0.6,
  fontSize: 12, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

s7.addNotes('Peta jalan transformasi ini kami rancang dalam 3 fase terukur. Pada periode 2026–2027, kami merapikan kelembagaan dan akreditasi lab. Masuk 2028, kami membentuk Strategic Business Unit (SBU) agar unit produksi bisa mandiri seperti korporasi swasta. Dan puncaknya pada 2030, Solo Technopark beroperasi layaknya holding company inovasi daerah yang sepenuhnya mandiri, membawa nama Kota Surakarta ke panggung teknologi global.');


// =========================================================================
// SLIDE 08: RE-SETTING TARGET FINANSIAL BLUD: RP 7,50 M ➔ RP 23,47 M
// =========================================================================
const s8 = pptx.addSlide();
addHeader(s8, 'RE-SETTING TARGET FINANSIAL: LOMPATAN RP 7,50 M ➔ RP 23,47 M', 'Data Riil Anggaran Pendapatan 2026–2030', 8);

// Left Column: Yearly Progression Table
const finTableData = [
  [
    { text: 'TAHUN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY }, align: 'center' } },
    { text: 'TARGET NOMINAL (RP)', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY }, align: 'right' } },
    { text: 'PERTUMBUHAN', options: { bold: true, color: C_WHITE, fill: { color: C_NAVY }, align: 'center' } }
  ],
  [
    { text: '2026', options: { bold: true, align: 'center' } },
    { text: 'Rp 7.503.392.698', options: { align: 'right' } },
    { text: 'Baseline', options: { align: 'center', bold: true } }
  ],
  [
    { text: '2027', options: { bold: true, align: 'center' } },
    { text: 'Rp 12.532.185.031', options: { align: 'right' } },
    { text: '+67.0%', options: { align: 'center', color: C_CYAN, bold: true } }
  ],
  [
    { text: '2028', options: { bold: true, align: 'center' } },
    { text: 'Rp 16.100.658.831', options: { align: 'right' } },
    { text: '+28.5%', options: { align: 'center', color: C_CYAN, bold: true } }
  ],
  [
    { text: '2029', options: { bold: true, align: 'center' } },
    { text: 'Rp 20.615.658.831', options: { align: 'right' } },
    { text: '+28.0%', options: { align: 'center', color: C_CYAN, bold: true } }
  ],
  [
    { text: '2030', options: { bold: true, align: 'center', color: C_AMBER } },
    { text: 'Rp 23.468.158.831', options: { align: 'right', bold: true, color: C_AMBER } },
    { text: 'MANDIRI (0% APBD)', options: { align: 'center', color: C_AMBER, bold: true } }
  ]
];

s8.addTable(finTableData, {
  x: 0.8, y: 1.25, w: 6.0, h: 2.8,
  colW: [1.1, 2.7, 2.2],
  border: { type: 'solid', pt: 1, color: C_BORDER },
  fill: { color: C_CARD },
  fontSize: 10,
  fontFace: 'Arial',
  valign: 'middle',
  rowH: [0.4, 0.45, 0.45, 0.45, 0.45, 0.55]
});

// Right Column: Breakdown Mesin Pendapatan 2030 (Full Height Symmetrical)
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 7.1, y: 1.25, w: 5.433, h: 5.5,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
});
s8.addShape(pptx.shapes.RECTANGLE, {
  x: 7.1, y: 1.25, w: 5.433, h: 0.45, fill: { color: C_BLUE }
});
s8.addText('RINCIAN MATA ANGGARAN PENDAPATAN 2030 (AUDIT RIIL)', {
  x: 7.25, y: 1.25, w: 5.15, h: 0.45,
  fontSize: 10.5, bold: true, color: C_WHITE, fontFace: 'Arial'
});

s8.addText([
  { text: '1. Kerjasama Program (51%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 12,82 Miliar\n', options: { bold: true, color: C_CYAN } },
  { text: '   • OGSCI Migas: Rp 6,25 M | Pemda e-Katalog: Rp 2,25 M\n   • Kemenperin: Rp 2,0 M | SMK GoGlobal: Rp 1,5 M\n\n', options: { color: C_MUTED, fontSize: 8.5 } },

  { text: '2. Layanan Vokasi & Sertifikasi (16%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 3,68 Miliar\n', options: { bold: true, color: C_CYAN } },
  { text: '   • Sertifikasi JFT: Rp 1,37 M | TUK Mandiri: Rp 1,35 M | Lab Data Center: Rp 950 jt\n\n', options: { color: C_MUTED, fontSize: 8.5 } },

  { text: '3. Diklat Mandiri Spesialis (11%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 2,47 Miliar\n', options: { bold: true, color: C_CYAN } },
  { text: '   • Welding Bawah Air: Rp 846 jt | Manufaktur: Rp 728 jt | AI & Cyber: Rp 225 jt\n\n', options: { color: C_MUTED, fontSize: 8.5 } },

  { text: '4. Sarana Kawasan & Event (10%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 2,45 Miliar\n', options: { bold: true, color: C_CYAN } },
  { text: '   • Pemanfaatan Event/Sarpras: Rp 1,5 M | Coworking & Virtual Office: Rp 900 jt\n\n', options: { color: C_MUTED, fontSize: 8.5 } },

  { text: '5. Rekayasa Manufaktur & Sparepart (6%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 1,46 Miliar\n', options: { bold: true, color: C_CYAN } },
  { text: '   • Komponen Presisi, Fabrikasi Mesin, & FabLab Industri\n\n', options: { color: C_MUTED, fontSize: 8.5 } },

  { text: '6. Kerjasama Pemanfaatan Lahan (3%): ', options: { bold: true, color: C_NAVY } },
  { text: 'Rp 579 Juta\n', options: { bold: true, color: C_MUTED } },
  { text: '   • Gedung Shopee & Kawasan Katulondi (Porsi sewa lahan ditekan minim)\n', options: { color: C_MUTED, fontSize: 8.5 } }
], {
  x: 7.3, y: 1.75, w: 5.05, h: 4.2,
  fontSize: 9.5, fontFace: 'Arial', lineSpacing: 11
});

// Bottom Note Box inside Right Card
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 7.25, y: 6.05, w: 5.133, h: 0.55,
  fill: { color: C_LIGHT_CYAN }, line: { color: C_CYAN, width: 1 }, rectRadius: 0.06
});
s8.addText('Audit Portofolio: Ketergantungan sewa lahan ditekan ke 3%, membuktikan STP bertransformasi menjadi pusat layanan riset & vokasi berdaya saing tinggi.', {
  x: 7.35, y: 6.08, w: 4.933, h: 0.48,
  fontSize: 8.5, italic: true, bold: true, color: C_BLUE, fontFace: 'Arial'
});

// Financial Growth Strategic Analysis Box (below left table) - Pure Typography
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 4.25, w: 6.0, h: 2.5,
  fill: { color: '0F1E36' }, line: { color: C_CYAN, width: 1.5 }, rectRadius: 0.1
});

// Header Pill inside Strategic Box
s8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 1.0, y: 4.4, w: 5.6, h: 0.38,
  fill: { color: C_BLUE }, rectRadius: 0.06
});
s8.addText('KESIMPULAN AUDIT FINANSIAL & KEMANDIRIAN BLUD', {
  x: 1.0, y: 4.4, w: 5.6, h: 0.38,
  fontSize: 10.5, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

// Big Metric Callout
s8.addText('LOMPATAN PENDAPATAN: +213% (3,1x LIPAT)', {
  x: 1.0, y: 4.88, w: 5.6, h: 0.4,
  fontSize: 15, bold: true, color: C_AMBER, align: 'center', fontFace: 'Arial'
});

// 3 Concrete Bullet Points
s8.addText([
  { text: '• Beban APBD 0% di 2030: ', options: { bold: true, color: C_LIGHT_AMBER } },
  { text: 'Seluruh biaya operasional UPTD STP tertutup penuh oleh pendapatan fungsional layanan tanpa subsidi kas daerah.\n', options: { color: 'E2E8F0' } },
  { text: '• Titik Impas (BEP) 2028: ', options: { bold: true, color: C_LIGHT_CYAN } },
  { text: 'Tercapai di target Rp 16,10 Miliar seiring beroperasinya Strategic Business Unit (SBU).\n', options: { color: 'E2E8F0' } },
  { text: '• Kualitas Pendapatan Berkelanjutan: ', options: { bold: true, color: C_LIGHT_GREEN } },
  { text: '88% disumbang keahlian riset & vokasi industri (program migas, uji lab, diklat spesialis).', options: { color: 'E2E8F0' } }
], {
  x: 1.0, y: 5.32, w: 5.6, h: 1.35,
  fontSize: 9.2, fontFace: 'Arial', lineSpacing: 13
});

s8.addNotes('Inilah angka komitmen kami, Bapak Walikota. Kami menaikkan target pendapatan BLUD STP lebih dari 300%—dari Rp 7,5 Miliar di 2026 menjadi Rp 23,47 Miliar di tahun 2030. Kenaikan ini bukan angan-angan, melainkan didorong diversifikasi nyata: program migas OGSCI menyumbang 6,2 miliar, jasa uji kompetensi keahlian menyumbang 3,6 miliar, dan pemanfaatan event menyumbang 1,5 miliar. STP siap menjadi BLUD sehat yang mandiri total tanpa perlu disubsidi biaya operasional dari APBD.');


// =========================================================================
// SLIDE 09: MANFAAT NYATA BAGI WARGA & KOTA SURAKARTA
// =========================================================================
const s9 = pptx.addSlide();
addHeader(s9, 'MANFAAT NYATA: DARI INOVASI MENUJU KESEJAHTERAAN MASYARAKAT', 'Dampak Sosial-Ekonomi bagi Kota Surakarta', 9);

const stats = [
  {
    num: '9.000',
    unit: 'Orang / Tahun',
    title: 'PENYERAPAN TENAGA KERJA',
    desc: 'Komitmen kuota afirmasi minimal 65% khusus pemegang KTP Surakarta, menekan angka pengangguran muda.',
    color: C_BLUE
  },
  {
    num: '80',
    unit: 'Startup & Unit Usaha',
    title: 'INKUBASI WIRAUSAHA BARU',
    desc: 'Melatih technopreneur muda, digitalisasi UMKM kuliner & batik, dan fasilitasi modal ventura.',
    color: C_CYAN
  },
  {
    num: '45',
    unit: 'Produk Inovasi Siap Pakai',
    title: 'KOMERSIALISASI RISET KOTA',
    desc: 'Solusi masalah kota: Smart Agriculture IoT, mesin tetas telur, alkes terjangkau, dan Smart City.',
    color: C_EMERALD
  },
  {
    num: 'Rp 180 M',
    unit: 'Investasi Swasta Masuk',
    title: 'PERTUMBUHAN EKONOMI DAERAH',
    desc: 'Menggerakkan multiplier effect ekonomi hotel MICE, kuliner, dan peningkatan setoran PAD Kota Surakarta.',
    color: C_AMBER
  }
];

stats.forEach((st, idx) => {
  const sx = 0.8 + idx * 2.95;
  s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: sx, y: 1.3, w: 2.8, h: 4.3,
    fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
  });
  
  // Top Accent Bar
  s9.addShape(pptx.shapes.RECTANGLE, {
    x: sx, y: 1.3, w: 2.8, h: 0.1, fill: { color: st.color }
  });

  // Big Number
  s9.addText(st.num, {
    x: sx + 0.1, y: 1.6, w: 2.6, h: 0.7,
    fontSize: 30, bold: true, color: st.color, align: 'center', fontFace: 'Arial'
  });
  s9.addText(st.unit, {
    x: sx + 0.1, y: 2.3, w: 2.6, h: 0.35,
    fontSize: 10.5, bold: true, color: C_MUTED, align: 'center', fontFace: 'Arial'
  });

  // Title
  s9.addText(st.title, {
    x: sx + 0.15, y: 2.85, w: 2.5, h: 0.55,
    fontSize: 11, bold: true, color: C_NAVY, align: 'center', fontFace: 'Arial'
  });

  // Description
  s9.addText(st.desc, {
    x: sx + 0.15, y: 3.5, w: 2.5, h: 1.9,
    fontSize: 9.5, color: C_TEXT, fontFace: 'Arial', lineSpacing: 14
  });
});

// Bottom Guarantee Banner
s9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 5.85, w: 11.733, h: 0.9,
  fill: { color: C_NAVY }, line: { color: C_CYAN, width: 1.2 }, rectRadius: 0.08
});
s9.addText('Tolak ukur keberhasilan tertinggi Solo Technopark bukan sekadar saldo kas BLUD, melainkan seberapa besar penurunan angka pengangguran dan peningkatan kesejahteraan riil bagi warga Kota Surakarta.', {
  x: 1.0, y: 5.85, w: 11.333, h: 0.9,
  fontSize: 11.5, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
});

s9.addNotes('Bapak Walikota, tolak ukur tertinggi kami bukan sekadar saldo kas BLUD, melainkan seberapa besar manfaat yang dirasakan langsung oleh warga Solo. Kami mengunci target penyerapan 9.000 tenaga kerja dengan kuota wajib minimal 65% bagi warga ber-KTP Surakarta. Kami mendampingi UMKM naik kelas dan menumbuhkan 80 wirausaha baru. Inilah bukti bahwa teknologi benar-benar hadir untuk menyejahterakan rakyat Solo.');


// =========================================================================
// SLIDE 10: TIGA PERMOHONAN KEBIJAKAN & RENCANA AKSI 100 HARI
// =========================================================================
const s10 = pptx.addSlide();
addHeader(s10, 'TIGA PERMOHONAN KEBIJAKAN KEPADA WALIKOTA & RENCANA AKSI 100 HARI', 'Dukungan Kepemimpinan & Komitmen Eksekusi', 10);

// Subtitle Notice
s10.addText('Manajemen Solo Technopark & BRIDA Tidak Memohon Tambahan Pagu APBD, Melainkan 3 Payung Regulasi Walikota:', {
  x: 0.8, y: 1.15, w: 11.733, h: 0.35,
  fontSize: 11.5, bold: true, color: C_BLUE, fontFace: 'Arial'
});

// 3 Leadership Asks Bento Cards
const asks = [
  {
    num: 'PERMOHONAN 1',
    title: 'INSTRUKSI WALIKOTA AFIRMASI e-KATALOG',
    color: C_BLUE,
    desc: 'Kebijakan afirmasi agar seluruh OPD/dinas Pemkot memprioritaskan belanja produk teknologi dan layanan binaan tenant STP melalui etalase e-Katalog Lokal (potensi serapan Rp 2,25 M).'
  },
  {
    num: 'PERMOHONAN 2',
    title: 'FLEKSIBILITAS KERJASAMA LAHAN BLUD',
    color: C_CYAN,
    desc: 'Regulasi pemanfaatan aset lahan dan sarana kawasan dengan skema Long-Term Revenue Sharing yang luwes bagi investor teknologi global untuk menanamkan modal di STP.'
  },
  {
    num: 'PERMOHONAN 3',
    title: 'SURAT EDARAN KONSORSIUM RISET KAMPUS',
    color: C_EMERALD,
    desc: 'Instruksi agar dana riset terapan perguruan tinggi se-Solo Raya (UNS, UMS, ISI, Poltek) diintegrasikan menyelesaikan permasalahan kota di STP sebagai pusat hilirisasi.'
  }
];

asks.forEach((a, idx) => {
  const ax = 0.8 + idx * 3.95;
  s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: ax, y: 1.6, w: 3.8, h: 2.8,
    fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
  });
  s10.addShape(pptx.shapes.RECTANGLE, {
    x: ax, y: 1.6, w: 3.8, h: 0.55, fill: { color: a.color }
  });
  s10.addText(a.num, {
    x: ax + 0.15, y: 1.65, w: 3.5, h: 0.22,
    fontSize: 9, bold: true, color: C_LIGHT_AMBER, align: 'center', fontFace: 'Arial'
  });
  s10.addText(a.title, {
    x: ax + 0.15, y: 1.88, w: 3.5, h: 0.25,
    fontSize: 9.5, bold: true, color: C_WHITE, align: 'center', fontFace: 'Arial'
  });
  s10.addText(a.desc, {
    x: ax + 0.2, y: 2.3, w: 3.4, h: 1.9,
    fontSize: 10, color: C_TEXT, fontFace: 'Arial', lineSpacing: 15
  });
});

// 100-Day Action Plan Strip
s10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.8, y: 4.65, w: 11.733, h: 2.1,
  fill: { color: C_CARD }, line: { color: C_BORDER, width: 1.2 }, rectRadius: 0.1
});
s10.addText('RENCANA AKSI 100 HARI PERTAMA (QUICK WINS PASCA-ARAHAN WALIKOTA SURAKARTA)', {
  x: 1.0, y: 4.8, w: 11.333, h: 0.35,
  fontSize: 11, bold: true, color: C_NAVY, fontFace: 'Arial'
});

const quickWins = [
  {
    day: 'HARI 1–30',
    title: 'HARMONISASI REGULASI',
    desc: 'Penyusunan draft Instruksi Walikota untuk e-Katalog Lokal Inovasi STP bersama Bagian Hukum Setda Kota Surakarta.'
  },
  {
    day: 'HARI 31–60',
    title: 'SHOWCASE 5 PRODUK UNGGULAN',
    desc: 'Gelar demonstrasi 5 produk hilirisasi (Smart Agriculture IoT & Alkes) di hadapan Walikota dan seluruh Kepala OPD.'
  },
  {
    day: 'HARI 61–100',
    title: 'MOU KONSORSIUM 5 KAMPUS',
    desc: 'Penandatanganan nota kesepahaman konsorsium riset terapan kampus bersama Kadin & Apindo di Solo Technopark.'
  }
];

quickWins.forEach((qw, idx) => {
  const qx = 1.0 + idx * 3.8;
  s10.addShape(pptx.shapes.RECTANGLE, {
    x: qx, y: 5.2, w: 3.6, h: 1.35,
    fill: { color: C_BG }, line: { color: C_BORDER, width: 1 }
  });
  s10.addText(qw.day, {
    x: qx + 0.15, y: 5.28, w: 3.3, h: 0.22,
    fontSize: 9.5, bold: true, color: C_EMERALD, fontFace: 'Arial'
  });
  s10.addText(qw.title, {
    x: qx + 0.15, y: 5.5, w: 3.3, h: 0.25,
    fontSize: 10, bold: true, color: C_NAVY, fontFace: 'Arial'
  });
  s10.addText(qw.desc, {
    x: qx + 0.15, y: 5.75, w: 3.3, h: 0.75,
    fontSize: 8.5, color: C_TEXT, fontFace: 'Arial', lineSpacing: 11
  });
});

s10.addNotes('Menutup paparan ini, Bapak Walikota: Manajemen Solo Technopark dan BRIDA tidak meminta tambahan alokasi belanja anggaran daerah. Kami hanya memohon tiga payung kepemimpinan kebijakan dari Bapak: perlindungan pasar lokal lewat e-Katalog, fleksibilitas kerja sama lahan, dan mandat sinergi kampus. Dengan tiga payung regulasi dari Bapak Walikota, tim kami siap langsung tancap gas mengeksekusi rencana aksi 100 hari pertama. Bersama Bapak, Solo Technopark siap menjadi kebanggaan nasional. Terima kasih, mohon arahan dan bimbingan Bapak Walikota.');

// Output paths
const outputPptxPathV2 = path.resolve(__dirname, 'PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pptx');
const outputPptxPathOut = path.resolve(__dirname, '../output_paparan/PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pptx');

pptx.writeFile({ fileName: outputPptxPathV2 })
  .then(() => {
    console.log(`[SUCCESS] PPTX Master V2 generated at: ${outputPptxPathV2}`);
    fs.copyFileSync(outputPptxPathV2, outputPptxPathOut);
    console.log(`[SUCCESS] Copied to output_paparan: ${outputPptxPathOut}`);
  })
  .catch(err => {
    console.error('[ERROR] Failed generating PPTX:', err);
    process.exit(1);
  });
