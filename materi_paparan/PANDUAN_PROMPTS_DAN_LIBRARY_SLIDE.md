# PANDUAN LIBRARY, PROMPT GAMBAR, DAN INTEGRASI CLARIO AI
## Bahan Paparan Roadmap Solo Technopark (STP) 2026–2030

Dokumen ini melengkapi file [materi_paparan_roadmap_stp.md](file:///d:/Project/GAWE/Paparan%20Roadmap%20STP/materi_paparan/materi_paparan_roadmap_stp.md) dengan menyediakan:
1. Rekomendasi library dan perangkat lunak untuk mengubah teks Markdown menjadi slide presentasi profesional.
2. Katalog prompt gambar AI fotorealistik per slide (Slide 1 s.d. Slide 10).
3. Galeri aset gambar yang telah di-generate ke dalam folder proyek.
4. Panduan pemanfaatan Clario Cloud AI (via MCP & Helper Script).

---

## 🛠️ BAGIAN 1: REKOMENDASI LIBRARY PEMBUATAN SLIDE

Bahan paparan yang telah dibuat dalam format Markdown (.md) dapat diubah menjadi slide presentasi elegan (PowerPoint .pptx, PDF, atau Slide Web) menggunakan pilihan library berikut:

| No | Library / Tool | Tipe & Teknologi | Keunggulan Utama | Cocok Digunakan Untuk |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Marp (Markdown Presentation Ecosystem)** | CLI / VS Code Extension | Konversi langsung file `.md` menjadi file PPTX, PDF, atau HTML dengan tema modern. Mendukung custom CSS warna Pemkot & STP. | **Paling Direkomendasikan** untuk workflow cepat dari file markdown yang sudah ada. |
| **2** | **Slidev** | Web Framework (Vite + Vue) | Slide berbasis web interaktif, mendukung animasi transisi halus, syntax highlighting, grafik Mermaid, dan responsif di layar sentuh/proyektor. | Presentasi digital interaktif di hadapan pimpinan atau ruang pameran. |
| **3** | **PptxGenJS / python-pptx** | Programmatic Library (JS / Python) | Meng-generate file PowerPoint (.pptx) asli secara otomatis dengan presisi tata letak koordinat pixel, shape, dan warna hex. | Otomasi generate slide korporat berskala banyak. |
| **4** | **Gamma App / Tome** | Platform Web Berbasis AI | Cukup *copy-paste* teks per slide dari file `.md`, AI akan otomatis menyusun kartu visual, kartu ringkasan, dan layout secara estetis. | Desain kilat dengan estetika modern tanpa perlu coding. |
| **5** | **Canva / Microsoft PowerPoint** | Desktop / Cloud App | Mengimpor poin teks ke dalam template bertema *Modern Government / Tech Innovation*. | Eksekusi manual oleh staf desainer grafis. |

> [!TIP]
> **Cara Cepat Menggunakan Marp**:
> 1. Pasang ekstensi **Marp for VS Code** di IDE Anda.
> 2. Buka file `.md`, tambahkan baris `marp: true` di paling atas.
> 3. Klik tombol *Export Slide to PPTX / PDF*.

---

## 🎨 BAGIAN 2: KATALOG PROMPT GAMBAR AI PER SLIDE (16:9)

Gunakan formula prompt berikut di generator gambar AI (seperti Midjourney, FLUX 2 Pro, DALL-E 3, atau tool bawaan):

### Slide 01: Cover & Visi Besar Transformasi
```text
Modern futuristic architecture of Solo Technopark building at twilight, sleek curved glass and blue metallic facade with glowing cyan and warm amber accent lighting, subtle elegant Indonesian batik geometric motifs subtly reflected on glass panels, lush green landscaped technology park campus, wide aerial drone shot, premium corporate executive presentation background, ultra-realistic, cinematic 8k resolution, 16:9
```

### Slide 02: Kedudukan Kelembagaan BRIDA – UPTD STP
```text
Sleek corporate boardroom in Indonesia, executive government officials and tech leaders in batik shirts and suits discussing policy around a high-tech smart glass table with glowing holographic data charts showing governance and research ecosystem, clean minimal interior design, soft morning sunlight through floor-to-ceiling windows, professional corporate photography, 16:9
```

### Slide 03: 4 Pilar Portofolio Layanan BLUD STP
```text
Split quad composition showing four specialized modern technical vocations: 1. Professional underwater welding with glowing blue electric arc underwater, 2. High-precision CNC mechanical engineering workshop, 3. Modern tech startup incubator with laptops and whiteboards, 4. Vocational high school students learning robotics. Premium industrial tech photography, sharp focus, 16:9
```

### Slide 04: Pipeline Hilirisasi Riset ke Pasar
```text
Smart agricultural technology laboratory in Solo Indonesia, young Indonesian engineer in lab coat inspecting an IoT-controlled robotic coffee sorting machine with glowing sensors, prototypes of green technology devices on clean workbench, blur background with hydroponics greenhouse, hyper-realistic, vivid color, 16:9
```

### Slide 05: Ekosistem Kemitraan & Tech Hubs (95+ Mitra)
```text
Futuristic technology innovation laboratory at Solo Technopark, Indonesian young male and female engineers collaborating on artificial intelligence, robotics, precision engineering, glowing holographic dashboards and cyber security data visualizations, clean professional lighting, premium corporate presentation slide background, cinematic 8k, 16:9
```

### Slide 06: Gap Analysis & Ruang Pembenahan
```text
Strategic planning concept, blueprint architectural drawings of a tech park transitioning from current state wireframe into vibrant glowing 3D completed buildings, glowing data graphs showing growth and gap bridging, navy blue background, sleek isometric infographic aesthetic, 16:9
```

### Slide 07: Roadmap Pengembangan STP 2026–2030
```text
Futuristic glowing roadmap bridge connecting three distinct innovation islands from 2026 to 2030, digital particles, sleek corporate style, illuminated steps leading towards a glowing global technopark city skyline at sunrise, clean minimalist corporate graphic, 16:9
```

### Slide 08: Re-Setting Target Finansial BLUD
```text
Ascending modern 3D translucent glass bar chart glowing in emerald green and vibrant gold against dark navy background, representing exponential financial revenue growth from 7.5B to 23.4B IDR, clean minimalist financial infographic, premium wealth and investment aesthetic, 8k, 16:9
```

### Slide 09: Manfaat Nyata bagi Masyarakat Solo
```text
Inspiring documentary-style photo of young Indonesian professionals and vocational students in Surakarta working proudly in high-tech digital studio and advanced manufacturing plant, smiling confidently, vibrant atmosphere of thriving economic opportunity and job creation, warm natural lighting, high resolution, 16:9
```

### Slide 10: Dukungan Kebijakan Walikota & Call to Action
```text
Prestigious city hall executive handshake between a visionary Indonesian city mayor and innovation park director, golden hour sunlight in a modern government office overlooking the panorama of Solo city with Mount Merapi in the far distance, inspiring leadership, cinematic corporate portrait, 16:9
```

---

## 🖼️ BAGIAN 3: ASET GAMBAR YANG TELAH DIGENERATE

Dua gambar utama berkualitas 16:9 telah selesai di-generate dan tersimpan langsung di folder proyek Anda:

1. **Cover Slide (Gedung Ikonik Solo Technopark Modern)**
   - Berkas: `materi_paparan/assets/slide_01_cover.jpg`
   - Karakter: Fasad gedung melengkung modern Solo Technopark di waktu senja dengan lampu aksen *cyan*, taman hijau, dan latar siluet pegunungan Surakarta.

2. **Lab Inovasi & Hilirisasi Riset (Slide 4 & 5)**
   - Berkas: `materi_paparan/assets/slide_04_05_innovation_lab.jpg`
   - Karakter: Insinyur muda Indonesia menguji lengan robotik, sensor IoT pertanian cerdas, dan visualisasi AI di dalam laboratorium berteknologi tinggi.

---

## 🤖 BAGIAN 4: CARA PENGGUNAAN CLARIO CLOUD AI

Anda telah memiliki API Key Clario yang aktif (`sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482`).

### 1. Menggunakan Script Helper Bawaan (`clario_helper.mjs`)
Skrip ini sudah tersedia di dalam folder `materi_paparan/` untuk menjalankan query AI langsung dari terminal:

```bash
# Tanya ide atau brainstorming (menggunakan model default clario/glm-5.3-flash)
node materi_paparan/clario_helper.mjs "Berikan 3 analogi menarik untuk menjelaskan BLUD kepada Walikota"

# Menggunakan model lain untuk coding atau analisis arsitektur
node materi_paparan/clario_helper.mjs "Beri saran optimasi layout slide" --model clario/gpt-5.6-sol
```

### 2. Menggunakan cURL Langsung (Semua Terminal)
```bash
curl.exe -s -X POST https://clario.apicloud.my.id/v1/chat/completions \
  -H "Authorization: Bearer sk-clario-8a0483a00917b34dab1b663416cf14242849cea57bd92482" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"clario/glm-5.3-flash\",\"messages\":[{\"role\":\"user\",\"content\":\"Halo Clario\"}]}"
```
