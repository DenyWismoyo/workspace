import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync } from 'fs';

const pdfPath1 = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\STRATEGI  PENGEMBANGAN KAWASAN STP 2027 - 2030-1 (1).pdf';
const pdfPath2 = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\Tambahan Slide Materi Walkot (1).pdf';

async function extractText(pdfPath, label) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`DOKUMEN: ${label}`);
  console.log('='.repeat(60));
  
  const buffer = readFileSync(pdfPath);
  const data = new Uint8Array(buffer);
  const doc = await getDocument({ data }).promise;
  console.log(`Total halaman: ${doc.numPages}`);
  
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map(item => item.str).join(' ');
    console.log(`\n--- HALAMAN ${i} ---`);
    console.log(strings.substring(0, 3000));
  }
}

try {
  await extractText(pdfPath1, 'STRATEGI PENGEMBANGAN KAWASAN STP 2027-2030');
  await extractText(pdfPath2, 'TAMBAHAN SLIDE MATERI WALKOT');
} catch(e) {
  console.error('Error:', e.message, e.stack);
}
