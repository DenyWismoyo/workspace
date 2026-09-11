import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync } from 'fs';

const pdfPath1 = 'D:\\Project\\GAWE\\Paparan Roadmap STP\\STRATEGI  PENGEMBANGAN KAWASAN STP 2027 - 2030-1 (1).pdf';

async function extractFull(pdfPath, label) {
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
    console.log(strings); // full text, no truncation
  }
}

try {
  await extractFull(pdfPath1, 'STRATEGI PENGEMBANGAN KAWASAN STP 2027-2030');
} catch(e) {
  console.error('Error:', e.message);
}
