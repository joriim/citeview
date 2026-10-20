// Node test: runs the same matcher + highlight geometry as the add-in on the sample draft and PDF.
// Usage: node test/run-test.mjs   (from the project root)  -> writes test/out/boxes.json
import { readFileSync, writeFileSync, mkdirSync, readdirSync } from 'node:fs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildIndex, findHighlights, itemPieces, extractDois } from '../src/matcher.js';

const dir = new URL('./sample/', import.meta.url);
const cites = JSON.parse(readFileSync(new URL('citations.json', dir)));
const srcDir = new URL('sources/', dir);
const files = readdirSync(srcDir).filter(f => f.endsWith('.pdf'));

async function load(name) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(new URL(name, srcDir))) }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    pages.push({ page, items: (await page.getTextContent()).items });
  }
  return { pages, index: buildIndex(pages) };
}

const loaded = {};
for (const f of files) {
  loaded[f] = await load(f);
  loaded[f].dois = extractDois(loaded[f].index.pages.slice(0, 2).map(p => p.raw).join(' '));
}

const SCALE = 1.5; // must match the pdftoppm render in draw_boxes.py (108 dpi)
const result = [];
let fail = 0;
for (const c of cites) {
  const base = decodeURIComponent(c.address).split(/[?#]/)[0].split(/[\\/]/).pop().toLowerCase();
  const dois = extractDois(decodeURIComponent(c.address));
  const name = files.find(f => f.toLowerCase() === base) || files.find(f => loaded[f].dois.some(d => dois.includes(d)));
  console.log(`\n${c.text}  ->  ${c.address}  ->  ${name || 'NOT RESOLVED'}`);
  if (!name) { fail++; continue; }
  const pdf = loaded[name];
  const { highlights, missing } = findHighlights(pdf.index, c.comment, c.paraText);
  for (const h of highlights) {
    const t = pdf.index.pages[h.page].raw.slice(h.start, h.end).replace(/\s+/g, ' ');
    console.log(`  [${h.level.padEnd(4)}] p${h.page + 1} ${h.how.padEnd(5)} ${h.score.toFixed(2)}  ${t.slice(0, 95)}`);
  }
  if (missing.length) { console.log('  MISSING:', missing); fail++; }
  const boxes = [];
  for (const h of highlights) {
    const { page, items } = pdf.pages[h.page];
    const vp = page.getViewport({ scale: SCALE });
    for (const pc of itemPieces(pdf.index.pages[h.page], h.start, h.end)) {
      const it = items[pc.item];
      const tx = pdfjs.Util.transform(vp.transform, it.transform);
      const fh = Math.hypot(tx[2], tx[3]);
      const fw = it.width * SCALE, len = it.str.length || 1;
      boxes.push({ page: h.page, level: h.level, x: tx[4] + fw * pc.from / len, y: tx[5] - fh * 0.9, w: fw * (pc.to - pc.from) / len, h: fh * 1.15 });
    }
  }
  result.push({ cite: c.text, file: name, boxes, levels: highlights.map(h => h.level) });
}
mkdirSync(new URL('./out/', import.meta.url), { recursive: true });
writeFileSync(new URL('./out/boxes.json', import.meta.url), JSON.stringify(result));
console.log(fail ? `\n${fail} problem(s)` : '\nAll citations resolved, all quotes found.');
process.exit(fail ? 1 : 0);
