// pdfview.js — renders a PDF with PDF.js and paints highlight boxes over it.
import * as pdfjsLib from '../vendor/pdfjs/pdf.min.mjs';
import { buildIndex, itemPieces, extractDois } from './matcher.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdfjs/pdf.worker.min.mjs', import.meta.url).href;

/** Parse a PDF (ArrayBuffer) into {doc, pages:[{page, items}], index}. */
export async function loadPdf(data) {
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(data.slice(0)) }).promise;
  const pages = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const tc = await page.getTextContent();
    pages.push({ page, items: tc.items, styles: tc.styles });
  }
  return { doc, pages, index: buildIndex(pages) };
}

/** DOIs printed on the first two pages — used to match doi.org links. */
export async function pdfDois(data) {
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(data.slice(0)) }).promise;
  let text = '';
  for (let i = 1; i <= Math.min(2, doc.numPages); i++) {
    const tc = await (await doc.getPage(i)).getTextContent();
    text += tc.items.map(it => it.str).join(' ') + ' ';
  }
  doc.destroy();
  return extractDois(text);
}

/**
 * Render all pages into `container` and paint highlights.
 * Returns the element of the first (most important) highlight, for scrolling.
 */
export async function renderPdf(container, pdf, highlights) {
  container.innerHTML = '';
  const width = Math.max(200, container.clientWidth - 8);
  const dpr = window.devicePixelRatio || 1;
  let first = null;

  for (let p = 0; p < pdf.pages.length; p++) {
    const { page, items, styles } = pdf.pages[p];
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: width / base.width });

    const wrap = document.createElement('div');
    wrap.className = 'page';
    wrap.style.width = `${viewport.width}px`;
    wrap.style.height = `${viewport.height}px`;
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;
    wrap.appendChild(canvas);
    const layer = document.createElement('div');
    layer.className = 'hl-layer';
    wrap.appendChild(layer);
    container.appendChild(wrap);

    await page.render({
      canvasContext: canvas.getContext('2d'),
      viewport,
      transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null,
    }).promise;

    for (const h of highlights.filter(h => h.page === p)) {
      for (const piece of itemPieces(pdf.index.pages[p], h.start, h.end)) {
        const box = itemBox(items[piece.item], piece, viewport, styles);
        if (!box) continue;
        const el = document.createElement('div');
        el.className = `hl hl-${h.level}`;
        Object.assign(el.style, { left: `${box.x}px`, top: `${box.y}px`, width: `${box.w}px`, height: `${box.h}px` });
        el.title = `${labelOf(h.level)}${h.how === 'auto' ? ` · match ${h.score.toFixed(2)}` : ''}`;
        layer.appendChild(el);
        if (!first) first = el;
      }
    }
  }
  return first;
}

const measureCtx = document.createElement('canvas').getContext('2d');

/** Width of a substring relative to the whole item, using the item's font family (proportional fonts). */
function fraction(item, styles, n) {
  if (n <= 0) return 0;
  if (n >= item.str.length) return 1;
  const fam = (styles && styles[item.fontName] && styles[item.fontName].fontFamily) || 'sans-serif';
  measureCtx.font = `100px ${fam}`;
  const full = measureCtx.measureText(item.str).width;
  return full > 0 ? measureCtx.measureText(item.str.slice(0, n)).width / full : n / item.str.length;
}

function itemBox(item, piece, viewport, styles) {
  if (!item || !item.str) return null;
  const tx = pdfjsLib.Util.transform(viewport.transform, item.transform);
  const fontH = Math.hypot(tx[2], tx[3]);
  const fullW = item.width * viewport.scale;
  const a = fraction(item, styles, piece.from);
  const b = fraction(item, styles, piece.to);
  const x = tx[4] + fullW * a;
  const w = fullW * (b - a);
  if (w <= 0 || fontH <= 0) return null;
  return { x, y: tx[5] - fontH * 0.9, w, h: fontH * 1.15 };
}

export function labelOf(level) {
  return { key: 'Key part', rel: 'Relevant', some: 'Some relevance' }[level] || level;
}
