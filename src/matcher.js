// matcher.js — finds which passages of a source PDF a citation comment refers to.
// Pure logic, no Office or DOM dependencies (also runs in Node for testing).
//
// Levels:  key  = purple  (explicit quotes in the comment)
//          rel  = green   (strong lexical match to the comment text)
//          some = yellow  (weaker match)
//
// Comment convention (all optional):
//   [key] "exact passage"     -> purple
//   [rel] "exact passage"     -> green
//   [some] "exact passage"    -> yellow
//   "untagged quote"          -> purple
//   any other free text       -> automatic green/yellow matching

export const LEVELS = ['key', 'rel', 'some'];

const TAG_ALIASES = {
  key: 'key', k: 'key', purple: 'key', avain: 'key',
  rel: 'rel', r: 'rel', green: 'rel', relevant: 'rel', olennainen: 'rel',
  some: 'some', s: 'some', yellow: 'some', jonkin: 'some', osittain: 'some',
};

const STOP = new Set(`a an the and or but if then than that this these those there their they them it its
is are was were be been being of in on at to for from by with as into onto over under between among
not no nor so such can could may might must shall should will would do does did done have has had
we our us you your he she his her i me my which who whom whose what when where why how also both
each other more most less least very much many some any all only own same just about above below
after before during while within without via per et al fig figure table eg ie however thus therefore
here using used use based study studies result results show shows shown found paper article`
  .split(/\s+/));

// ---------- text normalisation with offset map ----------

const CHAR_FIX = {
  '‘': "'", '’': "'", '“': '"', '”': '"',
  '–': '-', '—': '-', '−': '-', ' ': ' ',
};

/** Normalise raw text; returns {norm, map} where map[i] = raw index of norm[i]. */
export function normalizeWithMap(raw, eolSet = new Set()) {
  const out = [];
  const map = [];
  let lastSpace = true;
  for (let i = 0; i < raw.length; i++) {
    let c = raw[i];
    if (c === '­') continue; // soft hyphen
    // hyphenation at a line break: "nitro-" + EOL + "gen" -> "nitrogen"
    if (c === '-' && eolSet.has(i + 1) && /\p{L}/u.test(raw[i - 1] || '') && /\p{Ll}/u.test(raw[i + 2] || '')) {
      i++; // skip the EOL separator too
      continue;
    }
    c = CHAR_FIX[c] || c;
    if (/\s/.test(c)) {
      if (!lastSpace) { out.push(' '); map.push(i); lastSpace = true; }
      continue;
    }
    const n = c.normalize('NFKC').toLowerCase(); // ligatures: "ﬁ" -> "fi"
    for (const ch of n) { out.push(ch); map.push(i); }
    lastSpace = false;
  }
  return { norm: out.join(''), map };
}

export function normalize(text) {
  return normalizeWithMap(text).norm.trim();
}

function stem(w) {
  if (w.length > 5 && w.endsWith('ing')) return w.slice(0, -3);
  if (w.length > 4 && w.endsWith('ed')) return w.slice(0, -2);
  if (w.length > 4 && w.endsWith('es')) return w.slice(0, -2);
  if (w.length > 3 && w.endsWith('s') && !w.endsWith('ss')) return w.slice(0, -1);
  return w;
}

export function tokens(normText) {
  const words = normText.match(/[\p{L}\p{N}][\p{L}\p{N}-]*/gu) || [];
  return words.filter(w => w.length >= 3 && !STOP.has(w) && !/^\d+$/.test(w)).map(stem);
}

// ---------- index of a PDF ----------

/**
 * pages: [{ items: [{ str, hasEOL }] }]  (shape of pdf.js getTextContent().items)
 * Returns an index used by findHighlights and by the viewer to map offsets to items.
 */
export function buildIndex(pages) {
  const idxPages = [];
  const sentences = [];
  pages.forEach((p, pageNo) => {
    let raw = '';
    const itemSpans = [];
    const eolSet = new Set();
    for (const it of p.items) {
      const str = it.str || '';
      const start = raw.length;
      raw += str;
      itemSpans.push({ start, end: raw.length });
      if (it.hasEOL) { eolSet.add(raw.length); raw += ' '; }
    }
    const { norm, map } = normalizeWithMap(raw, eolSet);
    idxPages.push({ raw, norm, map, itemSpans });

    // sentence split on the normalised text
    const re = /[^.!?]+(?:[.!?]+|$)/g;
    let m;
    while ((m = re.exec(norm))) {
      const s = m[0];
      const lead = s.length - s.trimStart().length;
      const text = s.trim();
      if (text.length < 25) continue;
      if (/https?:|doi\.org|©|copyright|all rights reserved/.test(text)) continue; // metadata lines
      const toks = tokens(text);
      if (toks.length < 3) continue;
      sentences.push({ page: pageNo, nStart: m.index + lead, nEnd: m.index + lead + text.length, text, toks });
    }
  });

  // idf over sentences
  const df = new Map();
  for (const s of sentences) for (const t of new Set(s.toks)) df.set(t, (df.get(t) || 0) + 1);
  const N = sentences.length || 1;
  const idf = t => Math.log((N + 1) / ((df.get(t) || 0) + 1)) + 1;
  for (const s of sentences) s.vec = vectorize(s.toks, idf);

  return { pages: idxPages, sentences, idf };
}

function vectorize(toks, idf) {
  const tf = new Map();
  for (const t of toks) tf.set(t, (tf.get(t) || 0) + 1);
  const v = new Map();
  let norm2 = 0;
  for (const [t, c] of tf) { const w = (1 + Math.log(c)) * idf(t); v.set(t, w); norm2 += w * w; }
  const len = Math.sqrt(norm2) || 1;
  for (const [t, w] of v) v.set(t, w / len);
  return v;
}

function cosine(a, b) {
  let s = 0;
  const [small, big] = a.size < b.size ? [a, b] : [b, a];
  for (const [t, w] of small) { const o = big.get(t); if (o) s += w * o; }
  return s;
}

// ---------- comment parsing ----------

export function parseComment(text) {
  const quotes = [];
  const quoteRe = /(?:\[(\w+)\]\s*)?["“”«»]([^"“”«»]{8,})["“”«»]/g;
  let free = text || '';
  let m;
  while ((m = quoteRe.exec(text || ''))) {
    const level = TAG_ALIASES[(m[1] || 'key').toLowerCase()] || 'key';
    quotes.push({ level, text: m[2].trim() });
  }
  free = free.replace(quoteRe, ' ').replace(/\[(\w+)\]/g, ' ');
  return { quotes, free: free.trim() };
}

// ---------- matching ----------

function rawSpan(page, nStart, nEnd) {
  return { start: page.map[nStart], end: page.map[Math.max(nStart, nEnd - 1)] + 1 };
}

function locateQuote(index, quote) {
  const q = normalize(quote).replace(/^[\s.,;:…]+|[\s.,;:…]+$/g, '');
  if (!q) return null;
  // 1) exact substring
  for (let p = 0; p < index.pages.length; p++) {
    const at = index.pages[p].norm.indexOf(q);
    if (at >= 0) return { page: p, ...rawSpan(index.pages[p], at, at + q.length), score: 1, how: 'exact' };
  }
  // 2) fuzzy: best window of 1–3 consecutive sentences by token recall
  const qt = tokens(q);
  if (qt.length < 2) return null;
  const qset = new Set(qt);
  let best = null;
  const S = index.sentences;
  for (let i = 0; i < S.length; i++) {
    const seen = new Set();
    for (let w = 0; w < 3 && i + w < S.length && S[i + w].page === S[i].page; w++) {
      for (const t of S[i + w].toks) if (qset.has(t)) seen.add(t);
      const recall = seen.size / qset.size;
      const score = recall - 0.05 * w; // prefer shorter windows
      if (!best || score > best.score) best = { score, i, w, recall };
      if (recall === 1) break;
    }
  }
  if (!best || best.recall < 0.6) return null;
  const a = S[best.i], b = S[best.i + best.w];
  return { page: a.page, ...rawSpan(index.pages[a.page], a.nStart, b.nEnd), score: best.recall, how: 'fuzzy' };
}

/**
 * Returns { highlights: [{page,start,end,level,score,how}], missing: [quote texts not found] }
 * start/end are offsets into index.pages[page].raw.
 */
export function findHighlights(index, commentText, citingText = '', opts = {}) {
  const { maxRel = 6, maxSome = 12, relMin = 0.22, someMin = 0.12 } = opts;
  const { quotes, free } = parseComment(commentText || '');
  const highlights = [];
  const missing = [];

  for (const q of quotes) {
    const hit = locateQuote(index, q.text);
    if (hit) highlights.push({ ...hit, level: q.level });
    else missing.push(q.text);
  }

  // automatic matching from free text (+ the citing sentence, weighted lower)
  const freeToks = tokens(normalize(free));
  const citeToks = tokens(normalize(citingText));
  if (freeToks.length + citeToks.length >= 2 && index.sentences.length) {
    const qv = vectorize([...freeToks, ...freeToks, ...citeToks], index.idf);
    const taken = (s) => highlights.some(h => {
      if (h.page !== s.page) return false;
      const r = rawSpan(index.pages[s.page], s.nStart, s.nEnd);
      return r.start < h.end && h.start < r.end;
    });
    const scored = index.sentences
      .map(s => ({ s, score: cosine(qv, s.vec) }))
      .filter(x => x.score >= someMin)
      .sort((a, b) => b.score - a.score);
    const top = scored.length ? scored[0].score : 0;
    let nRel = 0, nSome = 0;
    for (const { s, score } of scored) {
      if (taken(s)) continue;
      const rel = score >= relMin && score >= 0.5 * top && nRel < maxRel;
      if (rel) nRel++;
      else if (nSome < maxSome) nSome++;
      else continue;
      highlights.push({ page: s.page, ...rawSpan(index.pages[s.page], s.nStart, s.nEnd), level: rel ? 'rel' : 'some', score, how: 'auto' });
    }
  }

  const order = { key: 0, rel: 1, some: 2 };
  highlights.sort((a, b) => order[a.level] - order[b.level] || a.page - b.page || a.start - b.start);
  return { highlights, missing };
}

/** Split a raw-offset highlight into per-item pieces: [{item, from, to}] (char offsets within item.str). */
export function itemPieces(indexPage, start, end) {
  const out = [];
  indexPage.itemSpans.forEach((sp, item) => {
    const a = Math.max(start, sp.start), b = Math.min(end, sp.end);
    if (b > a) out.push({ item, from: a - sp.start, to: b - sp.start });
  });
  return out;
}

/** Find DOIs in text (used to match "doi.org/..." links to loaded PDFs). */
export function extractDois(text) {
  const m = (text || '').match(/10\.\d{4,9}\/[^\s"<>]+/gi) || [];
  return [...new Set(m.map(d => d.replace(/[.,;)\]]+$/, '').toLowerCase()))];
}
