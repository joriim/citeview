// taskpane.js — one source pane (left or right, from ?slot= in the URL).
// Both panes look at the citations around the cursor and each takes its own one:
// left = citation at/after the cursor, right = the one after that.
// Each pane has its own colour, used both on the citation in the draft and on its source highlights.
/* global Office, Word */
import { loadPdf, pdfDois, renderPdf } from './pdfview.js';
import { findHighlights, extractDois } from './matcher.js';
import { SLOTS, assignSlots, citationKey, isSlotColor } from './slots.js';
import * as store from './store.js';

const SLOT = new URLSearchParams(location.search).get('slot') === 'right' ? 'right' : 'left';
const OTHER = SLOT === 'left' ? 'right' : 'left';
const PARAS_BEFORE = 2; // window of paragraphs around the cursor that counts as "on screen"
const PARAS_AFTER = 4;

const state = {
  current: null,   // {key, address, text, comment, paraText}
  marked: null,    // Word range we coloured in the draft
  files: [],
  pdfCache: new Map(),
  seq: 0,
  rendered: null,
  commentsOk: false,
  busy: false,
  again: false,
};

const $ = (id) => document.getElementById(id);
const status = (t) => { $('status').textContent = t; };
const ls = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch { /* ignore */ } },
};
const pinKey = (slot) => `citeview.pin.${slot}`;

document.documentElement.style.setProperty('--slot', SLOTS[SLOT].rgb);
document.documentElement.style.setProperty('--slot-ink', SLOTS[SLOT].ink);
document.documentElement.dataset.slot = SLOT;

Office.onReady(async (info) => {
  $('slotName').textContent = `${SLOTS[SLOT].name} source`;
  $('otherName').textContent = SLOTS[OTHER].name.toLowerCase();
  if (info.host !== Office.HostType.Word) { status('Open this add-in in Word.'); return; }
  state.commentsOk = Office.context.requirements.isSetSupported('WordApi', '1.4');
  bindUi();
  await refreshFiles();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, debounce(() => update(), 250));
  window.addEventListener('storage', (e) => {
    if (e.key === pinKey(OTHER) || e.key === pinKey(SLOT)) { syncPinUi(); update(true); }
    if (e.key === 'citeview.files') refreshFiles();
  });
  syncPinUi();
  status(state.commentsOk ? 'Ready' : 'Ready (this Word version cannot read comments)');
  update(true);
});

// ---------------- Word side ----------------

const AT = new Set(['Equal', 'Contains', 'ContainsStart', 'ContainsEnd', 'OverlapsBefore', 'OverlapsAfter', 'Inside', 'InsideStart', 'InsideEnd']);
const BEFORE = new Set(['Before', 'AdjacentBefore']);
const OVERLAP = new Set([...AT]);

/** Serialises updates: Word events can arrive faster than a round-trip. */
async function update(force = false) {
  if (state.busy) { state.again = true; return; }
  state.busy = true;
  try { await doUpdate(force); } finally {
    state.busy = false;
    if (state.again) { state.again = false; update(); }
  }
}

async function doUpdate(force) {
  const pins = { left: ls.get(pinKey('left')), right: ls.get(pinKey('right')) };
  if (pins[SLOT]) return; // this pane is pinned: keep showing what it shows

  let found;
  try {
    found = await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      const cur = sel.paragraphs.getFirst();

      // walk a few paragraphs back and forward from the cursor
      let first = cur, last = cur;
      for (let i = 0; i < Math.max(PARAS_BEFORE, PARAS_AFTER); i++) {
        const prev = i < PARAS_BEFORE ? first.getPreviousOrNullObject() : null;
        const next = i < PARAS_AFTER ? last.getNextOrNullObject() : null;
        if (prev) prev.load('isNullObject');
        if (next) next.load('isNullObject');
        await ctx.sync();
        if (prev && !prev.isNullObject) first = prev;
        if (next && !next.isNullObject) last = next;
      }
      const win = first.getRange('Start').expandTo(last.getRange('End'));
      const links = win.getHyperlinkRanges();
      links.load('items/hyperlink,items/text');
      await ctx.sync();
      if (!links.items.length) return { none: true };

      const rels = links.items.map((l) => l.compareLocationWith(sel));
      const paras = links.items.map((l) => { const p = l.paragraphs.getFirst(); p.load('text'); return p; });
      await ctx.sync();

      const cands = links.items.map((l, i) => {
        const r = rels[i].value;
        return {
          key: citationKey(l.hyperlink || '', l.text, paras[i].text),
          rel: AT.has(r) ? 'at' : BEFORE.has(r) ? 'before' : 'after',
        };
      });
      const idx = assignSlots(cands, pins)[SLOT];
      if (idx < 0) return { none: true };
      const link = links.items[idx];
      const key = cands[idx].key;
      if (!force && state.current && state.current.key === key) return null; // unchanged

      let comment = '';
      if (state.commentsOk) {
        const comments = paras[idx].getRange('Whole').getComments();
        comments.load('items/content');
        await ctx.sync();
        if (comments.items.length) {
          const where = comments.items.map((c) => c.getRange('Whole').compareLocationWith(link));
          const replies = comments.items.map((c) => { const r = c.replies; r.load('items/content'); return r; });
          await ctx.sync();
          const j = where.findIndex((r) => OVERLAP.has(r.value));
          if (j >= 0) comment = [comments.items[j].content, ...replies[j].items.map((r) => r.content)].join('\n');
        }
      }
      ctx.trackedObjects.add(link);
      return { link, key, address: link.hyperlink || '', text: link.text, comment, paraText: paras[idx].text };
    });
  } catch (e) {
    status(`Word error: ${e.message}`);
    return;
  }
  if (!found) return;

  if (found.none) {
    await unmark();
    state.current = null;
    showEmpty();
    return;
  }
  await unmark();
  await mark(found.link);
  state.current = { key: found.key, address: found.address, text: found.text, comment: found.comment, paraText: found.paraText };
  showCitation(state.current);
}

/** Remove our colour from the citation we marked — only if it still has our colour
 *  (the other pane may have taken that citation over and recoloured it). */
async function unmark() {
  const range = state.marked;
  state.marked = null;
  if (!range) return;
  try {
    await Word.run(range, async (ctx) => {
      range.load('font/highlightColor');
      await ctx.sync();
      if (isSlotColor(range.font.highlightColor, SLOT)) range.font.highlightColor = null;
      ctx.trackedObjects.remove(range);
      await ctx.sync();
    });
  } catch { /* range deleted meanwhile */ }
}

async function mark(link) {
  if (!$('markToggle').checked) {
    try { await Word.run(link, async (ctx) => { ctx.trackedObjects.remove(link); await ctx.sync(); }); } catch { /* ignore */ }
    return;
  }
  try {
    await Word.run(link, async (ctx) => {
      link.font.highlightColor = SLOTS[SLOT].word;
      await ctx.sync();
    });
    state.marked = link;
  } catch { /* read-only document */ }
}

// ---------------- Source side ----------------

function showEmpty() {
  state.seq++;
  $('card').hidden = true;
  $('viewer').innerHTML = '';
  state.rendered = null;
  $('hint').hidden = false;
  status(SLOT === 'right' ? 'No second citation near the cursor' : 'No citation near the cursor');
}

async function showCitation(c) {
  const seq = ++state.seq;
  $('hint').hidden = true;
  $('card').hidden = false;
  $('citeText').textContent = c.text || '(citation)';
  $('citeAddr').textContent = c.address;
  $('commentText').textContent = c.comment || '(no comment on this citation)';
  $('missing').hidden = true;
  $('nosource').hidden = true;
  setCounts([]);
  $('viewer').innerHTML = '';
  state.rendered = null;
  status('Looking for the source…');

  let file;
  try { file = await resolveSource(c.address); } catch { file = null; }
  if (seq !== state.seq) return;
  if (!file) {
    status('Source not found');
    $('nosourceMsg').textContent = /^https?:/i.test(c.address)
      ? 'This link could not be downloaded directly (most publisher sites block it). Add the PDF once, or choose it for this citation.'
      : 'No matching PDF among your sources. Add the folder, or choose the PDF for this citation.';
    $('nosource').hidden = false;
    return;
  }

  status(`Opening ${file.name}…`);
  let pdf = state.pdfCache.get(file.name);
  if (!pdf) {
    try { pdf = await loadPdf(file.data); } catch (e) { status(`Cannot read ${file.name}: ${e.message}`); return; }
    state.pdfCache.set(file.name, pdf);
  }
  if (seq !== state.seq) return;

  const { highlights, missing } = findHighlights(pdf.index, c.comment, c.paraText);
  setCounts(highlights);
  if (missing.length) {
    $('missing').textContent = `Not found in the PDF: ${missing.map((m) => `“${m.slice(0, 60)}${m.length > 60 ? '…' : ''}”`).join(', ')}`;
    $('missing').hidden = false;
  } else if (pdf.index.sentences.length === 0) {
    $('missing').textContent = 'This PDF has no text layer (scanned?). Run OCR on it to enable highlighting.';
    $('missing').hidden = false;
  }

  state.rendered = { pdf, highlights };
  const first = await renderPdf($('viewer'), pdf, highlights);
  if (seq !== state.seq) return;
  if (first) first.scrollIntoView({ block: 'center' });
  hlCursor = 0;
  status(`${file.name} · ${pdf.pages.length} pages`);
}

async function resolveSource(address) {
  await refreshFiles(); // the other pane may have added PDFs
  const assigned = store.getAssignment(address);
  if (assigned) { const f = await store.getFile(assigned); if (f) return f; }

  let decoded = address;
  try { decoded = decodeURIComponent(address); } catch { /* keep raw */ }
  const base = decoded.split(/[?#]/)[0].split(/[\\/]/).pop().toLowerCase();
  const stem = base.replace(/\.pdf$/, '');
  const byName = state.files.find((f) => f.name.toLowerCase() === base)
    || (stem.length > 3 && state.files.find((f) => f.name.toLowerCase().replace(/\.pdf$/, '') === stem));
  if (byName) return store.getFile(byName.name);

  const dois = extractDois(decoded);
  if (dois.length) {
    const byDoi = state.files.find((f) => (f.dois || []).some((d) => dois.includes(d)));
    if (byDoi) return store.getFile(byDoi.name);
  }

  if (/^https?:/i.test(address)) {
    const res = await fetch(address, { mode: 'cors' });
    if (!res.ok) return null;
    const data = await res.arrayBuffer();
    const head = new TextDecoder().decode(new Uint8Array(data.slice(0, 5)));
    if (head !== '%PDF-') return null;
    const name = base.endsWith('.pdf') ? base : `${(dois[0] || base || 'source').replace(/[^\w.-]+/g, '_')}.pdf`;
    const file = { name, data, dois: await safeDois(data) };
    await store.putFile(file);
    store.setAssignment(address, name);
    filesChanged();
    return file;
  }
  return null;
}

// ---------------- UI ----------------

function bindUi() {
  $('addPdfs').onclick = () => $('pdfInput').click();
  $('addFolder').onclick = () => $('folderInput').click();
  $('pdfInput').onchange = (e) => addFiles(e.target.files).finally(() => { e.target.value = ''; });
  $('folderInput').onchange = (e) => addFiles(e.target.files).finally(() => { e.target.value = ''; });
  $('clearPdfs').onclick = async () => {
    await store.clearFiles(); state.pdfCache.clear(); filesChanged(); status('Sources cleared');
  };
  const pickForCitation = () => $('assignInput').click();
  $('assign').onclick = pickForCitation;
  $('reassign').onclick = pickForCitation;
  $('assignInput').onchange = async (e) => {
    const added = await addFiles(e.target.files);
    if (added[0] && state.current) {
      store.setAssignment(state.current.address, added[0]);
      showCitation(state.current);
    }
    e.target.value = '';
  };
  $('openLink').onclick = () => {
    const url = state.current && state.current.address;
    if (!url) return;
    if (Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(url);
    else window.open(url, '_blank');
  };
  $('pin').onclick = () => {
    const pinned = !!ls.get(pinKey(SLOT));
    ls.set(pinKey(SLOT), pinned ? null : (state.current ? state.current.key : null));
    syncPinUi();
    if (pinned) update(true);
  };
  $('nextHl').onclick = nextHighlight;
  $('markToggle').checked = ls.get('citeview.mark') !== 'off';
  $('markToggle').onchange = async () => {
    ls.set('citeview.mark', $('markToggle').checked ? 'on' : 'off');
    if (!$('markToggle').checked) await unmark();
    update(true);
  };
  document.querySelectorAll('.legend input').forEach((cb) => {
    cb.onchange = () => $('viewer').classList.toggle(`hide-${cb.dataset.level}`, !cb.checked);
  });
  window.addEventListener('resize', debounce(() => {
    if (state.rendered) renderPdf($('viewer'), state.rendered.pdf, state.rendered.highlights);
  }, 300));
}

function syncPinUi() {
  const pinned = !!ls.get(pinKey(SLOT));
  $('pin').classList.toggle('on', pinned);
  $('pin').title = pinned ? 'Pinned — this pane keeps its source. Click to follow the cursor again.' : 'Pin: keep this source while the other pane follows the cursor';
  $('pin').setAttribute('aria-pressed', String(pinned));
}

async function addFiles(fileList) {
  const pdfs = [...fileList].filter((f) => /\.pdf$/i.test(f.name));
  const names = [];
  for (const [n, f] of pdfs.entries()) {
    status(`Adding ${n + 1}/${pdfs.length}: ${f.name}`);
    const data = await f.arrayBuffer();
    await store.putFile({ name: f.name, data, dois: await safeDois(data) });
    state.pdfCache.delete(f.name);
    names.push(f.name);
  }
  filesChanged();
  status(pdfs.length ? `Added ${pdfs.length} PDF${pdfs.length > 1 ? 's' : ''}` : 'No PDFs in selection');
  if (state.current && !$('nosource').hidden) showCitation(state.current);
  return names;
}

function filesChanged() {
  refreshFiles();
  ls.set('citeview.files', String(Date.now())); // tells the other pane
}

async function safeDois(data) {
  try { return await pdfDois(data); } catch { return []; }
}

async function refreshFiles() {
  try { state.files = (await store.listFiles()).map(({ name, dois }) => ({ name, dois })); } catch { state.files = []; }
  $('sourceCount').textContent = `${state.files.length} source${state.files.length === 1 ? '' : 's'}`;
}

function setCounts(h) {
  $('nKey').textContent = h.filter((x) => x.level === 'key').length;
  $('nRel').textContent = h.filter((x) => x.level === 'rel').length;
  $('nSome').textContent = h.filter((x) => x.level === 'some').length;
}

let hlCursor = 0;
function nextHighlight() {
  const els = [...$('viewer').querySelectorAll('.hl')].filter((el) => el.offsetParent !== null);
  if (!els.length) return;
  const groups = els.filter((el, i) => i === 0 || Math.abs(el.getBoundingClientRect().top - els[i - 1].getBoundingClientRect().top) > 40);
  hlCursor = (hlCursor + 1) % groups.length;
  document.querySelectorAll('.hl.flash').forEach((el) => el.classList.remove('flash'));
  groups[hlCursor].classList.add('flash');
  groups[hlCursor].scrollIntoView({ block: 'center', behavior: 'smooth' });
}

function debounce(fn, ms) {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}
