// taskpane.js — Word glue: watch the cursor, find the citation link and its comment,
// mark the citation in the draft, and show the source PDF with highlights.
/* global Office, Word */
import { loadPdf, pdfDois, renderPdf } from './pdfview.js';
import { findHighlights, extractDois } from './matcher.js';
import * as store from './store.js';

const MARK_COLOR = '#9EE7E3'; // highlight for the active citation in the draft (distinct from the 3 levels)

const state = {
  current: null,     // {address, text, comment, paraText}
  prev: null,        // {range, color} — the citation we marked in the draft
  files: [],         // [{name, dois}]
  pdfCache: new Map(),
  seq: 0,
  rendered: null,    // {pdf, highlights} for re-render on resize
  commentsOk: false,
};

const $ = (id) => document.getElementById(id);
const status = (t) => { $('status').textContent = t; };

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) { status('Open this add-in in Word.'); return; }
  state.commentsOk = Office.context.requirements.isSetSupported('WordApi', '1.4');
  bindUi();
  await refreshFiles();
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, debounce(() => onSelection(), 250));
  status(state.commentsOk ? 'Ready' : 'Ready (this Word version cannot read comments)');
  onSelection();
});

// ---------------- Word side ----------------

const HIT = new Set(['Equal', 'Contains', 'ContainsStart', 'ContainsEnd', 'OverlapsBefore', 'OverlapsAfter']);
const OVERLAP = new Set([...HIT, 'Inside', 'InsideStart', 'InsideEnd']);

async function onSelection(force = false) {
  let found;
  try {
    found = await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      const para = sel.paragraphs.getFirst();
      const paraRange = para.getRange('Whole');
      const links = paraRange.getHyperlinkRanges();
      links.load('items/hyperlink,items/text');
      para.load('text');
      await ctx.sync();
      if (!links.items.length) return null;

      const rels = links.items.map((l) => l.compareLocationWith(sel));
      await ctx.sync();
      const i = rels.findIndex((r) => HIT.has(r.value));
      if (i < 0) return null;
      const link = links.items[i];
      const address = link.hyperlink || '';
      if (!force && state.current && state.current.address === address && state.current.paraText === para.text) return null;

      let comment = '';
      if (state.commentsOk) {
        const comments = paraRange.getComments();
        comments.load('items/content');
        await ctx.sync();
        if (comments.items.length) {
          const where = comments.items.map((c) => c.getRange('Whole').compareLocationWith(link));
          const replies = comments.items.map((c) => { const r = c.replies; r.load('items/content'); return r; });
          await ctx.sync();
          let j = where.findIndex((r) => OVERLAP.has(r.value));
          if (j < 0) j = where.findIndex((r) => r.value === 'AdjacentAfter' || r.value === 'After');
          if (j < 0) j = 0;
          comment = [comments.items[j].content, ...replies[j].items.map((r) => r.content)].join('\n');
        }
      }

      link.load('font/highlightColor');
      ctx.trackedObjects.add(link);
      await ctx.sync();
      return { link, color: link.font.highlightColor, address, text: link.text, comment, paraText: para.text };
    });
  } catch (e) {
    status(`Word error: ${e.message}`);
    return;
  }
  if (!found) return;

  await markInDocument(found.link, found.color);
  state.current = { address: found.address, text: found.text, comment: found.comment, paraText: found.paraText };
  showCitation(state.current);
}

/** Restore the previously marked citation and mark the new one. */
async function markInDocument(link, color) {
  const enabled = $('markToggle').checked;
  if (state.prev) {
    const { range, color: old } = state.prev;
    try {
      await Word.run(range, async (ctx) => {
        range.font.highlightColor = old || null;
        ctx.trackedObjects.remove(range);
        await ctx.sync();
      });
    } catch { /* range was deleted meanwhile */ }
  }
  state.prev = null;
  if (!enabled) return;
  state.prev = { range: link, color };
  try {
    await Word.run(link, async (ctx) => {
      link.font.highlightColor = MARK_COLOR;
      await ctx.sync();
    });
  } catch { /* read-only document: skip marking */ }
}

// ---------------- Source side ----------------

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
  status('Looking for the source…');

  let file;
  try { file = await resolveSource(c.address); } catch (e) { file = null; }
  if (seq !== state.seq) return;
  if (!file) {
    status('Source not found');
    $('nosourceMsg').textContent = /^https?:/i.test(c.address)
      ? 'This link could not be downloaded directly (most publisher sites block it). Download the PDF once and add it, or choose it for this citation.'
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
  }
  if (!highlights.length && pdf.index.sentences.length === 0) {
    $('missing').textContent = 'This PDF has no text layer (scanned?). Run OCR on it to enable highlighting.';
    $('missing').hidden = false;
  }

  state.rendered = { pdf, highlights };
  const first = await renderPdf($('viewer'), pdf, highlights);
  if (seq !== state.seq) return;
  if (first) first.scrollIntoView({ block: 'center' });
  status(`${file.name} · ${pdf.pages.length} pages`);
}

async function resolveSource(address) {
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
    const name = (base.endsWith('.pdf') ? base : `${(dois[0] || base || 'source').replace(/[^\w.-]+/g, '_')}.pdf`);
    const file = { name, data, dois: await safeDois(data) };
    await store.putFile(file);
    store.setAssignment(address, name);
    await refreshFiles();
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
    await store.clearFiles(); state.pdfCache.clear(); await refreshFiles(); status('Sources cleared');
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
  $('refresh').onclick = () => onSelection(true);
  try { $('markToggle').checked = localStorage.getItem('citeview.mark') !== 'off'; } catch { /* default on */ }
  $('markToggle').onchange = () => {
    try { localStorage.setItem('citeview.mark', $('markToggle').checked ? 'on' : 'off'); } catch { /* ignore */ }
    if (!$('markToggle').checked && state.prev) markInDocument(state.prev.range, null).catch(() => {});
  };
  $('nextHl').onclick = nextHighlight;
  document.querySelectorAll('.legend input').forEach((cb) => {
    cb.onchange = () => $('viewer').classList.toggle(`hide-${cb.dataset.level}`, !cb.checked);
  });
  window.addEventListener('resize', debounce(() => {
    if (state.rendered) renderPdf($('viewer'), state.rendered.pdf, state.rendered.highlights);
  }, 300));
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
  await refreshFiles();
  status(pdfs.length ? `Added ${pdfs.length} PDF${pdfs.length > 1 ? 's' : ''}` : 'No PDFs in selection');
  if (state.current && !$('nosource').hidden) showCitation(state.current); // retry a missing source
  return names;
}

async function safeDois(data) {
  try { return await pdfDois(data); } catch { return []; }
}

async function refreshFiles() {
  try {
    state.files = (await store.listFiles()).map(({ name, dois }) => ({ name, dois }));
  } catch { state.files = []; }
  $('sourceCount').textContent = `${state.files.length} source${state.files.length === 1 ? '' : 's'}`;
}

function setCounts(h) {
  $('nKey').textContent = h.filter((x) => x.level === 'key').length;
  $('nRel').textContent = h.filter((x) => x.level === 'rel').length;
  $('nSome').textContent = h.filter((x) => x.level === 'some').length;
}

let hlCursor = -1;
function nextHighlight() {
  const els = [...$('viewer').querySelectorAll('.hl')].filter((el) => el.offsetParent !== null);
  if (!els.length) return;
  // jump between highlight groups, not every line fragment
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
