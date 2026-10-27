// slots.js — decides which citation each source pane shows. Pure logic (testable in Node).
//
// Word has no API for "what is visible on screen", so the visible area is approximated
// by the citations around the cursor (a few paragraphs before and after it).
//
// Rule: the citation at the cursor — or the next one after it, since a citation usually
// ends the sentence being read — goes to the LEFT pane; the citation after that goes to
// the RIGHT pane. Reading forward, citations move from right to left.

export const SLOTS = {
  left: { name: 'Left', word: '#00FFFF', rgb: '0, 170, 190', ink: '0, 115, 130' },   // Word "Turquoise"
  right: { name: 'Right', word: '#00FF00', rgb: '40, 175, 60', ink: '20, 115, 35' },  // Word "BrightGreen"
};

const WORD_NAMES = { turquoise: '#00FFFF', brightgreen: '#00FF00' };

/** True if a Word highlightColor value (hex or name) is the given slot's mark colour. */
export function isSlotColor(value, slot) {
  if (!value) return false;
  const v = String(value).trim().toLowerCase();
  const hex = WORD_NAMES[v] || v;
  return hex.toUpperCase() === SLOTS[slot].word;
}

/**
 * cands: citations in document order, each { key, rel } where rel is
 *        'at' (cursor inside), 'before' (citation before cursor) or 'after'.
 * pins:  { left: key|null, right: key|null } — a pinned pane keeps its citation.
 * Returns { left: index|-1, right: index|-1 } into cands.
 */
export function assignSlots(cands, pins = {}) {
  const out = { left: -1, right: -1 };
  if (!cands.length) return out;

  let primary = cands.findIndex((c) => c.rel === 'at');
  if (primary < 0) primary = cands.findIndex((c) => c.rel === 'after');
  if (primary < 0) primary = cands.length - 1;

  // order of preference: primary, then following citations, then preceding ones (nearest first)
  const order = [primary];
  for (let i = primary + 1; i < cands.length; i++) order.push(i);
  for (let i = primary - 1; i >= 0; i--) order.push(i);

  const pinnedIdx = (slot) => (pins[slot] ? cands.findIndex((c) => c.key === pins[slot]) : -1);
  const taken = new Set();
  for (const slot of ['left', 'right']) {
    if (pins[slot]) { const i = pinnedIdx(slot); if (i >= 0) taken.add(i); }
  }
  for (const slot of ['left', 'right']) {
    if (pins[slot]) { out[slot] = pinnedIdx(slot); continue; }
    const next = order.find((i) => !taken.has(i));
    if (next !== undefined) { out[slot] = next; taken.add(next); }
  }
  return out;
}

/** Stable-ish identity for a citation: link target + link text + its paragraph. */
export function citationKey(address, text, paraText) {
  let h = 0;
  const s = `${address}\u0000${text}\u0000${paraText}`;
  for (let i = 0; i < s.length; i++) h = (Math.imul(31, h) + s.charCodeAt(i)) | 0;
  return `${(h >>> 0).toString(36)}:${address}`;
}
