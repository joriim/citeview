# CiteView — Word add-in

A three-column reading view for checking citations:

```
 Left source pane      |   your draft    |   Right source pane
 (turquoise)           |                 |   (green)
 source of the         |  … soils (A).   |   source of the
 citation at the       |  … north (B).   |   next citation
 cursor  = A           |                 |   = B
```

- **Colour = which citation.** The left pane's citation is marked turquoise in the draft and its source passages are turquoise; the right pane's are green.
- **Strength = how relevant.** Strong + underline = key part, medium = relevant, light = some relevance.
- The panes follow the cursor: left = the citation at or just after the cursor, right = the next one. Reading forward, sources move from right to left. 📌 pins a pane.
- Word has no way to tell an add-in which text is visible on screen, so "visible" is approximated by the paragraphs around the cursor (2 before, 4 after).

No server, no sign-in, no AI service: everything runs inside Word's side panes.

**Setup:** open References → Left source and References → Right source, then drag the Left pane by its title bar to the left edge of the Word window.

## How to write the comment

Attach a Word comment to the citation (select the link → Review → New Comment).

```
Main evidence for the leaching claim:
[key] "Autumn application of cattle slurry increased nitrate leaching by 38 percent"
[rel] "On organic soils the leaching losses were nearly twice those of clay soils"
They also recommend timing manure by soil frost instead of calendar dates.
```

| In the comment | Becomes |
|---|---|
| `[key] "exact text"` or just `"exact text"` | strong + underline (key) |
| `[rel] "exact text"` | medium (relevant) |
| `[some] "exact text"` | light (some relevance) |
| any other free text | automatic medium/light: sentences in the PDF that share the most words with your comment and the citing paragraph |

Quotes don't need to be letter-perfect: if the exact text isn't found, the closest 1–3 sentences are used (≥ 60 % word match). Quotes that can't be found are listed in the pane. Tags also accept `[k] [r] [s]`, `[avain] [olennainen] [osittain]`.

## Where sources come from

Click **Add folder** (or **Add PDFs**) once and pick the folder with your source PDFs — e.g. the synced OneDrive/SharePoint folder next to the draft. They are stored in the add-in, so this survives restarts. A citation link is matched to a PDF by, in order:

1. a PDF you chose for that citation earlier (**Use a different PDF**),
2. file name — a link to `sources/Virtanen2024.pdf` (relative, `file://` or a SharePoint URL) matches `Virtanen2024.pdf`,
3. DOI — a link to `https://doi.org/10.xxxx/…` matches a PDF whose first two pages print that DOI,
4. direct download — works only for sites that allow it (e.g. some open repositories); downloaded PDFs are kept.

## Install (hosted)

The add-in is published at **https://joriim.github.io/citeview** — the page there has a one-click installer for Windows and steps for Mac and Word for the web. Nothing needs to run on your computer.

## Install (local development)

You need [Node.js](https://nodejs.org) 18 or newer.

```
npm install
npm run certs      # once: trusted localhost certificate for Office
npm start          # serves the add-in at https://localhost:3000 — keep it running
```

Then load `manifest.xml` into Word:

- **Word on Windows:** `npm run sideload` (opens Word with the add-in), or put `manifest.xml` in a shared network folder and add that folder under File → Options → Trust Center → Trusted Add-in Catalogs, restart Word, then Insert → My Add-ins → Shared Folder.
- **Word on Mac:** copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` and restart Word.
- **Word for the web:** Insert → Add-ins → Upload My Add-in → choose `manifest.xml`.

Open the panes from **References → Left source / Right source**.

### Host it instead of running it locally

Put the whole folder on any HTTPS host (GitHub Pages, SharePoint site, Azure Static Web Apps), then

```
npm run set-url -- https://your.host/citeview
```

and use the generated `manifest.hosted.xml`. An admin can deploy that manifest to the whole organisation from the Microsoft 365 admin center (Integrated apps).

## Requirements and limits

- Word 2019/2021, Microsoft 365 desktop, Mac or web. Reading comments needs WordApi 1.4 (current Microsoft 365); older Word still opens the PDF but cannot read the comment.
- PDFs only. Scanned PDFs need OCR first (no text layer = nothing to highlight).
- The automatic matching compares words, so a **Finnish comment won't match an English source** — use quoted passages for those.
- If Track Changes is on, marking the citation is recorded as a formatting change: untick **Mark in draft**.
- Publisher landing pages (most `doi.org` links) can't be downloaded from inside Word — add the PDF once.

## Files

```
manifest.xml          add-in manifest (localhost:3000)
src/taskpane.html     the side pane
src/taskpane.js       one pane (?slot=left|right): citations near cursor → its citation, comment, mark
src/slots.js          which citation goes to which pane; pane colours
src/matcher.js        finds passages for quotes and free text (pure JS, testable in Node)
src/pdfview.js        PDF.js rendering + highlight boxes
src/store.js          stored PDFs (IndexedDB) and citation→PDF choices
vendor/pdfjs/         PDF.js 4.10 (Apache-2.0)
test/                 sample draft + PDF, `npm test` runs the matcher on them
```
