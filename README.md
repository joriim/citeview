# CiteView — Word add-in

Put the cursor in a citation link in your draft. CiteView:

1. marks the citation in the draft (turquoise, removed again when you move on),
2. opens the cited PDF in the side pane,
3. highlights passages in three levels based on the comment attached to the citation:
   **purple** = key part, **green** = relevant, **yellow** = some relevance,
4. scrolls to the most important passage.

No server, no sign-in, no AI service: everything runs inside the Word task pane.

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
| `[key] "exact text"` or just `"exact text"` | purple |
| `[rel] "exact text"` | green |
| `[some] "exact text"` | yellow |
| any other free text | automatic green/yellow: sentences in the PDF that share the most words with your comment and the citing paragraph |

Quotes don't need to be letter-perfect: if the exact text isn't found, the closest 1–3 sentences are used (≥ 60 % word match). Quotes that can't be found are listed in the pane. Tags also accept `[k] [r] [s]`, `[purple] [green] [yellow]`, `[avain] [olennainen] [osittain]`.

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

Open it from **References → Show sources**.

### Host it instead of running it locally

Put the whole folder on any HTTPS host (GitHub Pages, SharePoint site, Azure Static Web Apps), then

```
npm run set-url -- https://your.host/citeview
```

and use the generated `manifest.hosted.xml`. An admin can deploy that manifest to the whole organisation from the Microsoft 365 admin center (Integrated apps).

## Requirements and limits

- Word 2019/2021, Microsoft 365 desktop, Mac or web. Reading comments needs WordApi 1.4 (current Microsoft 365); older Word still opens the PDF but cannot read the comment.
- PDFs only. Scanned PDFs need OCR first (no text layer = nothing to highlight).
- The automatic green/yellow matching compares words, so a **Finnish comment won't match an English source** — use quoted passages for those.
- If Track Changes is on, marking the citation is recorded as a formatting change: untick **Mark in draft**.
- Publisher landing pages (most `doi.org` links) can't be downloaded from inside Word — add the PDF once.

## Files

```
manifest.xml          add-in manifest (localhost:3000)
src/taskpane.html     the side pane
src/taskpane.js       Word side: cursor → citation → comment, marks the citation
src/matcher.js        finds passages for quotes and free text (pure JS, testable in Node)
src/pdfview.js        PDF.js rendering + highlight boxes
src/store.js          stored PDFs (IndexedDB) and citation→PDF choices
vendor/pdfjs/         PDF.js 4.10 (Apache-2.0)
test/                 sample draft + PDF, `npm test` runs the matcher on them
```
