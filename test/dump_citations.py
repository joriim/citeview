# Extract (link address, link text, comment, paragraph text) from a draft .docx — mirrors what the add-in reads via Office.js.
import json, sys
from docx import Document
d = Document(sys.argv[1])
comments = {c.comment_id: c.text for c in d.comments}
out = []
for p in d.paragraphs:
    for h in p.hyperlinks:
        ids = [int(e.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id'))
               for e in p._p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart')]
        out.append({'address': h.address, 'text': h.text, 'comment': '\n'.join(comments[i] for i in ids), 'paraText': p.text})
json.dump(out, sys.stdout, indent=1)
