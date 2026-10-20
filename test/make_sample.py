# Builds a sample source PDF and a draft .docx (hyperlinked citations + comments) for testing CiteView.
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import textwrap

pdf = canvas.Canvas('sample/sources/Virtanen2024.pdf', pagesize=A4)
W, H = A4
paras = [
 ("Nitrogen losses from boreal arable soils after autumn manure application", 'Helvetica-Bold', 13),
 ("A. Virtanen, B. Korhonen. Soil Science Example 12 (2024). https://doi.org/10.5555/ssex.2024.0123", 'Helvetica', 8),
 ("Abstract. We measured ammonia volatilisation, nitrous oxide emissions and nitrate leaching on twelve clay and "
  "organic soil plots in western Finland over two years. Autumn application of cattle slurry increased nitrate "
  "leaching by 38 percent compared with spring application. Nitrous oxide emissions peaked during freeze-thaw "
  "cycles in early spring. Injection reduced ammonia losses by more than half compared with broadcast spreading.", 'Helvetica', 10),
 ("1. Introduction. Agricultural nitrogen surplus remains a major source of nutrient loading to the Baltic Sea. "
  "Manure timing is regulated, but the evidence base for boreal conditions is thin. Previous studies were mostly "
  "conducted in temperate climates where soils rarely freeze. Snow cover and soil frost change the timing of "
  "microbial nitrogen transformations.", 'Helvetica', 10),
 ("2. Methods. Each plot was equipped with suction cups at 30 and 60 cm depth and with static chambers for gas "
  "sampling. Ammonia volatilisation was measured with wind tunnels during the first 72 hours after application. "
  "Soil mineral nitrogen was sampled monthly. Weather data came from an on-site station.", 'Helvetica', 10),
]
paras2 = [
 ("3. Results. Nitrate concentrations in drainage water were highest in November and in April. On organic soils "
  "the leaching losses were nearly twice those of clay soils. Freeze-thaw cycles explained most of the annual "
  "nitrous oxide emission, while summer fluxes were small. The plots with slurry injection had the lowest "
  "ammonia losses in all weather conditions.", 'Helvetica', 10),
 ("4. Discussion. Our results suggest that autumn application of manure should be avoided on organic soils "
  "in boreal regions. Delaying application to spring reduces leaching without lowering yields. Monitoring of soil "
  "frost could be used to time applications more precisely. Remote sensing of snow cover is a promising tool.", 'Helvetica', 10),
 ("5. Conclusions. Spring application combined with injection minimised total nitrogen losses. Policy should "
  "account for soil type and frost conditions rather than calendar dates only.", 'Helvetica', 10),
]

def draw(paras, page_no):
    y = H - 60
    for text, font, size in paras:
        pdf.setFont(font, size)
        width = 92 if size <= 10 else 70
        if size == 8: width = 120
        lines = textwrap.wrap(text, width)
        # force a hyphenated line break inside 'volatilisation' on page 1 abstract
        out = []
        for ln in lines:
            if 'ammonia volatilisation,' in ln and page_no == 1:
                a, b = ln.split('volatilisation,', 1)
                out.append(a + 'volatil-'); out.append('isation,' + b)
            else:
                out.append(ln)
        for ln in out:
            pdf.drawString(60, y, ln); y -= size * 1.45
        y -= 10

draw(paras, 1); pdf.showPage(); draw(paras2, 2); pdf.save()

# ---- draft docx with hyperlinked citations and comments ----
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
doc.add_heading('Draft: Manure timing on boreal soils', 1)

def add_link(par, url, text):
    r_id = par.part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    h = OxmlElement('w:hyperlink'); h.set(qn('r:id'), r_id)
    r = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    st = OxmlElement('w:rStyle'); st.set(qn('w:val'), 'Hyperlink'); rPr.append(st)
    c = OxmlElement('w:color'); c.set(qn('w:val'), '0563C1'); rPr.append(c)
    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    r.append(rPr); t = OxmlElement('w:t'); t.text = text; r.append(t); h.append(r)
    par._p.append(h)
    return h

p = doc.add_paragraph('Autumn manure application increases nitrate leaching on boreal soils, especially on organic soils ')
add_link(p, 'sources/Virtanen2024.pdf', '(Virtanen et al. 2024)')
p.add_run('.')
doc.add_paragraph('Freeze-thaw cycles dominate annual nitrous oxide emissions in the north ')
p2 = doc.paragraphs[-1]
add_link(p2, 'https://doi.org/10.5555/ssex.2024.0123', '(Virtanen et al. 2024)')
p2.add_run('.')
doc.save('sample/draft_tmp.docx')
