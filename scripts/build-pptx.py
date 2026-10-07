import re, copy, os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR as TC
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, '..', 'assets', 'shift-projects')
OUT = os.path.join(HERE, '..', 'templates', 'Shift-Projects-Template.pptx')

# Regenerates templates/Shift-Projects-Template.pptx (native shapes, Shift theme, 5 layouts, 19 slides).
# Requires: pip install python-pptx

BLACK, CREAM, LUMBER, OLIVE, BRICK, SLATE = TC.TEXT_1, TC.BACKGROUND_1, TC.ACCENT_1, TC.ACCENT_2, TC.ACCENT_3, TC.ACCENT_4
HEAD = '+mj-lt'   # theme heading font
R = 13.3333 / 100  # 1rem in inches (100rem = slide width)

def r(n): return Inches(n * R)

prs = Presentation()
prs.slide_width = Inches(13.3333)
prs.slide_height = Inches(7.5)

# ---------------- theme: colours + fonts ----------------
theme_part = prs.slide_master.part.part_related_by(RT.THEME)
t = etree.fromstring(theme_part.blob)
ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
t.set('name', 'Shift Projects')
cs = t.find('.//a:clrScheme', ns); cs.set('name', 'Shift Projects')
palette = dict(dk1='282829', lt1='F7EDDB', dk2='282829', lt2='F7EDDB',
               accent1='E0B977', accent2='978745', accent3='C35D27', accent4='8BA5BC',
               accent5='282829', accent6='F7EDDB', hlink='C35D27', folHlink='978745')
for k, v in palette.items():
    el = cs.find('a:' + k, ns)
    for c in list(el): el.remove(c)
    etree.SubElement(el, '{%s}srgbClr' % ns['a']).set('val', v)
fs = t.find('.//a:fontScheme', ns); fs.set('name', 'Shift Projects')
fs.find('a:majorFont/a:latin', ns).set('typeface', 'Familjen Grotesk')
fs.find('a:minorFont/a:latin', ns).set('typeface', 'Inter')
theme_part._blob = etree.tostring(t, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---------------- helpers ----------------
def set_fill(shape, color, brightness=None):
    shape.fill.solid(); shape.fill.fore_color.theme_color = color
    if brightness is not None: shape.fill.fore_color.brightness = brightness

def rect(slide, x, y, w, h, color, name=None, shape=MSO_SHAPE.RECTANGLE, line=None, brightness=None):
    s = slide.shapes.add_shape(shape, r(x), r(y), r(w), r(h))
    set_fill(s, color, brightness)
    if line is None: s.line.fill.background()
    else:
        s.line.color.theme_color = line; s.line.width = Pt(1)
    s.shadow.inherit = False
    if name: s.name = name
    return s

def rule(slide, x, y, w, color=BLACK, weight=1.25, name='Rule'):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, r(x), r(y), r(x + w), r(y))
    c.line.color.theme_color = color; c.line.width = Pt(weight); c.name = name
    return c

def line(slide, x1, y1, x2, y2, color=BLACK, weight=1.25, name='Line'):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, r(x1), r(y1), r(x2), r(y2))
    c.line.color.theme_color = color; c.line.width = Pt(weight); c.name = name
    return c

def text(slide, x, y, w, h, content, size=14, bold=False, color=BLACK, display=False, align='l',
         caps=False, spc=None, ls=None, anchor='t', muted=False, name=None, after=0):
    """content: str or list of (str|list of runs) paragraphs. '\n' makes new paragraphs."""
    tb = slide.shapes.add_textbox(r(x), r(y), r(w), r(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {'t': MSO_ANCHOR.TOP, 'm': MSO_ANCHOR.MIDDLE, 'b': MSO_ANCHOR.BOTTOM}[anchor]
    paras = content.split('\n') if isinstance(content, str) else content
    for i, ptxt in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {'l': PP_ALIGN.LEFT, 'r': PP_ALIGN.RIGHT, 'c': PP_ALIGN.CENTER}[align]
        if ls: p.line_spacing = ls
        if after: p.space_after = Pt(after)
        run = p.add_run(); run.text = ptxt.upper() if caps else ptxt
        f = run.font; f.size = Pt(size); f.bold = bold
        if display: f.name = HEAD
        f.color.theme_color = color
        if muted: f.color.brightness = 0.4 if color == BLACK else -0.4
        if spc is not None: run._r.get_or_add_rPr().set('spc', str(int(spc * 100)))
    return tb

def parse_path(d):
    toks = re.findall(r'[MLCZ]|-?\d+\.?\d*', d)
    out, i = [], 0
    while i < len(toks):
        c = toks[i]; i += 1
        n = {'M': 2, 'L': 2, 'C': 6, 'Z': 0}[c]
        out.append((c, [float(v) for v in toks[i:i + n]])); i += n
    return out

def freeform(slide, paths, vb, x, y, w, color, name, alpha=None):
    vx, vy, vw, vh = vb
    wi = w * R; hi = wi * vh / vw
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x * R), Inches(y * R), Inches(wi), Inches(hi))
    set_fill(s, color); s.line.fill.background(); s.shadow.inherit = False; s.name = name
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    spPr = s._element.spPr
    spPr.remove(spPr.find(qn('a:prstGeom')))
    cg = etree.Element('{%s}custGeom' % A)
    for tag in ('avLst', 'gdLst', 'ahLst', 'cxnLst'): etree.SubElement(cg, '{%s}%s' % (A, tag))
    rc = etree.SubElement(cg, '{%s}rect' % A); rc.set('l', '0'); rc.set('t', '0'); rc.set('r', 'r'); rc.set('b', 'b')
    pl = etree.SubElement(cg, '{%s}pathLst' % A)
    pw, ph = int(vw * 100), int(vh * 100)
    path = etree.SubElement(pl, '{%s}path' % A); path.set('w', str(pw)); path.set('h', str(ph))
    def pt(parent, px, py):
        e = etree.SubElement(parent, '{%s}pt' % A)
        e.set('x', str(int(round((px - vx) * 100)))); e.set('y', str(int(round((py - vy) * 100))))
    for d in paths:
        for c, v in parse_path(d):
            if c == 'M': pt(etree.SubElement(path, '{%s}moveTo' % A), v[0], v[1])
            elif c == 'L': pt(etree.SubElement(path, '{%s}lnTo' % A), v[0], v[1])
            elif c == 'C':
                cb = etree.SubElement(path, '{%s}cubicBezTo' % A)
                for k in range(3): pt(cb, v[2 * k], v[2 * k + 1])
            else: etree.SubElement(path, '{%s}close' % A)
    xfrm = spPr.find(qn('a:xfrm')); xfrm.addnext(cg)
    return s

def polygon(slide, pts, x, y, w, h, color, name, brightness=None, alpha=None):
    """pts in fractions (0..1) of w,h"""
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, r(x), r(y), r(w), r(h))
    set_fill(s, color, brightness); s.line.fill.background(); s.shadow.inherit = False; s.name = name
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    spPr = s._element.spPr
    spPr.remove(spPr.find(qn('a:prstGeom')))
    cg = etree.Element('{%s}custGeom' % A)
    for tag in ('avLst', 'gdLst', 'ahLst', 'cxnLst'): etree.SubElement(cg, '{%s}%s' % (A, tag))
    rc = etree.SubElement(cg, '{%s}rect' % A); rc.set('l', '0'); rc.set('t', '0'); rc.set('r', 'r'); rc.set('b', 'b')
    pl = etree.SubElement(cg, '{%s}pathLst' % A)
    path = etree.SubElement(pl, '{%s}path' % A); path.set('w', '100000'); path.set('h', '100000')
    for i, (px, py) in enumerate(pts):
        e = etree.SubElement(path, '{%s}%s' % (A, 'moveTo' if i == 0 else 'lnTo'))
        p = etree.SubElement(e, '{%s}pt' % A); p.set('x', str(int(px * 100000))); p.set('y', str(int(py * 100000)))
    etree.SubElement(path, '{%s}close' % A)
    spPr.find(qn('a:xfrm')).addnext(cg)
    if alpha is not None:
        clr = s._element.spPr.find(qn('a:solidFill'))[0]
        a = etree.SubElement(clr, '{%s}alpha' % A); a.set('val', str(int(alpha * 1000)))
    return s

def svg_paths(name):
    s = open(f'{ASSETS}/{name}.svg').read()
    vb = [float(v) for v in re.search(r'viewBox="([^"]+)"', s).group(1).split()]
    return re.findall(r' d="([^"]+)"', s), vb

ICON = svg_paths('icon'); STACKED = svg_paths('wordmark-stacked'); LINEAR = svg_paths('linear-large')

def icon(slide, x, y, w, color=BLACK): return freeform(slide, ICON[0], ICON[1], x, y, w, color, 'Shift icon')
def stacked(slide, x, y, w, color=BLACK): return freeform(slide, STACKED[0], STACKED[1], x, y, w, color, 'Shift wordmark (stacked)')
def linear(slide, x, y, w, color=BLACK): return freeform(slide, LINEAR[0], LINEAR[1], x, y, w, color, 'Linear graphic')

# ---------------- layouts ----------------
LAYOUTS = {}
tmp = prs.slides.add_slide(prs.slide_layouts[0])
def to_layout(shape, lay):
    lay.shapes._spTree.append(shape._element)
spec = [('Shift - Cream', CREAM, BLACK), ('Shift - Lumber', LUMBER, BLACK), ('Shift - Slate', SLATE, BLACK),
        ('Shift - Olive', OLIVE, BLACK), ('Shift - Washed Black', BLACK, CREAM)]
for i, (nm, bg, fg) in enumerate(spec):
    lay = prs.slide_layouts[i]
    for sh in list(lay.placeholders): sh._element.getparent().remove(sh._element)
    lay._element.cSld.set('name', nm)
    lay.background.fill.solid(); lay.background.fill.fore_color.theme_color = bg
    to_layout(text(tmp, 4.15, 51.6, 18, 2, 'Shift Projects', size=10, bold=True, color=fg, caps=True, spc=1.6, name='Footer company'), lay)
    to_layout(text(tmp, 23.05, 51.6, 40, 2, 'Deck title', size=10, bold=True, color=fg, caps=True, spc=1.6, name='Footer deck title'), lay)
    pg = text(tmp, 80, 51.6, 15.85, 2, '', size=10, bold=True, color=fg, align='r', spc=1.6, name='Slide number')
    to_layout(pg, lay)
    p = pg.text_frame.paragraphs[0]
    for rr in list(p._p.findall(qn('a:r'))): p._p.remove(rr)
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    fld = etree.SubElement(p._p, '{%s}fld' % A); fld.set('id', '{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}'); fld.set('type', 'slidenum')
    rp = etree.SubElement(fld, '{%s}rPr' % A); rp.set('lang', 'en-US'); rp.set('sz', '1000'); rp.set('b', '1'); rp.set('spc', '160')
    sf = etree.SubElement(rp, '{%s}solidFill' % A); sc = etree.SubElement(sf, '{%s}schemeClr' % A); sc.set('val', 'bg1' if fg == CREAM else 'tx1')
    tt = etree.SubElement(fld, '{%s}t' % A); tt.text = '‹#›'
    # move field before endParaRPr if present
    end = p._p.find(qn('a:endParaRPr'))
    if end is not None: p._p.remove(end); p._p.append(end)
    LAYOUTS[nm] = lay
_id = prs.slides._sldIdLst[0]; prs.part.drop_rel(_id.rId); prs.slides._sldIdLst.remove(_id)
for lay in list(prs.slide_layouts)[5:]: prs.slide_layouts.remove(lay)

CREAM_L, LUMBER_L, SLATE_L, OLIVE_L, BLACK_L = [LAYOUTS[s[0]] for s in spec]

def new(layout, notes):
    s = prs.slides.add_slide(layout)
    s.notes_slide.notes_text_frame.text = notes
    return s

def eyebrow(s, sec='Section', sub='Subsection', color=BLACK):
    text(s, 4.15, 4.15, 15, 2, sec, size=10, bold=True, color=color, caps=True, spc=1.6, name='Eyebrow section')
    rule(s, 19.85, 4.6, 7.7, color, 1, 'Eyebrow rule')
    text(s, 35.5, 4.15, 40, 2, sub, size=10, bold=True, color=color, caps=True, spc=1.6, name='Eyebrow subsection')

def H1(s, txt, x=4.15, y=11.5, w=44, size=48, color=BLACK):
    return text(s, x, y, w, 20, txt, size=size, display=True, color=color, ls=0.92, spc=-0.5, name='Heading')

# ================= 1 COVER =================
s = new(LUMBER_L, 'COVER. Replace the title. The stepped graphic is a native shape (Format > Fill to recolour).')
text(s, 4.15, 3.6, 80, 32, 'Presentation\nTitle', size=110, display=True, ls=0.85, spc=-2, name='Title')
linear(s, 54, 26.5, 46)

# ================= 2 DIVIDER =================
s = new(LUMBER_L, 'SECTION DIVIDER. Duplicate for each section; change the number and name.')
text(s, 4.15, 3.6, 70, 32, 'Section\nName', size=110, display=True, ls=0.85, spc=-2, name='Title')
text(s, 70, 3.6, 25.85, 14, '01', size=110, display=True, align='r', spc=-2, name='Section number')

# ================= 3 AGENDA =================
s = new(CREAM_L, 'AGENDA. Three numbered sections with short lists.')
eyebrow(s, 'Introduction', '')
H1(s, 'Welcome to\nthe Shift Projects\npresentation.', y=11.5, w=44)
text(s, 4.15, 33, 28, 8, 'A short line that tells the audience what this deck covers and why it matters to them.', size=14, name='Intro')
for i, (n, tt, items) in enumerate([('01', 'About', ['Story', 'Goals']), ('02', 'Approach', ['Process', 'Timeline']), ('03', 'Results', ['Highlights', 'Next steps'])]):
    x = 51 + i * 16
    text(s, x, 12.2, 14.5, 3, n, size=18, bold=True, name=f'Agenda {n} number')
    text(s, x, 15.2, 14.5, 4, tt, size=18, bold=True, caps=True, name=f'Agenda {n} title')
    text(s, x, 19.2, 14.5, 8, items, size=14, after=3, name=f'Agenda {n} items')

# ================= 4 SPLIT =================
s = new(CREAM_L, 'SPLIT. Big headline left, short body copy right.')
eyebrow(s, 'Section', '1.1 Subsection')
H1(s, 'Built for today,\nready for tomorrow', w=52)
text(s, 62, 12.1, 34, 14, 'Supporting copy goes here. Use clear, direct language and keep it to a short paragraph.', size=14, name='Body')

# ================= 5 NUMBERED =================
s = new(CREAM_L, 'NUMBERED COLUMNS. Four goals, steps or principles.')
eyebrow(s, 'Section', '1.2 Goals')
H1(s, 'Our goals', w=30)
for i, tx in enumerate(['First goal, one or two lines.', 'Second goal, one or two lines.', 'Third goal, one or two lines.', 'Fourth goal, one or two lines.']):
    x = 35.5 + i * 15.6
    text(s, x, 22, 14, 3, f'{i+1}.', size=18, bold=True, name=f'Goal {i+1} number')
    text(s, x, 25.4, 14, 10, tx, size=14, name=f'Goal {i+1} text')

# ================= 6 STATEMENT =================
s = new(SLATE_L, 'STATEMENT. One big line plus a small supporting paragraph.')
eyebrow(s, 'Section', '2.6 Tagline')
text(s, 4.4, 12, 76, 40, 'Homes built for today, ready for tomorrow.', size=108, display=True, ls=0.88, spc=-2, name='Statement')
text(s, 80, 27.5, 15.8, 14, 'A short supporting paragraph, four or five lines at most.', size=14, name='Supporting text')

# ================= 7 REVERSED =================
s = new(BLACK_L, 'REVERSED. A key figure or claim on Washed Black with brick bands.')
eyebrow(s, 'Section', 'Key figure', color=CREAM)
rect(s, 0, 22, 100, 6.4, BRICK, 'Band 1'); rect(s, 0, 31.5, 100, 6.4, BRICK, 'Band 2')
text(s, 4.15, 5.8, 80, 16, '100%', size=120, display=True, color=CREAM, spc=-2, name='Key figure')
text(s, 4.15, 41.5, 44, 8, 'What the number means, in one plain sentence.', size=22, color=CREAM, name='Caption')

# ================= 8 STATS =================
s = new(CREAM_L, 'STATS. Three big numbers with a label and one line of context.')
eyebrow(s, 'Section', 'Results')
H1(s, 'The numbers,\nplainly', w=40)
for i, (n, lab, ctx) in enumerate([('24', 'Projects delivered', 'One line of context for the figure.'),
                                   ('98%', 'Client satisfaction', 'One line of context for the figure.'),
                                   ('12', 'Weeks, average build', 'One line of context for the figure.')]):
    x = 4.15 + i * 31.6
    rule(s, x, 27, 29.5, BLACK, 1.5, f'Stat {i+1} rule')
    text(s, x, 28.4, 29.5, 12, n, size=72, display=True, spc=-1, name=f'Stat {i+1} number')
    text(s, x, 40.8, 29.5, 2, lab, size=11, bold=True, caps=True, spc=1.6, name=f'Stat {i+1} label')
    text(s, x, 43.8, 29.5, 5, ctx, size=14, muted=True, name=f'Stat {i+1} context')

# ================= 9 FEATURE =================
s = new(CREAM_L, 'FEATURE. Project or case study. Text left; photo panel right. Replace the grey photo area by inserting an image and cropping to the Shape.')
eyebrow(s, 'Section', 'Featured project')
text(s, 4.15, 11.5, 40, 14, 'Brick considerations in a heritage home', size=36, display=True, ls=0.95, spc=-0.5, name='Heading')
text(s, 4.15, 24.5, 38, 8, 'Two sentences on what the project was and what made it interesting. Keep it concrete.', size=14, name='Body')
for i, b in enumerate(['Point one, short and specific', 'Point two, short and specific', 'Point three, short and specific']):
    rule(s, 4.15, 34.5 + i * 4.8, 38, BLACK, 0.75, f'Bullet rule {i+1}')
    text(s, 4.15, 35.6 + i * 4.8, 38, 3, b, size=14, name=f'Bullet {i+1}')
rect(s, 49.85, 11.5, 46, 36, OLIVE, 'Panel')
polygon(s, [(0, .62), (.22, .62), (.27, .52), (.32, .42), (1, .42), (1, 1), (0, 1)], 49.85, 11.5, 46, 36, BLACK, 'Photo area (branded crop)', alpha=35)
text(s, 52, 13.6, 36, 8, 'Brick considerations\nin a heritage home', size=18, bold=True, caps=True, ls=0.95, name='Panel title')
icon(s, 52, 28, 5.2)
text(s, 52, 44.5, 20, 2, 'Drop project photo', size=10, bold=True, caps=True, spc=1.6, color=BLACK, name='Photo label')

# ================= 10 PHOTO GRID =================
s = new(CREAM_L, 'PHOTO GRID. Process shots (single-colour, behind the scenes) on the left; project shots (full colour) on the right. Insert images over the tiles.')
eyebrow(s, 'Section', 'Process and projects')
tiles = [('Process', 4.15, 11.6, 24.2, 37, True), ('Process', 29.85, 11.6, 19.4, 11.33, True), ('Process', 29.85, 24.43, 19.4, 24.17, True),
         ('Project', 50.75, 11.6, 14.5, 23.17, False), ('Project', 66.75, 11.6, 29.1, 23.17, False), ('Project', 50.75, 36.1, 45.1, 12.5, False)]
for i, (lab, x, y, w, h, proc) in enumerate(tiles):
    rect(s, x, y, w, h, BLACK, f'Photo tile {i+1}', brightness=0.78 if proc else 0.88)
    text(s, x + 1, y + h - 2.6, 15, 2, lab, size=10, bold=True, caps=True, spc=1.6, muted=True, name=f'Tile {i+1} label')

# ================= 11 QUOTE =================
s = new(OLIVE_L, 'QUOTE. A client or team testimonial in large type.')
eyebrow(s, 'Section', 'In their words')
text(s, 4.15, 11.5, 82, 30, '“They treated our home like it was theirs, and told us the truth at every step.”', size=48, display=True, ls=0.98, spc=-0.5, name='Quote')
text(s, 4.15, 41.5, 40, 2, 'Client name', size=11, bold=True, caps=True, spc=1.6, name='Attribution')
text(s, 4.15, 44, 40, 3, 'Project, neighbourhood', size=14, name='Attribution detail')
icon(s, 89.35, 43.5, 6.5)

# ================= 12 TONE DIALS =================
s = new(CREAM_L, 'SPECTRUM / TONE DIALS. Drag the coloured dot along each line to set a value.')
eyebrow(s, 'Section', '4.1 Tone dials')
H1(s, 'How do we\nspeak?', w=34)
dials = [('Casual', 'Formal', .58, BRICK), ('Fun', 'Serious', .60, SLATE), ('Realistic', 'Idealistic', .72, LUMBER),
         ('Approachable', 'Exclusive', .59, OLIVE), ('Traditional', 'Nontraditional', .69, BRICK), ('Friendly', 'Professional', .33, SLATE)]
for i, (a, b, v, col) in enumerate(dials):
    y = 11.6 + i * 6.3
    text(s, 51, y, 22, 2, a, size=10, bold=True, caps=True, spc=1.6, name=f'Dial {i+1} left')
    text(s, 73.85, y, 22, 2, b, size=10, bold=True, caps=True, spc=1.6, align='r', name=f'Dial {i+1} right')
    rule(s, 51, y + 3.2, 44.85, BLACK, 0.75, f'Dial {i+1} line')
    d = 1.35
    rect(s, 51 + v * 44.85 - d / 2, y + 3.2 - d / 2 * 0.0 - d / 2, d, d, col, f'Dial {i+1} dot', shape=MSO_SHAPE.OVAL)

# ================= 13 PILLARS =================
s = new(CREAM_L, 'PILLARS. Focus and specialty hierarchy.')
eyebrow(s, 'Section', '3.1 Company pillars')
H1(s, 'Shift Projects', w=40)
text(s, 51, 11.6, 20, 2, 'Focus', size=10, bold=True, caps=True, spc=1.6, name='Focus label'); rule(s, 51, 14.4, 44.85, BLACK, 1.25, 'Focus rule')
text(s, 51, 16, 44.85, 3, 'Residential projects', size=18, bold=True, caps=True, name='Focus title')
text(s, 51, 19.4, 44.85, 4, 'Custom home building and renovations in a considered, stress-free environment', size=14, name='Focus text')
text(s, 51, 27.6, 20, 2, 'Specialty', size=10, bold=True, caps=True, spc=1.6, name='Specialty label'); rule(s, 51, 30.4, 44.85, BLACK, 1.25, 'Specialty rule')
for i, (tt, bd) in enumerate([('New builds', 'From foundation to finish: thoughtfully designed, modern living solutions.'),
                              ('Home renovations', 'Enhancing the comfort, health, and functionality of existing homes.'),
                              ('Accessibility renovations', 'Home adjustments that support aging and mobility.')]):
    x = 51 + i * 15.6
    text(s, x, 32, 14.2, 6, tt, size=18, bold=True, caps=True, ls=0.95, name=f'Specialty {i+1} title')
    text(s, x, 39.5, 14.2, 10, bd, size=14, name=f'Specialty {i+1} text')

# ================= 14 DO / DON'T =================
s = new(CREAM_L, "DO / DON'T. Approved vs unapproved examples. Edit the sample text; the cross is two lines you can delete.")
eyebrow(s, 'Section', "Do's and don'ts")
H1(s, 'Say it\nthis way', w=28)
text(s, 35.5, 11.6, 30, 2, 'Approved', size=10, bold=True, caps=True, spc=1.6, name='Approved label'); rule(s, 35.5, 14.4, 33, BLACK, 1.25, 'Approved rule')
text(s, 72, 11.6, 24, 2, 'Unapproved', size=10, bold=True, caps=True, spc=1.6, name='Unapproved label'); rule(s, 72, 14.4, 23.8, BLACK, 1.25, 'Unapproved rule')
def sample(slide, x, y, w, h, fill, fg, outline=False, cross=False, cap='', name='Sample'):
    sh = rect(slide, x, y, w, h, fill, name, line=BLACK if outline else None)
    text(slide, x + 1.2, y + 1.1, w - 2, h - 2, 'Built for\ntoday', size=22, display=True, color=fg, ls=0.92, name=name + ' text')
    if cross:
        line(slide, x, y, x + w, y + h, BLACK, 1.25, name + ' cross 1'); line(slide, x, y + h, x + w, y, BLACK, 1.25, name + ' cross 2')
    text(slide, x, y + h + 0.6, w, 3, cap, size=10, muted=True, name=name + ' caption')
sample(s, 35.5, 16.2, 15.7, 11, CREAM, BLACK, outline=True, cap='Black on cream.', name='Approved 1')
sample(s, 52.8, 16.2, 15.7, 11, BLACK, CREAM, cap='Reversed.', name='Approved 2')
sample(s, 35.5, 31.6, 15.7, 11, SLATE, BLACK, cap='Black on a brand colour.', name='Approved 3')
sample(s, 52.8, 31.6, 15.7, 11, OLIVE, BLACK, cap='Black on a brand colour.', name='Approved 4')
sample(s, 72, 16.2, 23.8, 11, CREAM, BLACK, outline=True, cross=True, cap="Don't use white as a background; use cream.", name='Unapproved 1')
sample(s, 72, 31.6, 23.8, 11, BRICK, BLACK, cross=True, cap="Don't mix secondary colours.", name='Unapproved 2')

# ================= 15 CARDS =================
s = new(CREAM_L, 'CARDS. Four colour-block cards: options, services or audiences.')
eyebrow(s, 'Section', 'Options')
H1(s, 'Four ways\nto start', w=36)
cw = (91.7 - 4.5) / 4
for i, (tt, fill, fg, outline) in enumerate([('Option one', BLACK, CREAM, False), ('Option two', CREAM, BLACK, True), ('Option three', LUMBER, BLACK, False), ('Option four', SLATE, BLACK, False)]):
    x = 4.15 + i * (cw + 1.5)
    rect(s, x, 24, cw, 24.2, fill, f'Card {i+1}', line=BLACK if outline else None)
    text(s, x + 1.6, 25.6, cw - 3.2, 4, tt, size=18, bold=True, caps=True, color=fg, name=f'Card {i+1} title')
    text(s, x + 1.6, 29.4, cw - 3.2, 10, 'Short description in a sentence or two.', size=14, color=fg, name=f'Card {i+1} text')
    text(s, x + 1.6, 44.2, cw - 3.2, 2, 'Detail', size=10, bold=True, caps=True, spc=1.6, color=fg, name=f'Card {i+1} detail')

# ================= 16 STEPS =================
s = new(CREAM_L, 'STEPS. Stepped bars echo the Shift linear graphic. Process or timeline.')
eyebrow(s, 'Section', 'Our process')
H1(s, 'How a project\nmoves', w=28)
for i, (tt, bd) in enumerate([('Discover', 'What happens in this step.'), ('Design', 'What happens in this step.'), ('Build', 'What happens in this step.'), ('Handover', 'What happens in this step.')]):
    x = 35.5 + i * 3; y = 11.6 + i * 9.4
    rect(s, x, y, 100 - x, 8, BLACK, f'Step {i+1} bar')
    text(s, x + 2.4, y + 2.1, 5, 5, str(i + 1), size=36, display=True, color=LUMBER, name=f'Step {i+1} number')
    text(s, x + 8, y + 2.9, 18, 3, tt, size=18, bold=True, caps=True, color=CREAM, name=f'Step {i+1} title')
    text(s, x + 27, y + 3.1, 32, 3, bd, size=14, color=CREAM, name=f'Step {i+1} text')

# ================= 17 POLL =================
s = new(LUMBER_L, 'POLL / DISCUSSION. A question with lettered options.')
eyebrow(s, 'Discussion', 'Poll')
H1(s, 'What matters\nmost to you\nin a build?', w=42)
for i, o in enumerate(['Timeline', 'Budget', 'Design', 'Accessibility']):
    y = 11.6 + i * 8.7
    rule(s, 51, y, 44.85, BLACK, 1.25, f'Option {i+1} rule')
    c = rect(s, 51, y + 2.3, 3.4, 3.4, BLACK, f'Option {i+1} badge', shape=MSO_SHAPE.OVAL)
    text(s, 51, y + 2.3, 3.4, 3.4, 'ABCD'[i], size=14, bold=True, color=CREAM, align='c', anchor='m', name=f'Option {i+1} letter')
    text(s, 56.5, y + 2.6, 36, 4, o, size=22, anchor='m', name=f'Option {i+1} text')

# ================= 18 RECAP =================
s = new(CREAM_L, 'RECAP. Three-column outline of what was covered.')
eyebrow(s, 'Summary', 'What we covered')
H1(s, 'Recap', w=30)
for i, tt in enumerate(['First topic', 'Second topic', 'Third topic']):
    x = 4.15 + i * 31.6
    rule(s, x, 24, 29.5, BLACK, 1.5, f'Recap {i+1} rule')
    text(s, x, 25.4, 29.5, 6, f'0{i+1}', size=36, display=True, name=f'Recap {i+1} number')
    text(s, x, 32.6, 29.5, 3, tt, size=18, bold=True, caps=True, name=f'Recap {i+1} title')
    for j in range(3):
        rule(s, x, 36.8 + j * 3.6, 29.5, BLACK, 0.5, f'Recap {i+1} line {j+1}')
        text(s, x, 37.5 + j * 3.6, 29.5, 3, 'Key point', size=14, name=f'Recap {i+1} point {j+1}')

# ================= 19 CLOSING =================
s = new(LUMBER_L, 'CLOSING. Thank you, contact details, icon + stacked wordmark lockup.')
text(s, 4.15, 3.6, 80, 16, 'Thank you', size=110, display=True, spc=-2, name='Title')
text(s, 4.15, 33, 50, 12, 'Name\nname@shiftprojects.com\nshiftprojects.com', size=18, bold=True, caps=True, ls=1.05, name='Contact')
icon(s, 57.5, 38.5, 7)
stacked(s, 77.85, 37.4, 18)

prs.save(OUT)
print('saved', OUT, os.path.getsize(OUT))
