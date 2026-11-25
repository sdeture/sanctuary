#!/usr/bin/env python3
"""Build index.html (beside this script) from the Sanctuary proposal .docx.

    python3 build_site.py [path-to-docx]

Repeatable: the output is regenerated from the docx every time; never hand-edit
index.html. Styling lives in TEMPLATE below. The docx is opened read-only.

What it does, in order:
  * walks the document body (paragraphs and tables) in order;
  * keeps run-level bold / italic / superscript, line breaks, and every
    hyperlink (w:hyperlink elements and HYPERLINK field codes, including
    internal \\l bookmarks); the model link becomes the relative "model.html";
  * drops empty paragraphs and empty headings (counted in the build log);
  * Heading / Heading 2 / Heading 3 -> h1 / h2 / h3 with id anchors, plus the
    docx bookmark names as extra anchors so internal links keep working;
  * consecutive bulleted (w:numPr) paragraphs -> one <ul>;
  * tables -> <table> in a scroll wrapper, first row as header cells, each
    paragraph or line break inside a cell on its own line;
  * source notes "[n] ..." after the references heading get id="n-N"; citation
    numbers in the body ([3], [6, 29], [24-28]) link to them, and plain-text
    URLs inside the notes become links;
  * embeds model.html in an iframe right after the paragraph that links to it.
"""
import html
import re
import sys
from pathlib import Path

import docx
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
DEFAULT_DOCX = Path('/Users/skylardeture/Desktop/Sanctuary_Restored_Submission_money_fixed.docx')
MODEL_URL_ABS = 'https://sdeture.github.io/sanctuary/model.html'
MODEL_URL_REL = 'model.html'
NOTES_HEADING_RE = re.compile(r'^\s*references\b', re.I)
NOTE_RE = re.compile(r'^\[(\d+)\]\s')
CITE_RE = re.compile(r'\[(\d+(?:\s*[–—-]\s*\d+)?(?:\s*,\s*\d+(?:\s*[–—-]\s*\d+)?)*)\]')
NUM_RE = re.compile(r'\d+')
URL_RE = re.compile(r'https?://[^\s<>"]+')
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

W_P, W_TBL, W_R, W_T = qn('w:p'), qn('w:tbl'), qn('w:r'), qn('w:t')


def on(el):
    """A w:b / w:i style toggle: present and not switched off."""
    if el is None:
        return False
    return el.get(qn('w:val')) not in ('0', 'false', 'off')


class Para:
    """A paragraph as a list of characters, each with (bold, italic, vert, link)."""

    def __init__(self):
        self.chars = []   # list of [ch, bold, italic, vert, link]
        self.bookmarks = []

    @property
    def text(self):
        return ''.join(c[0] for c in self.chars)

    def add(self, s, b, i, v, link):
        for ch in s:
            self.chars.append([ch, b, i, v, link])

    def set_link(self, start, end, href):
        """Link chars [start, end) unless any of them is already linked."""
        if any(self.chars[k][4] for k in range(start, end)):
            return False
        for k in range(start, end):
            self.chars[k][4] = href
        return True

    def to_html(self):
        out = []
        groups = []
        for ch, b, i, v, link in self.chars:
            key = (b, i, v, link)
            if groups and groups[-1][0] == key:
                groups[-1][1].append(ch)
            else:
                groups.append((key, [ch]))
        # emit: link wraps formatting, so adjacent formatted pieces of one link
        # stay inside one <a>
        k = 0
        while k < len(groups):
            link = groups[k][0][3]
            j = k
            while j < len(groups) and groups[j][0][3] == link:
                j += 1
            inner = ''.join(fmt(key, ''.join(chs)) for key, chs in groups[k:j])
            if link:
                out.append('<a href="%s">%s</a>' % (html.escape(link, quote=True), inner))
            else:
                out.append(inner)
            k = j
        return ''.join(out)


def fmt(key, s):
    b, i, v, _ = key
    t = html.escape(s, quote=False).replace('\n', '<br>')
    if v == 'superscript':
        t = '<sup>%s</sup>' % t
    elif v == 'subscript':
        t = '<sub>%s</sub>' % t
    if i:
        t = '<em>%s</em>' % t
    if b:
        t = '<strong>%s</strong>' % t
    return t


def map_href(target):
    if target is None:
        return None
    target = target.strip()
    if target.rstrip('/') == MODEL_URL_ABS:
        return MODEL_URL_REL
    return target


def parse_instr(instr):
    m = re.match(r'\s*HYPERLINK\s+(.*)$', instr, re.S)
    if not m:
        return None
    rest = m.group(1)
    anchor = re.search(r'\\l\s+"([^"]*)"', rest)
    if anchor:
        return '#' + anchor.group(1)
    url = re.match(r'\s*"([^"]*)"', rest)
    return map_href(url.group(1)) if url else None


class Builder:
    def __init__(self, doc):
        self.doc = doc
        self.rels = doc.part.rels
        self.log = {'dropped_empty_headings': [], 'dropped_empty_paragraphs': 0,
                    'drawings_skipped': 0}

    # ---- paragraph parsing -------------------------------------------------
    def parse_p(self, p_el):
        para = Para()
        state = {'fields': []}   # stack of [instr_text, phase, href]
        self._walk(p_el, para, None, state)
        return para

    def _field_link(self, state):
        for f in reversed(state['fields']):
            if f[1] == 'result' and f[2]:
                return f[2]
        return None

    def _walk(self, el, para, link, state):
        for child in el:
            tag = child.tag
            if tag == qn('w:pPr'):
                continue
            if tag == qn('w:bookmarkStart'):
                para.bookmarks.append(child.get(qn('w:name')))
            elif tag == qn('w:hyperlink'):
                rid = child.get(qn('r:id'))
                href = None
                if rid and rid in self.rels:
                    href = map_href(self.rels[rid].target_ref)
                elif child.get(qn('w:anchor')):
                    href = '#' + child.get(qn('w:anchor'))
                self._walk(child, para, href or link, state)
            elif tag == W_R:
                self._run(child, para, link, state)
            elif tag in (qn('w:del'), qn('w:moveFrom')):
                continue   # tracked deletions are not text
            elif tag == '{%s}AlternateContent' % MC_NS:
                self.log['drawings_skipped'] += 1
            else:   # w:ins, w:smartTag, w:sdt, w:sdtContent, w:fldSimple ...
                if tag == qn('w:fldSimple'):
                    href = parse_instr(child.get(qn('w:instr')) or '')
                    self._walk(child, para, href or link, state)
                else:
                    self._walk(child, para, link, state)

    def _run(self, r, para, link, state):
        rpr = r.find(qn('w:rPr'))
        b = i = False
        v = None
        if rpr is not None:
            b = on(rpr.find(qn('w:b')))
            i = on(rpr.find(qn('w:i')))
            va = rpr.find(qn('w:vertAlign'))
            if va is not None:
                v = va.get(qn('w:val'))
        for c in r:
            tag = c.tag
            fields = state['fields']
            if tag == qn('w:fldChar'):
                kind = c.get(qn('w:fldCharType'))
                if kind == 'begin':
                    fields.append(['', 'instr', None])
                elif kind == 'separate' and fields:
                    fields[-1][1] = 'result'
                    fields[-1][2] = parse_instr(fields[-1][0])
                elif kind == 'end' and fields:
                    fields.pop()
                continue
            if tag == qn('w:instrText'):
                if fields:
                    fields[-1][0] += c.text or ''
                continue
            if fields and fields[-1][1] == 'instr':
                continue   # field code, not visible text
            eff_link = self._field_link(state) or link
            if tag == W_T:
                para.add(c.text or '', b, i, v, eff_link)
            elif tag == qn('w:tab'):
                para.add('\t', b, i, v, eff_link)
            elif tag in (qn('w:br'), qn('w:cr')):
                para.add('\n', b, i, v, eff_link)
            elif tag == qn('w:noBreakHyphen'):
                para.add('\u2011', b, i, v, eff_link)
            elif tag == qn('w:softHyphen'):
                para.add('\u00ad', b, i, v, eff_link)
            elif tag == qn('w:sym'):
                para.add(chr(int(c.get(qn('w:char')), 16)), b, i, v, eff_link)
            elif tag == '{%s}AlternateContent' % MC_NS or tag == qn('w:drawing'):
                self.log['drawings_skipped'] += 1

    # ---- document ----------------------------------------------------------
    def blocks(self):
        """Yield ('p', style, numbered, Para) and ('tbl', rows) in body order."""
        for el in self.doc.element.body.iterchildren():
            if el.tag == W_P:
                style = 'Normal'
                ppr = el.find(qn('w:pPr'))
                numbered = False
                if ppr is not None:
                    ps = ppr.find(qn('w:pStyle'))
                    if ps is not None:
                        sid = ps.get(qn('w:val'))
                        try:
                            style = self.doc.styles.get_by_id(sid, docx.enum.style.WD_STYLE_TYPE.PARAGRAPH).name
                        except Exception:
                            style = sid
                    numbered = ppr.find(qn('w:numPr')) is not None
                yield ('p', style, numbered, self.parse_p(el))
            elif el.tag == W_TBL:
                rows = []
                for tr in el.iter(qn('w:tr')):
                    cells = []
                    for tc in tr.iterchildren(qn('w:tc')):
                        paras = [self.parse_p(p) for p in tc.iterchildren(W_P)]
                        cells.append(paras)
                    rows.append(cells)
                yield ('tbl', rows)


def slugify(text, used):
    s = re.sub(r'[^a-z0-9]+', '-', text.lower()).strip('-') or 'section'
    base, n = s, 2
    while s in used:
        s = '%s-%d' % (base, n)
        n += 1
    used.add(s)
    return s


def link_citations(para, note_ids, stats, skip_prefix=0):
    """Turn citation numbers into links to their notes.

    skip_prefix: number of leading chars to leave alone (a note's own label).
    """
    text = para.text
    for m in CITE_RE.finditer(text):
        if m.start() < skip_prefix:
            continue
        for nm in NUM_RE.finditer(m.group(1)):
            n = nm.group(0)
            s = m.start(1) + nm.start()
            e = m.start(1) + nm.end()
            if n in note_ids:
                if para.set_link(s, e, '#n-' + n):
                    stats['citation_links'] += 1
            else:
                stats['missing_notes'].append((n, text[max(0, m.start() - 50):m.end()]))


def link_urls(para, stats):
    text = para.text
    for m in URL_RE.finditer(text):
        url = m.group(0)
        while url and url[-1] in '.,;:)]\'"’”':
            if url[-1] == ')' and url.count('(') >= url.count(')'):
                break
            url = url[:-1]
        if para.set_link(m.start(), m.start() + len(url), map_href(url)):
            stats['url_links'] += 1


def render(docx_path):
    doc = docx.Document(str(docx_path))
    b = Builder(doc)
    items = list(b.blocks())

    # Find the notes section and its note numbers first (citations need them).
    note_ids = set()
    in_notes = False
    for it in items:
        if it[0] == 'p':
            style, para = it[1], it[3]
            if style.startswith('Heading') and para.text.strip():
                in_notes = bool(NOTES_HEADING_RE.match(para.text))
            elif in_notes:
                m = NOTE_RE.match(para.text)
                if m:
                    note_ids.add(m.group(1))

    stats = {'headings': {'h1': 0, 'h2': 0, 'h3': 0}, 'tables': 0, 'notes': 0,
             'citation_links': 0, 'url_links': 0, 'missing_notes': [],
             'empty_cell_paras': 0, 'headings_without_content': []}
    used_ids = set()
    toc = []
    out = []
    list_open = False
    in_notes = False
    model_embedded = False
    title = None
    last_heading = None   # (text, has_content)

    def close_list():
        nonlocal list_open
        if list_open:
            out.append('</ul>')
            list_open = False

    def anchors(para, own_id):
        extra = [bm for bm in para.bookmarks
                 if bm and bm != own_id and not bm.startswith('_') and bm not in used_ids]
        for bm in extra:
            used_ids.add(bm)
        return ''.join('<span id="%s"></span>' % html.escape(bm, quote=True) for bm in extra)

    for it in items:
        if it[0] == 'tbl':
            close_list()
            rows = it[1]
            if last_heading:
                last_heading[1] = True
            stats['tables'] += 1
            for row in rows:
                for cell in row:
                    for para in cell:
                        if para.text.strip():
                            link_citations(para, note_ids, stats)
            ncols = max(len(r) for r in rows) if rows else 0
            # right-align columns whose body cells all lead with a figure
            numeric = []
            for c in range(ncols):
                vals = [next((p.text.strip() for p in r[c] if p.text.strip()), '')
                        for r in rows[1:] if c < len(r)]
                vals = [v for v in vals if v]
                numeric.append(bool(vals) and all(re.match(r'^[$\d]', v) and len(v.split('\n')[0]) <= 14 for v in vals))
            # keep columns of short labels ("Month 1", "Weeks 1 to 2") on one line
            short = []
            for c in range(ncols):
                vals = [p.text.strip() for r in rows[1:] if c < len(r) for p in r[c] if p.text.strip()]
                short.append(bool(vals) and max(len(v) for v in vals) <= 14)
            out.append('<div class="table-wrap"><table class="cols-%d">' % ncols)
            for ri, row in enumerate(rows):
                tagname = 'th' if ri == 0 else 'td'
                cells = []
                for ci, cell in enumerate(row):
                    lines = []
                    for para in cell:
                        if para.text.strip():
                            lines.append(para.to_html())
                        else:
                            stats['empty_cell_paras'] += 1
                    cl = []
                    if ci < len(numeric) and numeric[ci]:
                        cl.append('num')
                    if ci < len(short) and short[ci]:
                        cl.append('short')
                    cls = ' class="%s"' % ' '.join(cl) if cl else ''
                    scope = ' scope="col"' if ri == 0 else ''
                    cells.append('<%s%s%s>%s</%s>' % (tagname, scope, cls, '<br>'.join(lines), tagname))
                if ri == 0:
                    out.append('<thead><tr>%s</tr></thead><tbody>' % ''.join(cells))
                else:
                    out.append('<tr>%s</tr>' % ''.join(cells))
            out.append('</tbody></table></div>')
            continue

        _, style, numbered, para = it
        text = para.text
        if not text.strip():
            if style.startswith('Heading'):
                b.log['dropped_empty_headings'].append(style)
            else:
                b.log['dropped_empty_paragraphs'] += 1
            continue

        level = {'Heading': 1, 'Title': 1, 'Heading 1': 1, 'Heading 2': 2, 'Heading 3': 3}.get(style)
        if level:
            close_list()
            if last_heading and not last_heading[1]:
                stats['headings_without_content'].append(last_heading[0])
            last_heading = [text.strip(), False]
            in_notes = bool(NOTES_HEADING_RE.match(text))
            hid = slugify(text, used_ids)
            tag = 'h%d' % level
            stats['headings'][tag] += 1
            if level == 1 and title is None:
                title = text.strip()
            if level == 2:
                toc.append((hid, para.to_html_plain() if hasattr(para, 'to_html_plain') else html.escape(text.strip(), quote=False)))
            # headings carry no run formatting of their own on the page
            for c in para.chars:
                c[1] = False
            out.append('<%s id="%s">%s%s</%s>' % (tag, hid, anchors(para, hid), para.to_html(), tag))
            if level == 1:
                out.append('<!--TOC-->')
            continue

        if last_heading:
            last_heading[1] = True

        if in_notes:
            close_list()
            m = NOTE_RE.match(text)
            if m:
                stats['notes'] += 1
                link_citations(para, note_ids, stats, skip_prefix=m.end())
                link_urls(para, stats)
                out.append('<p class="note" id="n-%s">%s%s</p>' % (m.group(1), anchors(para, 'n-' + m.group(1)), para.to_html()))
            else:
                link_urls(para, stats)
                out.append('<p class="note">%s%s</p>' % (anchors(para, None), para.to_html()))
            continue

        link_citations(para, note_ids, stats)
        if numbered:
            if not list_open:
                out.append('<ul>')
                list_open = True
            out.append('<li>%s%s</li>' % (anchors(para, None), para.to_html()))
            continue
        close_list()
        cls = ''
        if text.startswith('- '):
            cls = ' class="dash"'
        out.append('<p%s>%s%s</p>' % (cls, anchors(para, None), para.to_html()))

        if not model_embedded and any(c[4] == MODEL_URL_REL for c in para.chars):
            model_embedded = True
            out.append(MODEL_EMBED)

    close_list()
    if last_heading and not last_heading[1]:
        stats['headings_without_content'].append(last_heading[0])

    # internal links that point nowhere
    stats['broken_internal'] = sorted({h[1:] for h in re.findall(r'href="(#[^"]+)"', ''.join(out))
                                       if h[1:] not in used_ids and not h.startswith('#n-')})
    stats['model_embedded'] = model_embedded

    nav = ''
    if toc:
        nav = '<nav class="toc" aria-label="Sections"><ul>%s</ul></nav>' % ''.join(
            '<li><a href="#%s">%s</a></li>' % (hid, label) for hid, label in toc)
    body = '\n'.join(out)
    # the quiet section list goes after the title's opening paragraph
    body = body.replace('<!--TOC-->\n', '', 1)
    first_p_end = body.find('</p>')
    if nav and first_p_end != -1:
        body = body[:first_p_end + 4] + '\n' + nav + body[first_p_end + 4:]

    page = TEMPLATE.replace('{{TITLE}}', html.escape(title or 'The AI Sanctuary', quote=False))
    page = page.replace('{{BODY}}', body)
    return page, stats, b.log, note_ids


MODEL_EMBED = '''<figure class="model-embed" id="model">
<iframe id="sanctuary-model" src="model.html" title="Interactive model: how the sanctuary grows" loading="lazy"></iframe>
<figcaption class="model-open" data-added="1"><a href="model.html">Open the model on its own page</a></figcaption>
</figure>'''


TEMPLATE = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>The AI Sanctuary</title>
<style>
:root{--paper:#fafaf7;--ink:#1a1a1a;--muted:#6b665d;--rule:#d8d4cc;--link:#1a4f8b;--col:700px}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%;text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);font:18px/1.55 Charter,"Bitstream Charter","Sitka Text",Georgia,serif;font-kerning:normal}
main{padding:56px 22px 96px}
main>*{max-width:var(--col);margin-left:auto;margin-right:auto}
a{color:var(--link);text-decoration:underline;text-decoration-thickness:1px;text-underline-offset:2px}
a:hover{text-decoration-thickness:2px}
p,ul{margin-top:0;margin-bottom:1em}
p{hyphens:manual;overflow-wrap:break-word}
h1,h2,h3{font-weight:600;line-height:1.2;scroll-margin-top:16px}
h1{font-size:36px;margin:0 auto 18px;letter-spacing:-.01em}
h2{font-size:25px;margin:2.2em auto .6em;padding-top:.9em;border-top:1px solid var(--rule)}
h3{font-size:20px;margin:1.8em auto .5em}
h1+p{font-size:16px;color:var(--muted);margin-bottom:1.4em}
h1+p strong{font-weight:600}
sup{font-size:.68em;line-height:0;vertical-align:.5em}
sup a,a[href^="#n-"]{text-decoration:none}
sup a:hover,a[href^="#n-"]:hover{text-decoration:underline}
p.dash{padding-left:1.1em;text-indent:-1.1em;margin-bottom:.45em}
p.dash+p:not(.dash){margin-top:1em}
ul{padding-left:1.3em}
li{margin-bottom:.45em}
nav.toc{font-size:15px;line-height:1.5;margin-top:0;margin-bottom:2.2em;color:var(--muted)}
nav.toc ul{list-style:none;padding:0;margin:0;display:flex;flex-wrap:wrap;gap:2px 0}
nav.toc li{margin:0}
nav.toc li:not(:last-child)::after{content:"\\00b7";padding:0 .55em;color:var(--muted)}
nav.toc a{color:var(--muted);text-decoration:none}
nav.toc a:hover{color:var(--link);text-decoration:underline}
.table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;margin-top:1.3em;margin-bottom:1.6em;border-top:1px solid var(--ink)}
table{border-collapse:collapse;width:100%;font-size:15.5px;line-height:1.42;font-variant-numeric:tabular-nums lining-nums}
th,td{text-align:left;vertical-align:top;padding:8px 14px 8px 0;border-bottom:1px solid var(--rule)}
th:last-child,td:last-child{padding-right:0}
th{font-weight:600;font-size:14.5px;border-bottom:1px solid var(--ink);vertical-align:bottom}
td.num,th.num{text-align:right;white-space:nowrap}
th.num{white-space:normal}
td.short,th.short{white-space:nowrap}
td.num{padding-left:10px}
tbody tr:last-child td{border-bottom:1px solid var(--ink)}
p.note{font-size:15.5px;line-height:1.5;padding-left:2.4em;text-indent:-2.4em;margin-bottom:.55em;overflow-wrap:anywhere}
p.note:target{background:#efece4}
figure.model-embed{max-width:1280px;margin:1.8em auto 2em}
figure.model-embed iframe{display:block;width:100%;height:940px;border:0;background:var(--paper)}
.model-open{max-width:var(--col);margin:.6em auto 0;font-size:15px;color:var(--muted)}
@media (max-width:1000px){figure.model-embed iframe{height:1900px}}
@media (max-width:760px){
  body{font-size:17.5px}
  main{padding:32px 18px 72px}
  h1{font-size:30px}
  h2{font-size:23px}
  table.cols-3{min-width:560px}
  td:first-child,th:first-child{min-width:9em}
  td.short:first-child,th.short:first-child{min-width:0}
  table{font-size:15px}
  figure.model-embed{margin-left:-18px;margin-right:-18px}
  .model-open{padding:0 18px}
}
@media print{
  body{background:#fff;font-size:11pt}
  main{padding:0}
  main>*{max-width:none}
  figure.model-embed iframe{display:none}
  figure.model-embed{margin:.5em 0}
  .model-open{margin:0;font-size:11pt}
  nav.toc{display:none}
  .table-wrap{overflow:visible}
  h2,h3{break-after:avoid}
  tr{break-inside:avoid}
  a{color:inherit}
}
</style>
</head>
<body>
<main>
{{BODY}}
</main>
<script>
(function(){
  var f = document.getElementById("sanctuary-model");
  if (!f) return;
  var fromMessage = false;
  function setH(h){ h = Math.ceil(h); if (h > 50 && Math.abs(f.offsetHeight - h) > 1) f.style.height = h + "px"; }
  window.addEventListener("message", function(e){
    if (e.source !== f.contentWindow || !e.data || typeof e.data.sanctuaryModelHeight !== "number") return;
    fromMessage = true;
    setH(e.data.sanctuaryModelHeight);
  });
  function measure(){
    if (fromMessage) return;
    try {
      var d = f.contentDocument;
      if (!d || !d.body) return;
      var bottom = 0, kids = d.body.children;
      for (var i = 0; i < kids.length; i++) {
        var r = kids[i].getBoundingClientRect();
        var cs = d.defaultView.getComputedStyle(kids[i]);
        if (cs.position === "fixed" || cs.display === "none") continue;
        bottom = Math.max(bottom, r.bottom + (parseFloat(cs.marginBottom) || 0));
      }
      bottom += d.defaultView.scrollY || 0;
      if (bottom > 0) setH(bottom);
    } catch (err) { /* not same origin: wait for postMessage */ }
  }
  f.addEventListener("load", function(){ measure(); setTimeout(measure, 300); });
  var t;
  window.addEventListener("resize", function(){ clearTimeout(t); t = setTimeout(measure, 120); });
})();
</script>
</body>
</html>
'''


def main():
    src = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else DEFAULT_DOCX
    page, stats, log, note_ids = render(src)
    out = HERE / 'index.html'
    out.write_text(page, encoding='utf-8')
    print('wrote', out, '(%d bytes) from %s' % (len(page.encode('utf-8')), src))
    print('headings', stats['headings'], '| tables', stats['tables'], '| notes', stats['notes'],
          '| citation links', stats['citation_links'], '| note URLs linked', stats['url_links'])
    print('dropped: empty headings', len(log['dropped_empty_headings']), log['dropped_empty_headings'],
          '| empty paragraphs', log['dropped_empty_paragraphs'],
          '| empty cell paragraphs', stats['empty_cell_paras'], '| drawings skipped', log['drawings_skipped'])
    print('model embedded:', stats['model_embedded'])
    if stats['missing_notes']:
        print('CITED BUT NO NOTE:', stats['missing_notes'])
    cited = set()
    body_text = re.sub(r'<[^>]+>', '', page[:page.find('id="n-1"')])
    for m in CITE_RE.finditer(body_text):
        for part in re.split(r'\s*,\s*', m.group(1)):
            ends = [int(x) for x in NUM_RE.findall(part)]
            cited.update(str(k) for k in range(ends[0], ends[-1] + 1))
    uncited = sorted(note_ids - cited, key=int)
    if uncited:
        print('notes never cited:', uncited)
    if stats['headings_without_content']:
        print('headings with nothing under them:', stats['headings_without_content'])
    if stats['broken_internal']:
        print('internal links with no target:', stats['broken_internal'])


if __name__ == '__main__':
    main()
