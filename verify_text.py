#!/usr/bin/env python3
"""Prove index.html carries the docx text exactly.

    python3 verify_text.py [path-to-docx]

Docx side: python-docx's own text extraction (not build_site's parser) of every
body paragraph and every table-cell paragraph, in document order.
HTML side: the text of every block element in <main> of index.html, in order.
Each side is reduced to non-empty lines with whitespace normalised, then diffed.
Elements marked data-added (the "Open the model" line), the section list
(<nav>), the iframe and scripts are reported separately as additions.
"""
import difflib
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

import docx
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
DEFAULT_DOCX = Path('/Users/skylardeture/Desktop/Sanctuary_Restored_Submission_money_fixed.docx')


def norm(s):
    return re.sub(r'\s+', ' ', s).strip()


def docx_lines(path):
    d = docx.Document(str(path))
    lines = []
    for el in d.element.body.iterchildren():
        if el.tag == qn('w:p'):
            lines.extend(docx.text.paragraph.Paragraph(el, d).text.split('\n'))
        elif el.tag == qn('w:tbl'):
            t = docx.table.Table(el, d)
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        lines.extend(p.text.split('\n'))   # in-cell line breaks
    return [norm(x) for x in lines if norm(x)]


BLOCK = {'p', 'h1', 'h2', 'h3', 'h4', 'li', 'th', 'td', 'figcaption'}
SKIP = {'script', 'style', 'iframe', 'nav', 'title', 'head'}


class Ext(HTMLParser):
    def __init__(self):
        super().__init__()
        self.lines, self.added = [], []
        self.cur, self.skip, self.added_depth, self.depth = [], 0, None, 0

    def flush(self):
        t = norm(''.join(self.cur))
        if t:
            (self.added if self.added_depth is not None else self.lines).append(t)
        self.cur = []

    def handle_starttag(self, tag, attrs):
        if tag in ('br', 'img', 'meta', 'link', 'input'):
            if tag == 'br':
                self.flush()
            return
        self.depth += 1
        if tag in SKIP:
            self.skip += 1
        if dict(attrs).get('data-added') and self.added_depth is None:
            self.flush()
            self.added_depth = self.depth
        if tag in BLOCK:
            self.flush()

    def handle_endtag(self, tag):
        if tag in ('br', 'img', 'meta', 'link', 'input'):
            return
        if tag in BLOCK:
            self.flush()
        if self.added_depth == self.depth:
            self.flush()
            self.added_depth = None
        if tag in SKIP:
            self.skip -= 1
        self.depth -= 1

    def handle_data(self, data):
        if not self.skip:
            self.cur.append(data)


def main():
    src = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else DEFAULT_DOCX
    a = docx_lines(src)
    page = (HERE / 'index.html').read_text(encoding='utf-8')
    body = page[page.index('<main>'):page.index('</main>')]
    e = Ext()
    e.feed(body)
    e.flush()
    b = e.lines
    diff = [l for l in difflib.unified_diff(a, b, 'docx', 'index.html', n=0, lineterm='')
            if not l.startswith(('---', '+++', '@@'))]
    print('docx lines: %d | html lines: %d | differing lines: %d' % (len(a), len(b), len(diff)))
    for l in diff[:40]:
        print('  ' + l[:200])
    print('words: docx %d, html %d' % (sum(len(x.split()) for x in a), sum(len(x.split()) for x in b)))
    print('additions outside his text (marked):', e.added)
    nav = re.search(r'<nav.*?</nav>', body, re.S)
    print('section list present:', bool(nav))


if __name__ == '__main__':
    main()
