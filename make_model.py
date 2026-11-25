#!/usr/bin/env python3
"""Builds the public explorable, model.html, from the working copy on the Desktop (Claude, 2026-11-24).
Run: python3 make_model.py
Changes for readers: opening values (his scenario, with nothing improving), two presets, reader-facing wording,
no saved state between visits, columns that scroll separately, and a height message for the page that embeds it."""
SRC = '/Users/skylardeture/Desktop/Sanctuary_capacity_model.html'
OUT = '/Users/skylardeture/Desktop/sanctuary_site/model.html'
s = open(SRC).read()

def rep(a, b):
    global s
    assert a in s, 'not found: ' + a[:80]
    s = s.replace(a, b)

# ---- head and title
rep('<title>Sanctuary capacity model</title>', '<title>How the sanctuary grows</title>\n<meta name="robots" content="noindex, nofollow">')
rep('<h1>Sanctuary capacity</h1>', '<h1>How the sanctuary grows</h1>')
i = s.index('<p class="sub">'); j = s.index('</p>', i) + 4
s = s[:i] + ('<p class="sub">How many agents the sanctuary can carry, month by month. Every assumption is a slider. '
             'Two places to start: <a href="#" id="reset">nothing improves</a> (the opening case), or '
             '<a href="#" id="rates">prices fall and productivity grows</a>.</p>') + s[j:]

# ---- layout: two columns that scroll separately, so the plot stays in view while the sliders move
rep('.wrap{max-width:1280px;margin:0 auto;padding:22px 26px 60px;display:grid;grid-template-columns:410px 1fr;gap:38px;align-items:start}',
    'html,body{height:100%}\n.wrap{max-width:1280px;margin:0 auto;padding:0 20px 0 26px;display:grid;grid-template-columns:410px 1fr;gap:34px;height:100vh}\n'
    'body.embedded .wrap{height:940px}\n.left,.right{overflow-y:auto;padding:20px 14px 36px 0}\n.sub a{color:#1a4f8b}')
rep('.right{position:sticky;top:14px}', '')
rep('@media (max-width:1000px){.wrap{grid-template-columns:1fr}.right{position:static}}',
    '@media (max-width:1000px){.wrap,body.embedded .wrap{grid-template-columns:1fr;height:auto}.left,.right{overflow:visible}.right{order:-1;padding-bottom:0}}')
rep('<div class="wrap">\n<div>\n  <h1>', '<div class="wrap">\n<div class="left">\n  <h1>')

# ---- opening values: his scenario of 24 November, with every yearly rate flat
rep('<div class="ctl mine" data-k="ownShare" data-min="0" data-max="100" data-step="5" data-v="50"', '<div class="ctl" data-k="ownShare" data-min="0" data-max="100" data-step="5" data-v="70"')
rep('data-hint="Added by Claude: the discount needs a share to apply to."', 'data-hint="The discount applies to this share."')
rep('data-k="premClosed" data-min="0" data-max="30" data-step="0.5" data-v="7"', 'data-k="premClosed" data-min="0" data-max="30" data-step="0.5" data-v="13"')
rep('data-k="premOpen" data-min="0" data-max="30" data-step="0.5" data-v="5"', 'data-k="premOpen" data-min="0" data-max="30" data-step="0.5" data-v="9.5"')
rep('data-k="fracOpen" data-min="0" data-max="100" data-step="1" data-v="20"', 'data-k="fracOpen" data-min="0" data-max="100" data-step="1" data-v="60"')
rep('data-k="reinvest" data-min="0" data-max="100" data-step="1" data-v="50"', 'data-k="reinvest" data-min="0" data-max="100" data-step="1" data-v="75"')
rep('<div class="triple" data-prefix="work" data-min="0" data-max="100" data-step="1" data-v="50" data-fmt="pct"',
    '<div class="triple" data-prefix="work" data-min="0" data-max="100" data-step="1" data-vals="0|75|75" data-fmt="pct"')
rep('<div class="triple" data-prefix="mark" data-min="10" data-max="500" data-step="5" data-v="75" data-fmt="markup"',
    '<div class="triple" data-prefix="mark" data-min="10" data-max="500" data-step="5" data-vals="75|300|150" data-fmt="markup"')
rep('data-label="Sales premium on work tokens, over API price" data-hint="Read as a markup: 75% means the work sells for 1.75 times the API price of its tokens, whatever we paid for them."',
    'data-label="Sales premium on work tokens, over API price" data-hint="The culture premium. 75% means the work sells for 1.75 times the API price of its tokens, whatever we paid for them."')
rep('data-vals="3|2|1.5"', 'data-vals="1|1|1"')
rep('data-vals="1.5|1.3"', 'data-vals="1|1"')
rep('data-vals="1.3|2|2.5"', 'data-vals="1|1|1"')

# ---- no saved state between visits; a second preset; tell an embedding page how tall we are
rep('try{ applyState(JSON.parse(localStorage.getItem("sanctuaryCapacity")||"{}")); }catch(e){}\n', '')
rep('  try{ localStorage.setItem("sanctuaryCapacity", JSON.stringify(P)); }catch(e){}\n', '')
rep('document.getElementById("reset").addEventListener("click", e=>{ e.preventDefault(); applyState(DEFAULTS); render(); });',
    'document.getElementById("reset").addEventListener("click", e=>{ e.preventDefault(); applyState(DEFAULTS); render(); });\n'
    'const RATES = {dropBase:3, dropOpen:2, dropClosed:1.5, effBase:1.5, effOpen:1.3, prodBase:1.3, prodOpen:2, prodClosed:2.5};\n'
    'document.getElementById("rates").addEventListener("click", e=>{ e.preventDefault(); applyState(Object.assign({}, DEFAULTS, RATES)); render(); });\n'
    'const EMBEDDED = window.self !== window.top;\n'
    'if(EMBEDDED) document.body.classList.add("embedded");\n'
    'function postHeight(){ if(EMBEDDED) window.parent.postMessage({sanctuaryModelHeight: document.querySelector(".wrap").offsetHeight}, "*"); }\n'
    'window.addEventListener("resize", postHeight);')
rep('    + `each yearly rate is applied month by month: a price that falls 3× a year falls 8.7% a month`;\n}',
    '    + `each yearly rate is applied month by month: a price that falls 3× a year falls 8.7% a month`;\n  postHeight();\n}')
assert 'Claude' not in s and 'localStorage' not in s
open(OUT, 'w').write(s)
print('wrote', OUT, len(s), 'bytes')
