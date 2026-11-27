"""Command line: archdiagrams render | add | icons | new | themes."""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

from . import __version__
from . import icons as _icons
from .diagram import DiagramError
from .spec import add_to_deck, load, new_deck
from .theme import THEMES

STARTERS = {
 'tiers': """# Left-to-right tiers: users -> edge -> application -> data. Coordinates are inches inside `size`.
theme: default
size: [14, 5.6]
slide:
  title: Solution architecture
  caption: One line that says what the reader should notice.
elements:
  - zone: {id: users, at: [0, 0, 2.8, 5.6], label: Users, color: teal,
           cards: [{id: staff, icon: users, title: Staff, sub: HTTPS},
                   {id: admin, icon: admin, title: Admins, sub: HTTPS · MFA},
                   {id: mob, icon: mobile, title: Mobile app, sub: HTTPS}]}
  - zone: {id: edge, at: [3.3, 0, 2.6, 5.6], label: Edge, color: orange,
           cards: [{id: waf, icon: firewall, title: WAF / proxy, sub: TLS · rate limit},
                   {id: iam, icon: identity, title: Identity, sub: OIDC · MFA}]}
  - zone: {id: app, at: [6.4, 0, 3.6, 5.6], label: Application, color: blue, cols: 1,
           cards: [{id: api, icon: api, title: API services, sub: REST},
                   {id: jobs, icon: queue, title: Background jobs, sub: queue}]}
  - zone: {id: data, at: [10.5, 0, 3.5, 5.6], label: Data, color: green,
           cards: [{id: db, icon: database, title: Database, sub: primary + replica},
                   {id: files, icon: storage, title: Object storage, sub: documents}]}
  - bus: {from: [staff, admin, mob], to: waf, color: teal}
  - arrow: {from: waf.r, to: api.l, color: blue}
  - arrow: {from: iam.r, to: api.l@0.7, color: orange, dashed: true}
  - arrow: {from: api.r, to: db.l, color: green}
  - arrow: {from: jobs.r, to: files.l, color: green}
""",
 'flow': """# A numbered flow: put badges on the cards in the order things happen.
theme: default
size: [14, 4.6]
slide:
  title: How an order flows
elements:
  - card: {id: a, at: [0, 1.6, 2.6, 1.2], icon: user, title: Customer, sub: places order}
  - card: {id: b, at: [3.8, 1.6, 2.6, 1.2], icon: api, title: Order API, sub: validates, color: blue}
  - card: {id: c, at: [7.6, 1.6, 2.6, 1.2], icon: queue, title: Queue, sub: decouples, color: blue}
  - card: {id: d, at: [11.4, 1.6, 2.6, 1.2], icon: database, title: Orders DB, sub: stores, color: green}
  - arrow: {from: a.r, to: b.l, label: HTTPS}
  - arrow: {from: b.r, to: c.l, label: event}
  - arrow: {from: c.r, to: d.l, label: write}
  - badges: [[a, 1], [b, 2], [c, 3], [d, 4]]
""",
 'compare': """# Two designs side by side, each with a question to think about.
theme: default
size: [14, 5.6]
slide:
  title: Two designs for the same need
elements:
  - zone: {id: A, at: [0, 0, 6.8, 5.6], label: Design A · central, color: navy}
  - card: {id: hub, at: [2.0, 0.6, 2.8, 1.0], icon: server, title: Central service, color: navy, emphasis: true}
  - card: {id: s1, at: [0.3, 3.0, 1.9, 0.9], icon: building, title: Site 1}
  - card: {id: s2, at: [2.45, 3.0, 1.9, 0.9], icon: building, title: Site 2}
  - card: {id: s3, at: [4.6, 3.0, 1.9, 0.9], icon: building, title: Site N}
  - arrow: {from: hub.b, to: s1.t, both: true, color: navy, route: straight}
  - arrow: {from: hub.b, to: s2.t, both: true, color: navy, route: straight}
  - arrow: {from: hub.b, to: s3.t, both: true, color: navy, route: straight}
  - note: {at: [0.3, 4.4, 6.2, 0.6], text: "Think about: what happens when sites double?", color: orange, bold: true, italic: false}
  - zone: {id: B, at: [7.2, 0, 6.8, 5.6], label: Design B · event-driven, color: teal}
  - card: {id: br, at: [9.2, 1.6, 2.8, 0.9], icon: stream, title: Event broker, color: blue, emphasis: true}
  - card: {id: t1, at: [7.5, 3.0, 1.9, 0.9], icon: gateway, title: Site 1}
  - card: {id: t2, at: [9.65, 3.0, 1.9, 0.9], icon: gateway, title: Site 2}
  - card: {id: t3, at: [11.8, 3.0, 1.9, 0.9], icon: gateway, title: Site N}
  - arrow: {from: t1.t, to: br.b, color: teal, route: straight}
  - arrow: {from: t2.t, to: br.b, color: teal, route: straight}
  - arrow: {from: t3.t, to: br.b, color: teal, route: straight}
  - note: {at: [7.5, 4.4, 6.2, 0.6], text: "Think about: how do you spot a silent device?", color: orange, bold: true, italic: false}
""",
}


def _floats(s, n, what):
    try:
        v = [float(x) for x in s.replace('x', ',').split(',')]
        assert len(v) == n
        return v
    except Exception:
        raise SystemExit(f'--{what} needs {n} numbers, e.g. ' + ('13.333x7.5' if n == 2 else '0.5,1.4,12.3,5.5'))


def soffice_bin():
    for c in (os.environ.get('ARCHDIAGRAMS_SOFFICE'), 'soffice', 'libreoffice',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice', r'C:\Program Files\LibreOffice\program\soffice.exe'):
        if c and (shutil.which(c) or Path(c).exists()):
            return shutil.which(c) or c
    return None


def to_png(pptx: Path, dpi=110):
    """Render every slide to PNG next to the deck (needs LibreOffice; uses pdftoppm or PyMuPDF if present)."""
    sb = soffice_bin()
    if not sb:
        print('PNG export skipped: LibreOffice not found (set ARCHDIAGRAMS_SOFFICE to its soffice path).', file=sys.stderr)
        return []
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([sb, '--headless', f'-env:UserInstallation=file://{tmp}/profile', '--convert-to', 'pdf',
                        '--outdir', tmp, str(pptx)], check=True, capture_output=True, timeout=180)
        pdf = Path(tmp) / (pptx.stem + '.pdf')
        prefix = pptx.with_suffix('')
        if shutil.which('pdftoppm'):
            subprocess.run(['pdftoppm', '-png', '-r', str(dpi), str(pdf), str(prefix)], check=True)
            return sorted(prefix.parent.glob(prefix.name + '-*.png'))
        try:
            import fitz  # PyMuPDF
        except ImportError:
            shutil.copy(pdf, pptx.with_suffix('.pdf'))
            print('Saved a PDF instead (install poppler pdftoppm or PyMuPDF for PNGs).', file=sys.stderr)
            return [pptx.with_suffix('.pdf')]
        out = []
        for i, page in enumerate(fitz.open(pdf), 1):
            p = Path(f'{prefix}-{i}.png'); page.get_pixmap(dpi=dpi).save(p); out.append(p)
        return out


def icon_sheet(refs, out, cols=8, tint='0A2244'):
    from PIL import Image, ImageDraw, ImageFont
    import io
    cell, pad = 150, 10
    rows = max(1, -(-len(refs) // cols))
    sheet = Image.new('RGB', (cols * cell, rows * cell), 'white')
    dr = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype('DejaVuSans.ttf', 11)
    except OSError:
        font = ImageFont.load_default()
    for i, ref in enumerate(refs):
        r, c = divmod(i, cols)
        im = Image.open(io.BytesIO(_icons.png_bytes(ref, tint))).convert('RGBA').resize((72, 72))
        sheet.paste(im, (c * cell + (cell - 72) // 2, r * cell + pad), im)
        label = ref if len(ref) <= 24 else '…' + ref[-23:]
        w = dr.textlength(label, font=font)
        dr.text((c * cell + (cell - w) / 2, r * cell + 95), label, fill=(60, 60, 60), font=font)
    sheet.save(out)


def main(argv=None):
    ap = argparse.ArgumentParser(prog='archdiagrams', description='Editable architecture diagrams for PowerPoint.')
    ap.add_argument('--version', action='version', version=__version__)
    sub = ap.add_subparsers(dest='cmd', required=True)

    r = sub.add_parser('render', help='spec(s) -> a new deck, one slide per diagram')
    r.add_argument('specs', nargs='+'); r.add_argument('-o', '--out', required=True)
    r.add_argument('--theme'); r.add_argument('--canvas', default='13.333x7.5', help='slide size in inches, WxH')
    r.add_argument('--template', help='use this deck\'s masters/layouts (its slides are not copied)')
    r.add_argument('--layout', help='layout name to use from the template (default: the emptiest)')
    r.add_argument('--png', action='store_true', help='also export PNGs (needs LibreOffice)')

    a = sub.add_parser('add', help='draw a spec onto an existing deck')
    a.add_argument('spec'); a.add_argument('--deck', required=True); a.add_argument('-o', '--out')
    g = a.add_mutually_exclusive_group(); g.add_argument('--slide', type=int, help='draw on this slide (1-based)')
    g.add_argument('--layout', help='add a new slide from this layout (default: emptiest layout)')
    a.add_argument('--box', help='x,y,w,h in inches on the slide (default: below the title)')
    a.add_argument('--theme'); a.add_argument('--png', action='store_true')

    i = sub.add_parser('icons', help='search icons; --sheet makes a contact sheet PNG')
    i.add_argument('query', nargs='?', default=''); i.add_argument('--limit', type=int, default=200)
    i.add_argument('--sheet', help='write a PNG contact sheet of the matches')

    n = sub.add_parser('new', help='write a starter spec'); n.add_argument('out')
    n.add_argument('--kind', choices=list(STARTERS), default='tiers')

    sub.add_parser('themes', help='list built-in themes')
    args = ap.parse_args(argv)

    try:
        if args.cmd == 'render':
            specs = [s for p in args.specs for s in load(p)]
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter('always')
                new_deck(specs, args.out, args.theme, _floats(args.canvas, 2, 'canvas'), args.template, args.layout)
            for x in w: print(f'warning: {x.message}', file=sys.stderr)
            print(f'wrote {args.out} ({len(specs)} slide{"s" * (len(specs) != 1)})')
            if args.png:
                for p in to_png(Path(args.out)): print(f'wrote {p}')
        elif args.cmd == 'add':
            spec = load(args.spec)[0]
            box = _floats(args.box, 4, 'box') if args.box else None
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter('always')
                add_to_deck(spec, args.deck, args.out, args.slide, args.layout, box, args.theme)
            for x in w: print(f'warning: {x.message}', file=sys.stderr)
            print(f'wrote {args.out or args.deck}')
            if args.png:
                for p in to_png(Path(args.out or args.deck)): print(f'wrote {p}')
        elif args.cmd == 'icons':
            refs = _icons.search(args.query, args.limit)
            if _icons.provider_root() is None:
                print('(provider icons missing from this install)', file=sys.stderr)
            if args.sheet:
                icon_sheet(refs, args.sheet); print(f'wrote {args.sheet} ({len(refs)} icons)')
            else:
                print('\n'.join(refs) if refs else 'no match')
        elif args.cmd == 'new':
            Path(args.out).write_text(STARTERS[args.kind], encoding='utf-8'); print(f'wrote {args.out}')
        elif args.cmd == 'themes':
            for k, t in THEMES.items():
                print(f'{k:12} font={t.font:10} ' + ' '.join(f'{r}={c}' for r, c in t.colors.items() if r not in ('white', 'black')))
    except (DiagramError, FileNotFoundError) as e:
        raise SystemExit(f'error: {e}')


if __name__ == '__main__':
    main()
