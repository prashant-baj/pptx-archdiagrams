"""YAML / JSON diagram specs -> slides.

A spec file holds one diagram, or several under ``slides:``::

    theme: navy                 # optional: built-in name, a theme file, or {base, font, colors}
    size: [14, 5.5]              # the diagram's own canvas in inches; elements use these coordinates
    slide:                       # optional text around the diagram
      eyebrow: PART 4 · FLOW
      title: One UPS event, end to end
      caption: 1 UPS on battery → 2 gateway → 3 history
      notes: Speaker notes ...
    elements:
      - zone:  {id: site, at: [0, 0, 3, 5.5], label: Site, color: navy, cards: [...], cols: 1}
      - card:  {id: gw, at: [3.4, 0.5, 2.4, 1], icon: gateway, title: Site gateway, sub: store and forward, color: teal}
      - arrow: {from: ups.r, to: gw.l, color: teal, label: events}
      - bus:   {from: [brk, ups, pdu], to: gw, color: navy}
      - badge: {on: ups, n: 1}
      - bar:   {at: [0, 6, 14, 0.7], text: One data model on every link}
      - note:  {at: [0, 5.2, 6], text: Customers see each site separately}
"""
from __future__ import annotations

import copy
import json
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from .diagram import Diagram, DiagramError
from .theme import load_theme

KINDS = ('zone', 'card', 'bar', 'note', 'arrow', 'bus', 'badge', 'badges')


def load(path) -> list[dict]:
    """Read a spec file; always returns a list of diagram specs."""
    text = Path(path).read_text(encoding='utf-8')
    if str(path).lower().endswith('.json'):
        data = json.loads(text)
    else:
        import yaml
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as e:
            raise DiagramError(f'{path}: YAML error - {e}. Tip: quote text that contains ? : # or , inside {{...}}') from None
    if not isinstance(data, dict):
        raise DiagramError(f'{path}: expected a mapping at the top level')
    if 'slides' in data:
        shared = {k: v for k, v in data.items() if k != 'slides'}
        return [{**shared, **s} for s in data['slides']]
    return [data]


def _at(e, n, kind):
    at = e.get('at')
    if not isinstance(at, (list, tuple)) or len(at) < n:
        raise DiagramError(f'{kind} {e.get("id", "")!r}: "at" must be [x, y, w, h]' if n == 4 else
                           f'{kind}: "at" must be [x, y, w] or [x, y, w, h]')
    return [float(v) for v in at]


def draw(d: Diagram, elements: list):
    """Draw spec elements onto a Diagram."""
    for raw in elements or []:
        if not isinstance(raw, dict):
            raise DiagramError(f'element must be a mapping, got {raw!r}')
        if 'type' in raw:
            kind, e = raw['type'], {k: v for k, v in raw.items() if k != 'type'}
        elif len(raw) == 1 and next(iter(raw)) in KINDS:
            kind, e = next(iter(raw.items()))
        else:
            raise DiagramError(f'cannot tell the element type of {raw!r}; use one of {", ".join(KINDS)}')
        e = copy.deepcopy(e) if isinstance(e, dict) else e
        if kind == 'zone':
            x, y, w, h = _at(e, 4, 'zone')
            e.pop('at')
            d.zone(e.pop('id', None), x, y, w, h, **e)
        elif kind == 'card':
            x, y, w, h = _at(e, 4, 'card')
            e.pop('at')
            d.card(e.pop('id', None), x, y, w, h, **e)
        elif kind == 'bar':
            x, y, w, h = _at(e, 4, 'bar')
            e.pop('at')
            d.bar(e.pop('id', None), x, y, w, h, e.pop('text'), **e)
        elif kind == 'note':
            at = _at(e, 3, 'note')
            e.pop('at')
            if len(at) == 4:
                e['h'] = at[3]
            d.note(at[0], at[1], at[2], e.pop('text'), **e)
        elif kind == 'arrow':
            src, dst = e.pop('from'), e.pop('to')
            d.arrow(src, dst, **e)
        elif kind == 'bus':
            d.bus(e.pop('from'), e.pop('to'), **e)
        elif kind == 'badge':
            d.badge(e.pop('on'), e.pop('n'), **e)
        elif kind == 'badges':
            for item in e:
                on, n = item[0], item[1]
                d.badge(on, n, *(item[2:]))
        else:
            raise DiagramError(f'unknown element type {kind!r}')
    return d


# ---------------------------------------------------------------------- slides
def _slide_size(prs):
    return prs.slide_width / 914400, prs.slide_height / 914400


def _header(slide, spec, theme, W):
    """Eyebrow + title text boxes for decks without a title placeholder. Returns the y where content can start."""
    s = spec.get('slide') or {}
    y = 0.4
    if s.get('eyebrow'):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(W - 1), Inches(0.35)); tb.name = 'Eyebrow'
        tb.text_frame.word_wrap = True
        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        r = tb.text_frame.paragraphs[0].add_run(); r.text = s['eyebrow'].upper()
        r.font.name = theme.font; r.font.size = Pt(theme.sizes['eyebrow']); r.font.bold = True; r.font.color.rgb = theme.rgb('orange')
        y += 0.4
    if s.get('title'):
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(W - 1), Inches(0.7)); tb.name = 'Title'
        tb.text_frame.word_wrap = True
        r = tb.text_frame.paragraphs[0].add_run(); r.text = s['title']
        r.font.name = theme.font; r.font.size = Pt(theme.sizes['slide_title']); r.font.bold = True; r.font.color.rgb = theme.rgb('navy')
        y += 0.85
    return y + 0.1 if y > 0.4 else 0.5


def _caption(slide, spec, theme, W, H):
    s = spec.get('slide') or {}
    if not s.get('caption'):
        return 0.5
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(H - 0.75), Inches(W - 1), Inches(0.45)); tb.name = 'Caption'
    tb.text_frame.word_wrap = True
    r = tb.text_frame.paragraphs[0].add_run(); r.text = s['caption']
    r.font.name = theme.font; r.font.size = Pt(theme.sizes['caption']); r.font.bold = True; r.font.color.rgb = theme.rgb('orange')
    return 0.85


def _notes(slide, spec):
    n = (spec.get('slide') or {}).get('notes')
    if n:
        slide.notes_slide.notes_text_frame.text = n


def _layout(prs, name):
    if name is None:
        return min(prs.slide_layouts, key=lambda l: len(l.placeholders))
    for l in prs.slide_layouts:
        if l.name.lower() == str(name).lower():
            return l
    raise DiagramError(f'layout {name!r} not found; layouts: {", ".join(l.name for l in prs.slide_layouts)}')


def _title_bottom(slide):
    t = slide.shapes.title
    if t is not None and t.has_text_frame:
        return (t.top + t.height) / 914400
    return None


def render_spec(slide, spec, box, theme=None):
    th = load_theme(theme or spec.get('theme'))
    size = spec.get('size') or [12, 5.5]
    d = Diagram(slide, size=size, box=box, theme=th, name=(spec.get('slide') or {}).get('title') or 'Diagram')
    draw(d, spec.get('elements'))
    d.check_overlaps()
    return d


def new_deck(specs, out, theme=None, canvas=(13.333, 7.5), template=None, layout=None):
    """One slide per spec in a new deck (or a copy of ``template`` with its slides removed)."""
    prs = Presentation(template) if template else Presentation()
    if template:
        lst = prs.slides._sldIdLst
        for sid in list(lst):
            prs.part.drop_rel(sid.rId); lst.remove(sid)
    else:
        prs.slide_width, prs.slide_height = Inches(canvas[0]), Inches(canvas[1])
    W, H = _slide_size(prs)
    diags = []
    for spec in specs:
        th = load_theme(theme or spec.get('theme'))
        slide = prs.slides.add_slide(_layout(prs, layout))
        tb = _title_bottom(slide)
        if tb is not None:
            slide.shapes.title.text = (spec.get('slide') or {}).get('title', '')
            top = tb + 0.25
        else:
            top = _header(slide, spec, th, W)
        bottom = _caption(slide, spec, th, W, H)
        diags.append(render_spec(slide, spec, (0.5, top, W - 1.0, H - top - bottom), th))
        _notes(slide, spec)
    prs.save(out)
    return diags


def add_to_deck(spec, deck, out=None, slide_no=None, layout=None, box=None, theme=None):
    """Draw onto slide ``slide_no`` (1-based) of an existing deck, or add a new slide from ``layout``."""
    prs = Presentation(deck)
    W, H = _slide_size(prs)
    th = load_theme(theme or spec.get('theme'))
    if slide_no:
        if not 1 <= slide_no <= len(prs.slides):
            raise DiagramError(f'slide {slide_no} out of range (deck has {len(prs.slides)})')
        slide = prs.slides[slide_no - 1]
    else:
        slide = prs.slides.add_slide(_layout(prs, layout))
        if slide.shapes.title is not None:
            slide.shapes.title.text = (spec.get('slide') or {}).get('title', '')
        _notes(slide, spec)
    if box is None:
        tb = _title_bottom(slide)
        top = (tb + 0.25) if tb is not None else 0.5
        box = (0.5, top, W - 1.0, H - top - 0.6)
    d = render_spec(slide, spec, box, th)
    prs.save(out or deck)
    return d
