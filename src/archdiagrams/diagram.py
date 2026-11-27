"""Draw architecture diagrams as native, editable PowerPoint shapes.

A diagram is drawn in its own coordinate space (inches, origin top-left, ``size`` = (w, h)) and mapped onto a box
on the slide, so the same diagram fits any slide size. Fonts, lines and icons scale with it.

Elements have ids. Arrows connect anchors such as ``"ups.r"`` (right edge middle) or ``"gw.t@0.25"`` (top edge,
a quarter of the way along) and route themselves with right angles unless told otherwise.

    from pptx import Presentation
    from archdiagrams import Diagram

    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    d = Diagram(slide, size=(12, 5), box=(0.6, 1.4, 12.1, 5.5), theme='navy')
    d.zone('site', 0, 0, 3, 5, 'Site', color='navy')
    d.card('ups', 0.2, 0.6, 2.6, 0.9, 'ups', 'UPS', 'SNMP')
    d.card('gw', 4, 0.6, 2.6, 0.9, 'gateway', 'Gateway', 'MQTT', color='teal')
    d.arrow('ups.r', 'gw.l', label='events')
    prs.save('out.pptx')
"""
from __future__ import annotations

import io
import math
import re
import warnings
from dataclasses import dataclass

from lxml import etree
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from . import icons as _icons
from .theme import Theme, load_theme

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
ALIGN = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
ANCHOR = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE, 'bottom': MSO_ANCHOR.BOTTOM}


class DiagramError(ValueError):
    pass


@dataclass
class Box:
    x: float
    y: float
    w: float
    h: float

    def point(self, side: str, frac: float = 0.5) -> tuple[float, float]:
        x, y, w, h = self.x, self.y, self.w, self.h
        return {'l': (x, y + h * frac), 'r': (x + w, y + h * frac), 't': (x + w * frac, y), 'b': (x + w * frac, y + h),
                'c': (x + w / 2, y + h / 2), 'tl': (x, y), 'tr': (x + w, y), 'bl': (x, y + h), 'br': (x + w, y + h)}[side]


def _text_width(text: str, size_pt: float) -> float:
    """Rough width in inches of a line of text (average glyph ~0.55 em, bold-ish)."""
    return len(text) * size_pt * 0.53 / 72


class Diagram:
    def __init__(self, slide, size=(12.0, 5.5), box=None, theme: str | dict | Theme | None = 'default',
                 group: bool = True, name: str = 'Diagram', fit: str = 'contain', warn: bool = True):
        """slide: a python-pptx slide. size: the diagram's own (w, h) in inches.
        box: (x, y, w, h) on the slide to draw into; default = the whole slide with 0.5 in margins.
        fit: 'contain' keeps the aspect ratio (centred in box), 'stretch' fills the box."""
        self.slide = slide
        self.theme = theme if isinstance(theme, Theme) else load_theme(theme)
        self.size = tuple(size)
        if box is None:
            prs_part = slide.part.package.presentation_part.presentation
            sw, sh = prs_part.slide_width / 914400, prs_part.slide_height / 914400
            box = (0.5, 0.5, sw - 1.0, sh - 1.0)
        bx, by, bw, bh = box
        sx, sy = bw / self.size[0], bh / self.size[1]
        if fit == 'stretch':
            self.sx, self.sy = sx, sy
            self.ox, self.oy = bx, by
        else:
            s = min(sx, sy)
            self.sx = self.sy = s
            self.ox = bx + (bw - self.size[0] * s) / 2
            self.oy = by + (bh - self.size[1] * s) / 2
        self.k = min(self.sx, self.sy)              # scale for fonts and line widths
        self.boxes: dict[str, Box] = {}
        self.warn = warn
        self.warnings: list[str] = []
        self.shapes = slide.shapes.add_group_shape().shapes if group else slide.shapes
        if group:
            self._group = slide.shapes[-1]
            self._group.name = name
        self._n = 0

    # ------------------------------------------------------------------ helpers
    def _X(self, x): return Inches(self.ox + x * self.sx)
    def _Y(self, y): return Inches(self.oy + y * self.sy)
    def _W(self, w): return Inches(w * self.sx)
    def _H(self, h): return Inches(h * self.sy)
    def _pt(self, p): return Pt(max(6.0, p * self.k))          # font sizes: never below 6 pt
    def _lw(self, p): return Pt(max(0.5, p * self.k))          # line widths

    def _name(self, kind, id_):
        self._n += 1
        return f'{kind} {id_ or self._n}'

    def _register(self, id_, x, y, w, h):
        if id_:
            if id_ in self.boxes:
                raise DiagramError(f'duplicate id {id_!r}')
            self.boxes[id_] = Box(x, y, w, h)

    def _warn(self, msg):
        self.warnings.append(msg)
        if self.warn:
            warnings.warn(msg, stacklevel=3)

    def _rrect(self, x, y, w, h, fill, line, width=1.0, dash=False, radius=0.08, name='Shape', shape=MSO_SHAPE.ROUNDED_RECTANGLE):
        s = self.shapes.add_shape(shape, self._X(x), self._Y(y), self._W(w), self._H(h))
        s.name = name
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            s.adjustments[0] = radius
        if fill is None:
            s.fill.background()
        else:
            s.fill.solid(); s.fill.fore_color.rgb = self.theme.rgb(fill)
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = self.theme.rgb(line); s.line.width = self._lw(width)
            if dash:
                s.line.dash_style = MSO_LINE.DASH
        s.shadow.inherit = False
        return s

    def _text(self, x, y, w, h, runs, anchor='top', align='left', name='Text', wrap=True):
        """runs: list of paragraphs; each paragraph = dict(text, size, bold, italic, color)."""
        tb = self.shapes.add_textbox(self._X(x), self._Y(y), self._W(w), self._H(h))
        tb.name = name
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = ANCHOR[anchor]
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        for i, r in enumerate(runs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = ALIGN[align]
            if r.get('space'):
                p.space_before = self._pt(r['space'])
            run = p.add_run(); run.text = r['text']
            f = run.font
            f.name = self.theme.font; f.size = self._pt(r.get('size', 11))
            f.bold = r.get('bold', False); f.italic = r.get('italic', False)
            f.color.rgb = self.theme.rgb(r.get('color', self.theme.text))
        return tb

    def _check_fit(self, label, text, size, avail_w, avail_h, max_lines=None):
        if not text or avail_w <= 0:
            return
        words, lines, cur = text.split(), 1, ''
        for wd in words:
            trial = (cur + ' ' + wd).strip()
            if _text_width(trial, size) > avail_w and cur:
                lines += 1; cur = wd
            else:
                cur = trial
            if _text_width(wd, size) > avail_w:
                self._warn(f'{label}: the word {wd!r} is wider than its box')
        need = lines * size * 1.25 / 72
        if (max_lines and lines > max_lines) or need > avail_h + 0.02:
            self._warn(f'{label}: text may overflow ({lines} lines in {avail_h:.2f} in): {text!r}')
        return lines

    # ------------------------------------------------------------------ elements
    def zone(self, id, x, y, w, h, label='', color='navy', dashed=True, fill=None, cards=None, cols=1,
             card_h=None, gap=0.2, pad=0.18, top=0.5, align='top', card_defaults=None, max_card_h=1.1):
        """Dashed rounded container with a caps label. Optional ``cards``: a list of card dicts laid out
        in a grid inside the zone (``cols`` columns, rows as needed). Cards are at most ``max_card_h`` tall
        unless ``card_h`` is given; ``align`` (top, middle, bottom, spread) places them vertically."""
        self._rrect(x, y, w, h, fill or f'pale-{color}', color, width=1.5, dash=dashed, radius=0.03,
                    name=self._name('Zone', id))
        if label:
            self._text(x + 0.2, y + 0.1, w - 0.4, 0.3,
                       [dict(text=label.upper(), size=self.theme.sizes['zone'], bold=True, color=color)],
                       name=self._name('Zone label', id))
        self._register(id, x, y, w, h)
        if cards:
            cols = max(1, int(cols))
            rows = math.ceil(len(cards) / cols)
            iw, ih = w - 2 * pad, h - top - pad
            cw = (iw - gap * (cols - 1)) / cols
            ch = card_h or min(max_card_h, (ih - gap * (rows - 1)) / rows)
            used = rows * ch + (rows - 1) * gap
            y0 = y + top + {'top': 0, 'middle': (ih - used) / 2, 'bottom': ih - used}.get(align, 0)
            if align == 'spread' and rows > 1:
                gap_v = (ih - rows * ch) / (rows - 1)
                y0 = y + top
            else:
                gap_v = gap
            for i, c in enumerate(cards):
                r, k = divmod(i, cols)
                kw = dict(card_defaults or {}); kw.update(c)
                kw.setdefault('color', kw.get('color'))
                self.card(kw.pop('id', None), x + pad + k * (cw + gap), y0 + r * (ch + gap_v), cw, ch, **kw)

    def card(self, id, x, y, w, h, icon=None, title='', sub=None, color=None, icon_color=None, emphasis=False,
             sub_color=None, title_size=None, sub_size=None, fill='white', align='left'):
        """White rounded card: icon left, bold title, small coloured line (protocol, role, detail).
        color: border role (default neutral). emphasis: thicker border."""
        t = self.theme
        self._rrect(x, y, w, h, fill, color or t.line, width=2.0 if emphasis else 1.0, radius=0.12,
                    name=self._name('Card', id or title))
        tx = x + 0.15
        if icon:
            isz = max(0.25, min(0.5, h - 0.3, w * 0.28))
            tint = t.hex(icon_color or color or t.text)
            try:
                data = _icons.png_bytes(icon, tint)
            except _icons.IconNotFound as e:
                raise DiagramError(str(e)) from None
            pic = self.shapes.add_picture(io.BytesIO(data), self._X(x + 0.15), self._Y(y + (h - isz) / 2),
                                          self._W(isz), self._H(isz))
            pic.name = self._name('Icon', id or title)
            tx = x + 0.15 + isz + 0.15
        ts, ss = title_size or t.sizes['title'], sub_size or t.sizes['sub']
        runs = [dict(text=title, size=ts, bold=True, color=t.text)]
        if sub:
            runs.append(dict(text=sub, size=ss, bold=True, color=sub_color or t.sub))
        avail_w = x + w - tx - 0.1
        self._text(tx, y + 0.06, avail_w, h - 0.12, runs, anchor='middle', align=align, name=self._name('Card text', id or title))
        # overflow estimate (in slide inches: text sizes scale with k)
        lt = self._check_fit(f'card {id or title!r} title', title, ts, avail_w, h - 0.12) or 1
        ls = self._check_fit(f'card {id or title!r} sub', sub, ss, avail_w, h - 0.12) or 0
        if (lt * ts + ls * ss) * 1.22 / 72 > h - 0.1:
            self._warn(f'card {id or title!r}: title + sub may not fit the card height ({h:.2f} in)')
        self._register(id, x, y, w, h)

    def bar(self, id, x, y, w, h, text, color='navy', text_color='white', size=None, align='left'):
        """Solid rounded band with one line of bold text (a principle, a rule, a shared concern)."""
        self._rrect(x, y, w, h, color, color, radius=0.12, name=self._name('Bar', id))
        self._text(x + 0.25, y, w - 0.5, h, [dict(text=text, size=size or self.theme.sizes['bar'], bold=True, color=text_color)],
                   anchor='middle', align=align, name=self._name('Bar text', id))
        self._register(id, x, y, w, h)

    def note(self, x, y, w, text, color=None, size=None, bold=False, italic=True, h=0.32, align='left', id=None):
        self._text(x, y, w, h, [dict(text=text, size=size or self.theme.sizes['note'], bold=bold, italic=italic,
                                     color=color or self.theme.muted)], align=align, name=self._name('Note', id))
        self._register(id, x, y, w, h)

    def badge(self, on, n, corner='tr', color='orange', d=0.38):
        """Numbered circle on an element's corner (or at an (x, y) point)."""
        if isinstance(on, str):
            px, py = self._box(on).point(corner)
        else:
            px, py = on
        s = self._rrect(px - d / 2, py - d / 2, d, d, color, 'white', width=1.5, name=f'Step {n}', shape=MSO_SHAPE.OVAL)
        tf = s.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = str(n)
        r.font.name = self.theme.font; r.font.size = self._pt(self.theme.sizes['badge']); r.font.bold = True
        r.font.color.rgb = self.theme.rgb('white')

    # ------------------------------------------------------------------ arrows
    def _box(self, id_):
        if id_ not in self.boxes:
            raise DiagramError(f'unknown id {id_!r}; known: {", ".join(self.boxes) or "(none)"}')
        return self.boxes[id_]

    def anchor(self, ref) -> tuple[tuple[float, float], str]:
        """'id.side[@frac]' -> ((x, y), side). Sides: l r t b c tl tr bl br. (x, y) tuples pass through."""
        if isinstance(ref, (list, tuple)):
            return (float(ref[0]), float(ref[1])), 'c'
        m = re.fullmatch(r'(.+?)\.(l|r|t|b|c|tl|tr|bl|br)(?:@([0-9.]+))?', ref)
        if not m:
            box = self._box(ref)
            return box.point('c'), 'c'
        id_, side, frac = m.group(1), m.group(2), float(m.group(3)) if m.group(3) else 0.5
        return self._box(id_).point(side, frac), side

    def route(self, src, dst, route='auto', via=None):
        (x1, y1), s1 = self.anchor(src)
        (x2, y2), s2 = self.anchor(dst)
        if via:
            return [(x1, y1)] + [tuple(v) for v in via] + [(x2, y2)]
        if route == 'straight' or abs(x1 - x2) < 0.01 or abs(y1 - y2) < 0.01:
            return [(x1, y1), (x2, y2)]
        h1, h2 = s1 in ('l', 'r'), s2 in ('l', 'r')
        if route == 'auto':
            route = 'hvh' if h1 and h2 else 'vhv' if not h1 and not h2 and s1 != 'c' else 'hv' if h1 else 'vh'
        if route == 'hvh':
            mx = (x1 + x2) / 2
            return [(x1, y1), (mx, y1), (mx, y2), (x2, y2)]
        if route == 'vhv':
            my = (y1 + y2) / 2
            return [(x1, y1), (x1, my), (x2, my), (x2, y2)]
        if route == 'hv':
            return [(x1, y1), (x2, y1), (x2, y2)]
        if route == 'vh':
            return [(x1, y1), (x1, y2), (x2, y2)]
        raise DiagramError(f'unknown route {route!r} (auto, straight, hvh, vhv, hv, vh)')

    def arrow(self, src, dst, color='grey', dashed=False, both=False, head=True, width=1.75, label=None,
              label_at=None, label_w=None, label_color=None, route='auto', via=None):
        """Arrow from src to dst. src/dst: 'id.side[@frac]', 'id', or (x, y).
        route: auto | straight | hvh | vhv | hv | vh; via: explicit waypoints [(x, y), ...]."""
        pts = self.route(src, dst, route, via)
        self.polyline(pts, color, dashed, both, head, width)
        if label:
            self._label(pts, label, label_at, label_w, label_color or self.theme.muted)

    def polyline(self, pts, color='grey', dashed=False, both=False, head=True, width=1.75):
        segs = [(a, b) for a, b in zip(pts[:-1], pts[1:]) if a != b]
        for i, ((x1, y1), (x2, y2)) in enumerate(segs):
            c = self.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, self._X(x1), self._Y(y1), self._X(x2), self._Y(y2))
            c.name = self._name('Arrow', None)
            c.line.color.rgb = self.theme.rgb(color); c.line.width = self._lw(width)
            if dashed:
                c.line.dash_style = MSO_LINE.DASH
            ln = c.line._get_or_add_ln()
            if head and i == len(segs) - 1:
                etree.SubElement(ln, A + 'tailEnd', type='triangle', w='med', len='med')
            if both and i == 0:
                etree.SubElement(ln, A + 'headEnd', type='triangle', w='med', len='med')

    def _label(self, pts, text, at, w, color):
        size = self.theme.sizes['label']
        w = w or _text_width(text, size) * 0.9 + 0.15
        if at:
            x, y = at
        else:
            segs = list(zip(pts[:-1], pts[1:]))
            seglen = lambda s: abs(s[0][0] - s[1][0]) + abs(s[0][1] - s[1][1])
            horiz = [s for s in segs if abs(s[0][1] - s[1][1]) < 0.01 and seglen(s) >= 0.45]
            # prefer a horizontal run (label sits above it); otherwise the longest segment
            (x1, y1), (x2, y2) = max(horiz, key=seglen) if horiz else max(segs, key=seglen)
            room = seglen(((x1, y1), (x2, y2)))
            if abs(y1 - y2) < 0.01 and room + 0.3 < _text_width(text, size) * 0.9:
                self._warn(f'arrow label {text!r} has little room ({room:.2f} in); widen the gap or set label_at')
            if abs(y1 - y2) < 0.01:
                w = min(w, max(seglen(((x1, y1), (x2, y2))) + 0.6, 0.8)) if not at else w
            if abs(y1 - y2) < 0.01:                      # horizontal: centred above the line
                x, y = (x1 + x2) / 2 - w / 2, y1 - 0.3
            else:                                         # vertical or diagonal: to the right of the midpoint
                x, y = (x1 + x2) / 2 + 0.08, (y1 + y2) / 2 - 0.15
        self._text(x, y, w, 0.3, [dict(text=text, size=size, italic=True, color=color)],
                   align='center', name=self._name('Arrow label', None), wrap=True)

    def bus(self, sources, dst, color='grey', side='r', gap=0.25, width=1.25, label=None):
        """Many sources into one target through a shared vertical (side r/l) or horizontal (side t/b) line."""
        pts = [self.anchor(f'{s}.{side}')[0] for s in sources]
        (tx, ty), tside = self.anchor(dst if '.' in dst else f'{dst}.{ {"r": "l", "l": "r", "t": "b", "b": "t"}[side] }')
        if side in ('r', 'l'):
            bx = (max(p[0] for p in pts) + gap) if side == 'r' else (min(p[0] for p in pts) - gap)
            for x, y in pts:
                self.polyline([(x, y), (bx, y)], color, head=False, width=width)
            ys = [p[1] for p in pts] + [ty]
            self.polyline([(bx, min(ys)), (bx, max(ys))], color, head=False, width=width)
            self.arrow((bx, ty), (tx, ty), color=color, width=width, label=label, route='straight')
        else:
            by = (max(p[1] for p in pts) + gap) if side == 'b' else (min(p[1] for p in pts) - gap)
            for x, y in pts:
                self.polyline([(x, y), (x, by)], color, head=False, width=width)
            xs = [p[0] for p in pts] + [tx]
            self.polyline([(min(xs), by), (max(xs), by)], color, head=False, width=width)
            self.arrow((tx, by), (tx, ty), color=color, width=width, label=label, route='straight')

    # ------------------------------------------------------------------ checks
    def check_overlaps(self, kinds=None):
        """Warn when two cards overlap (zones and bars may contain cards)."""
        ids = list(self.boxes)
        for i, a in enumerate(ids):
            for b in ids[i + 1:]:
                A_, B_ = self.boxes[a], self.boxes[b]
                inside = lambda P, Q: P.x >= Q.x - 1e-6 and P.y >= Q.y - 1e-6 and P.x + P.w <= Q.x + Q.w + 1e-6 and P.y + P.h <= Q.y + Q.h + 1e-6
                if inside(A_, B_) or inside(B_, A_):
                    continue
                if A_.x < B_.x + B_.w - 0.01 and B_.x < A_.x + A_.w - 0.01 and A_.y < B_.y + B_.h - 0.01 and B_.y < A_.y + A_.h - 0.01:
                    self._warn(f'{a!r} and {b!r} overlap')
        return self.warnings
