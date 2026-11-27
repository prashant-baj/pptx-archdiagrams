"""Themes: a palette of named roles, a font, and size defaults.

Every colour in a spec is a role name (``navy``, ``teal``, ``orange``...) or a hex string (``"0A7C59"``).
Pale fills for zones are derived automatically, so a theme only lists the strong colours.
"""
from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path

from pptx.dml.color import RGBColor


def _hex(c: str) -> str:
    c = c.strip().lstrip('#').upper()
    if len(c) != 6 or any(ch not in '0123456789ABCDEF' for ch in c):
        raise ValueError(f'not a 6-digit hex colour: {c!r}')
    return c


def mix(c: str, other: str = 'FFFFFF', t: float = 0.92) -> str:
    """Blend colour c towards other by t (0 = c, 1 = other)."""
    a, b = _hex(c), _hex(other)
    out = [round(int(a[i:i + 2], 16) * (1 - t) + int(b[i:i + 2], 16) * t) for i in (0, 2, 4)]
    return ''.join(f'{v:02X}' for v in out)


@dataclass
class Theme:
    name: str
    font: str
    colors: dict[str, str]                 # role -> hex
    text: str = 'navy'                     # role for card titles
    sub: str = 'teal'                      # role for the small line under a title
    line: str = 'C9D2DE'                   # neutral card border
    muted: str = 'grey'                    # arrow labels, notes
    pale: float = 0.93                     # how pale zone fills are
    sizes: dict[str, float] = field(default_factory=lambda: dict(
        title=13, sub=10.5, zone=11, label=10, note=11, badge=12, bar=14, slide_title=28, eyebrow=12, caption=14))

    def rgb(self, c: str | None) -> RGBColor:
        return RGBColor.from_string(self.hex(c))

    def hex(self, c: str | None) -> str:
        if c is None:
            c = self.line
        if c in self.colors:
            return _hex(self.colors[c])
        if c.startswith('pale-'):
            return mix(self.hex(c[5:]), t=self.pale)
        return _hex(c)

    def with_overrides(self, font: str | None = None, colors: dict | None = None) -> 'Theme':
        return replace(self, font=font or self.font, colors={**self.colors, **(colors or {})})


_BASE = dict(navy='0A2244', blue='2F6FD0', teal='1F9E8F', green='2E8B57', orange='ED5315', red='D94848',
             purple='6B4FBB', amber='C98A0B', grey='5B6472', white='FFFFFF', black='111111')

THEMES: dict[str, Theme] = {
    # Deep navy with an orange accent, Inter font
    'navy': Theme('navy', 'Inter', dict(_BASE)),
    # Portable default: fonts every Office install has
    'default': Theme('default', 'Calibri', dict(_BASE, navy='1F2A44', orange='E8590C', teal='0F766E', blue='1D4ED8')),
    # AWS-like palette
    'aws': Theme('aws', 'Arial', dict(_BASE, navy='232F3E', orange='FF9900', teal='01A88D', blue='527FFF', red='DD344C',
                                      purple='8C4FFF', green='7AA116'), sub='blue'),
    # Azure-like palette
    'azure': Theme('azure', 'Segoe UI', dict(_BASE, navy='243A5E', blue='0078D4', teal='008575', orange='D83B01',
                                             purple='5C2D91', green='107C10'), sub='blue'),
    # Calm monochrome for documents
    'mono': Theme('mono', 'Arial', dict(_BASE, navy='222222', blue='444444', teal='555555', orange='C0392B', red='C0392B',
                                        green='555555', purple='555555', amber='555555'), sub='grey'),
}


def load_theme(spec: str | dict | None) -> Theme:
    """A theme name, a dict ``{base, font, colors}``, or a path to a YAML/JSON file with the same keys."""
    if spec is None:
        return THEMES['default']
    if isinstance(spec, Theme):
        return spec
    if isinstance(spec, str):
        if spec in THEMES:
            return THEMES[spec]
        p = Path(spec)
        if p.exists():
            import yaml
            spec = yaml.safe_load(p.read_text(encoding='utf-8'))
        else:
            raise ValueError(f'unknown theme {spec!r}; built-in themes: {", ".join(THEMES)}')
    base = THEMES[spec.get('base', 'default')]
    t = base.with_overrides(spec.get('font'), spec.get('colors'))
    for k in ('text', 'sub', 'line', 'muted', 'pale'):
        if k in spec:
            t = replace(t, **{k: spec[k]})
    if 'sizes' in spec:
        t = replace(t, sizes={**t.sizes, **spec['sizes']})
    return t
