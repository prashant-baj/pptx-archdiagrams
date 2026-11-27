import glob
from pathlib import Path

import pytest
from pptx import Presentation

from archdiagrams import Diagram, DiagramError, add_to_deck, load, load_theme, new_deck
from archdiagrams import icons

ROOT = Path(__file__).resolve().parent.parent
EXAMPLES = sorted(glob.glob(str(ROOT / 'examples' / '*.yaml')))


def _shapes(slide):
    out = []
    for s in slide.shapes:
        out.append(s)
        if s.shape_type == 6:
            out.extend(s.shapes)
    return out


@pytest.mark.parametrize('spec', EXAMPLES, ids=lambda p: Path(p).stem)
def test_examples_render(spec, tmp_path):
    specs = load(spec)
    out = tmp_path / 'out.pptx'
    new_deck(specs, out)
    prs = Presentation(out)
    assert len(prs.slides) == len(specs)
    assert len(_shapes(prs.slides[0])) > 10


def test_general_icons_present():
    names = icons.general_names()
    assert len(names) >= 140 and {'ups', 'breaker', 'gateway', 'database', 'user'} <= set(names)
    assert icons.resolve('ups')[1] is True          # tintable
    assert icons.png_bytes('ups', 'ED5315')[:4] == b'\x89PNG'


def test_provider_icons_bundled():
    assert icons.provider_root() is not None
    refs = icons.search('', 5000)
    for prov, least in (('aws/', 500), ('azure/', 800), ('gcp/', 100), ('k8s/', 40)):
        assert sum(r.startswith(prov) for r in refs) >= least
    path, tintable = icons.resolve('aws/ec2')
    assert path.name == 'ec2.png' and tintable is False
    assert icons.resolve('azure/iot-hub')[0].exists()


def test_unknown_icon_and_id(tmp_path):
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    d = Diagram(slide, size=(10, 4), warn=False)
    with pytest.raises(DiagramError, match='not found'):
        d.card('x', 0, 0, 2, 1, 'no-such-icon', 'X')
    d.card('a', 0, 0, 2, 1, 'user', 'A')
    with pytest.raises(DiagramError, match='unknown id'):
        d.arrow('a.r', 'b.l')
    with pytest.raises(DiagramError, match='duplicate id'):
        d.card('a', 3, 0, 2, 1, None, 'A2')


def test_routes_and_overlap_warning():
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    d = Diagram(slide, size=(10, 4), warn=False)
    d.card('a', 0, 0, 2, 1, None, 'A'); d.card('b', 5, 2, 2, 1, None, 'B'); d.card('c', 1, 0.5, 2, 1, None, 'C')
    assert d.route('a.r', 'b.l') == [(2, 0.5), (3.5, 0.5), (3.5, 2.5), (5, 2.5)]
    assert d.route('a.b', 'b.t', 'vhv')[1] == (1, 1.5)
    assert any('overlap' in w for w in d.check_overlaps())


def test_add_to_existing_deck(tmp_path):
    base = tmp_path / 'base.pptx'
    prs = Presentation(); prs.slides.add_slide(prs.slide_layouts[5]).shapes.title.text = 'Existing'; prs.save(base)
    spec = {'size': [8, 3], 'elements': [{'card': {'id': 'a', 'at': [0, 0, 2, 1], 'icon': 'server', 'title': 'Server'}}]}
    add_to_deck(spec, base, tmp_path / 'one.pptx', slide_no=1)
    add_to_deck(spec, base, tmp_path / 'two.pptx', layout='Title Only')
    assert len(Presentation(tmp_path / 'one.pptx').slides) == 1
    assert len(Presentation(tmp_path / 'two.pptx').slides) == 2


def test_theme_overrides():
    t = load_theme({'base': 'aws', 'font': 'Arial', 'colors': {'brand': '00AA88'}})
    assert t.hex('brand') == '00AA88' and t.hex('orange') == 'FF9900' and t.hex('pale-brand') != '00AA88'
