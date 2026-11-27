"""Copy the provider icons (AWS, Azure, GCP, Kubernetes, ...) into the package, resized for slides.

Source: the icon files of the `diagrams` project (https://github.com/mingrammer/diagrams, MIT).
    pip install diagrams
    python tools/build_provider_icons.py            # or: python tools/build_provider_icons.py <resources folder>
Each icon is scaled to fit 128 px (about 280 dpi at the size it is drawn on a card) and saved as an optimised
256-colour PNG, which keeps the whole set near 7 MB.
"""
import shutil
import sys
from pathlib import Path

from PIL import Image

SIDE = 128
OUT = Path(__file__).resolve().parent.parent / 'src' / 'archdiagrams' / 'icons'


def source() -> Path:
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    import diagrams
    return Path(diagrams.__file__).resolve().parent.parent / 'resources'


def main():
    src = source()
    n = 0
    for prov in sorted(p for p in src.iterdir() if p.is_dir()):
        dst = OUT / prov.name
        if dst.exists():
            shutil.rmtree(dst)
        for p in prov.rglob('*.png'):
            out = dst / p.relative_to(prov)
            out.parent.mkdir(parents=True, exist_ok=True)
            im = Image.open(p).convert('RGBA')
            im.thumbnail((SIDE, SIDE), Image.LANCZOS)
            im.quantize(256, method=Image.Quantize.FASTOCTREE).save(out, 'PNG', optimize=True)
            n += 1
    print(f'{n} provider icons written under {OUT}')


if __name__ == '__main__':
    main()
