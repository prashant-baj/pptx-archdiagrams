"""Icon lookup.

Icon references:

* ``ups`` or ``general/ups`` - bundled general pack (tintable masks, Tabler MIT / Material Apache-2.0)
* ``aws/compute/ec2``, ``azure/compute/vm``, ``gcp/...``, ``k8s/...``, ``onprem/...``, ``saas/...``,
  ``programming/...``, ``ibm/...``, ``oci/...``, ``alibabacloud/...``, ``generic/...`` - provider packs, bundled
  with the package (from the ``diagrams`` project, MIT). Kept in their brand colours. Set
  ``ARCHDIAGRAMS_PROVIDER_ICONS`` to a folder laid out the same way to use your own copy instead.
* a file path to any PNG / JPG - used as is.

Provider names are matched loosely: ``aws/ec2`` finds ``aws/compute/ec2``; ``azure/vm`` finds the Azure VM icon.
"""
from __future__ import annotations

import functools
import io
import os
from pathlib import Path

from PIL import Image

HERE = Path(__file__).parent
GENERAL = HERE / 'icons' / 'general'
PROVIDERS = ('aws', 'azure', 'gcp', 'k8s', 'onprem', 'saas', 'programming', 'ibm', 'oci', 'alibabacloud',
             'generic', 'digitalocean', 'openstack', 'firebase', 'elastic', 'outscale', 'custom', 'gis')


class IconNotFound(LookupError):
    pass


@functools.lru_cache(maxsize=1)
def provider_root() -> Path | None:
    """Folder that holds the provider packs (``aws/``, ``azure/``, ...): the bundled one, or
    ``ARCHDIAGRAMS_PROVIDER_ICONS`` if set."""
    env = os.environ.get('ARCHDIAGRAMS_PROVIDER_ICONS')
    if env and Path(env).is_dir():
        return Path(env)
    root = HERE / 'icons'
    return root if any((root / p).is_dir() for p in PROVIDERS) else None


@functools.lru_cache(maxsize=1)
def _provider_index() -> dict[str, Path]:
    root = provider_root()
    if not root:
        return {}
    return {p.relative_to(root).with_suffix('').as_posix(): p for p in root.rglob('*.png')
            if p.relative_to(root).parts[0] in PROVIDERS}


def general_names() -> list[str]:
    return sorted(p.stem for p in GENERAL.glob('*.png'))


def all_refs() -> list[str]:
    return [f'general/{n}' for n in general_names()] + sorted(_provider_index())


def search(query: str = '', limit: int = 200) -> list[str]:
    """Icons whose reference contains every part of the query. A trailing '/' means "starts with":
    ``general/`` lists the general pack, ``aws/database/`` one AWS category."""
    q = query.lower().replace(' ', '-')
    if q.endswith('/'):
        return [r for r in all_refs() if r.startswith(q)][:limit]
    q = q.strip('/')
    return [r for r in all_refs() if all(part in r for part in q.split('/'))][:limit]


def resolve(ref: str) -> tuple[Path, bool]:
    """Return (path, tintable)."""
    p = Path(ref).expanduser()
    if p.suffix.lower() in ('.png', '.jpg', '.jpeg') and p.exists():
        return p, False
    name = ref.strip().strip('/')
    if name.startswith('general/'):
        name = name[len('general/'):]
    if '/' not in name:
        g = GENERAL / f'{name}.png'
        if g.exists():
            return g, True
        raise IconNotFound(_hint(ref))
    prov = name.split('/')[0]
    if prov not in PROVIDERS:
        raise IconNotFound(_hint(ref))
    idx = _provider_index()
    if not idx:
        raise IconNotFound(f'{ref!r}: provider icons are missing from this install (reinstall archdiagrams, '
                           'or set ARCHDIAGRAMS_PROVIDER_ICONS)')
    if name in idx:
        return idx[name], False
    leaf = name.split('/')[-1].replace('_', '-')
    hits = [k for k in idx if k.startswith(prov + '/') and k.split('/')[-1].replace('_', '-') == leaf]
    if not hits:
        hits = [k for k in idx if k.startswith(prov + '/') and leaf in k.split('/')[-1].replace('_', '-')]
    if hits:
        return idx[sorted(hits, key=len)[0]], False
    raise IconNotFound(_hint(ref))


def _hint(ref: str) -> str:
    leaf = ref.split('/')[-1]
    near = search(leaf, 8)
    return f'icon {ref!r} not found' + (f'; similar: {", ".join(near)}' if near else '; try: archdiagrams icons <word>')


@functools.lru_cache(maxsize=512)
def png_bytes(ref: str, tint: str | None) -> bytes:
    """PNG for an icon reference; general icons are recoloured to ``tint`` (hex)."""
    path, tintable = resolve(ref)
    if not (tintable and tint):
        return path.read_bytes()
    im = Image.open(path).convert('RGBA')
    r, g, b = (int(tint[i:i + 2], 16) for i in (0, 2, 4))
    alpha = im.getchannel('A')
    out = Image.new('RGBA', im.size, (r, g, b, 0))
    out.putalpha(alpha)
    buf = io.BytesIO()
    out.save(buf, 'PNG')
    return buf.getvalue()
