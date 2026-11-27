"""archdiagrams - editable architecture diagrams for PowerPoint, drawn as native shapes."""
__version__ = '0.1.0'

from .diagram import Box, Diagram, DiagramError  # noqa: E402
from .theme import THEMES, Theme, load_theme  # noqa: E402
from .spec import add_to_deck, draw, load, new_deck, render_spec  # noqa: E402

__all__ = ['Diagram', 'DiagramError', 'Box', 'Theme', 'THEMES', 'load_theme', 'load', 'draw', 'render_spec',
           'new_deck', 'add_to_deck', '__version__']
