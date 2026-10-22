"""編集可能な PowerPoint を作るための公開 API。"""

from .layout import align, distribute
from .model import Deck, Emphasis, Math, Rect, Text, key_id
from .tables import Table

__version__ = "0.2.0"
__all__ = ["Deck", "Rect", "Text", "Emphasis", "Math", "Table", "align", "distribute", "key_id", "__version__"]
