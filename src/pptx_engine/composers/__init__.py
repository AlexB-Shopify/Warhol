"""Slide composer registry.

Maps SlideType enum values to their corresponding composer classes.
Use get_composer() to look up the right composer for a slide type.
"""

import logging

from src.schemas.slide_schema import SlideType

from .base import BaseComposer
from .title import TitleComposer
from .section_header import SectionHeaderComposer
from .content import ContentComposer
from .two_column import TwoColumnComposer
from .quote import QuoteComposer
from .data_point import DataPointComposer
from .bullet_list import BulletListComposer
from .closing import ClosingComposer

logger = logging.getLogger(__name__)

# Singleton instances — composers are stateless, so one instance each is fine.
_title = TitleComposer()
_section = SectionHeaderComposer()
_content = ContentComposer()
_two_col = TwoColumnComposer()
_quote = QuoteComposer()
_data_point = DataPointComposer()
_bullet = BulletListComposer()
_closing = ClosingComposer()

COMPOSERS: dict[SlideType, BaseComposer] = {
    SlideType.TITLE: _title,
    SlideType.SECTION_HEADER: _section,
    SlideType.CONTENT: _content,
    SlideType.TWO_COLUMN: _two_col,
    SlideType.COMPARISON: _two_col,
    SlideType.QUOTE: _quote,
    SlideType.BULLET_LIST: _bullet,
    SlideType.CLOSING: _closing,
    SlideType.DATA_POINT: _data_point,
    SlideType.IMAGE_FULL: _content,       # Fallback to content with image area
    SlideType.IMAGE_WITH_TEXT: _content,   # Fallback to content with image area
    SlideType.CHART: _content,            # Fallback to content
    SlideType.TIMELINE: _bullet,          # Timeline can use numbered bullet layout
    SlideType.TEAM: _content,             # Team slides fallback to content
}


def get_composer(slide_type: SlideType) -> BaseComposer:
    """Look up the composer for a given slide type.

    Falls back to ContentComposer if no specific composer is registered.
    """
    composer = COMPOSERS.get(slide_type)
    if composer is None:
        logger.debug(f"No composer for {slide_type}, using ContentComposer")
        return _content
    return composer


__all__ = [
    "BaseComposer",
    "TitleComposer",
    "SectionHeaderComposer",
    "ContentComposer",
    "TwoColumnComposer",
    "QuoteComposer",
    "DataPointComposer",
    "BulletListComposer",
    "ClosingComposer",
    "get_composer",
    "COMPOSERS",
]
