"""Two-column / comparison slide composer.

Builds a two-column layout with card-style backgrounds, column titles,
body text, vertical divider, and section marker.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_card_background
from src.pptx_engine.text_operations import add_textbox, add_label


class TwoColumnComposer(BaseComposer):
    """Compose a two-column or comparison slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        heading_color = design.dark_slide_text
        text_color = design.text_secondary_resolved

        # 1. Background
        self.set_background(slide, design.content_bg)

        # 2. Title above columns
        title_text = spec.title or ""
        if title_text:
            add_textbox(
                slide,
                title_text,
                left=ml,
                top=0.45,
                width=w - ml - mr,
                height=0.6,
                font_name=design.emphasis_font_resolved,
                font_size=24,
                font_color=heading_color,
                alignment="left",
            )

        # 3. Divider line below title
        self.add_divider_line(slide, design, x1=ml, y=1.15, x2=w - mr)

        # 4. Split content into two columns
        blocks = spec.content_blocks
        mid = max(1, len(blocks) // 2)
        left_blocks = blocks[:mid]
        right_blocks = blocks[mid:]

        left_text = "\n\n".join(b.content for b in left_blocks)
        right_text = "\n\n".join(b.content for b in right_blocks)

        # Column geometry
        col_top = 1.35
        col_height = h - col_top - 0.65
        col_gap = 0.3
        usable_w = w - ml - mr - col_gap
        col_w = usable_w / 2.0
        left_x = ml
        right_x = ml + col_w + col_gap

        # 5. Card backgrounds
        add_card_background(
            slide,
            left=left_x,
            top=col_top,
            width=col_w,
            height=col_height,
            fill_color=design.surface_resolved,
            corner_radius=0.12,
        )
        add_card_background(
            slide,
            left=right_x,
            top=col_top,
            width=col_w,
            height=col_height,
            fill_color=design.surface_resolved,
            corner_radius=0.12,
        )

        # 6. Column content — extract column titles from first block if possible
        left_title = ""
        right_title = ""
        if left_blocks and left_blocks[0].type == "title":
            left_title = left_blocks[0].content
            left_text = "\n\n".join(b.content for b in left_blocks[1:])
        if right_blocks and right_blocks[0].type == "title":
            right_title = right_blocks[0].content
            right_text = "\n\n".join(b.content for b in right_blocks[1:])

        pad = 0.2  # Inner padding for cards

        # Left column title
        if left_title:
            add_textbox(
                slide, left_title,
                left=left_x + pad, top=col_top + 0.15,
                width=col_w - 2 * pad, height=0.4,
                font_name=design.medium_font_resolved,
                font_size=14,
                font_color=heading_color,
                bold=True,
            )
            left_body_top = col_top + 0.6
        else:
            left_body_top = col_top + 0.2

        # Left column body
        if left_text:
            add_textbox(
                slide, left_text,
                left=left_x + pad, top=left_body_top,
                width=col_w - 2 * pad,
                height=col_height - (left_body_top - col_top) - pad,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                line_spacing=design.paragraph.body_line_spacing,
            )

        # Right column title
        if right_title:
            add_textbox(
                slide, right_title,
                left=right_x + pad, top=col_top + 0.15,
                width=col_w - 2 * pad, height=0.4,
                font_name=design.medium_font_resolved,
                font_size=14,
                font_color=heading_color,
                bold=True,
            )
            right_body_top = col_top + 0.6
        else:
            right_body_top = col_top + 0.2

        # Right column body
        if right_text:
            add_textbox(
                slide, right_text,
                left=right_x + pad, top=right_body_top,
                width=col_w - 2 * pad,
                height=col_height - (right_body_top - col_top) - pad,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                line_spacing=design.paragraph.body_line_spacing,
            )

        # 7. Vertical divider between columns
        self.add_vertical_divider(
            slide, design,
            x=ml + col_w + col_gap / 2,
            y1=col_top + 0.3,
            y2=col_top + col_height - 0.3,
        )

        # 8. Section marker
        section_num = self.infer_section_number(spec)
        self.add_section_marker(
            slide, design,
            section_number=section_num,
            section_label=spec.title or "Comparison",
            color=heading_color,
        )

        # 9. Accent bar
        self.add_accent_element(slide, design, left=ml, top=h - 0.55, width=w * 0.10)
