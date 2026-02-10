"""Bullet list slide composer.

Builds a styled bullet list with numbered badge shapes alongside each point,
title with divider line, and section marker.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_badge_shape, add_image_placeholder
from src.pptx_engine.text_operations import add_textbox, add_label


class BulletListComposer(BaseComposer):
    """Compose a bullet list slide with numbered badges."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        heading_color = design.dark_slide_text
        text_color = design.text_secondary_resolved

        has_image = bool(spec.image_suggestions)
        content_right = w * 0.52 if has_image else w - mr

        # 1. Background
        self.set_background(slide, design.bullet_list_bg)

        # 2. Title
        title_text = spec.title or ""
        if title_text:
            add_textbox(
                slide,
                title_text,
                left=ml,
                top=0.5,
                width=content_right - ml,
                height=0.6,
                font_name=design.emphasis_font_resolved,
                font_size=24,
                font_color=heading_color,
                alignment="left",
            )

        # 3. Divider line below title
        self.add_divider_line(slide, design, x1=ml, y=1.2, x2=content_right)

        # 4. Bullet items with numbered badges
        bullets = self.get_bullets(spec)
        if not bullets:
            # Fallback: split body text into lines
            body_text = self.get_body_text(spec)
            if body_text:
                bullets = [s.strip() for s in body_text.split("\n") if s.strip()]

        badge_size = design.decoration.badge_size
        badge_x = ml
        text_x = badge_x + badge_size + 0.2
        text_w = content_right - text_x - 0.1

        # Calculate vertical spacing based on number of items
        start_y = 1.45
        max_items = min(len(bullets), 6)
        if max_items > 0:
            available_h = h - start_y - 0.65
            item_spacing = min(available_h / max_items, 0.85)
        else:
            item_spacing = 0.85

        for i, bullet in enumerate(bullets[:max_items]):
            y_pos = start_y + i * item_spacing

            # Badge
            add_badge_shape(
                slide,
                text=f"{i + 1:02d}",
                left=badge_x,
                top=y_pos,
                size=badge_size,
                fill_color=design.badge_fill_resolved,
                text_color=design.badge_text_resolved,
                font_name=design.emphasis_font_resolved,
                font_size=11,
            )

            # Bullet text
            add_textbox(
                slide,
                bullet,
                left=text_x,
                top=y_pos - 0.02,
                width=text_w,
                height=item_spacing - 0.08,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                alignment="left",
                line_spacing=1.15,
            )

        # 5. Image placeholder (right side if images suggested)
        if has_image:
            add_image_placeholder(
                slide,
                left=w * 0.56,
                top=0.5,
                width=w * 0.40,
                height=h * 0.70,
                fill_color=design.image_placeholder_fill_resolved,
                corner_radius=0.1,
            )

        # 6. Section marker
        section_num = self.infer_section_number(spec)
        self.add_section_marker(
            slide, design,
            section_number=section_num,
            section_label=spec.title or "Key Points",
            color=heading_color,
        )

        # 7. Accent bar
        self.add_accent_element(slide, design, left=ml, top=h - 0.55, width=w * 0.10)
