"""Section header slide composer.

Builds a visually dense section header with dark background, section number badge,
title, key point text blocks with badges/numbers, divider lines, descriptive text,
and an image placeholder area.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import (
    add_accent_bar,
    add_badge_shape,
    add_image_placeholder,
)
from src.pptx_engine.text_operations import add_textbox, add_label


class SectionHeaderComposer(BaseComposer):
    """Compose a section header / divider slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        text_color = design.dark_slide_text

        # 1. Dark background
        self.set_background(slide, design.section_header_bg)

        # 2. Section number badge (top-left)
        section_num = self.infer_section_number(spec)
        add_badge_shape(
            slide,
            text=f"{section_num:02d}",
            left=ml,
            top=0.5,
            size=design.decoration.badge_size,
            fill_color=design.badge_fill_resolved,
            text_color=design.badge_text_resolved,
            font_name=design.emphasis_font_resolved,
            font_size=13,
        )

        # 3. Section title — large and bold
        title_text = spec.title or ""
        if title_text:
            add_textbox(
                slide,
                title_text,
                left=ml,
                top=1.1,
                width=w * 0.50,
                height=1.0,
                font_name=design.emphasis_font_resolved,
                font_size=32,
                font_color=text_color,
                alignment="left",
                line_spacing=design.paragraph.title_line_spacing,
            )

        # 4. Subtitle / description below title
        subtitle_text = spec.subtitle or ""
        body_text = self.get_body_text(spec)
        desc_text = subtitle_text or body_text
        if desc_text:
            # Truncate if very long
            if len(desc_text) > 300:
                desc_text = desc_text[:297] + "..."
            add_textbox(
                slide,
                desc_text,
                left=ml,
                top=2.2,
                width=w * 0.45,
                height=1.2,
                font_name=design.light_font_resolved,
                font_size=11,
                font_color=text_color,
                alignment="left",
                line_spacing=1.3,
            )

        # 5. Key points with numbered badges (right side)
        bullets = self.get_bullets(spec)
        if not bullets and body_text and not subtitle_text:
            # Split body into pseudo-bullets if no explicit bullets
            bullets = [s.strip() for s in body_text.split("\n") if s.strip()][:3]

        badge_x = w * 0.55
        text_x = badge_x + design.decoration.badge_size + 0.15
        text_w = w - text_x - design.content_area.margin_right

        for i, bullet in enumerate(bullets[:3]):
            y_pos = 1.1 + i * 1.1

            # Badge
            add_badge_shape(
                slide,
                text=f"{i + 1:02d}",
                left=badge_x,
                top=y_pos,
                size=design.decoration.badge_size,
                fill_color=design.badge_fill_resolved,
                text_color=design.badge_text_resolved,
                font_name=design.emphasis_font_resolved,
                font_size=11,
            )

            # Bullet text next to badge
            add_textbox(
                slide,
                bullet,
                left=text_x,
                top=y_pos - 0.05,
                width=text_w,
                height=0.9,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                alignment="left",
                line_spacing=1.2,
            )

            # Divider line below each point (except last)
            if i < min(len(bullets), 3) - 1:
                self.add_divider_line(
                    slide, design,
                    x1=badge_x,
                    y=y_pos + 0.95,
                    x2=w - design.content_area.margin_right,
                )

        # 6. Image placeholder (bottom-right corner if images suggested)
        if spec.image_suggestions:
            add_image_placeholder(
                slide,
                left=w * 0.55,
                top=h * 0.65,
                width=w * 0.40,
                height=h * 0.28,
                fill_color=design.image_placeholder_fill_resolved,
                corner_radius=0.1,
            )

        # 7. Accent bar at bottom
        self.add_accent_element(
            slide, design,
            left=ml,
            top=h - 0.45,
            width=w * 0.12,
        )

        # 8. Footer
        self.add_slide_footer(slide, design, color=text_color)
