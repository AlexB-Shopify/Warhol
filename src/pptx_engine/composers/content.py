"""Content slide composer.

Builds a standard content slide with title, divider line, body text,
optional bullets, section marker, accent elements, and image placeholder.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_accent_bar, add_image_placeholder
from src.pptx_engine.text_operations import add_textbox, add_bullet_list, add_label


class ContentComposer(BaseComposer):
    """Compose a standard content slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        text_color = design.text_secondary_resolved
        heading_color = design.dark_slide_text

        has_image = bool(spec.image_suggestions)
        content_width = (w * 0.52) if has_image else (w - ml - mr)

        # 1. Background
        self.set_background(slide, design.content_bg)

        # 2. Title
        title_text = spec.title or ""
        title_top = 0.55
        if title_text:
            add_textbox(
                slide,
                title_text,
                left=ml,
                top=title_top,
                width=content_width,
                height=0.65,
                font_name=design.emphasis_font_resolved,
                font_size=24,
                font_color=heading_color,
                alignment="left",
                line_spacing=design.paragraph.title_line_spacing,
            )

        # 3. Divider line below title
        divider_y = title_top + 0.75
        self.add_divider_line(
            slide, design,
            x1=ml,
            y=divider_y,
            x2=ml + content_width,
        )

        # 4. Body text or bullets
        body_top = divider_y + 0.25
        body_height = h - body_top - 0.7

        bullets = self.get_bullets(spec)
        body_text = self.get_body_text(spec)

        if bullets:
            add_bullet_list(
                slide,
                items=bullets,
                left=ml + 0.1,
                top=body_top,
                width=content_width - 0.2,
                height=body_height,
                font_name=design.fonts.body_font,
                font_size=design.fonts.bullet_size,
                font_color=text_color,
                line_spacing=design.paragraph.bullet_line_spacing or 1.2,
            )
        elif body_text:
            add_textbox(
                slide,
                body_text,
                left=ml,
                top=body_top,
                width=content_width,
                height=body_height,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                alignment=design.paragraph.body_alignment,
                line_spacing=design.paragraph.body_line_spacing,
            )

        # 5. Image placeholder (right side if images suggested)
        if has_image:
            add_image_placeholder(
                slide,
                left=w * 0.56,
                top=0.55,
                width=w * 0.40,
                height=h * 0.70,
                fill_color=design.image_placeholder_fill_resolved,
                corner_radius=0.1,
            )

        # 6. Section marker (bottom-left)
        section_num = self.infer_section_number(spec)
        self.add_section_marker(
            slide, design,
            section_number=section_num,
            section_label=spec.title or "Content",
            color=heading_color,
        )

        # 7. Accent bar
        self.add_accent_element(
            slide, design,
            left=ml,
            top=h - 0.55,
            width=w * 0.10,
        )
