"""Closing slide composer.

Builds a closing slide with dark background, thank-you message,
contact info / next steps, accent bar, and brand elements.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_accent_bar
from src.pptx_engine.text_operations import add_textbox, add_label


class ClosingComposer(BaseComposer):
    """Compose a closing / thank-you slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        text_color = design.dark_slide_text

        # 1. Dark background
        self.set_background(slide, design.closing_bg)

        # 2. Main closing message — large ExtraLight
        title_text = spec.title or "Thank you"
        add_textbox(
            slide,
            title_text,
            left=ml,
            top=h * 0.20,
            width=w * 0.6,
            height=h * 0.35,
            font_name=design.extra_light_font_resolved,
            font_size=design.fonts.title_size,
            font_color=text_color,
            alignment="left",
            line_spacing=design.paragraph.title_line_spacing,
        )

        # 3. Subtitle / next steps
        subtitle_text = spec.subtitle or ""
        body_text = self.get_body_text(spec)
        secondary_text = subtitle_text or body_text

        if secondary_text:
            add_textbox(
                slide,
                secondary_text,
                left=ml,
                top=h * 0.58,
                width=w * 0.55,
                height=h * 0.22,
                font_name=design.light_font_resolved,
                font_size=design.fonts.body_size,
                font_color=text_color,
                alignment="left",
                line_spacing=1.3,
            )

        # 4. Accent bar
        self.add_accent_element(
            slide, design,
            left=ml,
            top=h * 0.55,
            width=w * 0.15,
        )

        # 5. Footer
        self.add_slide_footer(slide, design, color=text_color)
