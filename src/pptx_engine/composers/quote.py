"""Quote slide composer.

Builds a visually rich quote slide with dark background, large decorative
quotation mark, quote text, attribution, accent elements, and optional
image placeholder for speaker photo.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_accent_bar, add_image_placeholder
from src.pptx_engine.text_operations import add_textbox, add_label


class QuoteComposer(BaseComposer):
    """Compose a quote slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        text_color = design.dark_slide_text

        # 1. Dark background
        self.set_background(slide, design.quote_bg)

        # 2. Large decorative quotation mark
        add_textbox(
            slide,
            "\u201C",
            left=ml + 0.3,
            top=0.3,
            width=1.5,
            height=1.5,
            font_name=design.extra_light_font_resolved,
            font_size=120,
            font_color=design.accent_bar_color_resolved,
            alignment="left",
        )

        # 3. Quote text
        quote_text = self.get_quote_text(spec)
        if quote_text:
            add_textbox(
                slide,
                quote_text,
                left=ml + 0.8,
                top=1.4,
                width=w - ml - mr - 1.6,
                height=h * 0.40,
                font_name=design.quote_font_resolved,
                font_size=design.quote_size_resolved,
                font_color=text_color,
                italic=True,
                alignment="left",
                line_spacing=1.1,
            )

        # 4. Attribution text (from body/caption blocks or subtitle)
        body_text = self.get_body_text(spec)
        attribution = spec.subtitle or body_text
        if attribution:
            add_textbox(
                slide,
                f"\u2014 {attribution}",
                left=ml + 0.8,
                top=h * 0.68,
                width=w - ml - mr - 1.6,
                height=0.5,
                font_name=design.medium_font_resolved,
                font_size=design.caption_size_resolved,
                font_color=text_color,
                alignment="left",
            )

        # 5. Image placeholder for speaker photo (right side if images suggested)
        if spec.image_suggestions:
            add_image_placeholder(
                slide,
                left=w - mr - 1.5,
                top=h * 0.55,
                width=1.2,
                height=1.2,
                fill_color=design.image_placeholder_fill_resolved,
                corner_radius=0.6,  # Circle-ish for headshot
            )

        # 6. Accent bar
        self.add_accent_element(
            slide, design,
            left=ml + 0.8,
            top=h * 0.62,
            width=1.5,
        )

        # 7. Footer
        self.add_slide_footer(slide, design, color=text_color)
