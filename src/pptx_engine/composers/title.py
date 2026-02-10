"""Title slide composer.

Builds a visually rich title slide with dark background, large display title,
subtitle/date label, accent bar, image placeholder area, and footer.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_accent_bar, add_image_placeholder
from src.pptx_engine.text_operations import add_textbox, add_label


class TitleComposer(BaseComposer):
    """Compose a title/opening slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left

        # 1. Dark background
        self.set_background(slide, design.title_bg)

        # 2. Image placeholder area (right half)
        if spec.image_suggestions:
            add_image_placeholder(
                slide,
                left=w * 0.52,
                top=0.4,
                width=w * 0.45,
                height=h * 0.7,
                fill_color=design.image_placeholder_fill_resolved,
                corner_radius=0.12,
            )

        # 3. Title text — large ExtraLight in the left portion
        title_text = spec.title or ""
        if title_text:
            add_textbox(
                slide,
                title_text,
                left=ml,
                top=h * 0.22,
                width=w * 0.48 if spec.image_suggestions else w * 0.75,
                height=h * 0.45,
                font_name=design.extra_light_font_resolved,
                font_size=design.fonts.title_size,
                font_color=design.dark_slide_text,
                alignment="left",
                line_spacing=design.paragraph.title_line_spacing,
            )

        # 4. Subtitle / date label below title
        subtitle_text = spec.subtitle or ""
        if subtitle_text:
            add_label(
                slide,
                subtitle_text,
                left=ml,
                top=h * 0.72,
                width=w * 0.4,
                height=0.35,
                font_name=design.medium_font_resolved,
                font_size=design.fonts.subtitle_size // 2,
                font_color=design.dark_slide_text,
                alignment="left",
            )

        # 5. Accent bar near bottom
        self.add_accent_element(
            slide, design,
            left=ml,
            top=h - 0.55,
            width=w * 0.15,
        )

        # 6. Footer
        self.add_slide_footer(slide, design, color=design.dark_slide_text)
