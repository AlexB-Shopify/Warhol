"""Data point / hero number slide composer.

Builds a slide centered around a large hero number with ExtraLight font,
accent color treatment, decorative shapes, and supporting context text.
"""

from src.schemas.design_system import DesignSystem
from src.schemas.slide_schema import SlideSpec
from src.pptx_engine.composers.base import BaseComposer
from src.pptx_engine.shape_operations import add_accent_bar, add_rectangle
from src.pptx_engine.text_operations import add_textbox, add_hero_number, add_label


class DataPointComposer(BaseComposer):
    """Compose a data point / hero number slide."""

    def compose(self, slide, spec: SlideSpec, design: DesignSystem) -> None:
        w, h = self.get_dims(slide)
        ml = design.content_area.margin_left
        mr = design.content_area.margin_right
        heading_color = design.dark_slide_text
        text_color = design.text_secondary_resolved

        # 1. Background
        self.set_background(slide, design.content_bg)

        # 2. Extract data point and context
        number_text, context_text = self.get_data_point(spec)
        if not number_text:
            number_text = spec.title or ""

        # 3. Accent background shape behind the number
        add_rectangle(
            slide,
            left=ml - 0.1,
            top=0.5,
            width=w * 0.55,
            height=h * 0.55,
            fill_color=design.surface_accent_resolved,
            corner_radius=0.15,
        )

        # 4. Hero number — large ExtraLight with accent color
        add_hero_number(
            slide,
            number_text,
            left=ml + 0.3,
            top=0.8,
            width=w * 0.5,
            height=h * 0.42,
            font_name=design.extra_light_font_resolved,
            font_size=design.hero_number_size_resolved,
            font_color=design.data_point_accent,
            alignment="left",
        )

        # 5. Title / label above or below the number
        title_text = spec.title or ""
        if title_text and title_text != number_text:
            add_textbox(
                slide,
                title_text,
                left=ml + 0.3,
                top=h * 0.60,
                width=w * 0.5,
                height=0.5,
                font_name=design.emphasis_font_resolved,
                font_size=18,
                font_color=heading_color,
                alignment="left",
            )

        # 6. Context text (right side)
        if not context_text:
            context_text = self.get_body_text(spec)
        if context_text:
            add_textbox(
                slide,
                context_text,
                left=w * 0.58,
                top=1.0,
                width=w * 0.38,
                height=h * 0.50,
                font_name=design.fonts.body_font,
                font_size=design.fonts.body_size,
                font_color=text_color,
                alignment="left",
                line_spacing=1.3,
            )

        # 7. Accent bar below the context
        self.add_accent_element(
            slide, design,
            left=w * 0.58,
            top=h * 0.72,
            width=1.5,
        )

        # 8. Section marker
        section_num = self.infer_section_number(spec)
        self.add_section_marker(
            slide, design,
            section_number=section_num,
            section_label=title_text or "Data",
            color=heading_color,
        )
