"""Tests for Pydantic schema models."""

from pathlib import Path


from src.schemas.slide_schema import (
    ContentBlock,
    ContentInventory,
    ContentSection,
    DeckSchema,
    SlideSpec,
    SlideType,
)
from src.schemas.template_schema import PlaceholderInfo, TemplateRegistry, TemplateSlide
from src.schemas.design_system import ColorConfig, DesignSystem, FontConfig


class TestSlideSchema:
    def test_slide_type_enum(self):
        assert SlideType.TITLE == "title"
        assert SlideType.BULLET_LIST == "bullet_list"

    def test_content_block_defaults(self):
        block = ContentBlock(type="body", content="Hello")
        assert block.emphasis == "normal"

    def test_slide_spec_minimal(self):
        spec = SlideSpec(
            slide_number=1,
            slide_type=SlideType.TITLE,
            intent="Open the presentation",
        )
        assert spec.title is None
        assert spec.content_blocks == []
        assert spec.layout_hints == []

    def test_slide_spec_full(self):
        spec = SlideSpec(
            slide_number=2,
            slide_type=SlideType.BULLET_LIST,
            intent="Show key points",
            title="Key Points",
            subtitle="What matters",
            content_blocks=[
                ContentBlock(type="bullets", content="Point 1\nPoint 2\nPoint 3")
            ],
            speaker_notes="Talk about each point",
            layout_hints=["minimal_text"],
        )
        assert spec.slide_number == 2
        assert len(spec.content_blocks) == 1

    def test_deck_schema(self):
        schema = DeckSchema(
            title="Test Deck",
            target_audience="Engineers",
            key_message="Testing works",
            slides=[
                SlideSpec(
                    slide_number=1,
                    slide_type=SlideType.TITLE,
                    intent="Open",
                    title="Test Deck",
                )
            ],
        )
        assert len(schema.slides) == 1

    def test_content_inventory(self):
        inv = ContentInventory(
            main_topic="Testing",
            themes=["quality", "automation"],
            sections=[
                ContentSection(
                    heading="Unit Tests",
                    content="Write unit tests",
                    bullet_points=["Fast", "Isolated"],
                    importance="high",
                )
            ],
            key_data_points=["95% coverage"],
            quotes=["Testing is doubting"],
            summary="Testing is important",
        )
        assert inv.main_topic == "Testing"
        assert len(inv.sections) == 1

    def test_schema_json_roundtrip(self):
        schema = DeckSchema(
            title="Test",
            target_audience="All",
            key_message="Hello",
            slides=[
                SlideSpec(
                    slide_number=1,
                    slide_type=SlideType.TITLE,
                    intent="Open",
                )
            ],
        )
        json_str = schema.model_dump_json()
        restored = DeckSchema.model_validate_json(json_str)
        assert restored.title == schema.title
        assert len(restored.slides) == 1


class TestTemplateSchema:
    def test_placeholder_info(self):
        ph = PlaceholderInfo(
            name="Title 1",
            type="TITLE",
            position=(1.0, 0.5, 10.0, 1.5),
        )
        assert ph.position[2] == 10.0

    def test_template_slide(self):
        ts = TemplateSlide(
            template_file="test.pptx",
            slide_index=0,
            slide_type=SlideType.TITLE,
            tags=["corporate", "clean"],
            complexity=2,
        )
        assert ts.shape_count == 0
        assert len(ts.tags) == 2

    def test_registry_find_by_type(self):
        reg = TemplateRegistry(
            templates=[
                TemplateSlide(
                    template_file="a.pptx",
                    slide_index=0,
                    slide_type=SlideType.TITLE,
                ),
                TemplateSlide(
                    template_file="a.pptx",
                    slide_index=1,
                    slide_type=SlideType.CONTENT,
                ),
                TemplateSlide(
                    template_file="a.pptx",
                    slide_index=2,
                    slide_type=SlideType.TITLE,
                ),
            ]
        )
        titles = reg.find_by_type(SlideType.TITLE)
        assert len(titles) == 2

    def test_registry_find_by_tags(self):
        reg = TemplateRegistry(
            templates=[
                TemplateSlide(
                    template_file="a.pptx",
                    slide_index=0,
                    slide_type=SlideType.TITLE,
                    tags=["corporate", "clean"],
                ),
                TemplateSlide(
                    template_file="a.pptx",
                    slide_index=1,
                    slide_type=SlideType.CONTENT,
                    tags=["creative", "bold"],
                ),
            ]
        )
        results = reg.find_by_tags(["corporate"])
        assert len(results) == 1
        assert results[0].slide_index == 0

    def test_registry_save_load(self, tmp_path):
        reg = TemplateRegistry(
            templates=[
                TemplateSlide(
                    template_file="test.pptx",
                    slide_index=0,
                    slide_type=SlideType.TITLE,
                    tags=["test"],
                )
            ],
            source_files=["test.pptx"],
        )
        path = tmp_path / "registry.json"
        reg.save(path)
        loaded = TemplateRegistry.load(path)
        assert len(loaded.templates) == 1
        assert loaded.templates[0].tags == ["test"]


class TestDesignSystem:
    def test_defaults(self):
        ds = DesignSystem()
        assert ds.name == "Default"
        assert ds.fonts.title_font == "Arial"
        assert ds.colors.primary == "#1a73e8"

    def test_custom_config(self):
        ds = DesignSystem(
            name="Custom",
            fonts=FontConfig(title_font="Helvetica", title_size=48),
            colors=ColorConfig(primary="#ff0000"),
        )
        assert ds.fonts.title_font == "Helvetica"
        assert ds.fonts.body_font == "Arial"  # Default preserved
        assert ds.colors.primary == "#ff0000"

    def test_yaml_roundtrip(self, tmp_path):
        ds = DesignSystem(name="Test Brand")
        path = tmp_path / "test.yaml"
        ds.to_yaml(path)
        loaded = DesignSystem.from_yaml(path)
        assert loaded.name == "Test Brand"

    def test_load_default_yaml(self):
        path = Path(__file__).parent.parent / "design_systems" / "default.yaml"
        ds = DesignSystem.from_yaml(path)
        assert ds.name == "Default"
        assert ds.fonts.title_size == 44
