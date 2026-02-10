# Warhol

A Cursor-native PowerPoint presentation generator. Drop in a document, tell the Cursor agent what you want, and it builds a polished PPTX -- no API keys or external LLM services required.

## What is "Cursor-Native"?

This project uses [Cursor](https://cursor.sh/) (an AI-powered code editor) as its intelligence layer. The Cursor agent reads your documents, plans the narrative, writes slide content, and orchestrates the build. Python scripts handle the deterministic work: parsing files, constructing PPTX slides, and validating data.

There are no API keys to configure. The AI capabilities come from Cursor itself, guided by a set of workflow rules in `.cursor/rules/` that teach the agent how to build presentations.

## Quickstart

### 1. Clone and install

```bash
git clone <repo-url> && cd slide-builder-2
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

### 2. Add a template (optional but recommended)

Drop a branded `.pptx` file into `templates/` to use as a base. The builder will use its layouts, masters, and backgrounds. Without a template, slides are built on a blank canvas.

### 3. Open in Cursor and generate

Open the project in Cursor, then ask the agent in chat:

> Generate a presentation from `inputs/my-document.pdf`

The agent walks through the full pipeline -- parsing, content extraction, narrative planning, and PPTX construction -- and produces a finished deck. The output lands in `output/`.

You can guide the result with extra instructions:

> Generate a 10-slide presentation from `inputs/quarterly-review.pdf` using a corporate tone. Focus on the financial metrics and growth story.

> Build a presentation from `inputs/notes.md`. Keep it under 8 slides and make it punchy.

That's it. The agent handles the rest.

## How the Pipeline Works

The generation pipeline has six stages. The agent runs each one in order, writing intermediate files to `workspace/` so you can inspect or edit at any point.

```
Input Document
     |
     v
1. Parse .............. Extract text from PDF / DOCX / PPTX / TXT / MD
     |
     v
2. Assess ............. Evaluate content maturity (Level 1-4)
     |
     v
3. Develop ............ Create / refine slide content (adapted to maturity)
     |
     v
4. Architect .......... Plan slide-by-slide narrative and layout types
     |
     v
5. Build .............. Construct the actual PPTX file
     |
     v
6. Review ............. Programmatic quality checks + agent assessment
```

### Content Maturity (Adaptive Pipeline)

Not all inputs are equal. A polished draft needs less work than a list of raw ideas. The agent assesses input maturity and adapts the pipeline:

| Level | Input looks like... | Pipeline stages |
|-------|-------------------|-----------------|
| **1 -- Raw Ideas** | Bullet points, brainstorm notes, a topic sentence | Research -> Content Development -> Editor -> Design |
| **2 -- Outline** | Structured sections with key points, but no slide content | Content Development -> Editor -> Design |
| **3 -- Draft** | Slide-like content that needs tightening and polish | Editor -> Design |
| **4 -- Ready** | Presentation-ready content, just needs layout and build | Design only |

The agent reports the assessed level and you can override it ("this is just rough ideas, flesh it out").

## Project Structure

```
.cursor/rules/          -- Agent workflow rules (the "brain")
scripts/                -- Python scripts the agent calls
src/                    -- Python library (parsers, PPTX engine, schemas)
workspace/              -- Intermediate pipeline files (auto-generated)
design_systems/         -- YAML design system configs (fonts, colors)
templates/              -- PPTX template files (you provide these)
inputs/                 -- Input documents (you provide these)
output/                 -- Generated presentations land here
```

### Key Files

| File | Purpose |
|------|---------|
| `workspace/parsed_content.txt` | Raw text extracted from your input document |
| `workspace/content_inventory.json` | Structured content the agent extracted |
| `workspace/deck_schema.json` | Slide-by-slide plan (the blueprint for the deck) |
| `workspace/template_matches.json` | Which template slide maps to which deck slide |
| `workspace/quality_report.json` | Quality check results |
| `design_systems/*.yaml` | Font, color, and sizing configuration |
| `template_registry.json` | Metadata about available template slides (generated) |

## Setup

### Prerequisites

- **Python 3.11+**
- **[Cursor](https://cursor.sh/)** IDE (this is required -- the agent IS the LLM)

### Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

For development (linting + tests):

```bash
pip install -e ".[dev]"
```

### Provide Templates

Templates are `.pptx` files that supply branded layouts, masters, and backgrounds. The repo ships without templates (they're gitignored since they're large binary files). To use one:

1. Drop your `.pptx` file into `templates/`
2. Ask Cursor: *"Analyze the templates in `templates/`"*
3. The agent builds a `template_registry.json` that maps slide types to template layouts

Without templates, the builder creates slides on a blank canvas with your design system's fonts and colors.

### Configure a Design System

Design systems are YAML files that define your brand's fonts and colors:

```yaml
name: My Brand
fonts:
  title_font: Helvetica
  body_font: Helvetica
  title_size: 44
  subtitle_size: 28
  body_size: 18
  bullet_size: 16
colors:
  primary: "#1a73e8"
  secondary: "#34a853"
  accent: "#ea4335"
  text_dark: "#202124"
  text_light: "#5f6368"
  background: "#ffffff"
```

Save it to `design_systems/my_brand.yaml`. You can also extract one from an existing PPTX:

> Extract a design system from `templates/brand-deck.pptx`

## Workflows

All workflows are triggered by talking to the Cursor agent in chat.

### Generate a Presentation

> Generate a presentation from `inputs/document.pdf`

Options you can specify in natural language:
- Number of slides ("make it 12 slides")
- Tone ("keep it executive-level", "make it casual")
- Focus ("emphasize the financial metrics")
- Design system ("use the shopify design system")
- Template ("use the templates in `templates/corporate/`")

### Analyze Templates

> Analyze the templates in `templates/`

Extracts structural metadata from every slide in your template files and builds a searchable registry. The agent classifies each slide by type (title, content, two-column, etc.) so it can match them during generation.

### Extract a Design System

> Extract a design system from `templates/brand-deck.pptx`

Reads fonts, colors, and sizes from an existing PPTX and writes a design system YAML file.

### Run Pipeline Steps Manually (Tasks)

The project includes pre-configured Cursor/VS Code tasks you can run from the Command Palette (`Cmd+Shift+P` > "Tasks: Run Task"):

- **Parse Input Document** -- extract text from a file
- **Validate JSON Schema** -- check a workspace JSON file
- **Build Presentation** -- construct PPTX from the deck schema
- **Build with Templates** -- build using template matching
- **Quality Check** -- run programmatic quality checks
- **Analyze Templates** -- extract template metadata
- **Extract Design System** -- derive fonts/colors from a PPTX
- **Clean Workspace** -- remove intermediate files
- **Setup Environment** -- create venv and install dependencies

These are defined in `.cursor/workspace/tasks.json`.

### Edit After Generation

You don't have to regenerate from scratch. After a build, ask for targeted changes:

- *"Make slide 5 more visual"* -- edits the deck schema and rebuilds
- *"Add a slide about competitive landscape after slide 3"* -- inserts into the schema
- *"Use different fonts"* -- updates the design system and rebuilds
- *"The content needs more work"* -- re-runs content development at a lower maturity level

## Scripts Reference

All scripts run from the project root with the venv activated.

### `parse_input.py`

```bash
python scripts/parse_input.py <input_file> [-o workspace/parsed_content.txt]
```

Extracts text from PDF, DOCX, PPTX, TXT, or MD files.

### `validate_schema.py`

```bash
python scripts/validate_schema.py <json_file> <schema_name>
```

Validates a JSON file against a Pydantic schema. Schema names: `ContentInventory`, `DeckSchema`, `TemplateRegistry`, `DesignSystem`, `MatchResult`.

### `build_slides.py`

```bash
python scripts/build_slides.py <deck_schema.json> -o output.pptx [--design-system design_systems/default.yaml] [--matches workspace/template_matches.json]
```

Builds a PPTX from a deck schema and optional design system / template matches.

### `quality_check.py`

```bash
python scripts/quality_check.py <output.pptx> [--design-system design_systems/default.yaml] [--describe]
```

Runs programmatic quality checks: text overflow, font consistency, readability, pacing.

### `analyze_templates.py`

```bash
# Extract structural metadata from templates
python scripts/analyze_templates.py extract <template_dir> [-o workspace/template_descriptions.json]

# Merge with agent classifications into a registry
python scripts/analyze_templates.py merge <descriptions.json> <classifications.json> [-o template_registry.json]
```

### `extract_design_system.py`

```bash
python scripts/extract_design_system.py <file_or_dir> [-o design_systems/extracted.yaml] [--name "My Brand"]
```

### `match_templates.py`

```bash
python scripts/match_templates.py <deck_schema.json> <template_registry.json> [-o workspace/template_matches.json]
```

Matches deck slides to the best available template layouts.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Agent doesn't follow the workflow | Make sure you opened the project root in Cursor (not a subfolder). The `.cursor/rules/` must be at the workspace root. |
| `ModuleNotFoundError` when running scripts | Activate the venv: `source .venv/bin/activate` and run `pip install -e .` |
| PPTX has no backgrounds/branding | You need a template in `templates/`. Without one, slides are built on a blank canvas. |
| "Schema validation failed" | The agent produced malformed JSON. Ask it to read the error and fix the file. |
| Slides look cramped or text overflows | Run quality check: `python scripts/quality_check.py output/deck.pptx --describe` and ask the agent to fix the flagged slides. |

## License

MIT
