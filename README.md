# Warhol

A Cursor-native PowerPoint presentation generator. Drop in a document, tell the Cursor agent what you want, and it builds a polished PPTX -- no API keys or external LLM services required.

## What is "Cursor-Native"?

This project uses [Cursor](https://cursor.sh/) (an AI-powered code editor) as its intelligence layer. The Cursor agent reads your documents, plans the narrative, writes slide content, and orchestrates the build. Python scripts handle the deterministic work: parsing files, constructing PPTX slides, and validating data.

There are no API keys to configure. The AI capabilities come from Cursor itself, guided by a set of workflow rules in `.cursor/rules/` that teach the agent how to build presentations.

## Getting Started

### Prerequisites

- **[Cursor](https://cursor.sh/)** IDE (required -- this is the LLM)
- **Python 3.11+** ([python.org](https://www.python.org/downloads/))
- **Git**

### Setup (two options)

**Option A -- Ask Cursor to do it:**

1. Clone the repo and open it in Cursor:
   ```bash
   git clone <repo-url> warhol && cd warhol
   ```
2. Open the folder in Cursor (`File > Open Folder`)
3. In Cursor chat, type:
   > Set up warhol
4. Cursor reads the project rules, runs the setup script, and tells you when it's ready.

**Option B -- Do it yourself:**

```bash
git clone <repo-url> warhol && cd warhol
bash scripts/setup.sh
```

That's it. The script creates the virtual environment, installs dependencies, and verifies everything works.

### Generate a Presentation

1. Drop your input document (PDF, DOCX, PPTX, TXT, or MD) into `inputs/`
2. Optionally drop a branded `.pptx` template into `templates/`
3. Ask Cursor:

> Generate a presentation from `inputs/my-document.pdf`

The agent walks through the full pipeline -- parsing, content extraction, narrative planning, HTML composition, and PPTX construction -- and produces a finished deck. Output lands in `output/`.

You can guide the result with extra instructions:

> Generate a 10-slide presentation from `inputs/quarterly-review.pdf` using a corporate tone. Focus on the financial metrics and growth story.

> Build a presentation from `inputs/notes.md`. Keep it under 8 slides and make it punchy.

## How the Pipeline Works

The generation pipeline has eight stages. The agent runs each one in order, writing intermediate files to `workspace/` so you can inspect or edit at any point.

```
Input Document
     │
     ▼
1. Parse .............. Extract text from PDF / DOCX / PPTX / TXT / MD
     │
     ▼
2. Assess ............. Evaluate content maturity (Level 1-4)
     │
     ▼
3. Extract ............ Produce structured content inventory
     │
     ▼
4. Plan ............... Decide depth per section, sequence narrative
     │
     ▼
5. Develop ............ Create / refine slide content (adapted to maturity)
     │
     ▼
6. Compose HTML ....... Build fully branded HTML slide deck
     │
     ▼
7. Build PPTX ......... Mechanically reproduce HTML as PowerPoint
     │
     ▼
8. Review ............. Programmatic quality checks + agent assessment
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
.cursor/rules/          Agent workflow rules (the "brain")
scripts/                Python scripts the agent calls
src/                    Python library (parsers, PPTX engine, schemas)
workspace/              Intermediate pipeline files (auto-generated)
design_systems/         YAML design system configs (fonts, colors)
templates/              PPTX template files (you provide these)
inputs/                 Input documents (you provide these)
output/                 Generated presentations land here
```

### Key Files

| File | Purpose |
|------|---------|
| `workspace/parsed_content.txt` | Raw text extracted from your input document |
| `workspace/content_inventory.json` | Structured content the agent extracted |
| `workspace/deck_schema.json` | Slide-by-slide plan (the blueprint for the deck) |
| `workspace/deck_preview.html` | The HTML slide deck (primary creative output) |
| `workspace/template_matches.json` | Which template slide maps to which deck slide |
| `workspace/quality_report.json` | Quality check results |
| `design_systems/*.yaml` | Font, color, and sizing configuration |
| `template_registry.json` | Metadata about available template slides (generated) |

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

### Edit After Generation

You don't have to regenerate from scratch. After a build, ask for targeted changes:

- *"Make slide 5 more visual"* -- edits the HTML deck and rebuilds the PPTX
- *"Add a slide about competitive landscape after slide 3"* -- inserts into the HTML deck
- *"Use different fonts"* -- updates the design system and rebuilds
- *"The content needs more work"* -- re-runs content development
- *"Fix the text on slide 7"* -- edits the HTML directly and rebuilds

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

## Scripts Reference

All scripts run from the project root with the venv activated. The agent runs these automatically during the pipeline -- you normally don't need to run them by hand.

| Script | Purpose |
|--------|---------|
| `scripts/setup.sh` | Set up the project (venv + dependencies) |
| `scripts/parse_input.py` | Extract text from PDF/DOCX/PPTX/TXT/MD |
| `scripts/validate_schema.py` | Validate JSON against Pydantic schemas |
| `scripts/render_html.py` | Render branded HTML deck from schema + design system |
| `scripts/build_from_html.py` | Build PPTX from HTML slide deck (primary builder) |
| `scripts/match_templates.py` | Algorithmic template matching |
| `scripts/analyze_templates.py` | Extract structural metadata from templates |
| `scripts/extract_design_system.py` | Derive fonts/colors from PPTX |
| `scripts/quality_check.py` | Programmatic quality checks on PPTX |

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Agent doesn't follow the workflow | Make sure you opened the project root in Cursor (not a subfolder). The `.cursor/rules/` must be at the workspace root. |
| `ModuleNotFoundError` when running scripts | Run `bash scripts/setup.sh` or activate the venv: `source .venv/bin/activate && pip install -e .` |
| PPTX has no backgrounds/branding | You need a template in `templates/`. Without one, slides are built on a blank canvas. |
| "Schema validation failed" | The agent produced malformed JSON. Ask it to read the error and fix the file. |
| Slides look cramped or text overflows | Run quality check and ask the agent to fix flagged slides. |
| PPTX won't open in Google Slides | Run `python scripts/repair_pptx.py output/presentation.pptx` to fix compatibility issues. |

## License

MIT
