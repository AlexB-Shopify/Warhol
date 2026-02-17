# Template Bank

Drop `.pptx` files here for the Slide Builder to use as layout sources.

## What Goes Here

**Any `.pptx` file works** -- both blank templates and completed presentations:

- **Blank templates** from SlidesCarnival, SlidesGo, etc.
- **Completed decks** from colleagues, clients, or your own past work
- **Google Slides exports** (File > Download > Microsoft PowerPoint)

The Template Analyzer extracts the **layout skeleton** from each slide (shape positions, sizes, fonts, colors) and discards the actual text content. So a finished deck's slides become reusable blank layouts.

## Directory Structure

Organize by visual style:

```
templates/
  corporate/   -- Professional business style
  creative/    -- Bold colors, modern layouts
  minimal/     -- Clean, simple designs
```

## How to Use

1. Drop `.pptx` files into the appropriate subdirectory
2. Run the analyzer to index layouts:

   ```bash
   slide-builder analyze-templates templates/
   ```

3. This creates `template_registry.json` in the project root
4. Now `slide-builder generate` will use these templates

## Extracting a Design System

You can also derive a **design system** (fonts, colors, sizes) directly from your templates. This is useful when you have a preferred deck and want all generated presentations to match its style.

From a single preferred deck:

```bash
slide-builder extract-design-system "templates/corporate/Brand Deck.pptx" \
  -o design_systems/brand.yaml --name "My Brand"
```

Or aggregate across an entire directory:

```bash
slide-builder extract-design-system templates/corporate/ \
  -o design_systems/corporate.yaml
```

The extractor walks every slide, tallies fonts and colors by context (title vs. body), and picks the most common values. The output is a YAML file you can use with `--design-system` during generation.

## Tips

- **Start small**: Begin with 3-5 files from a single source, verify the registry looks right, then add more
- **Avoid SmartArt**: `python-pptx` cannot manipulate SmartArt or embedded charts -- those slides will be skipped with a warning
- **Deduplication is automatic**: If multiple files have similar "title + subtitle" slides, only the best variant is kept
- **Re-run after changes**: Always re-run `analyze-templates` after adding or removing files
- **Temp files**: Files starting with `~` or `.` are ignored (e.g., `~$template.pptx` lock files)

## Built-in Template

The `minimal/basic.pptx` file is a simple built-in template with basic slide layouts. It's generated programmatically so the app works out of the box without any downloads.
