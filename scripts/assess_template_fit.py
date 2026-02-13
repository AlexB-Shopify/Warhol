#!/usr/bin/env python3
"""Assess whether template-matched slides genuinely fit their content.

Runs after match_templates.py and before render_html.py.  For each slide
with match_type == "use_as_is", verifies:
  - The matched template has content zones
  - Each content block in the slide maps to a zone (role → zone_type)
  - Content fits within zone max_chars
  - A fit_score is computed; poor fits are demoted to compose mode

Demoted slides get match_type = "build_from_base" so render_html.py
builds them from scratch instead of cloning a mismatched template.

Usage:
    python scripts/assess_template_fit.py \
        workspace/deck_schema.json \
        workspace/template_matches.json \
        template_registry.json \
        -o workspace/template_matches.json \
        --fit-report workspace/fit_report.json
"""

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.schemas.slide_schema import DeckSchema
from src.schemas.template_schema import TemplateRegistry

# Minimum fit score to keep clone mode (0.0 - 1.0)
FIT_THRESHOLD = 0.5

# Map slide content roles to template zone types
ROLE_TO_ZONE = {
    "title":         {"title"},
    "subtitle":      {"subtitle"},
    "body":          {"body", "bullet_area"},
    "bullets":       {"body", "bullet_area"},
    "data_point":    {"data_point"},
    "quote":         {"body"},
    "caption":       {"caption", "subtitle"},
}


def _content_roles(spec) -> list[dict]:
    """Extract content roles and their char lengths from a SlideSpec."""
    roles = []
    if spec.title:
        roles.append({"role": "title", "chars": len(spec.title)})
    if spec.subtitle:
        roles.append({"role": "subtitle", "chars": len(spec.subtitle)})
    for block in spec.content_blocks:
        roles.append({"role": block.type, "chars": len(block.content)})
    return roles


def _assess_fit(spec, template) -> dict:
    """Compute a fit score for a slide against a matched template.

    Returns a dict with fit_score, zone_coverage, issues list, and pass/fail.
    """
    zones = template.content_zones
    content_roles = _content_roles(spec)

    # No content zones at all → cannot use clone mode
    if not zones:
        return {
            "fit_score": 0.0,
            "zone_coverage": 0.0,
            "char_fit": 0.0,
            "issues": ["Template has no content zones — cannot map content"],
            "pass": False,
        }

    # No content to place → trivially fits (section headers, etc.)
    if not content_roles:
        return {
            "fit_score": 1.0,
            "zone_coverage": 1.0,
            "char_fit": 1.0,
            "issues": [],
            "pass": True,
        }

    # Build zone lookup by type
    zone_by_type: dict[str, list] = {}
    for z in zones:
        zone_by_type.setdefault(z.zone_type, []).append(z)

    mapped = 0
    char_scores = []
    issues = []
    used_zones = set()

    for cr in content_roles:
        role = cr["role"]
        chars = cr["chars"]
        compatible_zone_types = ROLE_TO_ZONE.get(role, {"body"})

        # Find a matching zone
        best_zone = None
        for zt in compatible_zone_types:
            for z in zone_by_type.get(zt, []):
                if z.shape_name not in used_zones:
                    best_zone = z
                    break
            if best_zone:
                break

        if best_zone:
            mapped += 1
            used_zones.add(best_zone.shape_name)

            # Check character capacity
            if best_zone.max_chars and chars > 0:
                ratio = min(best_zone.max_chars / chars, 1.0) if chars > best_zone.max_chars else 1.0
                char_scores.append(ratio)
                if chars > best_zone.max_chars * 1.3:
                    issues.append(
                        f"{role}: {chars} chars exceeds zone '{best_zone.shape_name}' "
                        f"capacity ({best_zone.max_chars})"
                    )
            else:
                char_scores.append(1.0)
        else:
            issues.append(f"{role}: no compatible zone found (need {compatible_zone_types})")
            char_scores.append(0.0)

    # Compute scores
    zone_coverage = mapped / len(content_roles) if content_roles else 1.0
    char_fit = sum(char_scores) / len(char_scores) if char_scores else 1.0

    # Weighted fit score: zone coverage matters most
    fit_score = (zone_coverage * 0.7) + (char_fit * 0.3)

    return {
        "fit_score": round(fit_score, 3),
        "zone_coverage": round(zone_coverage, 3),
        "char_fit": round(char_fit, 3),
        "issues": issues,
        "pass": fit_score >= FIT_THRESHOLD,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Assess content-zone fit for template-matched slides"
    )
    parser.add_argument("deck_schema", type=Path, help="Path to deck_schema.json")
    parser.add_argument("matches", type=Path, help="Path to template_matches.json")
    parser.add_argument("registry", type=Path, help="Path to template_registry.json")
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="Output path for updated template_matches.json (default: overwrite input)")
    parser.add_argument("--fit-report", type=Path, default=Path("workspace/fit_report.json"),
                        help="Path to write fit report JSON")
    parser.add_argument("--threshold", type=float, default=FIT_THRESHOLD,
                        help=f"Minimum fit score to keep clone mode (default: {FIT_THRESHOLD})")
    args = parser.parse_args()

    # Load inputs
    deck = DeckSchema.model_validate_json(args.deck_schema.read_text())
    matches_data = json.loads(args.matches.read_text())
    registry = TemplateRegistry.load(args.registry)

    slide_lookup = {s.slide_number: s for s in deck.slides}
    matches = matches_data.get("matches", [])

    report_entries = []
    demoted = 0
    kept = 0
    skipped = 0

    for match in matches:
        slide_num = match["slide_number"]
        spec = slide_lookup.get(slide_num)
        if not spec:
            skipped += 1
            continue

        if match.get("match_type") != "use_as_is":
            # Already compose mode
            report_entries.append({
                "slide_number": slide_num,
                "match_type": "build_from_base",
                "fit_score": None,
                "action": "already_compose",
            })
            skipped += 1
            continue

        template_index = match.get("template_index", -1)
        if template_index < 0 or template_index >= len(registry.templates):
            match["match_type"] = "build_from_base"
            match["template_index"] = -1
            report_entries.append({
                "slide_number": slide_num,
                "match_type": "build_from_base",
                "fit_score": 0.0,
                "action": "demoted_invalid_index",
            })
            demoted += 1
            continue

        template = registry.templates[template_index]
        fit = _assess_fit(spec, template)

        if fit["pass"]:
            report_entries.append({
                "slide_number": slide_num,
                "match_type": "use_as_is",
                "template_index": template_index,
                "fit_score": fit["fit_score"],
                "zone_coverage": fit["zone_coverage"],
                "char_fit": fit["char_fit"],
                "issues": fit["issues"],
                "action": "kept_clone",
            })
            kept += 1
        else:
            match["match_type"] = "build_from_base"
            match["template_index"] = -1
            report_entries.append({
                "slide_number": slide_num,
                "match_type": "build_from_base",
                "fit_score": fit["fit_score"],
                "zone_coverage": fit["zone_coverage"],
                "char_fit": fit["char_fit"],
                "issues": fit["issues"],
                "action": "demoted_to_compose",
            })
            demoted += 1

    # Write updated matches
    output_path = args.output or args.matches
    output_path.write_text(json.dumps(matches_data, indent=2))

    # Write fit report
    fit_report = {
        "threshold": args.threshold,
        "total_slides": len(matches),
        "kept_clone": kept,
        "demoted_to_compose": demoted,
        "skipped": skipped,
        "entries": report_entries,
    }
    args.fit_report.parent.mkdir(parents=True, exist_ok=True)
    args.fit_report.write_text(json.dumps(fit_report, indent=2))

    # Summary
    print(f"Fit assessment: {len(matches)} slides (threshold={args.threshold})")
    print(f"  Clone (kept):  {kept}")
    print(f"  Compose (demoted): {demoted}")
    print(f"  Skipped/already compose: {skipped}")
    print(f"Updated matches: {output_path}")
    print(f"Fit report: {args.fit_report}")

    # Print demotions
    for entry in report_entries:
        if entry["action"] == "demoted_to_compose":
            issues_str = "; ".join(entry.get("issues", [])[:2])
            print(
                f"  Slide {entry['slide_number']}: "
                f"demoted (score={entry['fit_score']:.2f}) — {issues_str}"
            )


if __name__ == "__main__":
    main()
