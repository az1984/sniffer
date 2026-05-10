"""
sniffer_normalizer.py — Convert parsed SNIFFer structured runs to markdown.

Usage:
    python sniffer_normalizer.py <input> [--round <round>] [--output-dir <dir>]

    <input> is either a .docx file (parser runs internally) or a .json file (pre-parsed).

Output:
    <output-dir>/<round>/<band>/legend.json
    <output-dir>/<round>/<band>/<language>.md

Exit codes:
    0 — clean
    1 — parser hard-failed (only when input is .docx)
    2 — parser soft warnings (output still emitted)
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Optional

from sniffer_parser import parse_docx, to_dict, ParseResult


# ============================================================
# Markdown rendering
# ============================================================

def render_emphasis(text: str, bold: bool, italic: bool) -> str:
    """Wrap text in markdown emphasis markers.

    Handles whitespace correctly: markers go inside leading/trailing
    spaces so markdown renders properly.
    """
    if not text:
        return ""

    if not bold and not italic:
        return text

    # Split leading/trailing whitespace from content
    stripped = text.lstrip()
    leading = text[:len(text) - len(stripped)]
    stripped_r = stripped.rstrip()
    trailing = stripped[len(stripped_r):]
    content = stripped_r

    if not content:
        return text  # all whitespace, no emphasis to apply

    if bold and italic:
        return f"{leading}***{content}***{trailing}"
    elif bold:
        return f"{leading}**{content}**{trailing}"
    else:  # italic
        return f"{leading}*{content}*{trailing}"


def render_runs(runs: list[dict]) -> str:
    """Render a list of runs into a markdown string."""
    parts = []
    for run in runs:
        text = run["text"]
        bold = run.get("bold", False)
        italic = run.get("italic", False)
        # Token runs: no special treatment in markdown, just literal text
        parts.append(render_emphasis(text, bold, italic))
    return "".join(parts)


def render_paragraph(para: dict, numbered_index: Optional[int] = None) -> str:
    """Render a single paragraph to markdown."""
    style = para["style"]

    if style == "EL_SubHeading":
        # SubHeading is bold by definition; ## carries the weight.
        # Strip bold from runs to avoid ## **text** doubling.
        stripped_runs = [
            {**r, "bold": False} for r in para["runs"]
        ]
        content = render_runs(stripped_runs)
        return f"## {content}"

    content = render_runs(para["runs"])

    if style == "EL_Bullet":
        return f"- {content}"
    elif style == "EL_Numbered":
        idx = numbered_index if numbered_index is not None else 1
        return f"{idx}. {content}"
    else:  # EL_Body or fallback
        return content


def render_section(section: dict) -> str:
    """Render a full band section to markdown."""
    lines = []

    # Band heading
    lines.append(f"# Band {section['band']} - {section['language']}")
    lines.append("")

    # Subject line (if present)
    subject = section.get("subject")
    if subject:
        lines.append(f"**Subject:** {subject}")
        lines.append("")

    # Body paragraphs
    numbered_counter = 0
    for para in section["paragraphs"]:
        if para["style"] == "EL_Numbered":
            numbered_counter += 1
            lines.append(render_paragraph(para, numbered_counter))
        else:
            numbered_counter = 0  # reset on non-numbered
            lines.append(render_paragraph(para))
        lines.append("")  # blank line between paragraphs

    # Strip trailing blank lines
    while lines and lines[-1] == "":
        lines.pop()

    return "\n".join(lines) + "\n"


# ============================================================
# Legend serialization
# ============================================================

def extract_legend(data: dict) -> dict:
    """Extract legend data for JSON output."""
    return data["legend"]


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Convert parsed SNIFFer data to markdown projection"
    )
    parser.add_argument("input", help="Path to .docx or .json file")
    parser.add_argument("--round", "-r", default="draft",
                        help="Round identifier (default: 'draft')")
    parser.add_argument("--output-dir", "-o", default="sniffer",
                        help="Base output directory (default: 'sniffer')")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        sys.exit(4)

    # Load data
    exit_code = 0
    if input_path.suffix == ".json":
        with open(input_path, encoding="utf-8") as f:
            data = json.load(f)
    elif input_path.suffix == ".docx":
        result = parse_docx(str(input_path))

        if result.hard_errors:
            print(f"\nParser hard errors ({len(result.hard_errors)}):", file=sys.stderr)
            for err in result.hard_errors:
                print(f"  ✗ {err}", file=sys.stderr)
            print("\nParse failed. No output emitted.", file=sys.stderr)
            sys.exit(1)

        if result.soft_warnings:
            print(f"\nParser warnings ({len(result.soft_warnings)}):", file=sys.stderr)
            for warn in result.soft_warnings:
                print(f"  ⚠ {warn}", file=sys.stderr)
            exit_code = 2

        data = to_dict(result)
    else:
        print(f"Error: unsupported file type '{input_path.suffix}'. Use .docx or .json.",
              file=sys.stderr)
        sys.exit(4)

    # Determine output paths
    base = Path(args.output_dir)
    round_id = args.round

    # Process each section
    files_written = []
    bands_seen = set()

    for section in data["sections"]:
        band = section["band"]
        lang = section["language"].lower()
        band_dir = base / round_id / f"band{band}"
        band_dir.mkdir(parents=True, exist_ok=True)
        bands_seen.add((band, band_dir))

        # Write markdown
        md_path = band_dir / f"{lang}.md"
        md_content = render_section(section)
        md_path.write_text(md_content, encoding="utf-8")
        files_written.append(str(md_path))

    # Write legend.json once per band
    legend = extract_legend(data)
    for band, band_dir in bands_seen:
        legend_path = band_dir / "legend.json"
        legend_path.write_text(
            json.dumps(legend, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8"
        )
        files_written.append(str(legend_path))

    # Report
    print(f"Wrote {len(files_written)} files:", file=sys.stderr)
    for f in sorted(files_written):
        print(f"  {f}", file=sys.stderr)

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
