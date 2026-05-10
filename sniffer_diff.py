"""
sniffer_diff.py — Diff two rounds of the same band+language.

Compares paragraph-by-paragraph, including emphasis changes.
Accepts .md files directly or .docx files (parsed+normalized internally
using the kit's parser and normalizer).

Usage:
    python sniffer_diff.py <old> <new> [--output <path>]

    <old> and <new> are .md or .docx files.

Exit codes:
    0 — no differences
    1 — differences found
    2 — error (file not found, parse failure, etc.)
"""
from __future__ import annotations

import argparse
import difflib
import sys
import tempfile
from pathlib import Path

# Import kit tools (assumed to be in the same directory)
sys.path.insert(0, str(Path(__file__).parent))
from sniffer_parser import parse_docx, to_dict
from sniffer_normalizer import render_section


# ============================================================
# DOCX → markdown (via kit tools)
# ============================================================

def docx_to_markdown(docx_path: Path) -> str:
    """Parse a DOCX and normalize to markdown. Returns the markdown string.

    Uses the kit's parser and normalizer — no re-implementation.
    """
    result = parse_docx(str(docx_path))

    if result.hard_errors:
        print(f"Parser errors for {docx_path.name}:", file=sys.stderr)
        for err in result.hard_errors:
            print(f"  ✗ {err}", file=sys.stderr)
        raise RuntimeError(f"Parser failed on {docx_path.name}")

    data = to_dict(result)
    if not data["sections"]:
        raise RuntimeError(f"No sections found in {docx_path.name}")

    # Render first section (per-language mode = one section per file)
    return render_section(data["sections"][0])


def load_markdown(path: Path) -> str:
    """Load a file as markdown — handles both .md and .docx."""
    if path.suffix == ".docx":
        return docx_to_markdown(path)
    elif path.suffix == ".md":
        return path.read_text(encoding="utf-8")
    else:
        raise RuntimeError(f"Unsupported file type: {path.suffix}")


# ============================================================
# Paragraph extraction
# ============================================================

def extract_paragraphs(md: str) -> list[dict]:
    """Extract paragraphs from markdown with their type and content.

    Returns list of dicts: {"type": str, "raw": str, "content": str}
    - type: "heading", "subheading", "subject", "bullet", "numbered", "body", "blank"
    - raw: the original line
    - content: text without prefix markers
    """
    paragraphs = []
    for line in md.splitlines():
        stripped = line.strip()

        if not stripped:
            continue  # skip blank lines for diffing purposes

        if stripped.startswith("# "):
            paragraphs.append({
                "type": "heading",
                "raw": line,
                "content": stripped[2:]
            })
        elif stripped.startswith("## "):
            paragraphs.append({
                "type": "subheading",
                "raw": line,
                "content": stripped[3:]
            })
        elif stripped.startswith("**Subject:**"):
            paragraphs.append({
                "type": "subject",
                "raw": line,
                "content": stripped
            })
        elif stripped.startswith("- "):
            paragraphs.append({
                "type": "bullet",
                "raw": line,
                "content": stripped[2:]
            })
        elif len(stripped) > 2 and stripped[0].isdigit() and ". " in stripped[:5]:
            dot_pos = stripped.index(". ")
            paragraphs.append({
                "type": "numbered",
                "raw": line,
                "content": stripped[dot_pos + 2:]
            })
        else:
            paragraphs.append({
                "type": "body",
                "raw": line,
                "content": stripped
            })

    return paragraphs


# ============================================================
# Diff engine
# ============================================================

def diff_paragraphs(old_paras: list[dict], new_paras: list[dict]) -> list[dict]:
    """Diff two paragraph lists. Returns a list of change records.

    Each record: {"action": str, "old": dict|None, "new": dict|None, "detail": str}
    Actions: "unchanged", "added", "removed", "changed"
    """
    old_raws = [p["raw"] for p in old_paras]
    new_raws = [p["raw"] for p in new_paras]

    matcher = difflib.SequenceMatcher(None, old_raws, new_raws)
    changes = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i1, i2):
                changes.append({
                    "action": "unchanged",
                    "old": old_paras[k],
                    "new": new_paras[j1 + (k - i1)],
                    "detail": ""
                })

        elif tag == "replace":
            # Pair up old and new for inline comparison
            old_chunk = old_paras[i1:i2]
            new_chunk = new_paras[j1:j2]
            max_len = max(len(old_chunk), len(new_chunk))

            for k in range(max_len):
                if k < len(old_chunk) and k < len(new_chunk):
                    old_p = old_chunk[k]
                    new_p = new_chunk[k]
                    detail = inline_diff(old_p["raw"], new_p["raw"])
                    changes.append({
                        "action": "changed",
                        "old": old_p,
                        "new": new_p,
                        "detail": detail
                    })
                elif k < len(old_chunk):
                    changes.append({
                        "action": "removed",
                        "old": old_chunk[k],
                        "new": None,
                        "detail": ""
                    })
                else:
                    changes.append({
                        "action": "added",
                        "old": None,
                        "new": new_chunk[k],
                        "detail": ""
                    })

        elif tag == "delete":
            for k in range(i1, i2):
                changes.append({
                    "action": "removed",
                    "old": old_paras[k],
                    "new": None,
                    "detail": ""
                })

        elif tag == "insert":
            for k in range(j1, j2):
                changes.append({
                    "action": "added",
                    "old": None,
                    "new": new_paras[k],
                    "detail": ""
                })

    return changes


def inline_diff(old_line: str, new_line: str) -> str:
    """Produce a word-level inline diff between two lines."""
    old_words = old_line.split()
    new_words = new_line.split()

    matcher = difflib.SequenceMatcher(None, old_words, new_words)
    parts = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            parts.append(" ".join(old_words[i1:i2]))
        elif tag == "replace":
            parts.append(f"[-{' '.join(old_words[i1:i2])}-]")
            parts.append(f"[+{' '.join(new_words[j1:j2])}+]")
        elif tag == "delete":
            parts.append(f"[-{' '.join(old_words[i1:i2])}-]")
        elif tag == "insert":
            parts.append(f"[+{' '.join(new_words[j1:j2])}+]")

    return " ".join(parts)


# ============================================================
# Report rendering
# ============================================================

def render_report(changes: list[dict], old_name: str, new_name: str) -> str:
    """Render a diff report as markdown."""
    lines = []
    lines.append(f"# Diff Report")
    lines.append(f"")
    lines.append(f"**Old:** {old_name}")
    lines.append(f"**New:** {new_name}")
    lines.append(f"")

    # Summary
    added = sum(1 for c in changes if c["action"] == "added")
    removed = sum(1 for c in changes if c["action"] == "removed")
    changed = sum(1 for c in changes if c["action"] == "changed")
    unchanged = sum(1 for c in changes if c["action"] == "unchanged")

    if added == 0 and removed == 0 and changed == 0:
        lines.append("**No differences found.**")
        return "\n".join(lines) + "\n"

    lines.append(f"**Summary:** {changed} changed, {added} added, {removed} removed, {unchanged} unchanged")
    lines.append(f"")
    lines.append(f"---")
    lines.append(f"")

    # Detail
    para_num = 0
    for change in changes:
        action = change["action"]

        if action == "unchanged":
            para_num += 1
            continue

        para_num += 1

        if action == "added":
            lines.append(f"### ¶{para_num} — ADDED")
            lines.append(f"")
            lines.append(f"+ {change['new']['raw']}")
            lines.append(f"")

        elif action == "removed":
            lines.append(f"### ¶{para_num} — REMOVED")
            lines.append(f"")
            lines.append(f"- {change['old']['raw']}")
            lines.append(f"")

        elif action == "changed":
            lines.append(f"### ¶{para_num} — CHANGED")
            lines.append(f"")
            lines.append(f"- {change['old']['raw']}")
            lines.append(f"+ {change['new']['raw']}")
            if change["detail"]:
                lines.append(f"")
                lines.append(f"  *Inline:* {change['detail']}")
            lines.append(f"")

    return "\n".join(lines) + "\n"


def render_terminal(changes: list[dict], old_name: str, new_name: str) -> str:
    """Render a diff report for terminal output with ANSI colors."""
    RED = "\033[31m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    CYAN = "\033[36m"
    BOLD = "\033[1m"
    RESET = "\033[0m"

    lines = []
    lines.append(f"{BOLD}Diff: {old_name} → {new_name}{RESET}")
    lines.append("")

    added = sum(1 for c in changes if c["action"] == "added")
    removed = sum(1 for c in changes if c["action"] == "removed")
    changed = sum(1 for c in changes if c["action"] == "changed")
    unchanged = sum(1 for c in changes if c["action"] == "unchanged")

    if added == 0 and removed == 0 and changed == 0:
        lines.append(f"{GREEN}No differences found.{RESET}")
        return "\n".join(lines) + "\n"

    lines.append(
        f"{YELLOW}{changed} changed{RESET}, "
        f"{GREEN}{added} added{RESET}, "
        f"{RED}{removed} removed{RESET}, "
        f"{unchanged} unchanged"
    )
    lines.append("")

    para_num = 0
    for change in changes:
        action = change["action"]

        if action == "unchanged":
            para_num += 1
            continue

        para_num += 1

        if action == "added":
            lines.append(f"{CYAN}¶{para_num}{RESET} {GREEN}ADDED{RESET}")
            lines.append(f"  {GREEN}+ {change['new']['raw']}{RESET}")
            lines.append("")

        elif action == "removed":
            lines.append(f"{CYAN}¶{para_num}{RESET} {RED}REMOVED{RESET}")
            lines.append(f"  {RED}- {change['old']['raw']}{RESET}")
            lines.append("")

        elif action == "changed":
            lines.append(f"{CYAN}¶{para_num}{RESET} {YELLOW}CHANGED{RESET}")
            lines.append(f"  {RED}- {change['old']['raw']}{RESET}")
            lines.append(f"  {GREEN}+ {change['new']['raw']}{RESET}")
            if change["detail"]:
                lines.append(f"  {YELLOW}{change['detail']}{RESET}")
            lines.append("")

    return "\n".join(lines) + "\n"


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Diff two rounds of the same band+language"
    )
    parser.add_argument("old", help="Old round (.md or .docx)")
    parser.add_argument("new", help="New round (.md or .docx)")
    parser.add_argument("--output", "-o", help="Output markdown diff report to file")
    parser.add_argument("--markdown", "-m", action="store_true",
                        help="Force markdown output (no ANSI colors)")
    args = parser.parse_args()

    old_path = Path(args.old)
    new_path = Path(args.new)

    for p in [old_path, new_path]:
        if not p.exists():
            print(f"Error: file not found: {p}", file=sys.stderr)
            sys.exit(2)

    try:
        old_md = load_markdown(old_path)
        new_md = load_markdown(new_path)
    except RuntimeError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(2)

    old_paras = extract_paragraphs(old_md)
    new_paras = extract_paragraphs(new_md)

    changes = diff_paragraphs(old_paras, new_paras)

    has_diffs = any(c["action"] != "unchanged" for c in changes)

    if args.output or args.markdown:
        report = render_report(changes, old_path.name, new_path.name)
        if args.output:
            out = Path(args.output)
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(report, encoding="utf-8")
            print(f"Diff report written to {out}", file=sys.stderr)
        else:
            print(report)
    else:
        print(render_terminal(changes, old_path.name, new_path.name))

    sys.exit(1 if has_diffs else 0)


if __name__ == "__main__":
    main()
