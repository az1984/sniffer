"""
sniffer_validator.py — Semantic validation on parsed SNIFFer data.

Runs on top of the parser's structured output. Checks cross-section
and cross-field rules that the parser doesn't enforce.

Usage:
    python sniffer_validator.py <docx_or_json> [<docx_or_json> ...]

    Accepts one or more .docx files (parser runs internally) or .json files.
    Multiple per-language files are merged before validation.

Exit codes:
    0 — clean
    1 — hard errors (validation failed)
    2 — soft warnings only
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).parent))
from sniffer_parser import parse_docx, to_dict


# ============================================================
# Constants
# ============================================================

REQUIRED_LANGUAGES = {
    "English", "French", "German", "Italian", "Spanish",
    "Portuguese", "Japanese", "Chinese", "Korean"
}


# ============================================================
# Merge multiple parsed results
# ============================================================

def merge_parsed(datasets: list[dict]) -> dict:
    """Merge multiple parsed JSON outputs into one.

    Legend: uses the first non-empty legend found. Overrides are merged.
    Sections: concatenated.
    """
    merged = {
        "legend": {"dynamic_tokens": [], "link_tokens": []},
        "sections": []
    }

    all_overrides = {}

    for data in datasets:
        leg = data.get("legend", {})

        # Take the first non-empty legend as canonical
        if not merged["legend"]["dynamic_tokens"] and leg.get("dynamic_tokens"):
            merged["legend"]["dynamic_tokens"] = leg["dynamic_tokens"]
        if not merged["legend"]["link_tokens"] and leg.get("link_tokens"):
            merged["legend"]["link_tokens"] = leg["link_tokens"]

        # Merge overrides
        for lang, ov in leg.get("overrides", {}).items():
            if lang not in all_overrides:
                all_overrides[lang] = ov
            else:
                # Merge: later file's overrides win for same token
                existing = all_overrides[lang]
                if "dynamic_tokens" in ov:
                    existing_dt = {dt["token"]: dt for dt in existing.get("dynamic_tokens", [])}
                    for dt in ov["dynamic_tokens"]:
                        existing_dt[dt["token"]] = dt
                    existing["dynamic_tokens"] = list(existing_dt.values())
                if "link_tokens" in ov:
                    existing_lt = {lt["token"]: lt for lt in existing.get("link_tokens", [])}
                    for lt in ov["link_tokens"]:
                        existing_lt[lt["token"]] = lt
                    existing["link_tokens"] = list(existing_lt.values())

    if all_overrides:
        merged["legend"]["overrides"] = all_overrides

    # Merge sections
    for data in datasets:
        merged["sections"].extend(data.get("sections", []))

    return merged


# ============================================================
# Validation rules
# ============================================================

def validate(data: dict) -> tuple[list[str], list[str]]:
    """Run all validation rules. Returns (hard_errors, soft_warnings)."""
    errors: list[str] = []
    warnings: list[str] = []

    legend = data.get("legend", {})
    sections = data.get("sections", [])

    dynamic_tokens = legend.get("dynamic_tokens", [])
    link_tokens = legend.get("link_tokens", [])

    # ----------------------------------------------------------
    # Hard rule 1: All nine languages present
    # ----------------------------------------------------------
    present_languages = {s["language"] for s in sections}
    missing = REQUIRED_LANGUAGES - present_languages
    if missing:
        errors.append(
            f"Missing languages: {', '.join(sorted(missing))}. "
            f"Present: {', '.join(sorted(present_languages))}"
        )

    # ----------------------------------------------------------
    # Hard rule 2: Every dynamic token has non-empty example value
    # ----------------------------------------------------------
    for dt in dynamic_tokens:
        if not dt.get("example_value", "").strip():
            errors.append(
                f"Dynamic token '{dt['token']}' has empty example value"
            )

    # ----------------------------------------------------------
    # Hard rule 3: Every link token has non-empty display text
    # ----------------------------------------------------------
    for lt in link_tokens:
        if not lt.get("display_text", "").strip():
            errors.append(
                f"Link token '{lt['token']}' has empty display text"
            )

    # ----------------------------------------------------------
    # Hard rule 4: Every link token has a valid URL
    # ----------------------------------------------------------
    for lt in link_tokens:
        url = lt.get("target_url", "").strip()
        if not url:
            errors.append(
                f"Link token '{lt['token']}' has empty target URL"
            )
        else:
            parsed = urlparse(url)
            if not parsed.scheme or not parsed.netloc:
                errors.append(
                    f"Link token '{lt['token']}' has invalid URL: '{url}'"
                )

    # ----------------------------------------------------------
    # Hard rule 5: Band number consistency
    # ----------------------------------------------------------
    band_numbers = {s["band"] for s in sections}
    if len(band_numbers) > 1:
        errors.append(
            f"Inconsistent band numbers across sections: {sorted(band_numbers)}. "
            f"Expected all sections to reference the same band."
        )

    # ----------------------------------------------------------
    # Soft warning 6: Paragraph count asymmetry
    # ----------------------------------------------------------
    if len(sections) >= 2:
        counts = {s["language"]: len(s.get("paragraphs", [])) for s in sections}
        if counts:
            max_count = max(counts.values())
            min_count = min(counts.values())
            # Warn if any language has less than half the max
            if max_count > 0 and min_count < max_count * 0.5:
                sparse = [
                    f"{lang} ({c})" for lang, c in sorted(counts.items())
                    if c < max_count * 0.5
                ]
                full = [
                    f"{lang} ({c})" for lang, c in sorted(counts.items())
                    if c == max_count
                ]
                warnings.append(
                    f"Paragraph count asymmetry: {', '.join(sparse)} may be incomplete. "
                    f"Fullest: {', '.join(full)}"
                )

    # ----------------------------------------------------------
    # Soft warning 7: Token usage asymmetry
    # ----------------------------------------------------------
    all_defined = {dt["token"] for dt in dynamic_tokens} | {lt["token"] for lt in link_tokens}

    if len(sections) >= 2 and all_defined:
        usage_by_lang: dict[str, set[str]] = {}
        for s in sections:
            lang = s["language"]
            used = set()
            for para in s.get("paragraphs", []):
                for run in para.get("runs", []):
                    if run.get("token"):
                        used.add(run["text"])
            usage_by_lang[lang] = used

        # Find tokens used in some languages but not others
        all_used = set()
        for used in usage_by_lang.values():
            all_used |= used

        for token in sorted(all_used):
            langs_using = [lang for lang, used in usage_by_lang.items() if token in used]
            langs_missing = [lang for lang, used in usage_by_lang.items()
                             if token not in used and len(used) > 0]  # skip empty sections
            if langs_missing and langs_using:
                warnings.append(
                    f"Token '{token}' used in {', '.join(sorted(langs_using))} "
                    f"but missing from {', '.join(sorted(langs_missing))}"
                )

    # ----------------------------------------------------------
    # Soft warning 8: Subject line missing for populated section
    # ----------------------------------------------------------
    for s in sections:
        if s.get("paragraphs") and not s.get("subject"):
            warnings.append(
                f"Band {s['band']} - {s['language']} has body content but no subject line"
            )

    return errors, warnings


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate parsed SNIFFer data (semantic cross-checks)"
    )
    parser.add_argument("inputs", nargs="+",
                        help="One or more .docx or .json files")
    args = parser.parse_args()

    datasets = []
    parse_errors = []

    for input_path_str in args.inputs:
        input_path = Path(input_path_str)
        if not input_path.exists():
            print(f"Error: file not found: {input_path}", file=sys.stderr)
            sys.exit(4)

        if input_path.suffix == ".json":
            with open(input_path, encoding="utf-8") as f:
                datasets.append(json.load(f))
        elif input_path.suffix == ".docx":
            result = parse_docx(str(input_path))
            if result.hard_errors:
                parse_errors.extend(
                    f"[{input_path.name}] {e}" for e in result.hard_errors
                )
            datasets.append(to_dict(result))
        else:
            print(f"Error: unsupported file type '{input_path.suffix}'", file=sys.stderr)
            sys.exit(4)

    # Report parse errors first
    if parse_errors:
        print(f"\n{'='*60}", file=sys.stderr)
        print(f"PARSER ERRORS ({len(parse_errors)}) — fix these before validation",
              file=sys.stderr)
        print(f"{'='*60}", file=sys.stderr)
        for err in parse_errors:
            print(f"  ✗ {err}", file=sys.stderr)
        print(f"\nValidation aborted due to parser errors.", file=sys.stderr)
        sys.exit(1)

    # Merge datasets
    if len(datasets) == 1:
        merged = datasets[0]
    else:
        merged = merge_parsed(datasets)

    # Validate
    errors, warnings = validate(merged)

    if errors:
        print(f"\n{'='*60}", file=sys.stderr)
        print(f"VALIDATION ERRORS ({len(errors)})", file=sys.stderr)
        print(f"{'='*60}", file=sys.stderr)
        for err in errors:
            print(f"  ✗ {err}", file=sys.stderr)

    if warnings:
        print(f"\n{'='*60}", file=sys.stderr)
        print(f"VALIDATION WARNINGS ({len(warnings)})", file=sys.stderr)
        print(f"{'='*60}", file=sys.stderr)
        for warn in warnings:
            print(f"  ⚠ {warn}", file=sys.stderr)

    if errors:
        print(f"\nValidation failed with {len(errors)} error(s).", file=sys.stderr)
        sys.exit(1)
    elif warnings:
        print(f"\nValidation passed with {len(warnings)} warning(s).", file=sys.stderr)
        sys.exit(2)
    else:
        print(f"\nValidation passed — no errors, no warnings.", file=sys.stderr)
        sys.exit(0)


if __name__ == "__main__":
    main()
