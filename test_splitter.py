"""
test_splitter.py — Smoke tests for sniffer_splitter.py

Tests the unified-to-per-language DOCX splitter.

Usage:
    python test_splitter.py

Test fixtures are preserved in testing/splitter_fixtures/ for manual review.
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from sniffer_parser import parse_docx, to_dict
from sniffer_splitter import split_docx

# ============================================================
# Infra
# ============================================================

passed = 0
failed = 0

FIXTURE_DIR = Path(__file__).parent / "testing" / "splitter_fixtures"
UNIFIED_TEMPLATE = Path("/mnt/user-data/outputs/branch-b-template-band4.docx")


def report(name: str, ok: bool, detail: str = ""):
    global passed, failed
    if ok:
        passed += 1
        print(f"  ✓ {name}")
    else:
        failed += 1
        print(f"  ✗ {name}")
        if detail:
            print(f"    → {detail}")


# ============================================================
# Tests
# ============================================================

def test_01_produces_nine_files():
    """Splitting unified template produces 9 per-language files."""
    out_dir = FIXTURE_DIR / "01_nine_files"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    ok = len(files) == 9
    report("01: Produces 9 per-language files", ok,
           f"Got {len(files)} files" if not ok else "")


def test_02_all_files_parse_clean():
    """Every split file parses without hard errors."""
    out_dir = FIXTURE_DIR / "02_parse_clean"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    failures = []
    for f in files:
        result = parse_docx(f)
        if result.hard_errors:
            failures.append(f"{Path(f).name}: {result.hard_errors[0]}")

    ok = len(failures) == 0
    report("02: All split files parse clean", ok,
           f"Failures: {failures}" if not ok else "")


def test_03_each_file_has_one_section():
    """Each split file contains exactly one band section."""
    out_dir = FIXTURE_DIR / "03_one_section"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    multi_section = []
    for f in files:
        result = parse_docx(f)
        data = to_dict(result)
        if len(data["sections"]) != 1:
            multi_section.append(f"{Path(f).name}: {len(data['sections'])} sections")

    ok = len(multi_section) == 0
    report("03: Each file has exactly one section", ok,
           f"Multi-section: {multi_section}" if not ok else "")


def test_04_language_matches_filename():
    """The language in the band heading matches the filename."""
    out_dir = FIXTURE_DIR / "04_lang_match"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    mismatches = []
    for f in files:
        result = parse_docx(f)
        data = to_dict(result)
        section = data["sections"][0]
        expected_lang = section["language"].lower()
        filename = Path(f).stem  # e.g., "band4-english"
        if expected_lang not in filename:
            mismatches.append(f"{filename} vs {section['language']}")

    ok = len(mismatches) == 0
    report("04: Language matches filename", ok,
           f"Mismatches: {mismatches}" if not ok else "")


def test_05_legend_preserved():
    """Each split file has the full legend (3 dynamic + 3 link tokens)."""
    out_dir = FIXTURE_DIR / "05_legend"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    issues = []
    for f in files:
        result = parse_docx(f)
        data = to_dict(result)
        dt_count = len(data["legend"]["dynamic_tokens"])
        lt_count = len(data["legend"]["link_tokens"])
        if dt_count != 3 or lt_count != 3:
            issues.append(f"{Path(f).name}: {dt_count} dynamic, {lt_count} link")

    ok = len(issues) == 0
    report("05: Legend preserved in all files (3 dynamic + 3 link)", ok,
           f"Issues: {issues}" if not ok else "")


def test_06_chinese_override_folded():
    """Chinese file has overrides folded into the legend (no separate override section)."""
    out_dir = FIXTURE_DIR / "06_override"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    chinese_file = [f for f in files if "chinese" in Path(f).name.lower()]
    if not chinese_file:
        report("06: Chinese override folded in", False, "No Chinese file found")
        return

    result = parse_docx(chinese_file[0])
    data = to_dict(result)

    # Should have no overrides section (folded in)
    has_overrides = "overrides" in data["legend"] and data["legend"]["overrides"]

    # Display text should be Chinese, not English
    link_tokens = data["legend"]["link_tokens"]
    device_return = [lt for lt in link_tokens if lt["token"] == "<Device Return Process>"]

    chinese_display = False
    if device_return:
        # Should contain Chinese characters, not "device return process"
        display = device_return[0]["display_text"]
        chinese_display = any('\u4e00' <= c <= '\u9fff' for c in display)

    ok = not has_overrides and chinese_display
    report("06: Chinese override folded into legend", ok,
           f"has_overrides={has_overrides}, chinese_display={chinese_display}" if not ok else "")


def test_07_english_no_override():
    """English file has the global legend values (no override)."""
    out_dir = FIXTURE_DIR / "07_no_override"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    english_file = [f for f in files if "english" in Path(f).name.lower()]
    if not english_file:
        report("07: English has global legend values", False, "No English file found")
        return

    result = parse_docx(english_file[0])
    data = to_dict(result)

    link_tokens = data["legend"]["link_tokens"]
    device_return = [lt for lt in link_tokens if lt["token"] == "<Device Return Process>"]

    ok = False
    if device_return:
        ok = device_return[0]["display_text"] == "device return process"

    report("07: English has global legend values (not overridden)", ok,
           f"Display: {device_return[0]['display_text'] if device_return else '?'}" if not ok else "")


def test_08_paragraph_count_preserved():
    """Populated sections maintain their paragraph count after splitting."""
    out_dir = FIXTURE_DIR / "08_para_count"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    # Get counts from unified
    unified_result = parse_docx(str(UNIFIED_TEMPLATE))
    unified_data = to_dict(unified_result)
    unified_counts = {
        s["language"]: len(s["paragraphs"])
        for s in unified_data["sections"]
    }

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))
    mismatches = []
    for f in files:
        result = parse_docx(f)
        data = to_dict(result)
        section = data["sections"][0]
        lang = section["language"]
        split_count = len(section["paragraphs"])
        unified_count = unified_counts.get(lang, -1)
        if split_count != unified_count:
            mismatches.append(f"{lang}: unified={unified_count}, split={split_count}")

    ok = len(mismatches) == 0
    report("08: Paragraph counts preserved after split", ok,
           f"Mismatches: {mismatches}" if not ok else "")


def test_09_round_trip_split_then_validate():
    """Split files can be fed back to the validator as a set."""
    out_dir = FIXTURE_DIR / "09_round_trip"
    if out_dir.exists():
        shutil.rmtree(out_dir)

    files = split_docx(str(UNIFIED_TEMPLATE), str(out_dir))

    # Parse all and merge (simulating what the validator does)
    from sniffer_validator import validate, merge_parsed
    datasets = []
    for f in files:
        result = parse_docx(f)
        datasets.append(to_dict(result))

    merged = merge_parsed(datasets)
    errors, warnings = validate(merged)

    # Should have no hard errors (all 9 languages present)
    lang_errors = [e for e in errors if "Missing languages" in e]
    ok = len(lang_errors) == 0
    report("09: Split files pass validator (all 9 languages present)", ok,
           f"Errors: {errors}" if not ok else "")


# ============================================================
# Runner
# ============================================================

def main():
    if not UNIFIED_TEMPLATE.exists():
        print(f"Error: unified template not found at {UNIFIED_TEMPLATE}", file=sys.stderr)
        sys.exit(2)

    print("\n" + "=" * 60)
    print("SNIFFer Splitter — Smoke Tests")
    print("=" * 60)

    print("\n--- Core splitting ---")
    test_01_produces_nine_files()
    test_02_all_files_parse_clean()
    test_03_each_file_has_one_section()
    test_04_language_matches_filename()

    print("\n--- Legend handling ---")
    test_05_legend_preserved()
    test_06_chinese_override_folded()
    test_07_english_no_override()

    print("\n--- Fidelity ---")
    test_08_paragraph_count_preserved()
    test_09_round_trip_split_then_validate()

    print("\n" + "=" * 60)
    print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
    print("=" * 60)

    if FIXTURE_DIR.exists():
        fixture_count = sum(1 for _ in FIXTURE_DIR.rglob("*.docx"))
        print(f"\nFixtures preserved at: {FIXTURE_DIR}/")
        print(f"  {fixture_count} split DOCX files across test directories")

    print()
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
