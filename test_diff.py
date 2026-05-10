"""
test_diff.py — Smoke tests for sniffer_diff.py

Generates markdown fixture pairs, diffs them, and verifies the output.

Usage:
    python test_diff.py

Test fixtures are preserved in testing/diff_fixtures/ for manual review.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from sniffer_diff import (
    extract_paragraphs, diff_paragraphs, render_report,
    inline_diff, load_markdown
)

# ============================================================
# Infra
# ============================================================

passed = 0
failed = 0

FIXTURE_DIR = Path(__file__).parent / "testing" / "diff_fixtures"


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


def save_fixture(name: str, content: str) -> Path:
    FIXTURE_DIR.mkdir(parents=True, exist_ok=True)
    path = FIXTURE_DIR / name
    path.write_text(content, encoding="utf-8")
    return path


# ============================================================
# Fixture content
# ============================================================

ROUND_3 = """\
# Band 4 - English

**Subject:** Action Required: Your Device Has Been Quarantined

Hello [DEVICE_USER_NAME],

Your Lilly-issued device, [DEVICE_MODEL] with serial number [DEVICE_SERIAL_NUMBER], has been offline for over 30 days.

## Required action

Return this device immediately by following the <Device Return Process>.

## Exceptions — No Action Required

- Worker on leave: If you are on leave, no action is required.

- Lost or stolen device: If the device has already been reported as lost or stolen, please disregard this email.

***Note:*** *Your supervisor has been copied on this notification.*

## Why it matters

- Quarantined devices must be recovered to protect Lilly data

- Ensures compliance with security policies

## Need help?

Contact <ChatNow> for immediate support

We appreciate your prompt action.

Tech@Lilly Device Management

*— Please do not reply to this automated email. —*
"""

ROUND_4 = """\
# Band 4 - English

**Subject:** Urgent: Your Lilly Device Has Been Quarantined

Hello [DEVICE_USER_NAME],

Your Lilly-issued device, [DEVICE_MODEL] with serial number [DEVICE_SERIAL_NUMBER], has been offline for over 30 days. As a result, it has been quarantined to protect Lilly's data and maintain security compliance.

## Required action

Return this device immediately by following the <Device Return Process>.

## Exceptions — No Action Required

- Worker on leave: If you are on leave, no action is required. When you return, submit a <Request to Unblock Device> to restore access.

- Lost or stolen device: If the device has already been reported as lost or stolen, please disregard this email.

***Note:*** *Your supervisor has been copied on this notification.*

## Why it matters

- Quarantined devices must be recovered to protect Lilly data

- Ensures compliance with security policies

- Completes the device lifecycle inventory records

## Need help?

Contact <ChatNow> for immediate support

We appreciate your prompt action.

Tech@Lilly Device Management

*— Please do not reply to this automated email. —*
"""

IDENTICAL_A = """\
# Band 1 - French

**Subject:** Test

Bonjour [NAME].

## Section

- Point un

- Point deux
"""

IDENTICAL_B = IDENTICAL_A  # exact copy

EMPHASIS_OLD = """\
# Band 1 - English

**Subject:** Test

This has **bold text** in it.

This has *italic text* in it.
"""

EMPHASIS_NEW = """\
# Band 1 - English

**Subject:** Test

This has bold text in it.

This has ***bold italic text*** in it.
"""


# ============================================================
# Tests
# ============================================================

def test_01_identical_no_diff():
    """Identical files produce no differences."""
    old_path = save_fixture("01_identical_a.md", IDENTICAL_A)
    new_path = save_fixture("01_identical_b.md", IDENTICAL_B)

    old_paras = extract_paragraphs(IDENTICAL_A)
    new_paras = extract_paragraphs(IDENTICAL_B)
    changes = diff_paragraphs(old_paras, new_paras)

    has_diffs = any(c["action"] != "unchanged" for c in changes)
    report("01: Identical files → no differences", not has_diffs,
           f"Found {sum(1 for c in changes if c['action'] != 'unchanged')} diffs" if has_diffs else "")


def test_02_subject_changed():
    """Subject line change detected."""
    save_fixture("02_round3.md", ROUND_3)
    save_fixture("02_round4.md", ROUND_4)

    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)

    subject_changed = any(
        c["action"] == "changed" and c["old"]["type"] == "subject"
        for c in changes
    )
    report("02: Subject line change detected", subject_changed)


def test_03_paragraph_text_changed():
    """Body paragraph text change detected (R3→R4 added quarantine explanation)."""
    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)

    body_changed = any(
        c["action"] == "changed" and c["old"]["type"] == "body"
        and "offline" in c["old"]["content"]
        for c in changes
    )
    report("03: Body paragraph text change detected", body_changed)


def test_04_bullet_expanded():
    """Bullet expanded with additional text (leave policy gained unblock instructions)."""
    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)

    bullet_changed = any(
        c["action"] == "changed" and c["old"]["type"] == "bullet"
        and "leave" in c["old"]["content"]
        for c in changes
    )
    report("04: Bullet expanded with new text", bullet_changed)


def test_05_paragraph_added():
    """New bullet added (R4 added 'device lifecycle' bullet)."""
    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)

    added = any(
        c["action"] == "added" and c["new"]["type"] == "bullet"
        and "lifecycle" in c["new"]["content"]
        for c in changes
    )
    report("05: New paragraph added", added)


def test_06_unchanged_preserved():
    """Unchanged paragraphs marked as unchanged."""
    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)

    unchanged_count = sum(1 for c in changes if c["action"] == "unchanged")
    ok = unchanged_count > 0
    report("06: Unchanged paragraphs preserved", ok,
           f"Unchanged: {unchanged_count}" if not ok else "")


def test_07_emphasis_change_detected():
    """Bold removed / italic→bold+italic change detected."""
    save_fixture("07_emphasis_old.md", EMPHASIS_OLD)
    save_fixture("07_emphasis_new.md", EMPHASIS_NEW)

    old_paras = extract_paragraphs(EMPHASIS_OLD)
    new_paras = extract_paragraphs(EMPHASIS_NEW)
    changes = diff_paragraphs(old_paras, new_paras)

    emphasis_changes = [c for c in changes if c["action"] == "changed"]
    ok = len(emphasis_changes) == 2  # bold→plain and italic→bold+italic
    report("07: Emphasis changes detected", ok,
           f"Changed count: {len(emphasis_changes)}" if not ok else "")


def test_08_inline_diff_words():
    """Inline diff shows word-level changes."""
    result = inline_diff(
        "Return this device immediately",
        "Return this device as soon as possible"
    )
    ok = "[-immediately-]" in result and "[+as soon as possible+]" in result
    save_fixture("08_inline_example.txt",
                 f"Old: Return this device immediately\n"
                 f"New: Return this device as soon as possible\n"
                 f"Diff: {result}\n")
    report("08: Inline diff shows word-level [-old-] [+new+]", ok,
           f"Got: {result}" if not ok else "")


def test_09_report_markdown_format():
    """Rendered report is valid markdown with expected sections."""
    old_paras = extract_paragraphs(ROUND_3)
    new_paras = extract_paragraphs(ROUND_4)
    changes = diff_paragraphs(old_paras, new_paras)
    report_md = render_report(changes, "round3.md", "round4.md")

    save_fixture("09_sample_report.md", report_md)

    checks = {
        "has_title": "# Diff Report" in report_md,
        "has_old_name": "round3.md" in report_md,
        "has_new_name": "round4.md" in report_md,
        "has_summary": "changed" in report_md and "added" in report_md,
        "has_changed_marker": "CHANGED" in report_md,
        "has_added_marker": "ADDED" in report_md,
    }
    all_ok = all(checks.values())
    failures = [k for k, v in checks.items() if not v]
    report("09: Report has expected markdown structure", all_ok,
           f"Failed: {failures}" if not all_ok else "")


def test_10_docx_input():
    """DOCX files can be loaded and diffed via kit tools."""
    docx_path = Path("/mnt/user-data/outputs/band4-english.docx")
    if not docx_path.exists():
        report("10: DOCX input loads via kit tools", False, "English template not found")
        return

    try:
        md = load_markdown(docx_path)
        paras = extract_paragraphs(md)
        ok = len(paras) > 5  # should have heading, subject, body, etc.
        report("10: DOCX input loads via kit tools", ok,
               f"Paragraphs: {len(paras)}" if not ok else "")
    except Exception as e:
        report("10: DOCX input loads via kit tools", False, str(e))


def test_11_paragraph_removed():
    """Removed paragraph detected when new version is shorter."""
    old_md = """\
# Band 1 - English

**Subject:** Test

Paragraph one.

Paragraph two.

Paragraph three.
"""
    new_md = """\
# Band 1 - English

**Subject:** Test

Paragraph one.

Paragraph three.
"""
    save_fixture("11_removed_old.md", old_md)
    save_fixture("11_removed_new.md", new_md)

    old_paras = extract_paragraphs(old_md)
    new_paras = extract_paragraphs(new_md)
    changes = diff_paragraphs(old_paras, new_paras)

    removed = any(
        c["action"] == "removed" and "two" in c["old"]["content"]
        for c in changes
    )
    report("11: Removed paragraph detected", removed)


def test_12_subheading_change():
    """Sub-heading text change detected."""
    old_md = """\
# Band 1 - English

## Required action

Do something.
"""
    new_md = """\
# Band 1 - English

## Action required immediately

Do something.
"""
    save_fixture("12_subheading_old.md", old_md)
    save_fixture("12_subheading_new.md", new_md)

    old_paras = extract_paragraphs(old_md)
    new_paras = extract_paragraphs(new_md)
    changes = diff_paragraphs(old_paras, new_paras)

    heading_changed = any(
        c["action"] == "changed" and c["old"]["type"] == "subheading"
        for c in changes
    )
    report("12: Sub-heading text change detected", heading_changed)


# ============================================================
# Runner
# ============================================================

def main():
    print("\n" + "=" * 60)
    print("SNIFFer Diff Tool — Smoke Tests")
    print("=" * 60)

    print("\n--- Core diffing ---")
    test_01_identical_no_diff()
    test_02_subject_changed()
    test_03_paragraph_text_changed()
    test_04_bullet_expanded()
    test_05_paragraph_added()
    test_06_unchanged_preserved()

    print("\n--- Emphasis & inline ---")
    test_07_emphasis_change_detected()
    test_08_inline_diff_words()

    print("\n--- Output ---")
    test_09_report_markdown_format()
    test_10_docx_input()

    print("\n--- Edge cases ---")
    test_11_paragraph_removed()
    test_12_subheading_change()

    print("\n" + "=" * 60)
    print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
    print("=" * 60)

    if FIXTURE_DIR.exists():
        fixtures = sorted(FIXTURE_DIR.iterdir())
        print(f"\nFixtures preserved at: {FIXTURE_DIR}/")
        print(f"  {len(fixtures)} fixture files")

    print()
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
