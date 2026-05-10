# SNIFFer Usage Guide
## Standardized Normalization & Interchange Format Framework

SNIFFer is a translation governance kit for managing multilingual email content. It provides a structured DOCX-based authoring workflow with automated parsing, validation, diffing, and format conversion.

Part of the same family as **LAPdog**.

---

## Quick Start

```bash
# Dependencies
pip install python-docx --break-system-packages

# Run all tests (72 tests across 5 suites)
python test_parser.py
python test_normalizer.py
python test_validator.py
python test_diff.py
python test_splitter.py
```

---

## Tools

### Parser (`sniffer_parser.py`)

Parses a DOCX file into structured runs (JSON). Enforces all `EL_*` style rules, token definitions, and structural requirements.

```bash
# Parse to stdout
python sniffer_parser.py band4-english.docx

# Parse to file
python sniffer_parser.py band4-english.docx -o parsed.json
```

**Exit codes:** 0 = clean, 1 = hard errors (no output), 2 = soft warnings (output emitted)

### Normalizer (`sniffer_normalizer.py`)

Converts parsed data to markdown projection — one `.md` per language per band, plus `legend.json`.

```bash
# From DOCX (runs parser internally)
python sniffer_normalizer.py band4-english.docx -r 2026-Q2 -o output/

# From pre-parsed JSON
python sniffer_normalizer.py parsed.json -r 2026-Q2 -o output/
```

**Output structure:**
```
output/
  2026-Q2/
    band4/
      english.md
      legend.json
```

### Validator (`sniffer_validator.py`)

Semantic cross-checks on top of parsed data. Accepts multiple files (per-language or unified).

```bash
# Validate a full language set
python sniffer_validator.py band4-english.docx band4-french.docx band4-german.docx ...

# Validate a unified doc
python sniffer_validator.py unified-band4.docx
```

**Hard rules:** all 9 languages present, non-empty example values, non-empty display text, valid URLs, consistent band numbers.

**Soft warnings:** paragraph count asymmetry, token usage asymmetry, missing subject lines.

### Diff Tool (`sniffer_diff.py`)

Paragraph-level diff between two rounds of the same band+language. Word-level inline diffs for changed paragraphs.

```bash
# Terminal output (ANSI colors)
python sniffer_diff.py round3/english.md round4/english.md

# Markdown report
python sniffer_diff.py round3/english.md round4/english.md -m

# Save report to file
python sniffer_diff.py round3/english.md round4/english.md -o diff_report.md

# Compare DOCXs directly
python sniffer_diff.py old.docx new.docx
```

### Splitter (`sniffer_splitter.py`)

Splits a unified multi-language DOCX into per-language files. Folds per-language legend overrides into each output file's legend.

```bash
python sniffer_splitter.py unified-band4.docx -o split_output/
```

**Output:**
```
split_output/
  band4-english.docx
  band4-french.docx
  band4-chinese.docx
  ... (9 files)
```

---

## DOCX Template Structure

Each DOCX (per-language mode) follows this structure:

```
Instructions preamble       ← free-form, parser ignores
LEGEND                      ← global token reference (EL_LegendHeading)
  Dynamic Data Tokens       ← [TOKEN] → example value + description
  Link Tokens               ← <TOKEN> → display text + URL + description
LEGEND - <Language>         ← optional: locale-specific overrides
SUBJECT                     ← email subject line (EL_SubjectHeading + EL_Subject)
Band N - <Language>         ← body content (EL_BandHeading)
  paragraphs...             ← EL_Body, EL_SubHeading, EL_Bullet, EL_Numbered
```

### Reserved Styles (`EL_*`)

| Style | Purpose |
|---|---|
| `EL_LegendHeading` | LEGEND / LEGEND - \<Language\> markers |
| `EL_BandHeading` | Band N - \<Language\> markers |
| `EL_SubjectHeading` | SUBJECT marker |
| `EL_Subject` | Email subject line text |
| `EL_SubHeading` | Bold structural labels (e.g., "Required action") |
| `EL_LegendLabel` | Sub-table labels ("Dynamic Data Tokens", "Link Tokens") |
| `EL_Rule` | Horizontal rule bounding legend tables |
| `EL_Body` | Standard body paragraphs |
| `EL_Bullet` | Bullet list items |
| `EL_Numbered` | Numbered list items |

### Token Conventions

- **Dynamic data:** `[BRACKETS]` — e.g., `[DEVICE_USER_NAME]`. Substituted at render time with runtime values.
- **Links:** `<ANGLE BRACKETS>` — e.g., `<ChatNow>`. Become clickable links at render time.
- Tokens stay literal in DOCX, structured runs, and markdown. Substitution is a downstream concern.
- Every token in body text must be defined in the LEGEND. The parser hard-fails on undefined tokens.

---

## Workflow

### Typical round

1. **Author** — Product owner edits per-language DOCX templates in Word
2. **Parse** — `sniffer_parser.py` validates structure and extracts content
3. **Normalize** — `sniffer_normalizer.py` projects to markdown for diffing and review
4. **Validate** — `sniffer_validator.py` cross-checks all languages
5. **Diff** — `sniffer_diff.py` compares against previous round

### Receiving imperfect work

The parser is strict on `EL_*` styles. If the product owner returns a doc with wrong styles:

1. Open in Word
2. Fix the styles (re-apply the correct `EL_*` style from the style picker)
3. Re-parse

The parser is strict; the operator is the soft layer.

### Unified → per-language conversion

If the product owner prefers working in a single file:

1. Author all languages in one unified DOCX
2. Run `sniffer_splitter.py` to produce per-language files
3. Parse/validate/normalize the per-language outputs

---

## Testing

See `testing/README.md` for the full testing guide with fixture descriptions and manual review instructions.

```
testing/
  README.md                    ← 250+ line testing guide
  parser_fixtures/             ← 18 DOCX files
  normalizer_fixtures/         ← 19 DOCX files
  validator_fixtures/          ← 27 DOCX files
  diff_fixtures/               ← 12 files (markdown pairs + sample report)
  splitter_fixtures/           ← 81 DOCX files across 9 test directories
```

**72 tests total, 0 failures.**

---

## File Inventory

```
sniffer_parser.py              ← DOCX → structured runs (JSON)
sniffer_normalizer.py          ← structured runs → markdown projection
sniffer_validator.py           ← semantic cross-checks
sniffer_diff.py                ← round-over-round paragraph diff
sniffer_splitter.py            ← unified DOCX → per-language DOCXs
sniffer-format.md              ← format specification

test_parser.py                 ← 18 tests
test_normalizer.py             ← 20 tests
test_validator.py              ← 13 tests
test_diff.py                   ← 12 tests
test_splitter.py               ← 9 tests

band4-english.docx             ← per-language template (populated)
band4-chinese.docx             ← per-language template (populated, with override)
band4-spanish.docx             ← per-language template (populated)
branch-b-template-band4.docx   ← legacy unified template (all 9 languages)

testing/                       ← fixtures and testing guide
```

---

## v1.1 Backlog

- **TUI master interface** — single entry point to pick tools, choose files, run workflows
