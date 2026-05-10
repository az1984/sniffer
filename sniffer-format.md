# SNIFFer Format Specification
## Standardized Normalization & Interchange Format Framework

**Status:** v1.0 — all core tools built and tested.
**Owners:** Operator (kit author), Product Owner (translation review).
**Supersedes:** Ad-hoc spreadsheet representation of approved translations.

---

## 1. Purpose

SNIFFer is the **approved translation source of truth**. It feeds:

- Sample-email rendering for product-owner review
- Validator cross-checks against deployed Logic App text (Branch A)
- Future automated drift detection round-over-round

This spec defines the interchange format between the product owner (who authors and signs off on translations) and the kit (which consumes SNIFFer for rendering and validation).

---

## 2. Three-format architecture

SNIFFer exists in three forms with deterministic, one-directional conversion:

| Layer | Audience | Format |
|---|---|---|
| **Authoring surface** | Product owner | DOCX, one file per band, edited in Word |
| **Internal representation** | Kit | Structured runs (paragraph + bold/italic per text span), parsed via python-docx |
| **Diff / audit / review surface** | Operator, future reviewers | Markdown projection, one `.md` per band per language |

Conversion flow: **DOCX → structured runs → markdown**. Markdown is regenerated from DOCX on every parse and is never hand-edited.

### 2.1 Why DOCX

The product owner validates bold/italic emphasis as part of the review surface. Plain-text formats cannot carry emphasis without ad-hoc markup conventions that add learning cost without solving the underlying problem. Word preserves emphasis natively and is the product owner's existing tool. Format complexity moves to the parser, where it belongs.

### 2.2 Why markdown for the audit layer

Markdown is plain-text diffable, git-friendly, and viewable without Word. It carries enough emphasis information (`**bold**`, `*italic*`) to reflect the formatting decisions that matter for translation review without dragging the full DOCX binary through diffs.

---

## 3. Document structure

One DOCX per band per language (for apps that use bands). Not all apps partition into bands; for single-band apps, the DOCX is simply one file per language. The multi-language-per-file format is retained as a legacy mode but is not the primary workflow.

Each DOCX contains, in order:

1. An optional **instructions preamble** (free-form, any non-`EL_` style)
2. A **global LEGEND section** (exactly once)
3. Zero or more **per-language LEGEND override sections**
4. A **subject line section** (`EL_SubjectHeading` + `EL_Subject`)
5. One **band section** containing the email body content

### 3.0 Instructions preamble

The parser ignores all content before the first `EL_LegendHeading` paragraph. This space is reserved for human-readable instructions on how to use the template. The preamble may use any non-`EL_` styles (Normal, Heading 1, etc.) and is never parsed, validated, or emitted.

This allows the template to be self-documenting without polluting the structured output.

### 3.1 Section markers

All section markers are Heading 2 base style, distinguished by `EL_*` style name.

| Section type | Style | Text pattern | Required |
|---|---|---|---|
| Global legend | `EL_LegendHeading` | `LEGEND` (literal) | Yes, exactly once, must be first `EL_*` element |
| Per-language legend override | `EL_LegendHeading` | `LEGEND - <Language>` | Optional, zero or more |
| Subject line heading | `EL_SubjectHeading` | `SUBJECT` (literal) | Yes, exactly once |
| Band section | `EL_BandHeading` | `Band N - <Language>` | Yes, one per language in scope |

`<Language>` is the canonical English name: `English`, `French`, `German`, `Italian`, `Spanish`, `Portuguese`, `Japanese`, `Chinese`, `Korean`.

### 3.2 Section ordering (per-language mode)

```
(instructions preamble)         ← optional, free-form, parser ignores
LEGEND                          ← global, exactly once, first EL_* element
LEGEND - <Language>             ← optional override for this language
SUBJECT                         ← email subject line
Band N - <Language>             ← body content
```

### 3.3 Section ordering (legacy multi-language mode)

```
(instructions preamble)         ← optional, free-form, parser ignores
LEGEND                          ← global, exactly once, first EL_* element
LEGEND - French                 ← optional overrides, any order
LEGEND - Japanese
SUBJECT                         ← subject lines (one EL_Subject per language, labeled)
Band N - English                ← one per language, any order
Band N - French
Band N - German
...
```

---

## 4. The LEGEND section

The legend maps token labels (which appear literally in body text) to their semantic meaning, example values, and downstream substitution targets.

### 4.1 Token conventions

Two token classes, distinguished by delimiter:

| Class | Delimiter | Example | Substituted with |
|---|---|---|---|
| Dynamic data | `[]` | `[QUARANTINE_DATE]` | A runtime value (date, number, identifier) |
| Link | `<>` | `<Self-Service Portal>` | Anchor with display text + URL target |

Tokens appear **literally** in body text across all three layers (DOCX, structured runs, markdown). Substitution happens only at render time (e.g., HTML email preview). Authoring and review surfaces never see substituted values.

### 4.1.1 Token provenance

The `[]` tokens in the DOCX are a human-readable convention for the authoring surface. In the deployed email templates, some of these correspond to Logic App variable inserts, identified by the `@{...}` syntax (e.g., `@{variables('device_user_name')}`). The kit detects Logic App tokens by matching the `@{...}` pattern, not by assuming all dynamic data tokens are Logic App variables — other insert mechanisms may exist. The mapping from DOCX `[]` tokens to their deployed syntax is a downstream rendering concern. The DOCX never contains `@{...}` syntax.

### 4.2 Sub-table structure

The LEGEND section contains two sub-tables, each preceded by a label paragraph and bounded by horizontal rules.

**Table detection pattern.** docx-js (the template generator) cannot produce custom table styles, so `EL_LegendTable` cannot be applied as a Word table style in the template. Instead, the parser identifies legend tables by position: the `EL_Rule` → table → `EL_Rule` sandwich, immediately following an `EL_LegendLabel` paragraph. The full sequence for each sub-table is:

```
EL_LegendLabel   →  "Dynamic Data Tokens"
EL_Rule          →  ────────────────────
                     (table)
EL_Rule          →  ────────────────────
```

`EL_Rule` is an empty paragraph with a visible bottom border. It serves as both a visual separator for the author and a structural marker for the parser. The parser hard-fails if a table appears without the `EL_Rule` sandwich, or if `EL_Rule` appears without an adjacent table.

**Dynamic Data Tokens** (label styled `EL_LegendLabel`, table bounded by `EL_Rule`)

| Token | Example Value | Description |
|---|---|---|
| `[DEVICE_USER_NAME]` | `Jane Smith` | Display name of the device's assigned user |
| `[DEVICE_MODEL]` | `MacBook Pro 14"` | Model name of the Lilly-issued device |
| `[DEVICE_SERIAL_NUMBER]` | `ABC-12345` | Device serial number |

- **Token**: literal as it appears in body text, including brackets.
- **Example Value**: representative substituted value. Used in preview rendering and as documentation.
- **Description**: human-readable explanation of what the token represents.

**Link Tokens** (label styled `EL_LegendLabel`, table bounded by `EL_Rule`)

| Token | Display Text | Target URL | Description |
|---|---|---|---|
| `<Device Return Process>` | device return process | `https://lilly.service-now.com/kb_view.do?sysparm_article=KB2053555` | KB article for returning Lilly-issued devices |
| `<Request to Unblock Device>` | Request to Unblock Device from Quarantine | `https://lilly.service-now.com/ec?id=sc_cat_item&sys_id=a532e7d92b18f290c50df629fe91bfa4` | ServiceNow catalog item to restore quarantined device |
| `<ChatNow>` | ChatNow in Teams | `https://teams.microsoft.com/l/entity/d1417b70-54e4-441c-925a-d0c9d9234c34/conversations?tenantId=18a59a81-eea8-4c30-948a-d8824cdc2580` | Teams-based live support channel |

- **Token**: literal as it appears in body text, including angle brackets.
- **Display Text**: clickable text shown to the end user. Translatable per language via overrides.
- **Target URL**: resolved URL. Generally not translatable; recorded here for review and audit.
- **Description**: human-readable explanation.

### 4.3 Per-language legend overrides

A `LEGEND - <Language>` section contains the same two sub-tables (same `EL_LegendLabel` + `EL_LegendTable` styling) but includes **only the rows that differ from global**.

Rules:

- An override row **MUST** reference a token defined in the global LEGEND. New tokens cannot be introduced in overrides. (Hard rule.)
- An override may supply any subset of non-token columns. Empty columns inherit the global value.
- A token without an override entry uses global values for that language.

Example: a French override for `[QUARANTINE_DATE]` might supply `15 mars 2026` as the example value while inheriting the global description.

---

## 5. Subject line section

The subject line section is introduced by an `EL_SubjectHeading` paragraph containing the literal text `SUBJECT`. The next paragraph, styled `EL_Subject`, contains the email subject line text for this language.

```
EL_SubjectHeading  →  SUBJECT
EL_Subject         →  Action Required: Your Lilly Device Has Been Quarantined
```

In per-language mode, there is one subject line per file. In legacy multi-language mode, multiple `EL_Subject` paragraphs may follow the heading, each implicitly ordered to match the band section language order.

The parser emits the subject line as a `subject` field on the band section output, not as a body paragraph.

Token labels (`[]`, `<>`) may appear in subject lines and are subject to the same legend cross-check rules as body text.

---

## 6. Band sections

Each band section (`Band N - <Language>`, styled `EL_BandHeading`) contains the body content of one band's email in one language.

### 6.1 Paragraph styles

| Content type | Style | Notes |
|---|---|---|
| Standard body paragraph | `EL_Body` | Default for prose content |
| Sub-heading | `EL_SubHeading` | Bold label line that introduces a logical block within a band section (e.g., "Required action", "Exceptions", "Why it matters"). Always its own paragraph, always entirely bold. |
| Bullet list item | `EL_Bullet` | One paragraph per bullet |
| Numbered list item | `EL_Numbered` | One paragraph per numbered item |

**Emphasis: paragraph-level vs. character-level.** `EL_SubHeading` is a paragraph-level style — the entire paragraph is bold and functions as a structural label. This is distinct from character-level bold/italic applied to words or phrases *within* an `EL_Body`, `EL_Bullet`, or `EL_Numbered` paragraph. Bolding three important words mid-sentence is character formatting; a standalone bold line above a paragraph is `EL_SubHeading`. The parser distinguishes these by style, not by inspecting whether the bold spans the full paragraph.

Token labels (`[]`, `<>`) appear literally and may be inside or outside emphasis runs.

### 6.2 Compound elements in source HTML

The source HTML sometimes combines a sub-heading and its body text in a single `<p>` with a `<br>` between them (e.g., `<strong>Required action</strong><br>Return this device...`). In the DOCX, these are **always split** into two paragraphs: one `EL_SubHeading`, one `EL_Body`. The DOCX is a source of truth, not a 1:1 HTML mirror.

### 6.3 Token usage in body

Every `[TOKEN]` and `<TOKEN>` in body text **MUST** be defined in the global LEGEND. The parser hard-fails on undefined tokens.

A token may appear multiple times in a band section. A globally-defined token unused in any band section produces a soft warning (likely stale).

---

## 7. Markdown projection

The kit emits one `.md` per band per language at `sniffer/<round>/<band>/<lang>.md`.

### 7.1 Conventions

| DOCX element | Markdown |
|---|---|
| `EL_BandHeading` text | `# <heading text>` |
| `EL_Subject` paragraph | `**Subject:** <text>` (first line, before body) |
| `EL_SubHeading` paragraph | `## <text>` |
| `EL_Body` paragraph | Plain paragraph |
| `EL_Bullet` paragraph | `- <text>` |
| `EL_Numbered` paragraph | `1. <text>` (renumbered sequentially) |
| Bold run | `**text**` |
| Italic run | `*text*` |
| Bold-italic run | `***text***` |
| `[TOKEN]` | `[TOKEN]` (literal, no escaping) |
| `<TOKEN>` | `<TOKEN>` (literal, no escaping) |

The LEGEND section is **not** projected to per-band markdown files. It is parsed into a separate `legend.json` at `sniffer/<round>/<band>/legend.json` for kit consumption and audit.

### 7.2 Intentionally discarded

The markdown projection deliberately drops:

- Font, size, color (downstream styling concern)
- Track Changes / revision marks (rejected at parse time; see §8)
- Comments (rejected at parse time)
- Paragraph alignment

---

## 8. Parser contract

### 8.1 Strict on styles

The parser **requires** `EL_*` styles on all structural elements. A paragraph or table using a non-`EL_` style where an `EL_*` style is expected produces a **warning naming the offending paragraph** and the parser **rejects the document**.

No fallback heuristics. No silent guessing from Word's underlying XML. The parser stays strict; the operator is the soft layer. Workflow: receive imperfect work from the product owner → re-apply matching `EL_*` styles → re-parse.

### 8.2 Hard rules (parser rejects)

| Rule | Rationale |
|---|---|
| Global LEGEND section missing | No token reference; nothing to validate against |
| Global LEGEND not first section | Ordering contract |
| Required band section missing | Incomplete translation set |
| Section heading text doesn't match documented patterns (§3.1) | Structural ambiguity |
| Token in body text not defined in global LEGEND | Undefined token; likely typo |
| Per-language override defines token absent from global LEGEND | Override can't introduce new tokens |
| Track Changes present | Revision marks must be accepted/rejected before parse |
| Comments present | Must be resolved/removed before parse |
| Non-`EL_` style on a structural element | Style contract; operator re-styles and re-parses |
| Table in LEGEND section without `EL_Rule` sandwich | Table detection relies on `EL_Rule` → table → `EL_Rule` pattern (see §4.2) |
| `EL_Rule` paragraph without adjacent table | Orphaned rule; structural ambiguity |

On hard-rule failure: non-zero exit, **all** violations reported in one pass (parser does not stop at first error), no output emitted.

### 8.3 Soft rules (parser warns, continues)

| Rule | Rationale |
|---|---|
| Globally-defined token unused in any band section | Likely stale legend entry |
| Per-language override row identical to global row | Override is a no-op |
| Empty band section (heading present, no body) | Possibly incomplete; worth flagging |

### 8.4 Output

On success:

- Structured runs (in-memory or serialized to JSON)
- Markdown projection (one `.md` per language)
- Parsed legend (`legend.json`)

On hard-rule failure:

- Non-zero exit code
- All violations listed
- No structured output, no markdown, no legend emitted

---

## 9. Validator (separate concern)

The parser produces structured output. A separate validator runs on top of structured output for semantic cross-checks beyond schema:

- All nine languages present
- Every token example value is non-empty
- Link token target URLs parse as valid URLs
- (Future) cross-band token consistency

The validator is a separate tool with its own spec; documented here only to clarify the boundary with the parser.

---

## 10. Reserved styles reference

| Style | Base | Purpose |
|---|---|---|
| `EL_LegendHeading` | Heading 2 | `LEGEND` and `LEGEND - <Language>` section markers |
| `EL_BandHeading` | Heading 2 | `Band N - <Language>` section markers |
| `EL_SubjectHeading` | Heading 2 | `SUBJECT` section marker |
| `EL_Subject` | Normal | Email subject line text |
| `EL_SubHeading` | Normal (bold) | Structural label within a band section (e.g., "Required action", "Why it matters") |
| `EL_LegendLabel` | Normal (bold) | Sub-table labels: "Dynamic Data Tokens", "Link Tokens" |
| `EL_Rule` | Normal (empty, bottom border) | Horizontal rule bounding legend tables above and below. Visual separator for the author, structural marker for the parser. |
| `EL_Body` | Normal | Body paragraphs in band sections |
| `EL_Bullet` | List Bullet | Bullet list items in band sections |
| `EL_Numbered` | List Number | Numbered list items in band sections |

The `EL_` prefix is the parser's signal that a style is kit-reserved. It avoids collisions with Eli Lilly corporate template styles and Word defaults by construction.

---

## 11. Out of scope

- Parser implementation (separate session)
- Markdown normalizer implementation (separate session)
- Validator implementation (separate session)
- Legacy spreadsheet → template DOCX converter (separate session, depends on spreadsheet)
- HTML rendering pipeline (downstream; consumes parsed legend + structured runs)
- Branch A reconciliation (separate concern)

---

## 12. Open questions

1. ~~**Canonical language list.**~~ **Resolved.** English, French, German, Italian, Spanish, Portuguese, Japanese, Chinese, Korean. Stable; unlikely to vary between rounds.
2. ~~**Link token URL representation.**~~ **Resolved.** Legend stores literal URL when possible. When the URL is config-driven and the literal isn't available, the description column notes this. No config-key indirection in the legend itself.
3. **Nested emphasis.** Word allows bold-inside-italic and vice versa. The markdown projection handles this (`***text***`), but the validator may want a rule about whether nested emphasis is intentional or accidental. Deferred until real-world examples surface.
