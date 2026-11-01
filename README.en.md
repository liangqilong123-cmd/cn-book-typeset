# cn-book-typeset — Markdown → Print-Ready DOCX for Chinese Books

[中文](README.md) · English

> **Turn a folder of Markdown into an actual book.** Standard Chinese 32-mo trim size (147×210 mm), mirrored margins, running heads that alternate between chapter title and book title, page numbers on the outer edge, and a live Word TOC field.

Not "a Word doc with a heading" — it follows the typographic conventions real publishers use.

## Why this exists

Laying out a book by hand in Word breaks in predictable places:

| Problem | Manual Word | This tool |
|---|---|---|
| Binding margin | Set inner/outer per section, redo everything on edit | `mirrorMargins` on, margins mirror automatically |
| Running heads (odd = chapter, even = book) | Odd/even headers + one section per chapter + unlink each | One section per chapter, headers follow the chapter |
| Page numbers at the trim edge | Right on recto, left on verso, easy to misalign | Page field, side chosen automatically |
| TOC page numbers | Insert TOC, forget F9, numbers are wrong | TOC field inserted; open, Ctrl+A, F9 |
| **Real phone numbers / emails in the manuscript** | Nobody checks page by page | Auto-redaction + a pre-save assertion that fails the build |

That last one is what makes this unusual — see [Redaction](#redaction--self-check).

## Specs

| Item | Value |
|---|---|
| Trim size | 147 × 210 mm (Chinese 32-mo; ≈ A5) |
| Margins | top 20 / bottom 18 / **inner 20 / outer 16 mm** (mirrored) |
| Body | SimSun 10.5pt, 1.5 line spacing, 2-em first-line indent |
| H1 (chapter) | SimHei 16pt, starts a new page |
| H2 (section) | SimHei 13pt |
| Epigraph / quote | KaiTi, centered (Chinese publishing convention) |
| Running head | 9pt grey with a hairline rule |
| Folio | 9pt, bottom-right on recto, bottom-left on verso |

## Quick start

```bash
pip install python-docx pyyaml

# Run the bundled example (two public-domain classical essays)
python build_book.py --config examples/book.yaml

# Sample only — first chapter
python build_book.py --config examples/book.yaml --sample
```

Output: `examples/out/样张_32开.docx`. Open it, press **Ctrl+A then F9** to generate the TOC page numbers.

## Using it for your own book

```
mybook/
├─ book.yaml
├─ 01_chapter-one.md
├─ 02_chapter-two.md
└─ img/
   └─ img_001.png
```

```yaml
book:
  title: MY  BOOK
  subtitle: —— subtitle ——
  running_head_even: MY  BOOK
  running_head_odd_template: "{chapter}"
  publisher_notes:
    - This edition is for internal circulation only.

chapters:
  - file: 01_chapter-one.md
  - file: 02_chapter-two.md

layout:
  body_size_pt: 10.5
  line_spacing: 1.5

pii:
  whitelist: ["13800001111"]

verify:
  must_contain: ["for internal circulation only"]

output: out/mybook_32mo.docx
```

**Chapter titles are read from each file's H1** — you don't repeat them in the config, so you can't drift out of sync (a mistake the author made once and won't make again).

## Supported Markdown

Deliberately narrow — book layout doesn't need more.

| Syntax | Result |
|---|---|
| `# Title` | H1, page break before |
| `## Title` | H2 |
| paragraph | body text, 2-em indent |
| `> quote` / `- item` | body paragraph |
| `![alt](img_005.png)` | figure + caption "图 5" |
| **short line** (≤14 chars, no terminal punctuation) | auto-detected as epigraph, KaiTi centered |

The last rule exists because classical Chinese texts put aphorisms and inscriptions on their own short lines, and they only look right in KaiTi, centered.

## Redaction & self-check

Manuscripts routinely contain real contact details — memoirs, family histories, interview transcripts especially. Once those land in a PDF, they're public.

`pii_guard.py` gives you two layers:

```python
from pii_guard import mask, audit

clean, n = mask(text, whitelist={'13800001111'})   # mobile / email / landline / ID → placeholder
leaks = audit(clean, whitelist={'13800001111'})
assert not leaks, f'still contains PII: {leaks}'
```

The build pipeline calls it three times:

1. **While parsing body Markdown**, line by line
2. **On front matter** — title, subtitle, publisher notes, running heads (the part everyone forgets)
3. **Before saving**, scanning the entire document XML; any leak aborts the save rather than writing a half-clean file

Plus `verify.must_contain` checks that required legal text survived:

```
✗ 必含内容缺失（可能是免责声明被删）：['不构成任何医疗建议']
```

## Pitfalls already handled

These are from real typesetting sessions, not theory:

1. **`add_section()` inherits the previous `pgNumType`** — new sections default to `start=1`, so continuous pagination requires explicitly removing that element, or every chapter restarts at page 1
2. **`mirrorMargins` must be written into `settings.xml`** — setting left/right margins alone does not mirror them
3. **Alternate headers need `odd_and_even_pages_header_footer`**, and every section's `even_page_header` must be unlinked separately
4. **Chinese fonts need two properties**: `run.font.name` (Latin) *and* `rFonts.set(qn('w:eastAsia'))` (CJK). Miss one and it silently falls back
5. **Read chapter titles from the source, never hand-write them** — a hand-written list once mislabelled four consecutive chapters
6. **First-line indent as `Pt(size*2)`, not `Cm(0.74)`** — points track the font size; centimeters don't

## Limitations

- **DOCX only**, no direct PDF. Let Word / WPS / LibreOffice do that conversion — fidelity is better
- **TOC needs one keystroke (F9)** — it's a Word field; page numbers are computed by Word
- **No cover design** — do that in Word or a design tool
- **No footnotes, index, or cross-references** — add those in the output DOCX
- **Epigraph detection is heuristic.** For precise control, put epigraphs under a `## 题识` heading

## License

Code: MIT. Example texts are public domain (Tao Yuanming, *The Peach Blossom Spring*; Fan Zhongyan, *Record of Yueyang Tower*).
