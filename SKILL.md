---
name: markdown-to-docx
description: Convert Markdown articles into standardized DOCX/Word documents using Pandoc by default, installing Pandoc for the user when it is missing, then applying academic-style typography, spacing, headings, abstracts, keywords, captions, tables, references, and Chinese-first formatting rules. Use when Codex needs to create a .docx from .md or Markdown text and apply a repeatable Word layout standard, especially for Chinese articles; English-only article rules are not finalized yet.
---

# Markdown To DOCX

## Workflow

Use Pandoc as the default Markdown-to-DOCX engine.

1. Check Pandoc first:

```powershell
pandoc --version
```

2. If Pandoc is missing, install it for the user. On Windows, prefer:

```powershell
winget install --id JohnMacFarlane.Pandoc -e --silent --accept-package-agreements --accept-source-agreements
```

If `winget` is unavailable, try `choco install pandoc -y` or `scoop install pandoc` when those package managers exist. Request escalation/approval when the install command needs network or system access.

3. Convert with the bundled wrapper, which runs Pandoc and then applies the Chinese article DOCX style using `python-docx`:

```powershell
C:\Users\10124\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe C:\Users\10124\.codex\skills\markdown-to-docx\scripts\markdown_to_docx.py input.md output.docx
```

The wrapper also attempts to install Pandoc automatically if it is missing. Use `--no-install-pandoc` only when the user explicitly wants a read-only check or no installation.

Prefer the Codex bundled Python because it includes `python-docx`. If it is unavailable, use any Python environment with `python-docx` installed.

## Input Conventions

Accept ordinary Markdown:

- `#` is the main Chinese title unless a document already has a title before section headings.
- An optional English title can be marked as a second `#` line immediately after the Chinese title, or as `英文题目：...`.
- `## 摘要` / `## 中文摘要` starts the Chinese abstract.
- `## Abstract` / `## 英文摘要` starts the English abstract.
- `关键词：...` and `Keywords: ...` are formatted as keyword paragraphs.
- `## 目录` inserts a table-of-contents heading. When a live Word TOC is required, generate or preserve the Pandoc/Word field and tell the user Word may need right-click > update field after opening.
- `#`, `##`, and `###` after the title map to level 1, level 2, and level 3 headings.
- Figure and table captions beginning with `图`, `Figure`, `表`, or `Table` are centered caption paragraphs.
- `## 参考文献` / `## References` starts the references section; following paragraphs use the reference style.
- Markdown tables are converted by Pandoc and then table text spacing is normalized.

## Formatting Standard

Load `references/chinese-article-format.md` when exact spacing, font size, or paragraph behavior is needed. The wrapper implements that Chinese article standard after Pandoc conversion:

- Chinese body font: SimSun/宋体; English body font: Times New Roman; black text.
- Body: 小四/12 pt, left aligned, first-line indent 2 Chinese characters, fixed 22 pt line spacing.
- References: 五号/10.5 pt, fixed 20 pt line spacing, hanging indent 2.5 characters.
- Headings, abstracts, keywords, captions, tables, and title rules follow the reference file.

## Quality Checks

After conversion, open or inspect the DOCX enough to verify:

- The file is not empty and contains the expected title, sections, captions, tables, and references.
- Markdown syntax artifacts such as `#`, `**`, and table pipes do not remain in normal prose.
- Pandoc handled Markdown features such as tables, images, lists, links, and code blocks correctly.
- English-only document formatting may need user confirmation because the English article standard is still pending.
