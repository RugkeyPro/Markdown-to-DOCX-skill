#!/usr/bin/env python
"""Convert Markdown to DOCX with Pandoc, then apply the Chinese article style."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


FONT_CN = "SimSun"
FONT_CN_HEADING = "SimHei"
FONT_EN = "Times New Roman"
BLACK = RGBColor(0, 0, 0)


def run(cmd: list[str], *, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, check=check, text=True)


def find_pandoc() -> str | None:
    found = shutil.which("pandoc")
    if found:
        return found

    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Pandoc" / "pandoc.exe",
        Path(os.environ.get("ProgramFiles", "")) / "Pandoc" / "pandoc.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "Pandoc" / "pandoc.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def install_pandoc() -> str:
    if os.name == "nt":
        if shutil.which("winget"):
            run([
                "winget",
                "install",
                "--id",
                "JohnMacFarlane.Pandoc",
                "-e",
                "--silent",
                "--accept-package-agreements",
                "--accept-source-agreements",
            ])
        elif shutil.which("choco"):
            run(["choco", "install", "pandoc", "-y"])
        elif shutil.which("scoop"):
            run(["scoop", "install", "pandoc"])
        else:
            raise RuntimeError("Pandoc is missing and no supported Windows package manager was found.")
    elif sys.platform == "darwin":
        if not shutil.which("brew"):
            raise RuntimeError("Pandoc is missing and Homebrew was not found.")
        run(["brew", "install", "pandoc"])
    else:
        if shutil.which("apt-get"):
            run(["sudo", "apt-get", "update"])
            run(["sudo", "apt-get", "install", "-y", "pandoc"])
        elif shutil.which("dnf"):
            run(["sudo", "dnf", "install", "-y", "pandoc"])
        elif shutil.which("pacman"):
            run(["sudo", "pacman", "-S", "--noconfirm", "pandoc"])
        else:
            raise RuntimeError("Pandoc is missing and no supported Linux package manager was found.")

    pandoc = find_pandoc()
    if not pandoc:
        raise RuntimeError("Pandoc installation completed, but pandoc was not found on PATH.")
    return pandoc


def set_run_font(run, size_pt=12, cn_font=FONT_CN, en_font=FONT_EN, bold=False):
    run.font.name = en_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = BLACK


def char_width(chars: float, size_pt: float = 12):
    return Pt(chars * size_pt)


def set_para(paragraph, align=None, before=0, after=0, line=None, first=None, hanging=None):
    fmt = paragraph.paragraph_format
    if align is not None:
        paragraph.alignment = align
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    if line is None:
        fmt.line_spacing = 1
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        fmt.line_spacing = Pt(line)
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    fmt.first_line_indent = first
    if hanging is not None:
        fmt.left_indent = hanging
        fmt.first_line_indent = -hanging


def apply_font(paragraph, size=12, cn_font=FONT_CN, en_font=FONT_EN, bold=False):
    for run in paragraph.runs:
        set_run_font(run, size, cn_font, en_font, bold)


def configure_style(doc: Document, name: str, *, size=12, cn_font=FONT_CN, en_font=FONT_EN,
                    bold=False, before=0, after=0, line=22, align=None):
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = en_font
    style._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = BLACK
    fmt = style.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    if line is None:
        fmt.line_spacing = 1
        fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        fmt.line_spacing = Pt(line)
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    if align is not None:
        fmt.alignment = align


def create_reference_docx(path: Path):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.1)
    section.right_margin = Inches(1.1)

    configure_style(doc, "Normal", size=12, cn_font=FONT_CN, en_font=FONT_EN, line=22)
    configure_style(doc, "Body Text", size=12, cn_font=FONT_CN, en_font=FONT_EN, line=22)
    configure_style(doc, "Title", size=22, cn_font=FONT_CN, en_font=FONT_EN, bold=True,
                    line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
    configure_style(doc, "Subtitle", size=18, cn_font=FONT_CN, en_font=FONT_EN, bold=True,
                    line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
    configure_style(doc, "Heading 1", size=16, cn_font=FONT_CN_HEADING, en_font=FONT_EN,
                    bold=True, before=12, after=12, line=None, align=WD_ALIGN_PARAGRAPH.CENTER)
    configure_style(doc, "Heading 2", size=12, cn_font=FONT_CN_HEADING, en_font=FONT_EN,
                    bold=True, before=6, after=6, line=None, align=WD_ALIGN_PARAGRAPH.LEFT)
    configure_style(doc, "Heading 3", size=12, cn_font=FONT_CN, en_font=FONT_EN,
                    bold=True, before=6, after=6, line=None, align=WD_ALIGN_PARAGRAPH.LEFT)
    configure_style(doc, "Caption", size=10.5, cn_font=FONT_CN, en_font=FONT_EN,
                    line=None, before=6, after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
    configure_style(doc, "TOC Heading", size=18, cn_font=FONT_CN, en_font=FONT_EN,
                    bold=True, line=None, before=6, after=18, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph("Reference document for Pandoc styles.")
    doc.save(path)


def normalized(text: str) -> str:
    return re.sub(r"\s+", "", text).lower()


def paragraph_kind(paragraph, mode: str) -> tuple[str, str]:
    text = paragraph.text.strip()
    norm = normalized(text)
    style = paragraph.style.name if paragraph.style else ""

    if style == "Title":
        return "title_cn", mode
    if style == "Subtitle":
        return "title_en", mode
    if norm in {"摘要", "中文摘要", "摘要:"}:
        return "abstract_title_cn", "abstract_cn"
    if norm in {"abstract", "英文摘要"}:
        return "abstract_title_en", "abstract_en"
    if norm in {"目录", "contents", "tableofcontents"}:
        return "toc_title", "body"
    if norm in {"参考文献", "references"}:
        return "heading1", "refs"
    if re.match(r"^(关键词|关键字)[:：]", text):
        return "keywords_cn", mode
    if re.match(r"^(keywords|key words)\s*[:：]", text, re.I):
        return "keywords_en", mode
    if re.match(r"^(图|figure|表|table)\s*[\d一二三四五六七八九十]+", text, re.I):
        return "caption", mode
    if style == "Heading 1":
        return "heading1", "body"
    if style == "Heading 2":
        return "heading2", "body"
    if style == "Heading 3":
        return "heading3", "body"
    if mode == "abstract_cn":
        return "abstract_body_cn", mode
    if mode == "abstract_en":
        return "abstract_body_en", mode
    if mode == "refs":
        return "reference", mode
    return "body", mode


def apply_chinese_article_format(docx_path: Path):
    doc = Document(docx_path)
    mode = "body"
    for paragraph in doc.paragraphs:
        if not paragraph.text.strip():
            continue
        kind, mode = paragraph_kind(paragraph, mode)
        if kind == "title_cn":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=12, line=None)
            apply_font(paragraph, 22, FONT_CN, FONT_EN, True)
        elif kind == "title_en":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=12, line=None)
            apply_font(paragraph, 18, FONT_CN, FONT_EN, True)
        elif kind == "abstract_title_cn":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=18, line=None)
            apply_font(paragraph, 18, FONT_CN, FONT_EN, True)
        elif kind == "abstract_title_en":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=18, line=None)
            apply_font(paragraph, 18, FONT_CN, FONT_EN, True)
        elif kind == "keywords_cn":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=12, after=0, line=22)
            apply_font(paragraph, 12, FONT_CN, FONT_EN, False)
            if paragraph.runs:
                set_run_font(paragraph.runs[0], 12, FONT_CN, FONT_EN, True)
        elif kind == "keywords_en":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=12, after=0, line=22)
            apply_font(paragraph, 12, FONT_CN, FONT_EN, False)
            if paragraph.runs:
                set_run_font(paragraph.runs[0], 12, FONT_CN, FONT_EN, True)
        elif kind == "toc_title":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=18, line=None)
            apply_font(paragraph, 18, FONT_CN, FONT_EN, True)
        elif kind == "heading1":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=12, after=12, line=None)
            apply_font(paragraph, 16, FONT_CN_HEADING, FONT_EN, True)
        elif kind == "heading2":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=6, after=6, line=None)
            apply_font(paragraph, 12, FONT_CN_HEADING, FONT_EN, True)
        elif kind == "heading3":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=6, after=6, line=None)
            apply_font(paragraph, 12, FONT_CN, FONT_EN, True)
        elif kind == "caption":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.CENTER, before=6, after=6, line=None)
            apply_font(paragraph, 10.5, FONT_CN, FONT_EN, False)
        elif kind == "reference":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0, line=20,
                     hanging=char_width(2.5, 10.5))
            apply_font(paragraph, 10.5, FONT_CN, FONT_EN, False)
        elif kind == "abstract_body_en":
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0, line=22,
                     first=char_width(1, 12))
            apply_font(paragraph, 12, FONT_CN, FONT_EN, False)
        else:
            set_para(paragraph, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0, line=22,
                     first=char_width(2, 12))
            apply_font(paragraph, 12, FONT_CN, FONT_EN, False)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    set_para(paragraph, before=3, after=3, line=None)
                    apply_font(paragraph, 10.5, FONT_CN, FONT_EN, False)

    doc.save(docx_path)


def convert(markdown_path: Path, output_path: Path, *, install_if_missing: bool = True):
    pandoc = find_pandoc()
    if not pandoc:
        if not install_if_missing:
            raise RuntimeError("Pandoc is missing. Re-run without --no-install-pandoc to install it.")
        pandoc = install_pandoc()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        reference_doc = Path(tmp) / "reference.docx"
        create_reference_docx(reference_doc)
        run([
            pandoc,
            str(markdown_path),
            "--from",
            "markdown+pipe_tables+implicit_figures",
            "--to",
            "docx",
            "--standalone",
            "--reference-doc",
            str(reference_doc),
            "-o",
            str(output_path),
        ])
    apply_chinese_article_format(output_path)


def main():
    parser = argparse.ArgumentParser(description="Convert Markdown to standardized DOCX with Pandoc.")
    parser.add_argument("input", type=Path, help="Input UTF-8 Markdown file.")
    parser.add_argument("output", type=Path, help="Output DOCX file.")
    parser.add_argument("--no-install-pandoc", action="store_true",
                        help="Fail instead of installing Pandoc when it is missing.")
    args = parser.parse_args()
    convert(args.input, args.output, install_if_missing=not args.no_install_pandoc)


if __name__ == "__main__":
    main()
