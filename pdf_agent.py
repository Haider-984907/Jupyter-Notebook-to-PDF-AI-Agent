#!/usr/bin/env python3
"""
pdf_agent.py — Jupyter Notebook → Professional PDF Report Generator

Converts any .ipynb file into a polished, ready-to-submit PDF report
with syntax-highlighted code, formatted markdown, styled output blocks,
a table of contents, header/footer, and page numbers.

Usage:
    python pdf_agent.py notebook.ipynb [--output report.pdf] [--title "My Report"]
"""

import json
import re
import sys
import os
import argparse
import base64
import io
from pathlib import Path
from datetime import datetime

# ── ReportLab ────────────────────────────────────────────────────────────────
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, KeepTogether, Image as RLImage, Preformatted
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── Pygments (syntax highlighting) ───────────────────────────────────────────
from pygments import highlight
from pygments.lexers import PythonLexer, get_lexer_by_name
from pygments.formatters import HtmlFormatter
from pygments.styles import get_style_by_name

# ── PIL ───────────────────────────────────────────────────────────────────────
from PIL import Image as PILImage

# ─────────────────────────────────────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────
C_BRAND      = colors.HexColor("#2D3A8C")   # deep indigo — headings / header
C_ACCENT     = colors.HexColor("#4F6FD8")   # medium blue — subheadings / rules
C_CODE_BG    = colors.HexColor("#1E1E2E")   # dark bg for code blocks
C_CODE_FG    = colors.HexColor("#CDD6F4")   # light text for code
C_OUTPUT_BG  = colors.HexColor("#F0F4FF")   # pale blue — output blocks
C_OUTPUT_BD  = colors.HexColor("#C5D0F0")   # output border
C_ERROR_BG   = colors.HexColor("#FFF0F0")   # pale red — error blocks
C_ERROR_BD   = colors.HexColor("#F0B8B8")   # error border
C_ERROR_TXT  = colors.HexColor("#C0392B")   # red text for errors
C_TABLE_HDR  = colors.HexColor("#2D3A8C")   # table header bg
C_TABLE_ALT  = colors.HexColor("#EEF1FB")   # alternating table row
C_MUTED      = colors.HexColor("#6B7280")   # muted grey — footer / meta
C_WHITE      = colors.white
C_BLACK      = colors.HexColor("#1A1A2E")
C_RULE       = colors.HexColor("#D1D5DB")


# ─────────────────────────────────────────────────────────────────────────────
# STYLE FACTORY
# ─────────────────────────────────────────────────────────────────────────────

def build_styles():
    """Return a dict of all named ParagraphStyles used in the PDF."""
    base = getSampleStyleSheet()

    def S(name, **kw):
        return ParagraphStyle(name, **kw)

    return {
        # ── Document title (cover page) ──────────────────────────────────
        "cover_title": S("cover_title",
            fontName="Helvetica-Bold", fontSize=28,
            textColor=C_BRAND, spaceAfter=6, leading=34,
            alignment=TA_CENTER),

        "cover_sub": S("cover_sub",
            fontName="Helvetica", fontSize=13,
            textColor=C_ACCENT, spaceAfter=4,
            alignment=TA_CENTER),

        "cover_meta": S("cover_meta",
            fontName="Helvetica", fontSize=10,
            textColor=C_MUTED, spaceAfter=2,
            alignment=TA_CENTER),

        # ── Headings ─────────────────────────────────────────────────────
        "h1": S("h1",
            fontName="Helvetica-Bold", fontSize=18,
            textColor=C_BRAND, spaceBefore=18, spaceAfter=6, leading=22),

        "h2": S("h2",
            fontName="Helvetica-Bold", fontSize=14,
            textColor=C_ACCENT, spaceBefore=14, spaceAfter=4, leading=18),

        "h3": S("h3",
            fontName="Helvetica-Bold", fontSize=12,
            textColor=C_BLACK, spaceBefore=10, spaceAfter=3, leading=15),

        "h4": S("h4",
            fontName="Helvetica-BoldOblique", fontSize=11,
            textColor=C_BLACK, spaceBefore=8, spaceAfter=2, leading=14),

        # ── Body text ────────────────────────────────────────────────────
        "body": S("body",
            fontName="Helvetica", fontSize=10,
            textColor=C_BLACK, spaceAfter=6, leading=15,
            alignment=TA_JUSTIFY),

        "body_bold": S("body_bold",
            fontName="Helvetica-Bold", fontSize=10,
            textColor=C_BLACK, spaceAfter=6, leading=15),

        "blockquote": S("blockquote",
            fontName="Helvetica-Oblique", fontSize=10,
            textColor=colors.HexColor("#374151"),
            leftIndent=18, borderPadding=(4, 8, 4, 8),
            spaceAfter=8, leading=14,
            borderColor=C_ACCENT, borderWidth=2,
            backColor=colors.HexColor("#F8F9FF")),

        # ── Code ─────────────────────────────────────────────────────────
        "code_label": S("code_label",
            fontName="Helvetica-Bold", fontSize=8,
            textColor=colors.HexColor("#A6ADC8"),
            spaceAfter=0, leading=10),

        "inline_code": S("inline_code",
            fontName="Courier", fontSize=9,
            textColor=C_CODE_FG, backColor=C_CODE_BG,
            borderPadding=2, leading=12),

        # ── Output ───────────────────────────────────────────────────────
        "output_text": S("output_text",
            fontName="Courier", fontSize=8.5,
            textColor=colors.HexColor("#1F2937"),
            backColor=C_OUTPUT_BG,
            borderPadding=(4, 6, 4, 6),
            leading=12, spaceAfter=4),

        "error_text": S("error_text",
            fontName="Courier-Bold", fontSize=8.5,
            textColor=C_ERROR_TXT,
            backColor=C_ERROR_BG,
            borderPadding=(4, 6, 4, 6),
            leading=12, spaceAfter=4),

        # ── Table of Contents ─────────────────────────────────────────────
        "toc_h1": S("toc_h1",
            fontName="Helvetica-Bold", fontSize=11,
            textColor=C_BRAND, spaceAfter=3, leading=14,
            leftIndent=0),

        "toc_h2": S("toc_h2",
            fontName="Helvetica", fontSize=10,
            textColor=C_BLACK, spaceAfter=2, leading=13,
            leftIndent=14),

        "toc_h3": S("toc_h3",
            fontName="Helvetica-Oblique", fontSize=9,
            textColor=C_MUTED, spaceAfter=1, leading=12,
            leftIndent=28),

        # ── List items ────────────────────────────────────────────────────
        "li": S("li",
            fontName="Helvetica", fontSize=10,
            textColor=C_BLACK, spaceAfter=3, leading=14,
            leftIndent=16, firstLineIndent=0),

        "li2": S("li2",
            fontName="Helvetica", fontSize=10,
            textColor=C_BLACK, spaceAfter=2, leading=13,
            leftIndent=32, firstLineIndent=0),

        # ── Misc ──────────────────────────────────────────────────────────
        "caption": S("caption",
            fontName="Helvetica-Oblique", fontSize=8,
            textColor=C_MUTED, alignment=TA_CENTER, spaceAfter=6),

        "toc_title": S("toc_title",
            fontName="Helvetica-Bold", fontSize=16,
            textColor=C_BRAND, spaceAfter=12, leading=20),
    }


# ─────────────────────────────────────────────────────────────────────────────
# NOTEBOOK PARSER
# ─────────────────────────────────────────────────────────────────────────────

class NotebookParser:
    """Parse a .ipynb JSON file into structured cell data."""

    def __init__(self, path: str):
        with open(path, "r", encoding="utf-8") as f:
            self.nb = json.load(f)
        self.nbformat = self.nb.get("nbformat", 4)
        self.metadata  = self.nb.get("metadata", {})
        self.cells     = self.nb.get("cells", [])
        self.kernel    = (self.metadata.get("kernelspec", {})
                                       .get("display_name", "Python 3"))
        self.lang_info = (self.metadata.get("language_info", {})
                                       .get("name", "python"))

    def parse(self):
        """Return list of dicts, one per cell."""
        parsed = []
        for idx, cell in enumerate(self.cells):
            ctype   = cell.get("cell_type", "")
            source  = self._join(cell.get("source", []))
            outputs = cell.get("outputs", [])
            ec      = cell.get("execution_count")

            parsed.append({
                "idx":    idx,
                "type":   ctype,
                "source": source,
                "outputs": [self._parse_output(o) for o in outputs],
                "execution_count": ec,
            })
        return parsed

    def _join(self, src):
        if isinstance(src, list):
            return "".join(src)
        return src or ""

    def _parse_output(self, out):
        otype = out.get("output_type", "")
        result = {"type": otype}

        if otype in ("stream",):
            result["text"] = self._join(out.get("text", []))
            result["stream_name"] = out.get("name", "stdout")

        elif otype in ("execute_result", "display_data"):
            data = out.get("data", {})
            result["text"]  = self._join(data.get("text/plain", []))
            result["html"]  = self._join(data.get("text/html",  []))
            result["image"] = data.get("image/png") or data.get("image/jpeg")
            result["img_fmt"] = "png" if "image/png" in data else "jpeg"

        elif otype == "error":
            result["ename"]     = out.get("ename", "Error")
            result["evalue"]    = out.get("evalue", "")
            result["traceback"] = out.get("traceback", [])

        return result


# ─────────────────────────────────────────────────────────────────────────────
# MARKDOWN → REPORTLAB FLOWABLES
# ─────────────────────────────────────────────────────────────────────────────

class MarkdownRenderer:
    """
    Lightweight Markdown → ReportLab flowable converter.
    Supports: H1–H4, bold/italic/inline-code, bullet lists,
    numbered lists, blockquotes, horizontal rules, and tables.
    """

    def __init__(self, styles):
        self.styles = styles

    def render(self, md_text: str) -> list:
        flowables = []
        lines = md_text.splitlines()
        i = 0

        while i < len(lines):
            line = lines[i]

            # ── Blank line ───────────────────────────────────────────────
            if not line.strip():
                i += 1
                continue

            # ── Horizontal rule ──────────────────────────────────────────
            if re.match(r"^[-*_]{3,}\s*$", line):
                flowables.append(Spacer(1, 4))
                flowables.append(HRFlowable(width="100%", thickness=0.8,
                                            color=C_RULE))
                flowables.append(Spacer(1, 4))
                i += 1
                continue

            # ── ATX Headings ─────────────────────────────────────────────
            m = re.match(r"^(#{1,4})\s+(.*)", line)
            if m:
                level = len(m.group(1))
                text  = self._inline(m.group(2).strip())
                sname = f"h{level}"
                flowables.append(Paragraph(text, self.styles[sname]))
                if level == 1:
                    flowables.append(
                        HRFlowable(width="100%", thickness=1.5,
                                   color=C_BRAND))
                elif level == 2:
                    flowables.append(
                        HRFlowable(width="60%", thickness=0.6,
                                   color=C_ACCENT))
                i += 1
                continue

            # ── Blockquote ───────────────────────────────────────────────
            if line.startswith(">"):
                bq_lines = []
                while i < len(lines) and lines[i].startswith(">"):
                    bq_lines.append(lines[i].lstrip("> ").strip())
                    i += 1
                text = self._inline(" ".join(bq_lines))
                flowables.append(Paragraph(text, self.styles["blockquote"]))
                continue

            # ── Markdown table ───────────────────────────────────────────
            if "|" in line and i + 1 < len(lines) and re.match(
                    r"^\|?[-| :]+\|?\s*$", lines[i + 1]):
                tbl_lines = [line]
                i += 2                         # skip separator row
                while i < len(lines) and "|" in lines[i]:
                    tbl_lines.append(lines[i])
                    i += 1
                flowables.extend(self._table(tbl_lines))
                continue

            # ── Unordered list ────────────────────────────────────────────
            if re.match(r"^(\s*)[-*+]\s+", line):
                while i < len(lines) and re.match(r"^(\s*)[-*+]\s+", lines[i]):
                    m2    = re.match(r"^(\s*)[-*+]\s+(.*)", lines[i])
                    indent = len(m2.group(1))
                    text   = self._inline(m2.group(2))
                    bullet = "•" if indent == 0 else "◦"
                    sname  = "li" if indent == 0 else "li2"
                    flowables.append(
                        Paragraph(f"{bullet}&nbsp;&nbsp;{text}",
                                  self.styles[sname]))
                    i += 1
                continue

            # ── Ordered list ──────────────────────────────────────────────
            if re.match(r"^\d+\.\s+", line):
                num = 1
                while i < len(lines) and re.match(r"^\d+\.\s+", lines[i]):
                    m2   = re.match(r"^\d+\.\s+(.*)", lines[i])
                    text = self._inline(m2.group(1))
                    flowables.append(
                        Paragraph(f"{num}.&nbsp;&nbsp;{text}",
                                  self.styles["li"]))
                    num += 1
                    i   += 1
                continue

            # ── Normal paragraph ──────────────────────────────────────────
            para_lines = []
            while i < len(lines) and lines[i].strip() and not (
                lines[i].startswith("#") or
                lines[i].startswith(">") or
                re.match(r"^[-*+]\s", lines[i]) or
                re.match(r"^\d+\.\s", lines[i]) or
                ("|" in lines[i] and i + 1 < len(lines) and
                 re.match(r"^\|?[-| :]+\|?\s*$", lines[i + 1]))
            ):
                para_lines.append(lines[i])
                i += 1

            text = self._inline(" ".join(para_lines))
            if text.strip():
                flowables.append(Paragraph(text, self.styles["body"]))

        return flowables

    def _inline(self, text: str) -> str:
        """Convert inline markdown (bold, italic, code, links) to ReportLab XML."""
        # Escape XML special chars first (except our own tags)
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        # Bold-italic
        text = re.sub(r"\*\*\*(.*?)\*\*\*",
                      r"<b><i>\1</i></b>", text)
        # Bold
        text = re.sub(r"\*\*(.*?)\*\*",
                      r"<b>\1</b>", text)
        text = re.sub(r"__(.*?)__",
                      r"<b>\1</b>", text)
        # Italic
        text = re.sub(r"\*(.*?)\*",
                      r"<i>\1</i>", text)
        text = re.sub(r"_(.*?)_",
                      r"<i>\1</i>", text)
        # Inline code
        text = re.sub(r"`([^`]+)`",
                      r'<font name="Courier" size="9">\1</font>', text)
        # Links → just show text
        text = re.sub(r"\[([^\]]+)\]\([^\)]+\)",
                      r'<u>\1</u>', text)

        return text

    def _table(self, raw_lines: list) -> list:
        """Convert markdown table lines into a ReportLab Table flowable."""
        rows = []
        for ln in raw_lines:
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            rows.append(cells)

        if not rows:
            return []

        col_count = max(len(r) for r in rows)
        # Pad rows
        data = [r + [""] * (col_count - len(r)) for r in rows]

        # Header row styling
        header = [Paragraph(f"<b>{c}</b>", ParagraphStyle(
                    "th", fontName="Helvetica-Bold", fontSize=9,
                    textColor=C_WHITE, alignment=TA_CENTER))
                  for c in data[0]]
        body_rows = []
        for ri, row in enumerate(data[1:]):
            bg = C_TABLE_ALT if ri % 2 == 0 else C_WHITE
            body_rows.append([
                Paragraph(self._inline(c), ParagraphStyle(
                    "td", fontName="Helvetica", fontSize=9,
                    textColor=C_BLACK, alignment=TA_LEFT))
                for c in row
            ])

        table_data = [header] + body_rows
        col_w = (160 * mm) / col_count

        tbl = Table(table_data, colWidths=[col_w] * col_count, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1,  0), C_TABLE_HDR),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_TABLE_ALT, C_WHITE]),
            ("GRID",        (0, 0), (-1, -1), 0.4, C_RULE),
            ("FONTNAME",    (0, 0), (-1,  0), "Helvetica-Bold"),
            ("FONTSIZE",    (0, 0), (-1, -1), 9),
            ("ALIGN",       (0, 0), (-1, -1), "LEFT"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",  (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_TABLE_ALT, C_WHITE]),
        ]))
        return [Spacer(1, 6), tbl, Spacer(1, 8)]


# ─────────────────────────────────────────────────────────────────────────────
# SYNTAX-HIGHLIGHTED CODE BLOCK
# ─────────────────────────────────────────────────────────────────────────────

def make_code_block(source: str, exec_count, styles) -> list:
    """
    Render a Python code cell as a dark-themed styled block.
    Uses a Table to create a rounded-corner-like dark background.
    """
    flowables = []

    # Label row: "In [N]:"
    label = f"In [{exec_count}]:" if exec_count is not None else "In [ ]:"
    flowables.append(Spacer(1, 6))

    # Apply Pygments token colouring → simplified colour mapping
    token_colors = {
        "Token.Keyword":            "#CBA6F7",   # purple
        "Token.Keyword.Namespace":  "#CBA6F7",
        "Token.Name.Builtin":       "#89B4FA",   # blue
        "Token.Name.Function":      "#89DCEB",   # cyan
        "Token.Name.Class":         "#F9E2AF",   # yellow
        "Token.Literal.String":     "#A6E3A1",   # green
        "Token.Literal.String.Doc": "#A6E3A1",
        "Token.Literal.Number":     "#FAB387",   # peach
        "Token.Comment":            "#6C7086",   # overlay0
        "Token.Operator":           "#89DCEB",
        "Token.Punctuation":        "#CDD6F4",
    }

    lines = source.splitlines()

    # Build label + code in one Table cell for background
    label_para = Paragraph(
        f'<font name="Helvetica-Bold" size="8" color="#A6ADC8">{label}</font>',
        ParagraphStyle("cl", leading=11))

    # Colour each line with basic pygments
    from pygments.lexers import PythonLexer
    from pygments.token import Token

    lexer  = PythonLexer()
    tokens = list(lexer.get_tokens(source))

    def tok_to_color(ttype):
        key = str(ttype)
        for pattern, color in token_colors.items():
            if key.startswith(pattern):
                return color
        return "#CDD6F4"          # default text

    # Build coloured lines as XML strings
    xml_lines = []
    current_line_parts = []

    for ttype, value in tokens:
        col = tok_to_color(ttype)
        # Escape XML
        val = (value.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;"))
        # Split on newlines
        parts = val.split("\n")
        for pi, part in enumerate(parts):
            if part:
                current_line_parts.append(
                    f'<font color="{col}">{part}</font>')
            if pi < len(parts) - 1:
                xml_lines.append("".join(current_line_parts))
                current_line_parts = []

    if current_line_parts:
        xml_lines.append("".join(current_line_parts))

    code_style = ParagraphStyle(
        "codeinner",
        fontName="Courier", fontSize=8.5,
        textColor=C_CODE_FG,
        leading=13, spaceAfter=0, spaceBefore=0,
        leftIndent=0, rightIndent=0,
        backColor=C_CODE_BG
    )

    code_paras = []
    for ln in xml_lines:
        p = Paragraph(ln or "&nbsp;", code_style)
        code_paras.append(p)

    # Wrap in a Table for background + border
    inner_content = [label_para] + code_paras
    tbl = Table([[inner_content]], colWidths=[160 * mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), C_CODE_BG),
        ("BOX",          (0, 0), (-1, -1), 0.5, colors.HexColor("#45475A")),
        ("TOPPADDING",   (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
        ("LEFTPADDING",  (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
    ]))
    flowables.append(tbl)
    return flowables


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT BLOCK RENDERER
# ─────────────────────────────────────────────────────────────────────────────

def _strip_ansi(text: str) -> str:
    """Remove ANSI escape sequences."""
    return re.sub(r"\x1b\[[0-9;]*[mGKHF]", "", text)


def make_output_block(outputs: list, exec_count, styles) -> list:
    """Render cell outputs as styled blocks below the code cell."""
    flowables = []
    if not outputs:
        return flowables

    out_label = f"Out [{exec_count}]:" if exec_count is not None else "Out [ ]:"

    for out in outputs:
        otype = out["type"]

        # ── Stream / execute_result text ────────────────────────────────
        if otype in ("stream", "execute_result", "display_data"):
            text = out.get("text", "")
            if not text and out.get("html"):
                text = "[HTML output — see notebook]"
            if not text:
                continue

            text = _strip_ansi(text)
            label_para = Paragraph(
                f'<font name="Helvetica-Bold" size="8"'
                f' color="#6B7280">{out_label}</font>',
                ParagraphStyle("ol", leading=11))

            out_style = ParagraphStyle(
                "outinner",
                fontName="Courier", fontSize=8.5,
                textColor=colors.HexColor("#1F2937"),
                leading=12.5, spaceAfter=0, spaceBefore=0,
            )
            out_lines = []
            for ln in text.splitlines():
                ln_esc = (ln.replace("&", "&amp;")
                            .replace("<", "&lt;")
                            .replace(">", "&gt;"))
                out_lines.append(Paragraph(ln_esc or "&nbsp;", out_style))

            tbl = Table([[([label_para] + out_lines)]],
                        colWidths=[160 * mm])
            tbl.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), C_OUTPUT_BG),
                ("BOX",           (0, 0), (-1, -1), 0.5, C_OUTPUT_BD),
                ("TOPPADDING",    (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ]))
            flowables.append(tbl)
            flowables.append(Spacer(1, 4))

            # ── Embedded image ───────────────────────────────────────────
            if out.get("image"):
                img_data = base64.b64decode(out["image"])
                img      = PILImage.open(io.BytesIO(img_data))
                w_px, h_px = img.size
                max_w = 155 * mm
                ratio  = max_w / w_px
                img_h  = h_px * ratio
                rl_img = RLImage(io.BytesIO(img_data), width=max_w, height=img_h)
                flowables.append(rl_img)
                flowables.append(Spacer(1, 4))

        # ── Error / traceback ────────────────────────────────────────────
        elif otype == "error":
            ename = out.get("ename", "Error")
            evalue = out.get("evalue", "")
            tb_raw = out.get("traceback", [])

            err_label = Paragraph(
                f'<font name="Helvetica-Bold" size="8"'
                f' color="#C0392B">&#9888; {ename}: {evalue}</font>',
                ParagraphStyle("el", leading=12))

            tb_lines_clean = []
            for ln in tb_raw:
                clean = _strip_ansi(ln)
                for sub in clean.splitlines():
                    esc = (sub.replace("&", "&amp;")
                              .replace("<", "&lt;")
                              .replace(">", "&gt;"))
                    tb_lines_clean.append(esc)

            err_style = ParagraphStyle(
                "errinner",
                fontName="Courier", fontSize=8,
                textColor=C_ERROR_TXT,
                leading=12, spaceAfter=0, spaceBefore=0,
            )
            tb_paras = [Paragraph(ln or "&nbsp;", err_style)
                        for ln in tb_lines_clean]

            inner = [err_label] + tb_paras
            tbl = Table([[inner]], colWidths=[160 * mm])
            tbl.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), C_ERROR_BG),
                ("BOX",           (0, 0), (-1, -1), 0.8, C_ERROR_BD),
                ("LINEAFTER",     (0, 0), (0,  -1), 3, C_ERROR_TXT),
                ("TOPPADDING",    (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("LEFTPADDING",   (0, 0), (-1, -1), 10),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
                ("VALIGN",        (0, 0), (-1, -1), "TOP"),
            ]))
            flowables.append(tbl)
            flowables.append(Spacer(1, 6))

    return flowables


# ─────────────────────────────────────────────────────────────────────────────
# PAGE TEMPLATE (header + footer)
# ─────────────────────────────────────────────────────────────────────────────

class NBDocTemplate(BaseDocTemplate):
    """Custom BaseDocTemplate with running header / footer and TOC support."""

    def __init__(self, filename, nb_title="Notebook Report", **kw):
        super().__init__(filename, **kw)
        self.nb_title  = nb_title
        self.generated = datetime.now().strftime("%B %d, %Y")
        self._toc      = None            # set by caller

        frame = Frame(
            self.leftMargin, self.bottomMargin,
            self.width, self.height,
            id="normal"
        )
        self.addPageTemplates([
            PageTemplate(id="normal", frames=[frame],
                         onPage=self._on_page)
        ])

    def afterFlowable(self, flowable):
        """Register headings with the TOC."""
        if isinstance(flowable, Paragraph):
            style = flowable.style.name
            text  = flowable.getPlainText()
            if style == "h1" and self._toc:
                self._toc.notify("TOCEntry", (0, text, self.page))
            elif style == "h2" and self._toc:
                self._toc.notify("TOCEntry", (1, text, self.page))
            elif style == "h3" and self._toc:
                self._toc.notify("TOCEntry", (2, text, self.page))

    def _on_page(self, canvas, doc):
        canvas.saveState()
        w, h = A4

        # ── Header bar ────────────────────────────────────────────────
        canvas.setFillColor(C_BRAND)
        canvas.rect(0, h - 16 * mm, w, 16 * mm, fill=1, stroke=0)

        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(C_WHITE)
        canvas.drawString(20 * mm, h - 10 * mm, self.nb_title)

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#A5B4FC"))
        canvas.drawRightString(w - 20 * mm, h - 10 * mm,
                               f"Generated {self.generated}")

        # ── Footer bar ────────────────────────────────────────────────
        canvas.setFillColor(colors.HexColor("#F3F4F6"))
        canvas.rect(0, 0, w, 12 * mm, fill=1, stroke=0)

        canvas.setStrokeColor(C_RULE)
        canvas.setLineWidth(0.5)
        canvas.line(20 * mm, 12 * mm, w - 20 * mm, 12 * mm)

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(C_MUTED)
        canvas.drawString(20 * mm, 4 * mm, "nb2pdf_agent.py — Notebook PDF Report")
        canvas.drawRightString(w - 20 * mm, 4 * mm,
                               f"Page {doc.page}")

        canvas.restoreState()


# ─────────────────────────────────────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────────────────────────────────────

def make_cover(title: str, nb_path: str, styles, kernel: str) -> list:
    """Build a cover page flowable list."""
    flowables = []
    w, h = A4

    flowables.append(Spacer(1, 45 * mm))

    # Logo / badge strip (drawn via a Table with coloured cells)
    badge = Table([[""]], colWidths=[40 * mm], rowHeights=[2 * mm])
    badge.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_ACCENT),
    ]))
    flowables.append(badge)
    flowables.append(Spacer(1, 8 * mm))

    flowables.append(Paragraph(title, styles["cover_title"]))
    flowables.append(Spacer(1, 4))
    flowables.append(Paragraph("Jupyter Notebook Analysis Report",
                                styles["cover_sub"]))
    flowables.append(Spacer(1, 18))

    # Thin rule
    flowables.append(HRFlowable(width="50%", thickness=1,
                                color=C_ACCENT, hAlign="CENTER"))
    flowables.append(Spacer(1, 18))

    meta = [
        ("Source file", Path(nb_path).name),
        ("Kernel",      kernel),
        ("Generated",   datetime.now().strftime("%B %d, %Y at %H:%M")),
    ]
    for label, value in meta:
        flowables.append(
            Paragraph(f"<b>{label}:</b> {value}", styles["cover_meta"]))

    flowables.append(PageBreak())
    return flowables


# ─────────────────────────────────────────────────────────────────────────────
# TABLE OF CONTENTS PAGE
# ─────────────────────────────────────────────────────────────────────────────

def make_toc_page(toc_obj, styles) -> list:
    flowables = []
    flowables.append(Paragraph("Table of Contents", styles["toc_title"]))
    flowables.append(HRFlowable(width="100%", thickness=1.5, color=C_BRAND))
    flowables.append(Spacer(1, 8))
    flowables.append(toc_obj)
    flowables.append(PageBreak())
    return flowables


# ─────────────────────────────────────────────────────────────────────────────
# MAIN AGENT
# ─────────────────────────────────────────────────────────────────────────────

def extract_title(cells: list) -> str:
    """Try to find a H1 heading in the first markdown cell."""
    for cell in cells:
        if cell["type"] == "markdown":
            for line in cell["source"].splitlines():
                m = re.match(r"^#\s+(.*)", line)
                if m:
                    return m.group(1).strip()
    return "Notebook Report"


def run(nb_path: str, out_path: str, title: str | None = None):
    print(f"📖  Parsing notebook: {nb_path}")
    parser = NotebookParser(nb_path)
    cells  = parser.parse()

    # Infer title from notebook content if not supplied
    doc_title = title or extract_title(cells)
    print(f"📝  Title: {doc_title}")

    styles = build_styles()
    md_renderer = MarkdownRenderer(styles)

    # ── Build TOC object ──────────────────────────────────────────────────
    toc = TableOfContents()
    toc.levelStyles = [
        styles["toc_h1"],
        styles["toc_h2"],
        styles["toc_h3"],
    ]
    toc.dotsMinLevel = 0

    # ── Assemble story ────────────────────────────────────────────────────
    story = []

    # Cover page
    story += make_cover(doc_title, nb_path, styles, parser.kernel)

    # TOC page
    story += make_toc_page(toc, styles)

    total = len(cells)
    for ci, cell in enumerate(cells):
        ctype = cell["type"]
        src   = cell["source"]
        ec    = cell["execution_count"]

        if not src.strip():
            continue

        print(f"  [{ci+1}/{total}] {ctype} cell")

        if ctype == "markdown":
            story += md_renderer.render(src)
            story.append(Spacer(1, 4))

        elif ctype == "code":
            story += make_code_block(src, ec, styles)
            story += make_output_block(cell["outputs"], ec, styles)
            story.append(Spacer(1, 8))

        elif ctype == "raw":
            raw_style = ParagraphStyle(
                "raw", fontName="Courier", fontSize=8.5,
                textColor=C_MUTED, leading=12)
            story.append(Paragraph(src.replace("\n", "<br/>"), raw_style))

    # ── Create PDF ────────────────────────────────────────────────────────
    doc = NBDocTemplate(
        out_path,
        nb_title=doc_title,
        pagesize=A4,
        leftMargin=25 * mm,
        rightMargin=25 * mm,
        topMargin=22 * mm,
        bottomMargin=18 * mm,
    )
    doc._toc = toc

    print(f"🖨️   Building PDF → {out_path}")
    # Two-pass build for TOC page numbers
    doc.multiBuild(story)
    print(f"✅  Done!  PDF saved to: {out_path}")
    return out_path


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert a Jupyter Notebook (.ipynb) to a professional PDF report.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python nb2pdf_agent.py notebook.ipynb
  python nb2pdf_agent.py notebook.ipynb --output report.pdf
  python nb2pdf_agent.py notebook.ipynb --output report.pdf --title "My Lab Report"
        """
    )
    parser.add_argument("notebook",
                        help="Path to the .ipynb file")
    parser.add_argument("-o", "--output",
                        default=None,
                        help="Output PDF path (default: <notebook_name>.pdf)")
    parser.add_argument("-t", "--title",
                        default=None,
                        help="Override the report title")

    args = parser.parse_args()

    nb_path = args.notebook
    if not os.path.isfile(nb_path):
        print(f"❌  File not found: {nb_path}", file=sys.stderr)
        sys.exit(1)
    if not nb_path.endswith(".ipynb"):
        print("⚠️   Warning: file does not have .ipynb extension",
              file=sys.stderr)

    out_path = args.output or Path(nb_path).stem + ".pdf"
    run(nb_path, out_path, title=args.title)


if __name__ == "__main__":
    main()
 