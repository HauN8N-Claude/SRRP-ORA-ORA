#!/usr/bin/env python3
"""md_to_pdf.py — Convertit un guide Markdown en PDF paginé (reportlab + police DejaVu).

Usage : python build/md_to_pdf.py build/05_guide.md build/Guide_Assistantes_ORA_ORA.pdf

Gère : titres (#, ##, ###), paragraphes, listes (- / 1.), tableaux (| … |),
citations (>), règles (---), gras (**), code (`). Les emojis décoratifs sont
remplacés par des libellés lisibles (DejaVu ne contient pas les emojis couleur).
"""
import re, os, sys
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                HRFlowable, ListFlowable, ListItem)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

FD = "/usr/share/fonts/truetype/dejavu/"
pdfmetrics.registerFont(TTFont("DJ", FD + "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DJ-B", FD + "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("DJM", FD + "DejaVuSansMono.ttf"))
registerFontFamily("DJ", normal="DJ", bold="DJ-B", italic="DJ", boldItalic="DJ-B")

REPL = {"⭐": "●", "💡": "Astuce :", "➡️": "→", "➡": "→", "🔄": "", "🔎": "", "🚫": "Interdit :",
        "❓": "", "✅": "", "🗂️": "", "🗂": "", "⚠️": "(!)", "⚠": "(!)", "👉": "→"}

def sanitize(s):
    for k, v in REPL.items():
        s = s.replace(k, v)
    return "".join(ch for ch in s if not (0x1F000 <= ord(ch) <= 0x1FAFF or 0x2600 <= ord(ch) <= 0x27BF
                   or 0x2B00 <= ord(ch) <= 0x2BFF or ord(ch) == 0xFE0F))

def inline(s):
    s = sanitize(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    s = re.sub(r"`(.+?)`", r'<font face="DJM" size=9>\1</font>', s)
    return s

def st(name, **kw):
    base = dict(fontName="DJ", fontSize=10, leading=14, spaceAfter=5)
    base.update(kw)
    return ParagraphStyle(name, **base)

H1 = st("h1", fontName="DJ-B", fontSize=18, leading=22, spaceAfter=10, textColor=colors.HexColor("#1F4E79"))
H2 = st("h2", fontName="DJ-B", fontSize=13.5, leading=18, spaceBefore=12, spaceAfter=6, textColor=colors.HexColor("#2E75B6"))
H3 = st("h3", fontName="DJ-B", fontSize=11.5, leading=15, spaceBefore=8, spaceAfter=4, textColor=colors.HexColor("#2E75B6"))
P = st("p")
QUOTE = st("quote", leftIndent=10, textColor=colors.HexColor("#555555"),
           backColor=colors.HexColor("#F4F4F4"), borderPadding=4, spaceBefore=3, spaceAfter=6)
CELL = st("cell", fontSize=9, leading=12)
CELLH = st("cellh", fontSize=9, leading=12, fontName="DJ-B", textColor=colors.white)


def convert(src, out):
    md = open(src, encoding="utf-8").read().splitlines()
    story, i, n = [], 0, len(md)

    def flush_table(rows):
        data = [[Paragraph(inline(c), CELLH if r == 0 else CELL) for c in row]
                for r, row in enumerate(rows)]
        t = Table(data, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E75B6")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAAAAA")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F7FC")]),
            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
        story.append(t); story.append(Spacer(1, 6))

    while i < n:
        line = md[i].rstrip()
        if not line.strip():
            story.append(Spacer(1, 4)); i += 1; continue
        if line.startswith("# "): story.append(Paragraph(inline(line[2:]), H1))
        elif line.startswith("## "): story.append(Paragraph(inline(line[3:]), H2))
        elif line.startswith("### "): story.append(Paragraph(inline(line[4:]), H3))
        elif set(line) <= set("-") and len(line) >= 3:
            story.append(HRFlowable(width="100%", thickness=0.7, color=colors.HexColor("#CCCCCC"),
                                    spaceBefore=4, spaceAfter=8))
        elif line.startswith("|"):
            rows = []
            while i < n and md[i].lstrip().startswith("|"):
                cells = [c.strip() for c in md[i].strip().strip("|").split("|")]
                if not all(set(c) <= set("-: ") for c in cells):
                    rows.append(cells)
                i += 1
            flush_table(rows); continue
        elif line.startswith(">"):
            buf = []
            while i < n and md[i].startswith(">"):
                buf.append(md[i].lstrip(">").strip()); i += 1
            story.append(Paragraph(inline(" ".join(buf)), QUOTE)); continue
        elif re.match(r"^[-*] ", line):
            items = []
            while i < n and re.match(r"^[-*] ", md[i].rstrip()):
                items.append(ListItem(Paragraph(inline(md[i].rstrip()[2:]), P))); i += 1
            story.append(ListFlowable(items, bulletType="bullet", start="•", leftIndent=14)); continue
        elif re.match(r"^\d+\. ", line):
            items = []
            while i < n and re.match(r"^\d+\. ", md[i].rstrip()):
                items.append(ListItem(Paragraph(inline(re.sub(r'^\d+\. ', '', md[i].rstrip())), P))); i += 1
            story.append(ListFlowable(items, bulletType="1", leftIndent=14)); continue
        elif line.startswith("*") and line.endswith("*") and len(line) > 2:
            story.append(Paragraph(inline(line.strip("*")), st("foot", fontSize=8.5, textColor=colors.grey)))
        else:
            story.append(Paragraph(inline(line), P))
        i += 1

    SimpleDocTemplate(out, pagesize=A4, topMargin=16 * mm, bottomMargin=16 * mm,
                      leftMargin=16 * mm, rightMargin=16 * mm,
                      title="Guide assistantes - ORA ORA SSR").build(story)
    print("PDF ->", out, os.path.getsize(out), "octets")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "build/05_guide.md"
    out = sys.argv[2] if len(sys.argv) > 2 else "build/Guide_Assistantes_ORA_ORA.pdf"
    convert(src, out)
