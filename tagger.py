"""
tagger.py
---------
Converts a plain / untagged MCQ Word document into the "tagged" docx
format used for question-bank uploads:

    $total: <N>
    @que
      @type
        Option
      !type
      @body
        <question text>
      !body
      @options
        <option 1>
        <option 2>
        ...
      !options
      @ans
        <bold correct letter>
      !ans
      @ep
        <explanation text>
      !ep
    !que

Tag lines are Arial 12pt, colour C00000 (dark red).
Content lines are Times New Roman 12pt, black.
The correct-answer letter is bold.
Options render as an auto-lettered list (a, b, c, d...) that restarts
at "a" for every question, matching the reference document.

Two entry points:
    parse_plain_docx(path)         -> list[Question]
    build_tagged_docx(questions, out_path)
"""

from __future__ import annotations

import re
import zipfile
import os
import base64
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

import docx  # python-docx, used only for reading the plain input file

from assets_data import (
    CONTENT_TYPES_B64,
    PACKAGE_RELS_B64,
    CORE_PROPS_B64,
    APP_PROPS_B64,
    DOCUMENT_RELS_B64,
    STYLES_B64,
    FONT_TABLE_B64,
    SETTINGS_B64,
    WEB_SETTINGS_B64,
    THEME1_B64,
)


# --------------------------------------------------------------------------
# Data model
# --------------------------------------------------------------------------

@dataclass
class Question:
    body: list[str] = field(default_factory=list)          # question text, 1+ paragraphs
    options: list[str] = field(default_factory=list)       # option texts, in order a,b,c,...
    answer_letter: str | None = None                       # e.g. "C"
    explanation: list[str] = field(default_factory=list)   # explanation text, 1+ paragraphs
    qtype: str = "Option"                                  # "Option" or "Numerical"


# --------------------------------------------------------------------------
# Parsing the plain / untagged input document
# --------------------------------------------------------------------------

_QUESTION_START = re.compile(
    r"^\s*(?:Q(?:uestion)?\.?\s*(\d+)\s*[.):]?|(\d+)\s*[.):])\s*(.*)$",
    re.IGNORECASE,
)
_OPTION_LINE = re.compile(r"^\s*\(?([a-dA-D])\)?[.):]\s*(.*)$")
_ANSWER_LINE = re.compile(
    r"^\s*(?:correct\s+answer|ans(?:wer)?|key)\s*[:\-]?\s*\(?([a-dA-D])\)?\b.*$",
    re.IGNORECASE,
)
_EXPLANATION_LABEL = re.compile(r"^\s*(?:explanation|solution|exp)\s*[:\-]?\s*(.*)$", re.IGNORECASE)


def parse_plain_docx(path: str) -> list[Question]:
    """Read a plain MCQ docx and split it into structured Question objects.

    Expected input shape per question (labels are case-insensitive and the
    punctuation after numbers/letters is flexible):

        1. Question text goes here, can wrap to a second line.
        a) Option one
        b) Option two
        c) Option three
        d) Option four
        Answer: C
        Explanation: Why C is correct.

    Blank lines are ignored. Multi-paragraph question bodies (e.g.
    assertion/reason style) are supported by anything appearing between the
    question-number line and the first option line.
    """
    doc = docx.Document(path)
    raw_lines: list[str] = []
    for p in doc.paragraphs:
        # Options (and sometimes other content) are sometimes packed into a
        # single paragraph using manual line breaks rather than separate
        # paragraphs, e.g. "A. foo\nB. bar\nC. baz\nD. qux". Split those out
        # so each option/statement is its own line for the parser below.
        for sub in p.text.split("\n"):
            sub = sub.strip()
            if sub:
                raw_lines.append(sub)
    lines = raw_lines

    questions: list[Question] = []
    current: Question | None = None
    mode = None  # "body" | "options" | "explanation"

    def flush():
        if current is not None and (current.body or current.options):
            questions.append(current)

    for line in lines:
        qmatch = _QUESTION_START.match(line)
        omatch = _OPTION_LINE.match(line)
        amatch = _ANSWER_LINE.match(line)
        ematch = _EXPLANATION_LABEL.match(line)

        # A new question only starts on a numbered line that isn't actually
        # an option line (options are also "letter + punctuation").
        if qmatch:
            flush()
            current = Question()
            rest = (qmatch.group(3) or "").strip()
            if rest:
                current.body.append(rest)
            mode = "body"
            continue

        if current is None:
            # Content before the first recognised question number - skip.
            continue

        if amatch:
            current.answer_letter = amatch.group(1).upper()
            mode = "explanation"
            continue

        if ematch:
            mode = "explanation"
            if ematch.group(1).strip():
                current.explanation.append(ematch.group(1).strip())
            continue

        if omatch and mode in ("body", "options"):
            current.options.append(omatch.group(2).strip())
            mode = "options"
            continue

        # Plain continuation line - goes to whichever section we're in.
        if mode == "body":
            current.body.append(line)
        elif mode == "options" and current.options:
            # Wrapped continuation of the last option.
            current.options[-1] += " " + line
        elif mode == "explanation":
            current.explanation.append(line)
        else:
            current.body.append(line)

    flush()

    for q in questions:
        q.qtype = "Option" if len(q.options) >= 2 else "Numerical"

    return questions


# --------------------------------------------------------------------------
# Building the tagged output document
# --------------------------------------------------------------------------

_TAG_RPR = (
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:color w:val="C00000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
)
_BODY_RPR = (
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
    'w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
)
_BODY_BOLD_RPR = (
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
    'w:cs="Times New Roman"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
)


def _para(text: str, rpr: str, *, indent="left", num_id: int | None = None,
          leading_tab: bool = False, pstyle: str | None = None) -> str:
    """Build a single <w:p> paragraph as an XML string."""
    ind = ""
    if indent == "left":
        ind = '<w:ind w:left="720"/>'
    elif indent == "firstLine":
        ind = '<w:ind w:firstLine="720"/>'
    elif indent == "hanging":
        ind = '<w:ind w:left="720" w:hanging="720"/>'

    style = f'<w:pStyle w:val="{pstyle}"/>' if pstyle else ""
    numpr = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>' if num_id else ""
    tabs = '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>' if leading_tab else ""

    ppr = f"<w:pPr>{style}{numpr}{tabs}<w:spacing w:after=\"20\"/>{ind}<w:jc w:val=\"both\"/></w:pPr>"
    tab_run = "<w:tab/>" if leading_tab else ""
    run = f'<w:r>{rpr}{tab_run}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>' if text else ""
    return f"<w:p>{ppr}{run}</w:p>"


def _tag_line(text: str, indent="left", leading_tab=False) -> str:
    return _para(text, _TAG_RPR, indent=indent, leading_tab=leading_tab)


def _body_line(text: str, bold: bool = False) -> str:
    return _para(text, _BODY_BOLD_RPR if bold else _BODY_RPR, indent="left")


def _option_line(text: str, num_id: int) -> str:
    return _para(text, _BODY_RPR, indent=None, num_id=num_id, pstyle="ListParagraph")


NUMBERING_HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
)
NUMBERING_FOOTER = "</w:numbering>"

ABSTRACT_NUM = (
    '<w:abstractNum w:abstractNumId="0">'
    '<w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
    '<w:lvlText w:val="%1)"/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
    '<w:lvl w:ilvl="1" w:tentative="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
    '<w:lvlText w:val="%2."/><w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
    '<w:lvl w:ilvl="2" w:tentative="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>'
    '<w:lvlText w:val="%3."/><w:lvlJc w:val="right"/>'
    '<w:pPr><w:ind w:left="2160" w:hanging="180"/></w:pPr></w:lvl>'
    '</w:abstractNum>'
)


def _build_numbering_xml(num_lists: int) -> str:
    # Each question gets its own <w:num> instance with an explicit
    # startOverride so the a)/b)/c)/d) lettering restarts at "a" every time,
    # instead of continuing from the previous question's list.
    nums = "".join(
        f'<w:num w:numId="{i + 1}"><w:abstractNumId w:val="0"/>'
        f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
        f'</w:num>'
        for i in range(num_lists)
    )
    return NUMBERING_HEADER + ABSTRACT_NUM + nums + NUMBERING_FOOTER


def _question_block_xml(q: Question, num_id: int) -> str:
    parts = []
    parts.append(_tag_line("@que", indent="firstLine"))
    parts.append(_tag_line("@type", leading_tab=True))
    parts.append(_body_line(q.qtype))
    parts.append(_tag_line("!type"))
    parts.append(_tag_line("@body"))
    for line in q.body:
        parts.append(_body_line(line))
    parts.append(_tag_line("!body", leading_tab=True))

    if q.qtype == "Option" and q.options:
        parts.append(_tag_line("@options"))
        for opt in q.options:
            parts.append(_option_line(opt, num_id))
        parts.append(_tag_line("!options"))

    parts.append(_tag_line("@ans"))
    parts.append(_body_line(q.answer_letter or "", bold=True))
    parts.append(_tag_line("!ans"))

    parts.append(_tag_line("@ep", leading_tab=True))
    for line in q.explanation:
        parts.append(_body_line(line))
    parts.append(_tag_line("!ep"))

    parts.append(_tag_line("!que", indent="firstLine"))
    return "".join(parts)


_SECT_PR_FALLBACK = (
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)

_DOCUMENT_HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
)
_DOCUMENT_FOOTER = "</w:body></w:document>"


def build_tagged_docx(questions: list[Question], out_path: str) -> None:
    """Write `questions` out as a tagged .docx at `out_path`.

    Every package part EXCEPT document.xml and numbering.xml is copied
    byte-for-byte from the known-good reference tagging template (embedded
    in assets_data.py as base64), so the file matches exactly what the
    target upload system already accepts."""
    body_xml = [_tag_line(f"$total: {len(questions)}", indent="firstLine")]
    for i, q in enumerate(questions, start=1):
        body_xml.append(_question_block_xml(q, num_id=i))

    document_xml = _DOCUMENT_HEADER + "".join(body_xml) + _SECT_PR_FALLBACK + _DOCUMENT_FOOTER
    numbering_xml = _build_numbering_xml(len(questions))
    _write_full_docx(document_xml, numbering_xml, out_path)


# --------------------------------------------------------------------------
# Package assembly. Static parts (content types, relationships, docProps,
# styles, fonts, settings, webSettings, theme) are decoded from the
# reference template embedded in assets_data.py, guaranteeing they match a
# file already known to be accepted by the target system. Only
# document.xml (the questions) and numbering.xml (per-question restart-at-a
# lettered lists) are generated per upload.
# --------------------------------------------------------------------------


def _write_full_docx(document_xml: str, numbering_xml: str, out_path: str) -> None:
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", base64.b64decode(CONTENT_TYPES_B64))
        zf.writestr("_rels/.rels", base64.b64decode(PACKAGE_RELS_B64))
        zf.writestr("docProps/core.xml", base64.b64decode(CORE_PROPS_B64))
        zf.writestr("docProps/app.xml", base64.b64decode(APP_PROPS_B64))
        zf.writestr("word/_rels/document.xml.rels", base64.b64decode(DOCUMENT_RELS_B64))
        zf.writestr("word/document.xml", document_xml.encode("utf-8"))
        zf.writestr("word/numbering.xml", numbering_xml.encode("utf-8"))
        zf.writestr("word/styles.xml", base64.b64decode(STYLES_B64))
        zf.writestr("word/fontTable.xml", base64.b64decode(FONT_TABLE_B64))
        zf.writestr("word/settings.xml", base64.b64decode(SETTINGS_B64))
        zf.writestr("word/webSettings.xml", base64.b64decode(WEB_SETTINGS_B64))
        zf.writestr("word/theme/theme1.xml", base64.b64decode(THEME1_B64))


def convert(input_path: str, output_path: str) -> int:
    """Convenience wrapper used by the Streamlit app. Returns question count."""
    questions = parse_plain_docx(input_path)
    build_tagged_docx(questions, output_path)
    return len(questions)
