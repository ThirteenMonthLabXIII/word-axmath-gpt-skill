from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from inspect_layout_rules import inspect_layout

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
)
FLAGS_OFF = (
    '<w:widowControl w:val="0"/><w:keepNext w:val="0"/>'
    '<w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/>'
)
STYLES = f'''<w:styles {NS}>
  <w:docDefaults><w:pPrDefault><w:pPr>{FLAGS_OFF}</w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:ind w:left="360"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
    <w:basedOn w:val="Base"/><w:pPr><w:keepLines/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="InheritedList"><w:name w:val="List"/>
    <w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="NextHeading"><w:name w:val="Next Heading Candidate"/>
    <w:basedOn w:val="Normal"/><w:next w:val="Heading1"/></w:style>
</w:styles>'''


def make_docx(path: Path) -> None:
    document = f'''<w:document {NS}><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:keepNext w:val="0"/></w:pPr>
        <w:r><w:t>标题</w:t></w:r></w:p>
      <w:p><w:r><w:t xml:space="preserve">1. 设 </w:t></w:r>
        <w:r><w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>
        <w:r><w:t xml:space="preserve"> 满足条件</w:t></w:r></w:p>
      <w:p><w:r><w:t>（</w:t></w:r><w:r><w:t>１）</w:t></w:r><w:r><w:t>求</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="InheritedList"/></w:pPr><w:r><w:t>自动列表</w:t></w:r></w:p>
      <w:p><w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>
      <w:p><w:r><w:t>第 1 题图</w:t></w:r></w:p>
      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing/></w:r></w:p>
      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>第 2 题图</w:t></w:r></w:p>
      <w:p><w:r><w:t>2.1章节</w:t></w:r></w:p>
      <w:p><w:r><w:t>1.(1) 小问； (2) 小问</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>
      <w:tbl><w:tr><w:tc><w:p><w:pPr><w:pageBreakBefore/></w:pPr>
        <w:r><w:t>表格文字</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>大纲段落</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="NextHeading"/></w:pPr><w:r><w:t>后续样式关联标题</w:t></w:r></w:p>
    </w:body></w:document>'''
    header = f'''<w:hdr {NS}><w:p><w:pPr><w:widowControl/></w:pPr>
      <w:r><w:t>页眉</w:t></w:r></w:p></w:hdr>'''
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/document.xml", document)
        z.writestr("word/header1.xml", header)


class LayoutInspectionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.docx = Path(self.temp.name) / "math.docx"
        make_docx(self.docx)

    def test_detects_inherited_pagination_and_direct_overrides_across_stories(self) -> None:
        result = inspect_layout(self.docx)
        findings = result["confirmed_or_candidate_findings"]
        pagination = {(x["story"], x["paragraph"], x["property"], x["origin"])
                      for x in findings if x["kind"] == "PAGINATION_ENABLED"}
        self.assertIn(("word/document.xml", 1, "keepLines", "style:Heading1"), pagination)
        self.assertNotIn(("word/document.xml", 1, "keepNext", "style:Base"), pagination)
        self.assertIn(("word/document.xml", 11, "pageBreakBefore", "paragraph"), pagination)
        self.assertIn(("word/header1.xml", 1, "widowControl", "paragraph"), pagination)
        self.assertFalse(any(x["kind"] == "PAGINATION_DEFAULT_UNSPECIFIED"
                             for x in result["unresolved_require_word_or_source_review"]))
        self.assertTrue(any(x["kind"] == "NONZERO_INDENT" and x["origin"] == "style:Base" for x in findings))
        self.assertTrue(any(x["kind"] == "NUMBERING_OR_HEADING" and x["automatic_numbering"] for x in findings))
        self.assertTrue(any(x["kind"] == "NUMBERING_OR_HEADING" and x["outline_level"] == "0" for x in findings))
        self.assertTrue(any(x["kind"] == "NUMBERING_OR_HEADING" and x["heading_style_reference_candidate"] for x in findings))

    def test_detects_number_and_formula_gaps_without_classifying_equation_as_image(self) -> None:
        result = inspect_layout(self.docx)
        findings = result["confirmed_or_candidate_findings"]
        self.assertTrue(any(x["kind"] == "NUMBER_FOLLOWED_BY_WHITESPACE" and x["paragraph"] == 2 for x in findings))
        self.assertEqual({x["side"] for x in findings if x["kind"] == "TEXT_FORMULA_WHITESPACE"}, {"before", "after"})
        self.assertTrue(any(x["kind"] == "NON_ASCII_NUMBER_LABEL_CANDIDATE" and x["paragraph"] == 3 for x in findings))
        self.assertFalse(any(x["kind"] == "NUMBER_FOLLOWED_BY_WHITESPACE" and x["paragraph"] == 9 for x in findings))
        inline_labels = {x["label"] for x in result["unresolved_require_word_or_source_review"]
                         if x["kind"] == "INLINE_NUMBER_WHITESPACE_CANDIDATE"}
        self.assertEqual(inline_labels, {"(1)", "(2)"})
        self.assertEqual(len(result["graphics_inventory"]), 2)
        self.assertFalse(any(x["paragraph"] == 2 for x in result["graphics_inventory"]))

    def test_figures_captions_and_manual_break_require_actual_review(self) -> None:
        result = inspect_layout(self.docx)
        findings = result["confirmed_or_candidate_findings"]
        self.assertTrue(any(x["kind"] == "GRAPHIC_NOT_CENTERED_CANDIDATE" and x["paragraph"] == 5 for x in findings))
        self.assertTrue(any(x["kind"] == "CAPTION_NOT_CENTERED_CANDIDATE" and x["paragraph"] == 6 for x in findings))
        self.assertFalse(any(x["kind"] in {"GRAPHIC_NOT_CENTERED_CANDIDATE", "CAPTION_NOT_CENTERED_CANDIDATE"}
                             and x["paragraph"] in {7, 8} for x in findings))
        self.assertTrue(any(x["kind"] == "MANUAL_PAGE_BREAK_REVIEW" for x in result["unresolved_require_word_or_source_review"]))
        self.assertTrue(result["visual_and_editability_review_required"])

    def test_report_is_read_only_and_refuses_to_overwrite(self) -> None:
        before = hashlib.sha256(self.docx.read_bytes()).hexdigest()
        report = Path(self.temp.name) / "report.json"
        script = Path(__file__).with_name("inspect_layout_rules.py")
        command = [sys.executable, str(script), str(self.docx), "--report", str(report)]
        run = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertEqual(json.loads(report.read_text(encoding="utf-8"))["sha256"], before)
        again = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
        self.assertNotEqual(again.returncode, 0)
        self.assertEqual(hashlib.sha256(self.docx.read_bytes()).hexdigest(), before)


if __name__ == "__main__":
    unittest.main()
