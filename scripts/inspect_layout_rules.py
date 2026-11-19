"""Read-only locator for the user's math handout layout requirements.

DOCX structure cannot establish visual placement or all effective Word values.
The report separates confirmed issues from items requiring human/Word review.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from inspect_question_numbering import inspect as inspect_numbering

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
O = "{urn:schemas-microsoft-com:office:office}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
STORY = re.compile(r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$")
PAGINATION = ("widowControl", "keepNext", "keepLines", "pageBreakBefore")
INDENT = ("left", "right", "start", "end", "firstLine", "hanging", "leftChars",
          "rightChars", "startChars", "endChars", "firstLineChars", "hangingChars")
SPACE = " \t\u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u202f\u205f\u3000"
NUMERIC = r"[0-9０-９]+"
LABEL = re.compile(
    r"^(?P<lead>[" + SPACE + r"]*)(?P<star>[*＊]*)(?P<label>"
    r"[（(][" + SPACE + r"]*" + NUMERIC + r"[" + SPACE + r"]*[）)]|"
    + NUMERIC + r"[.．、](?![0-9０-９])|"
    r"(?:例题|例|练习|习题)[" + SPACE + r"]*" + NUMERIC + r")"
    r"(?P<gap>[" + SPACE + r"]*)"
)
INLINE_WIDE_LABEL = re.compile(r"(?:^|[;；:：\n])[" + SPACE + r"]*（[" + SPACE + r"]*[0-9０-９]+[" + SPACE + r"]*）")
INLINE_LABEL = re.compile(r"(?:[;；:：]|[.．])[" + SPACE + r"]*(?P<label>[（(][" + SPACE + r"]*[0-9０-９]+[" + SPACE + r"]*[）)])(?P<gap>[" + SPACE + r"]+)")
CAPTION = re.compile(r"^\s*(?:第\s*\d+\s*题图|题\s*\d+\s*图|例(?:题)?\s*\d+\s*图|图\s*\d+(?:[.\-]\d+)?)\s*$")


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def child(node: ET.Element | None, tag: str) -> ET.Element | None:
    return None if node is None else node.find(W + tag)


def value(node: ET.Element | None) -> str | None:
    return None if node is None else node.get(W + "val")


def enabled(node: ET.Element) -> bool:
    return (node.get(W + "val") or "true").strip().lower() not in {"0", "false", "off", "no"}


def style_layers(p: ET.Element, styles: dict[str, ET.Element], default_id: str | None,
                 defaults: ET.Element | None) -> tuple[list[tuple[str, ET.Element]], str | None]:
    direct = child(p, "pPr")
    style_id = value(child(direct, "pStyle")) or default_id
    layers: list[tuple[str, ET.Element]] = []
    if direct is not None:
        layers.append(("paragraph", direct))
    current = style_id
    seen: set[str] = set()
    while current and current not in seen and current in styles:
        seen.add(current)
        st = styles[current]
        props = child(st, "pPr")
        if props is not None:
            layers.append(("style:" + current, props))
        current = value(child(st, "basedOn"))
    if defaults is not None:
        layers.append(("docDefaults", defaults))
    return layers, style_id


def effective(layers: list[tuple[str, ET.Element]], name: str) -> tuple[ET.Element | None, str]:
    for origin, props in layers:
        node = child(props, name)
        if node is not None:
            return node, origin
    return None, "unspecified"


def effective_indent(layers: list[tuple[str, ET.Element]]) -> dict[str, dict[str, str]]:
    found: dict[str, dict[str, str]] = {}
    # firstLine and hanging are mutually exclusive at the same precedence level.
    first_line_claimed = False
    char_line_claimed = False
    for origin, props in layers:
        node = child(props, "ind")
        if node is None:
            continue
        for key in INDENT:
            if key in found:
                continue
            if key in ("firstLine", "hanging") and first_line_claimed:
                continue
            if key in ("firstLineChars", "hangingChars") and char_line_claimed:
                continue
            val = node.get(W + key)
            if val is not None:
                found[key] = {"value": val, "origin": origin}
        if node.get(W + "firstLine") is not None or node.get(W + "hanging") is not None:
            first_line_claimed = True
        if node.get(W + "firstLineChars") is not None or node.get(W + "hangingChars") is not None:
            char_line_claimed = True
    return found


def events(paragraph: ET.Element) -> list[tuple[str, str, bool]]:
    """Return text, controls, formula and graphic atoms belonging to this paragraph."""
    result: list[tuple[str, str, bool]] = []

    def walk(node: ET.Element) -> None:
        for item in node:
            tag = item.tag
            if tag in {W + "pPr", W + "rPr", W + "p"}:
                continue
            if tag == W + "t":
                result.append(("text", item.text or "", False))
            elif tag in {W + "tab", W + "ptab"}:
                result.append(("text", "\t", False))
            elif tag in {W + "br", W + "cr"}:
                result.append(("page_break" if item.get(W + "type") == "page" else "break", "\n", False))
            elif tag in {M + "oMath", M + "oMathPara"}:
                result.append(("formula", "", False))
            elif tag == W + "object":
                ole = item.find(".//" + O + "OLEObject")
                progid = (ole.get("ProgID") or "").lower() if ole is not None else ""
                if progid in {"equation.axmath", "equation.3", "equation.dsmt4"}:
                    result.append(("formula", "", False))
                else:
                    result.append(("graphic", "OLE:" + (progid or "unknown"), False))
            elif tag in {W + "drawing", W + "pict"}:
                result.append(("graphic", tag.rsplit("}", 1)[-1], item.find(".//" + WP + "anchor") is not None))
            else:
                walk(item)

    walk(paragraph)
    return result


def projection(atoms: list[tuple[str, str, bool]]) -> str:
    return "".join(content if kind in {"text", "break", "page_break"} else "\ufffc" for kind, content, _ in atoms)


def formula_gaps(atoms: list[tuple[str, str, bool]]) -> list[dict[str, str]]:
    gaps = []
    for index, (kind, _, _) in enumerate(atoms):
        if kind != "formula":
            continue
        for direction in (-1, 1):
            pos = index + direction
            spaces = ""
            while 0 <= pos < len(atoms) and atoms[pos][0] == "text":
                text = atoms[pos][1]
                if not text:
                    pos += direction
                    continue
                if direction < 0:
                    stripped = text.rstrip(SPACE)
                    spaces = text[len(stripped):] + spaces
                else:
                    stripped = text.lstrip(SPACE)
                    spaces += text[:len(text) - len(stripped)]
                if stripped:
                    if spaces:
                        gaps.append({"side": "before" if direction < 0 else "after", "whitespace": repr(spaces)})
                    break
                pos += direction
    return gaps


def inspect_layout(docx: Path) -> dict:
    docx = docx.resolve(strict=True)
    before = digest(docx)
    issues: list[dict] = []
    unresolved: list[dict] = []
    graphics: list[dict] = []
    paragraphs = 0
    with zipfile.ZipFile(docx) as z:
        names = z.namelist()
        if "word/document.xml" not in names:
            raise ValueError("word/document.xml is missing")
        styles: dict[str, ET.Element] = {}
        default_id = None
        defaults = None
        if "word/styles.xml" in names:
            root = ET.fromstring(z.read("word/styles.xml"))
            defaults_root = child(root, "docDefaults")
            defaults = child(child(defaults_root, "pPrDefault"), "pPr")
            for style in root.findall(W + "style"):
                if style.get(W + "type") != "paragraph":
                    continue
                sid = style.get(W + "styleId")
                if sid:
                    styles[sid] = style
                    if style.get(W + "default", "").lower() in {"1", "true", "on"}:
                        default_id = sid
        else:
            unresolved.append({"kind": "STYLES_MISSING", "detail": "Verify Word's effective defaults."})

        for story in sorted(n for n in names if STORY.fullmatch(n)):
            root = ET.fromstring(z.read(story))
            for index, p in enumerate(root.iter(W + "p"), 1):
                paragraphs += 1
                ref = {"story": story, "paragraph": index}
                layers, sid = style_layers(p, styles, default_id, defaults)
                atoms = events(p)
                raw = projection(atoms)
                excerpt = raw.replace("\ufffc", "[对象]").replace("\n", " ")[:100]
                for name in PAGINATION:
                    node, origin = effective(layers, name)
                    if node is None:
                        unresolved.append({**ref, "kind": "PAGINATION_DEFAULT_UNSPECIFIED", "property": name,
                                           "style": sid, "text": excerpt})
                    elif enabled(node):
                        issues.append({**ref, "kind": "PAGINATION_ENABLED", "property": name,
                                       "origin": origin, "style": sid, "text": excerpt})

                if story == "word/document.xml":
                    indents = effective_indent(layers)
                    for key, data in indents.items():
                        if data["value"] not in {"0", "0.0", "+0", "-0"}:
                            issues.append({**ref, "kind": "NONZERO_INDENT", "property": key,
                                           "value": data["value"], "origin": data["origin"], "text": excerpt})
                    jc, origin = effective(layers, "jc")
                    align = value(jc) or "left"
                    has_graphic = any(k == "graphic" for k, _, _ in atoms)
                    text_only = "".join(v for k, v, _ in atoms if k == "text").strip(SPACE)
                    looks_caption = bool(CAPTION.fullmatch(text_only))
                    if has_graphic and not text_only and align != "center":
                        issues.append({**ref, "kind": "GRAPHIC_NOT_CENTERED_CANDIDATE", "alignment": align,
                                       "origin": origin, "text": excerpt})
                    if has_graphic and text_only:
                        unresolved.append({**ref, "kind": "MIXED_TEXT_AND_GRAPHIC_REVIEW", "text": excerpt})
                    if not has_graphic and not looks_caption and text_only and align not in {"left", "start"}:
                        issues.append({**ref, "kind": "BODY_ALIGNMENT_CANDIDATE", "alignment": align,
                                       "origin": origin, "text": excerpt})
                    if looks_caption and align != "center":
                        issues.append({**ref, "kind": "CAPTION_NOT_CENTERED_CANDIDATE", "alignment": align,
                                       "origin": origin, "text": excerpt})
                    if looks_caption:
                        unresolved.append({**ref, "kind": "CAPTION_ASSOCIATION_REVIEW", "text": excerpt})

                for line in raw.split("\n"):
                    if line and line[0] in SPACE and line.strip(SPACE):
                        issues.append({**ref, "kind": "LEADING_LINE_WHITESPACE_CANDIDATE", "text": excerpt})
                    match = LABEL.match(line)
                    if not match:
                        continue
                    label = match.group("label")
                    if match.group("lead"):
                        issues.append({**ref, "kind": "LEADING_LABEL_WHITESPACE", "text": excerpt})
                    if match.group("gap"):
                        issues.append({**ref, "kind": "NUMBER_FOLLOWED_BY_WHITESPACE", "label": label, "text": excerpt})
                    if any(c in label for c in "（ ）．、０１２３４５６７８９"):
                        issues.append({**ref, "kind": "NON_ASCII_NUMBER_LABEL_CANDIDATE", "label": label, "text": excerpt})
                if INLINE_WIDE_LABEL.search(raw):
                    unresolved.append({**ref, "kind": "INLINE_FULLWIDTH_LABEL_CANDIDATE", "text": excerpt})
                for match in INLINE_LABEL.finditer(raw):
                    unresolved.append({**ref, "kind": "INLINE_NUMBER_WHITESPACE_CANDIDATE",
                                       "label": match.group("label"), "text": excerpt})
                if any(k == "page_break" for k, _, _ in atoms):
                    unresolved.append({**ref, "kind": "MANUAL_PAGE_BREAK_REVIEW", "text": excerpt})
                for gap in formula_gaps(atoms):
                    issues.append({**ref, "kind": "TEXT_FORMULA_WHITESPACE", **gap, "text": excerpt})
                if any(k == "formula" for k, _, _ in atoms):
                    auto_space, origin = effective(layers, "autoSpaceDE")
                    if auto_space is not None and enabled(auto_space):
                        unresolved.append({**ref, "kind": "EAST_ASIAN_AUTO_SPACING_REVIEW", "origin": origin, "text": excerpt})
                graphic_no = 0
                for kind, description, floating in atoms:
                    if kind != "graphic":
                        continue
                    graphic_no += 1
                    graphics.append({**ref, "graphic": graphic_no, "type": description, "floating": floating,
                                     "text": excerpt})
                    unresolved.append({**ref, "kind": "GRAPHIC_ROLE_AND_CENTER_REVIEW", "graphic": graphic_no,
                                       "type": description, "floating": floating, "text": excerpt})

    numbering = inspect_numbering(docx)
    for item in numbering["candidates"]:
        issues.append({"story": item["story"], "paragraph": item["paragraph"], "kind": "NUMBERING_OR_HEADING",
                       "automatic_numbering": item["automatic_numbering"], "numbering_fields": item["numbering_fields"],
                       "heading_style_candidate": item["heading_style_candidate"],
                       "heading_style_reference_candidate": item["heading_style_reference_candidate"],
                       "outline_level": item["outline_level"], "outline_source": item["outline_source"],
                       "text": item["text_preview"]})
    after = digest(docx)
    if before != after:
        raise RuntimeError("Input changed during inspection; repeat with a frozen copy")
    return {
        "schema": "axmath-layout-inspection/v1", "docx": str(docx), "sha256": before,
        "paragraphs": paragraphs, "confirmed_or_candidate_findings": issues,
        "unresolved_require_word_or_source_review": unresolved, "graphics_inventory": graphics,
        "finding_counts": dict(Counter(x["kind"] for x in issues)),
        "unresolved_counts": dict(Counter(x["kind"] for x in unresolved)),
        "visual_and_editability_review_required": True,
        "limitations": ["Absent pagination values require Word effective-value verification.",
                        "Graphic role, caption association, floating positions and actual visual spacing require source and Word review.",
                        "Math internals and content completeness are not verified by this tool."],
    }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", help="Existing DOCX; it is never modified")
    parser.add_argument("--report", help="Write a new UTF-8 JSON report; refuses to overwrite")
    args = parser.parse_args()
    report = inspect_layout(Path(args.docx))
    encoded = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        path = Path(args.report).resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as stream:
            stream.write(encoded + "\n")
        print(json.dumps({"docx": report["docx"], "sha256": report["sha256"],
                          "finding_counts": report["finding_counts"], "unresolved_counts": report["unresolved_counts"],
                          "graphics": len(report["graphics_inventory"]), "report": str(path),
                          "visual_and_editability_review_required": True}, ensure_ascii=False, indent=2))
    else:
        print(encoded)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
