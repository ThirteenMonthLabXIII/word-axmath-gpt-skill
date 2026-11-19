"""Fast static OMML risk inventory; never rewrites formula semantics.

Flags derivative prime tokens and collection operators for targeted review.
A source marker is only a risk signal, never evidence that AxMath is broken.
"""
from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

from normalize_axmath_tex import PRIME_SOURCE_MARKERS, contains_prime_or_derivative_marker

NS = {
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
M = "{" + NS["m"] + "}"
SET_TOKENS = set("∪∩⊂⊆⊃⊇∈∉∅∖⋃⋂")
_PRIME_RUN_RE = re.compile(r"'+|′+|ʹ+|’+|‘+|[″ʺ]|‴|⁗")
_RELATION_SYMBOLS = {
    "≤": "le", "≥": "ge", "≠": "ne", "<": "lt", ">": "gt", "=": "eq",
    "+": "plus", "−": "minus", "-": "minus", "±": "pm", "∓": "mp",
    "×": "times", "÷": "divide", "·": "cdot", "⋅": "cdot",
}
_OPERATOR_SYMBOLS = {
    "∫": "integral", "∬": "double_integral", "∭": "triple_integral", "∮": "contour_integral",
    "∑": "sum", "∏": "product", "∐": "coproduct",
    "∪": "union", "∩": "intersection", "∈": "membership", "∉": "nonmembership",
    "→": "arrow", "↦": "mapsto", "⇒": "implies", "⇔": "equivalence",
    "∞": "infinity",
}


def _math_text(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return "".join((item.text or "") for item in node.findall(".//m:t", NS))


def _omml_hide_enabled(node: ET.Element | None) -> bool:
    if node is None:
        return False
    return node.get(M + "val", "1").strip().lower() not in {"0", "false", "off"}


def _source_prime_orders(text: str) -> list[int]:
    orders: list[int] = []
    for match in _PRIME_RUN_RE.finditer(text):
        token = match.group(0)
        if token in {"′", "ʹ", "’", "‘"}:
            orders.append(1)
        elif token in {"″", "ʺ"}:
            orders.append(2)
        elif token == "‴":
            orders.append(3)
        elif token == "⁗":
            orders.append(4)
        else:
            orders.append(len(token))
    return orders


def _limit_script_candidates(formula: ET.Element, ordinal: int) -> list[dict]:
    rows: list[dict] = []
    for node in formula.iter():
        local = node.tag[len(M):] if node.tag.startswith(M) else ""
        if local not in {"limLow", "limUpp", "sSub", "sSup", "sSubSup"}:
            continue
        base = node.find("m:e", NS)
        base_text = _math_text(base).strip()
        if base_text.lower() != "lim":
            continue
        if local in {"limLow", "sSub"}:
            placements = ["below"]
            script_tags = ["m:lim" if local == "limLow" else "m:sub"]
        elif local in {"limUpp", "sSup"}:
            placements = ["above"]
            script_tags = ["m:lim" if local == "limUpp" else "m:sup"]
        else:
            placements = ["below", "above"]
            script_tags = ["m:sub", "m:sup"]
        script_texts = [_math_text(node.find(tag, NS))[:120] for tag in script_tags]
        rows.append({
            "ordinal": ordinal,
            "source_node": local,
            "operator": "lim",
            "placements": placements,
            "script_texts": script_texts,
            "script_text": " / ".join(x for x in script_texts if x),
            "base_text": base_text,
            "reason": "lower_lim_script_must_use_underset_for_axmath_roundtrip",
        })
    return rows


def _formula_structure_candidate(formula: ET.Element, ordinal: int, inspected: str) -> dict | None:
    node_counts = Counter(
        node.tag[len(M):]
        for node in formula.iter()
        if node.tag.startswith(M)
    )
    relation_counts = Counter(
        token for char, token in _RELATION_SYMBOLS.items() if char in inspected
        for _ in range(inspected.count(char))
    )
    operator_counts = Counter(
        token for char, token in _OPERATOR_SYMBOLS.items() if char in inspected
        for _ in range(inspected.count(char))
    )
    delimiters = []
    for node in formula.iter(M + "d"):
        props = node.find("m:dPr", NS)
        if props is None:
            delimiters.append({"begin": "(", "end": ")"})
            continue
        begin = props.find("m:begChr", NS)
        end = props.find("m:endChr", NS)
        begin_value = begin.get(M + "val") if begin is not None else None
        end_value = end.get(M + "val") if end is not None else None
        # An explicitly empty OMML delimiter is intentional (e.g. an open
        # brace for a system of equations); only an absent value gets a default.
        delimiters.append({
            "begin": "(" if begin_value is None else begin_value,
            "end": ")" if end_value is None else end_value,
        })

    # OMML n-ary limits use m:sub / m:sup rather than sSub / sSup.
    # Count only nonempty visible limit content; an empty slot is not a script.
    nary_sub = nary_sup = nary_pair = 0
    for nary in formula.iter(M + "nary"):
        props = nary.find("m:naryPr", NS)
        sub_hide = props.find("m:subHide", NS) if props is not None else None
        sup_hide = props.find("m:supHide", NS) if props is not None else None
        has_sub = bool(_math_text(nary.find("m:sub", NS)).strip()) and not _omml_hide_enabled(sub_hide)
        has_sup = bool(_math_text(nary.find("m:sup", NS)).strip()) and not _omml_hide_enabled(sup_hide)
        nary_sub += int(has_sub)
        nary_sup += int(has_sup)
        nary_pair += int(has_sub and has_sup)

    # Word often represents a degree mark as a superscript circle, while
    # AxMath exports the same visually raised mark as a standalone \degree.
    degree_superscripts = sum(
        1 for node in formula.iter()
        if node.tag in {M + "sSup", M + "sSubSup"}
        and _math_text(node.find("m:sup", NS)).strip() in {"°", "∘"}
    )

    signature = {
        "fractions": node_counts["f"],
        "radicals": node_counts["rad"],
        "matrices": node_counts["m"],
        "equation_arrays": node_counts["eqArr"],
        "delimiters": delimiters,
        "subscripts": node_counts["sSub"] + node_counts["limLow"] + node_counts["sSubSup"] + nary_sub,
        "superscripts": node_counts["sSup"] + node_counts["limUpp"] + node_counts["sSubSup"] + nary_sup,
        "degree_superscripts": degree_superscripts,
        "subsuperscripts": node_counts["sSubSup"] + nary_pair,
        "relations": dict(sorted(relation_counts.items())),
        "operators": dict(sorted(operator_counts.items())),
    }
    if not any((
        signature["fractions"], signature["radicals"], signature["matrices"],
        signature["equation_arrays"], signature["delimiters"],
        signature["subscripts"], signature["superscripts"],
        signature["subsuperscripts"], signature["relations"], signature["operators"],
    )):
        return None
    return {
        "ordinal": ordinal,
        "source_text": inspected[:180],
        "structure": signature,
        "reason": "source_formula_structure_requires_internal_axmath_roundtrip_review",
    }


def scan_source_math(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    math_nodes = root.findall(".//m:oMath", NS)
    primes, sets, standalone, limits, structures = [], [], [], [], []
    for ordinal, formula in enumerate(math_nodes, 1):
        text = "".join((node.text or "") for node in formula.findall(".//m:t", NS))
        # OMML may carry operators/accents in attribute values, not m:t text.
        chars = [
            node.get(M + "val") or ""
            for node in formula.iter()
            if node.tag == M + "chr"
        ]
        # OMML n-ary operators default to U+222B when naryPr/chr is omitted;
        # Word commonly omits the explicit integral character on save.
        for nary in formula.iter(M + "nary"):
            props = nary.find("m:naryPr", NS)
            if props is None or props.find("m:chr", NS) is None:
                chars.append("∫")
        inspected = text + "".join(chars)
        limits.extend(_limit_script_candidates(formula, ordinal))
        structure = _formula_structure_candidate(formula, ordinal, inspected)
        if structure is not None:
            structures.append(structure)
        if contains_prime_or_derivative_marker(inspected):
            primes.append({
                "ordinal": ordinal,
                "source_text": text[:180],
                "prime_symbols": sorted({c for c in inspected if c in PRIME_SOURCE_MARKERS}),
                "prime_orders": _source_prime_orders(inspected),
                "reason": "source_prime_semantics_require_axmath_contract_review",
            })
        tokens = sorted(set(inspected) & SET_TOKENS)
        if tokens:
            sets.append({
                "ordinal": ordinal,
                "source_text": text[:180],
                "symbols": tokens,
                "reason": "collection_symbol_size_and_glyph_visual_review",
            })
    # Word may implicitly center naked standalone OMML while the resulting OLE
    # follows the paragraph's left alignment. COM x_pt can misleadingly match.
    ordinal = 0
    for pi, paragraph in enumerate(root.findall(".//w:p", NS), 1):
        equations = paragraph.findall(".//m:oMath", NS)
        if (
            len(equations) == 1
            and len(paragraph.findall("./m:oMath", NS)) == 1
            and not paragraph.findall(".//m:oMathPara", NS)
            and not paragraph.findall(".//w:t", NS)
            and not paragraph.findall(".//w:object", NS)
        ):
            standalone.append({
                "ordinal": ordinal + 1,
                "paragraph_index": pi,
                "reason": "standalone_native_math_can_center_while_axmath_ole_is_left",
            })
        ordinal += len(equations)
    return {
        "schema": "axmath-source-semantic-risks/v1",
        "source": str(Path(path).resolve()),
        "formula_count": len(math_nodes),
        "prime_semantic_candidates": primes,
        "limit_subscript_candidates": limits,
        "limit_subscript_ordinals": sorted({row["ordinal"] for row in limits}),
        "math_structure_candidates": structures,
        "math_structure_ordinals": [row["ordinal"] for row in structures],
        "set_symbol_visual_candidates": sets,
        "standalone_display_alignment_visual_candidates": standalone,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--out")
    args = parser.parse_args()
    result = scan_source_math(Path(args.source))
    data = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(data, encoding="utf-8")
    print(data)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
