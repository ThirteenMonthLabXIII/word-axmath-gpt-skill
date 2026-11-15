"""Check that mapped body text was replaced by AxMath without text drift.

This check is deliberately limited to main-document text and mapped inline
AxMath objects. The formula-map and complete internal-TeX checks remain
separate: this module only verifies the positions and ordinary-text projection.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from audit_docx import M, NS, W, analyze
from math_layout_contract import _semantic_tokens

O = "{urn:schemas-microsoft-com:office:office}"
SCHEMA = "axmath-source-final-map/v1"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def compare_body_math_text(source: Path, final: Path, formula_map: Path) -> dict:
    source, final, formula_map = (p.resolve() for p in (source, final, formula_map))
    data = json.loads(formula_map.read_text(encoding="utf-8-sig"))
    errors: list[str] = []
    if data.get("schema") != SCHEMA:
        errors.append("unsupported formula-map schema")
    for label, path in (("source", source), ("final", final)):
        meta = data.get(label) or {}
        if Path(str(meta.get("path") or "")).resolve() != path:
            errors.append(f"{label} path differs from formula map")
        if str(meta.get("sha256") or "").lower() != _sha256(path):
            errors.append(f"{label} hash differs from formula map")

    targets = data.get("targets")
    if not isinstance(targets, list):
        targets = []
        errors.append("formula map has no targets array")

    body_by_ordinal: dict[int, str] = {}
    owners: dict[int, str] = {}
    target_ids: set[str] = set()
    for target in targets:
        if not isinstance(target, dict):
            errors.append("formula-map target is not an object")
            continue
        target_id = str(target.get("target_id") or "")
        if not target_id or target_id in target_ids:
            errors.append(f"missing or duplicate target_id: {target_id!r}")
        target_ids.add(target_id)
        ordinals = target.get("final_ordinals")
        if not isinstance(ordinals, list) or not ordinals:
            errors.append(f"{target_id}: missing final ordinals")
            continue
        for ordinal in ordinals:
            if isinstance(ordinal, bool) or not isinstance(ordinal, int) or ordinal < 1:
                errors.append(f"{target_id}: invalid final ordinal")
                continue
            if ordinal in owners:
                errors.append(f"final ordinal {ordinal} has multiple owners")
            owners[ordinal] = target_id
        if target.get("kind") != "added_body_math":
            continue
        source_text = target.get("source_text")
        if not isinstance(source_text, str) or not source_text:
            errors.append(f"{target_id}: missing exact source_text")
            continue
        if len(ordinals) != 1 or not isinstance(ordinals[0], int):
            errors.append(f"{target_id}: text projection requires one final ordinal")
            continue
        expected_tex = str(target.get("expected_tex") or "").strip()
        if len(source_text) == 1 and source_text.isalnum():
            if _semantic_tokens(expected_tex) != _semantic_tokens(source_text):
                errors.append(f"{target_id}: single source character differs from expected_tex")
        body_by_ordinal[ordinals[0]] = source_text

    with zipfile.ZipFile(final) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    axmath = [node for node in root.iter(O + "OLEObject") if node.get("ProgID") == "Equation.AxMath"]
    if set(owners) != set(range(1, len(axmath) + 1)):
        errors.append("formula-map owners do not cover the final AxMath objects exactly")
    ordinal_by_element = {id(node): ordinal for ordinal, node in enumerate(axmath, 1)}
    projected: list[str] = []
    found_body: set[int] = set()

    for paragraph in root.findall(".//w:p", NS):
        chunks: list[str] = []

        def walk(node: ET.Element, in_math: bool = False) -> None:
            if node.tag in (M + "oMath", M + "oMathPara"):
                in_math = True
            if node.tag == W + "object":
                objects = [x for x in node.iter(O + "OLEObject") if id(x) in ordinal_by_element]
                for obj in objects:
                    ordinal = ordinal_by_element[id(obj)]
                    if ordinal in body_by_ordinal:
                        chunks.append(body_by_ordinal[ordinal])
                        found_body.add(ordinal)
                return
            if node.tag == W + "t" and not in_math:
                chunks.append(node.text or "")
            for child in node:
                walk(child, in_math)

        walk(paragraph)
        projected.append("".join(chunks))

    if found_body != set(body_by_ordinal):
        errors.append("a mapped body AxMath object is not an inline w:object in the main document")
    source_plain = analyze(source)["plain"]
    differences = [
        {"paragraph": i + 1, "source": a, "projected_final": b}
        for i, (a, b) in enumerate(zip(source_plain, projected)) if a != b
    ]
    if len(source_plain) != len(projected):
        errors.append("source and final paragraph counts differ")
    if differences:
        errors.append("ordinary text or mapped body-object positions differ")
    return {
        "schema": "axmath-body-math-text-contract/v1",
        "source": str(source),
        "final": str(final),
        "formula_map": str(formula_map),
        "body_target_count": len(body_by_ordinal),
        "passed": not errors,
        "errors": errors,
        "differences": differences[:50],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--final", type=Path, required=True)
    parser.add_argument("--formula-map", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args()
    result = compare_body_math_text(args.source, args.final, args.formula_map)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"passed": result["passed"], "errors": result["errors"]}, ensure_ascii=False))
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
