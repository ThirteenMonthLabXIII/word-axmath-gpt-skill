from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import Mock, patch
from xml.sax.saxutils import escape

from body_math import _cleanup_owned_word
from body_math_text_contract import compare_body_math_text
from math_layout_contract import _semantic_tokens

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"


def docx(path: Path, parts: list[str | None]) -> None:
    body = []
    for part in parts:
        if part is None:
            body.append('<w:r><w:object><o:OLEObject ProgID="Equation.AxMath"/></w:object></w:r>')
        else:
            body.append(f'<w:r><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    xml = f'<w:document xmlns:w="{W}" xmlns:o="{O}"><w:body><w:p>{"".join(body)}</w:p></w:body></w:document>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", xml)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class BodyMathTextContractTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.source = root / "source.docx"
        self.final = root / "final.docx"
        self.map = root / "map.json"
        docx(self.source, ["1. 设x=1，求x。"])
        docx(self.final, ["1. 设", None, "=", None, "，求", None, "。"])
        self.data = {
            "schema": "axmath-source-final-map/v1",
            "source": {"path": str(self.source), "sha256": sha(self.source)},
            "final": {"path": str(self.final), "sha256": sha(self.final)},
            "targets": [
                {"target_id": f"BODY-{i}", "kind": "added_body_math", "source_text": char,
                 "expected_tex": char, "final_ordinals": [i]}
                for i, char in enumerate(("x", "1", "x"), 1)
            ],
        }

    def compare(self) -> dict:
        self.map.write_text(json.dumps(self.data, ensure_ascii=False), encoding="utf-8")
        return compare_body_math_text(self.source, self.final, self.map)

    def test_repeated_letters_and_digit_preserve_question_number(self) -> None:
        self.assertTrue(self.compare()["passed"])

    def test_duplicate_character_or_changed_explanation_fails(self) -> None:
        for pieces in (
            ["1. 设x", None, "=", None, "，求", None, "。"],
            ["1. 已知", None, "=", None, "，求", None, "。"],
        ):
            with self.subTest(pieces=pieces):
                docx(self.final, pieces)
                self.data["final"]["sha256"] = sha(self.final)
                self.assertFalse(self.compare()["passed"])

    def test_wrong_character_or_object_location_fails(self) -> None:
        self.data["targets"][1]["source_text"] = "2"
        self.assertFalse(self.compare()["passed"])
        self.data["targets"][1]["source_text"] = "1"
        docx(self.final, ["1. ", None, "设=", None, "，求", None, "。"])
        self.data["final"]["sha256"] = sha(self.final)
        self.assertFalse(self.compare()["passed"])

    def test_missing_text_duplicate_ownership_and_stale_hash_fail(self) -> None:
        self.data["targets"][0].pop("source_text")
        self.assertFalse(self.compare()["passed"])
        self.data["targets"][0]["source_text"] = "x"
        self.data["targets"][1]["final_ordinals"] = [1]
        self.assertFalse(self.compare()["passed"])
        self.data["targets"][1]["final_ordinals"] = [2]
        self.data["final"]["sha256"] = "0" * 64
        self.assertFalse(self.compare()["passed"])

    def test_existing_axmath_before_new_body_targets(self) -> None:
        docx(self.source, ["1. 设", None, "=1，求x。"])
        docx(self.final, ["1. 设", None, "=", None, "，求", None, "。"])
        self.data["source"]["sha256"] = sha(self.source)
        self.data["final"]["sha256"] = sha(self.final)
        self.data["targets"][0] = {"target_id": "SRC-1", "kind": "source_formula", "final_ordinals": [1]}
        self.assertTrue(self.compare()["passed"])

    def test_axmath_export_delimiters_are_not_semantic_tokens(self) -> None:
        self.assertEqual(_semantic_tokens("$x$\r"), _semantic_tokens("x"))
        self.assertEqual(_semantic_tokens("$1$\r"), _semantic_tokens("1"))
        self.assertNotEqual(_semantic_tokens("$x$\r"), _semantic_tokens("y"))
        self.assertNotEqual(_semantic_tokens(r"\$"), _semantic_tokens(""))

    def test_greek_source_character_must_match_expected_tex(self) -> None:
        docx(self.source, ["设α为参数。"])
        docx(self.final, ["设", None, "为参数。"])
        self.data["source"]["sha256"] = sha(self.source)
        self.data["final"]["sha256"] = sha(self.final)
        self.data["targets"] = [{"target_id": "BODY-G", "kind": "added_body_math",
                                 "source_text": "α", "expected_tex": r"\alpha", "final_ordinals": [1]}]
        self.assertTrue(self.compare()["passed"])
        self.data["targets"][0]["expected_tex"] = r"\beta"
        self.assertFalse(self.compare()["passed"])


class OwnedWordCleanupTests(unittest.TestCase):
    def test_cleanup_requires_matching_word_name_and_creation_time(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            identity = Path(folder) / "identity.json"
            unix_time = 1_700_000_000
            identity.write_text(json.dumps({
                "pid": 123, "start_utc_ticks": (62_135_596_800 + unix_time) * 10_000_000,
            }), encoding="utf-8")
            process = Mock()
            process.name.return_value = "wps.exe"
            process.create_time.return_value = unix_time
            with patch("body_math.psutil.Process", return_value=process):
                self.assertFalse(_cleanup_owned_word(identity))
                process.terminate.assert_not_called()
                process.name.return_value = "WINWORD.EXE"
                process.create_time.return_value = unix_time + 2
                self.assertFalse(_cleanup_owned_word(identity))
                process.terminate.assert_not_called()
                process.create_time.return_value = unix_time
                self.assertTrue(_cleanup_owned_word(identity))
                process.terminate.assert_called_once()


if __name__ == "__main__":
    unittest.main()
