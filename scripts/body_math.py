"""Inspect and replace reviewed single-character body math with AxMath."""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import psutil
from execution_lock import conversion_lock
from math_layout_contract import _semantic_tokens


def _cleanup_owned_word(identity_file: Path) -> bool:
    """On timeout, stop only the Word PID and creation time proven by the child."""
    if not identity_file.is_file():
        return False
    try:
        identity = json.loads(identity_file.read_text(encoding="utf-8-sig"))
        process = psutil.Process(int(identity["pid"]))
        expected_start = int(identity["start_utc_ticks"]) / 10_000_000 - 62_135_596_800
        if process.name().lower() != "winword.exe" or abs(process.create_time() - expected_start) > 0.1:
            return False
        process.terminate()
        try:
            process.wait(timeout=10)
        except psutil.TimeoutExpired:
            process.kill()
            process.wait(timeout=10)
        return True
    except (KeyError, ValueError, OSError, json.JSONDecodeError, psutil.Error):
        return False


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="mode", required=True)
    for name in ("inspect", "apply"):
        command = sub.add_parser(name)
        command.add_argument("--input", type=Path, required=True)
        command.add_argument("--report", type=Path, required=True)
        command.add_argument("--timeout", type=int, default=1800)
        if name == "apply":
            command.add_argument("--plan", type=Path, required=True)
            command.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("This command requires Windows, Word, and AxMath.")
    if args.timeout < 1:
        parser.error("--timeout must be positive")
    if args.mode == "apply":
        try:
            plan = json.loads(args.plan.read_text(encoding="utf-8-sig"))
            for row in plan.get("targets") or []:
                source_text = str(row.get("source_text") or "")
                tex = str(row.get("expected_tex") or "")
                if len(source_text) == 1 and source_text.isalnum():
                    if _semantic_tokens(source_text) != _semantic_tokens(tex):
                        raise ValueError(f"{row.get('target_id')}: source character and expected_tex differ")
        except (OSError, ValueError, AttributeError, json.JSONDecodeError) as exc:
            parser.error(f"Invalid body-math plan: {exc}")

    script = Path(__file__).with_name("insert_body_math.ps1")
    with conversion_lock(), tempfile.TemporaryDirectory(prefix="axmath-body-word-") as temp_dir:
        identity = Path(temp_dir) / "identity.json"
        cmd = [
            "powershell.exe", "-NoProfile", "-File", str(script),
            "-Mode", args.mode, "-InputDocx", str(args.input.resolve()),
            "-ReportPath", str(args.report.resolve()), "-IdentityPath", str(identity),
        ]
        if args.mode == "apply":
            cmd += ["-PlanPath", str(args.plan.resolve()), "-OutputDocx", str(args.output.resolve())]
        try:
            result = subprocess.run(cmd, capture_output=True, timeout=args.timeout)
        except subprocess.TimeoutExpired:
            stopped = _cleanup_owned_word(identity)
            print(f"Body-math operation timed out; verified task-owned Word stopped: {stopped}", file=sys.stderr)
            return 1
    for stream, payload in ((sys.stdout, result.stdout), (sys.stderr, result.stderr)):
        if payload:
            stream.write(payload.decode("utf-8", errors="replace"))
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
