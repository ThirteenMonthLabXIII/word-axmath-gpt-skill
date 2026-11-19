from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import tempfile
import time
import unittest
from ctypes import wintypes
from pathlib import Path


CONVERTER = Path(__file__).with_name("convert_officemath_to_axmath.ps1")
POWERSHELL = shutil.which("pwsh") or shutil.which("powershell")
WORKER = r'''param([string]$Converter,[string]$Control,[string]$Ready)
$ErrorActionPreference = 'Stop'
$tokens = $null
$errors = $null
$tree = [System.Management.Automation.Language.Parser]::ParseFile($Converter,[ref]$tokens,[ref]$errors)
if ($errors.Count -ne 0) { throw 'Converter has PowerShell parse errors.' }
$function = $tree.Find({ param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Write-ControlState'
}, $true)
if ($null -eq $function) { throw 'Write-ControlState was not found.' }
. ([scriptblock]::Create($function.Extent.Text))
$ControlPath = $Control
$wordPidOwned = $false
$wordPid = $null
$wordCreationUnix = $null
$result = [ordered]@{ control_write_retries = 0 }
try {
    $probe = [IO.File]::Open($Control,[IO.FileMode]::Open,[IO.FileAccess]::Write,[IO.FileShare]::ReadWrite)
    $probe.Dispose()
    throw 'Expected the control file to be locked.'
} catch [IO.IOException] {
    [IO.File]::WriteAllText($Ready,'locked')
}
try {
    Write-ControlState -Phase 'retry_test' -Batch 1 -Before 1 -After 0
    [Console]::Out.WriteLine('retry_count=' + $result.control_write_retries)
    exit 0
} catch {
    [Console]::Error.WriteLine('retry_count=' + $result.control_write_retries)
    [Console]::Error.WriteLine($_.Exception.GetType().FullName + ': ' + $_.Exception.Message)
    exit 3
}
'''


@unittest.skipUnless(os.name == "nt" and POWERSHELL, "Windows and PowerShell are required")
class ControlStateRetryTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="axmath_control_retry_")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.worker = self.folder / "worker.ps1"
        self.worker.write_text(WORKER, encoding="utf-8")
        self.control = self.folder / "control.json"
        self.control.write_text('{"phase":"before"}', encoding="utf-8")
        self.ready = self.folder / "ready"
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        self.kernel.CreateFileW.argtypes = [
            wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID,
            wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
        ]
        self.kernel.CreateFileW.restype = wintypes.HANDLE
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.CloseHandle.restype = wintypes.BOOL

    def _lock(self):
        # Real Windows share denial, matching the collision Set-Content sees.
        handle = self.kernel.CreateFileW(str(self.control), 0x80000000, 0, None, 3, 0x80, None)
        if handle == ctypes.c_void_p(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        return handle

    def _worker(self):
        process = subprocess.Popen(
            [POWERSHELL, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(self.worker),
             "-Converter", str(CONVERTER), "-Control", str(self.control), "-Ready", str(self.ready)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.addCleanup(lambda: process.kill() if process.poll() is None else None)
        deadline = time.monotonic() + 8
        while not self.ready.exists() and time.monotonic() < deadline:
            if process.poll() is not None:
                stdout, stderr = process.communicate()
                self.fail(f"Worker exited before verifying the lock: {stdout} {stderr}")
            time.sleep(0.02)
        self.assertTrue(self.ready.exists(), "Worker did not confirm the lock")
        return process

    def test_retries_until_an_exclusive_lock_is_released(self) -> None:
        handle = self._lock()
        try:
            process = self._worker()
            time.sleep(0.6)
            self.assertIsNone(process.poll(), "Worker stopped while a transient lock was held")
        finally:
            self.kernel.CloseHandle(handle)
        stdout, stderr = process.communicate(timeout=8)
        self.assertEqual(process.returncode, 0, f"{stdout} {stderr}")
        self.assertIn("retry_count=", stdout)
        self.assertGreater(int(stdout.strip().split("retry_count=")[-1]), 0)
        self.assertEqual(json.loads(self.control.read_text(encoding="utf-8-sig"))["phase"], "retry_test")

    def test_persistent_lock_ends_with_an_error_after_retries(self) -> None:
        handle = self._lock()
        try:
            process = self._worker()
            start = time.monotonic()
            stdout, stderr = process.communicate(timeout=8)
            elapsed = time.monotonic() - start
            self.assertEqual(process.returncode, 3, f"{stdout} {stderr}")
            self.assertIn("retry_count=30", stderr)
            self.assertGreater(elapsed, 2.0, "Persistent lock should be retried before failing")
        finally:
            self.kernel.CloseHandle(handle)
        self.assertEqual(json.loads(self.control.read_text(encoding="utf-8"))["phase"], "before")


if __name__ == "__main__":
    unittest.main()
