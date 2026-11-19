param(
    [Parameter(Mandatory = $true)][string]$InputDocx,
    [Parameter(Mandatory = $true)][string]$Report
)

$ErrorActionPreference = 'Stop'
$inputPath = [System.IO.Path]::GetFullPath($InputDocx)
$reportPath = [System.IO.Path]::GetFullPath($Report)
if (-not [System.IO.File]::Exists($inputPath)) { throw "Input does not exist: $inputPath" }
if ($inputPath.Equals($reportPath, [StringComparison]::OrdinalIgnoreCase)) { throw 'Report must not be the input DOCX.' }
if ([System.IO.File]::Exists($reportPath)) { throw "Report already exists: $reportPath" }
$inputHash = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash
$priorWordIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$startedAt = Get-Date
$word = $null
$document = $null
$ownWord = $false
$wordProcId = 0
$results = [System.Collections.Generic.List[object]]::new()
$unresolved = [System.Collections.Generic.List[object]]::new()
$stories = [System.Collections.Generic.List[object]]::new()
$fields = @('WidowControl', 'KeepWithNext', 'KeepTogether', 'PageBreakBefore')

if (-not ('AxMathLayoutWordProcess' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AxMathLayoutWordProcess {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

try {
    $word = New-Object -ComObject Word.Application
    $newAutomation = @()
    for ($attempt = 0; $attempt -lt 8; $attempt++) {
        $newAutomation = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
            Where-Object { $priorWordIds -notcontains [int]$_.ProcessId -and
                $_.CommandLine -match '(^|\s)/Automation(\s|$)' -and $_.CommandLine -match '-Embedding' })
        if ($newAutomation.Count -gt 0) { break }
        Start-Sleep -Milliseconds 250
    }
    if ($newAutomation.Count -ne 1) { throw 'Could not identify exactly one new automation Word process; the document was not opened.' }
    $wordProcId = [int]$newAutomation[0].ProcessId
    $process = Get-Process -Id $wordProcId -ErrorAction Stop
    if ($process.ProcessName -ne 'WINWORD' -or $process.StartTime -lt $startedAt.AddSeconds(-2)) {
        throw 'Word process identity could not be verified; the document was not opened.'
    }
    $ownWord = $true
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    $windowHandle = [IntPtr]([int64]$document.ActiveWindow.Hwnd)
    [uint32]$windowProcess = 0
    [void][AxMathLayoutWordProcess]::GetWindowThreadProcessId($windowHandle, [ref]$windowProcess)
    if ([int]$windowProcess -ne $wordProcId) {
        $ownWord = $false
        throw 'The document window did not belong to the new Word process; effective-value inspection stopped.'
    }
    $totalParagraphs = 0

    foreach ($initialStory in $document.StoryRanges) {
        $range = $initialStory
        $segment = 0
        while ($null -ne $range) {
            $segment++
            $storyType = [int]$range.StoryType
            $paragraphCount = [int]$range.Paragraphs.Count
            $stories.Add([pscustomobject]@{ story_type = $storyType; segment = $segment; paragraphs = $paragraphCount })
            for ($i = 1; $i -le $paragraphCount; $i++) {
                $totalParagraphs++
                $paragraph = $range.Paragraphs.Item($i)
                $format = $paragraph.Range.ParagraphFormat
                foreach ($field in $fields) {
                    try {
                        $raw = [int]$format.$field
                        if ($raw -eq -1 -or $raw -eq 1) {
                            $results.Add([pscustomobject]@{
                                story_type = $storyType; segment = $segment; paragraph = $i
                                range_start = [int]$paragraph.Range.Start; property = $field; value = $raw
                            })
                        } elseif ($raw -ne 0) {
                            $unresolved.Add([pscustomobject]@{
                                story_type = $storyType; segment = $segment; paragraph = $i
                                range_start = [int]$paragraph.Range.Start; property = $field
                                reason = "unexpected Word value: $raw"
                            })
                        }
                    } catch {
                        $unresolved.Add([pscustomobject]@{
                            story_type = $storyType; segment = $segment; paragraph = $i
                            range_start = [int]$paragraph.Range.Start; property = $field
                            reason = $_.Exception.Message
                        })
                    }
                }
            }
            $range = $range.NextStoryRange
        }
    }
    $document.Close(0)
    $document = $null
    if ((Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash -ne $inputHash) {
        throw 'Input changed during the inspection; repeat on a frozen copy.'
    }
    $payload = [ordered]@{
        schema = 'axmath-word-pagination-inspection/v1'
        input = $inputPath
        input_sha256 = $inputHash
        word_process_id = $wordProcId
        paragraphs_examined = $totalParagraphs
        stories = @($stories.ToArray())
        enabled = @($results.ToArray())
        unresolved = @($unresolved.ToArray())
        pagination_values_pass = ($results.Count -eq 0 -and $unresolved.Count -eq 0 -and $totalParagraphs -gt 0)
        note = 'Read-only effective values only. Validate figure placement, numbering, content and visual layout separately.'
    }
    $json = $payload | ConvertTo-Json -Depth 8
    $parent = Split-Path -Parent $reportPath
    if (-not [System.IO.Directory]::Exists($parent)) { [void][System.IO.Directory]::CreateDirectory($parent) }
    $writer = [System.IO.StreamWriter]::new([System.IO.FileStream]::new($reportPath, [System.IO.FileMode]::CreateNew), [System.Text.UTF8Encoding]::new($false))
    try { $writer.Write($json) } finally { $writer.Dispose() }
    [pscustomobject]@{
        paragraphs_examined = $totalParagraphs; enabled = $results.Count; unresolved = $unresolved.Count
        pagination_values_pass = $payload.pagination_values_pass; report = $reportPath
    } | ConvertTo-Json
    if (-not $payload.pagination_values_pass) { exit 2 }
    exit 0
} finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $word -and $ownWord) { $word.Quit(0) }
    if ($null -ne $word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
