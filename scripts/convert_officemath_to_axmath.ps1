param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputDocx,
  [string]$TemplatePath='',
  [string]$ReportPath='',
  [string]$ControlPath='',
  [switch]$OverwriteOutput,
  [switch]$ResumeOutput
)
$ErrorActionPreference='Stop'
$InputFull=[IO.Path]::GetFullPath($InputDocx)
$OutputFull=[IO.Path]::GetFullPath($OutputDocx)
if([string]::Equals($InputFull,$OutputFull,[StringComparison]::OrdinalIgnoreCase)){throw 'Refusing to overwrite the input DOCX.'}
if(-not (Test-Path -LiteralPath $InputFull)){throw "Input DOCX not found: $InputFull"}
if($ResumeOutput -and $OverwriteOutput){throw 'ResumeOutput and OverwriteOutput cannot be combined.'}
if($ResumeOutput -and -not (Test-Path -LiteralPath $OutputFull)){throw "Resume working DOCX does not exist: $OutputFull"}
if(-not $ResumeOutput -and (Test-Path -LiteralPath $OutputFull) -and -not $OverwriteOutput){throw "Output already exists: $OutputFull. Use -OverwriteOutput only for an intentional intermediate replacement."}
function Resolve-AxMathTemplate {
  param([string]$Requested)
  if(-not [string]::IsNullOrWhiteSpace($Requested)){
    if(Test-Path -LiteralPath $Requested){ return (Resolve-Path -LiteralPath $Requested).Path }
    throw "AxMath.dotm not found: $Requested"
  }
  $candidates=@()
  if(${env:ProgramFiles(x86)}){$candidates += (Join-Path ${env:ProgramFiles(x86)} 'AxMath\MSOffice\AxMath.dotm')}
  if($env:ProgramFiles){$candidates += (Join-Path $env:ProgramFiles 'AxMath\MSOffice\AxMath.dotm')}
  foreach($candidate in $candidates){
    if(Test-Path -LiteralPath $candidate){ return $candidate }
  }
  throw 'AxMath.dotm not found. Install/activate AxMath or pass -TemplatePath explicitly.'
}
$TemplatePath=Resolve-AxMathTemplate $TemplatePath
if([string]::IsNullOrWhiteSpace($ReportPath)){$ReportPath=$OutputDocx+'.conversion.json'}
if([string]::IsNullOrWhiteSpace($ControlPath)){$ControlPath=$OutputDocx+'.control.json'}
$ReportFull=[IO.Path]::GetFullPath($ReportPath)
$ControlFull=[IO.Path]::GetFullPath($ControlPath)
foreach($pair in @(@('ReportPath',$ReportFull),@('ControlPath',$ControlFull))){
  if([string]::Equals($pair[1],$InputFull,[StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($pair[1],$OutputFull,[StringComparison]::OrdinalIgnoreCase)){
    throw "$($pair[0]) must not point to the input or output DOCX."
  }
}
if([string]::Equals($ReportFull,$ControlFull,[StringComparison]::OrdinalIgnoreCase)){throw 'ReportPath and ControlPath must be different files.'}

$key='HKCU:\Software\AxMath\WordCmds'
$result=[ordered]@{
  source=$InputDocx
  output=$OutputDocx
  macro='AMSMML2AM'
  strategy='official-batch-loop-mode0-guarded'
  success=$false
  complete=$false
  batches=@()
  control_write_retries=0
  word_forced_cleanup=$false
}
$total=[Diagnostics.Stopwatch]::StartNew()
$word=$null
$doc=$null
$oldNo=$null
$oldWait=$null
$beforeDumps=@()
$knownDumps=@()
$wordPid=$null
$wordPidOwned=$false
$candidateWordPidValid=$false
$normalWasSaved=$null
$checkpointPath=$OutputFull+'.conversion-checkpoint.json'
$sourceSha=$null
$resumeState=$null
$wordCreationUnix=$null
$wordStartTicks=$null

function Get-SharedSha256 {
  param([string]$Path)
  $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
  try {
    $hash=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($hash.ComputeHash($stream))).Replace('-','').ToLowerInvariant() }
    finally { $hash.Dispose() }
  } finally { $stream.Dispose() }
}

function Test-TaskWordIdentity {
  param([int]$TaskPid,[long]$StartTicks)
  if($TaskPid -le 0 -or $StartTicks -le 0){return $false}
  $process=Get-Process -Id $TaskPid -ErrorAction SilentlyContinue
  if($null -eq $process){return $false}
  try { return ($process.ProcessName -eq 'WINWORD' -and $process.StartTime.ToUniversalTime().Ticks -eq $StartTicks) }
  catch { return $false }
}

function Write-ConversionCheckpoint {
  param([int]$Remaining,[int]$LastBatch,[string]$Status='saved')
  # Save() finished before this checkpoint is published. The output hash makes
  # a crash between a Word save and its journal update fail closed on resume.
  $state=[ordered]@{
    schema='axmath-conversion-checkpoint/v1'
    source_path=$InputFull
    source_sha256=$sourceSha
    working_path=$OutputFull
    working_sha256=(Get-SharedSha256 $OutputFull)
    remaining_omath=$Remaining
    last_batch=$LastBatch
    status=$Status
    updated_at=(Get-Date).ToUniversalTime().ToString('o')
  }
  $tmp=$checkpointPath+'.'+[Guid]::NewGuid().ToString('N')+'.tmp'
  try {
    [IO.File]::WriteAllText($tmp,($state|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
    if(Test-Path -LiteralPath $checkpointPath){
      # File.Replace requires a non-empty backup path on .NET/PowerShell 7.
      # Keep the previous journal safe until the atomic replacement succeeds.
      $backup=$checkpointPath+'.'+[Guid]::NewGuid().ToString('N')+'.bak'
      [IO.File]::Replace($tmp,$checkpointPath,$backup)
      if(Test-Path -LiteralPath $backup){Remove-Item -LiteralPath $backup -Force}
    } else {
      [IO.File]::Move($tmp,$checkpointPath)
    }
  } finally {
    if(Test-Path -LiteralPath $tmp){Remove-Item -LiteralPath $tmp -Force}
  }
}

function Write-ControlState {
  param(
    [string]$Phase,
    [string]$Token='',
    [int]$Batch=0,
    [object[]]$Preexisting=@(),
    [int]$Before=0,
    [int]$After=0
  )
  $state=[ordered]@{
    phase=$Phase
    token=$Token
    batch=$Batch
    preexisting_dialog_pids=@($Preexisting)
    before_omath=$Before
    after_omath=$After
    owned_word_pid= $(if($wordPidOwned){$wordPid}else{$null})
    owned_word_create_unix=$wordCreationUnix
    updated_at=(Get-Date).ToString('o')
  }
  $json=$state|ConvertTo-Json -Depth 4
  # The dialog watcher may briefly hold the state file while reading it. A
  # sharing violation after a successful checkpoint should not abort the next
  # conversion batch; fail closed if the collision persists.
  for($attempt=0;$attempt -lt 30;$attempt++){
    try {
      Set-Content -LiteralPath $ControlPath -Value $json -Encoding UTF8
      break
    } catch [System.IO.IOException] {
      $result.control_write_retries=[int]$result.control_write_retries+1
      if($attempt -eq 29){throw}
      Start-Sleep -Milliseconds 100
    }
  }
}

try {
  $sourceSha=Get-SharedSha256 $InputFull
  if($ResumeOutput){
    if(-not (Test-Path -LiteralPath $checkpointPath)){throw "Resume refused: no saved conversion checkpoint at $checkpointPath"}
    $resumeState=Get-Content -LiteralPath $checkpointPath -Raw -Encoding UTF8|ConvertFrom-Json
    if(
      $resumeState.schema -ne 'axmath-conversion-checkpoint/v1' -or
      -not [string]::Equals([string]$resumeState.source_path,$InputFull,[StringComparison]::OrdinalIgnoreCase) -or
      -not [string]::Equals([string]$resumeState.working_path,$OutputFull,[StringComparison]::OrdinalIgnoreCase) -or
      $resumeState.source_sha256 -ne $sourceSha -or
      $resumeState.working_sha256 -ne (Get-SharedSha256 $OutputFull)
    ){throw 'Resume refused: source/working SHA-256 or checkpoint provenance changed.'}
    $result.resumed_from_checkpoint=$true
    $result.resumed_at_batch=[int]$resumeState.last_batch
  } else {
    if(Test-Path -LiteralPath $checkpointPath){throw "Refusing fresh conversion with a pre-existing recovery checkpoint: $checkpointPath"}
    $result.resumed_from_checkpoint=$false
  }
  $existingWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $result.preexisting_word_pids=@($existingWord)

  if(-not $ResumeOutput){
    if(Test-Path -LiteralPath $OutputDocx){Remove-Item -LiteralPath $OutputDocx -Force}
    Copy-Item -LiteralPath $InputDocx -Destination $OutputDocx
  }
  # A frozen source may intentionally carry the Windows read-only attribute.
  # The working copy must be writable or Word will divert Save() into an
  # interactive Save As dialog after the first AxMath batch.
  $outputItem=Get-Item -LiteralPath $OutputFull
  $result.output_readonly_cleared=[bool]$outputItem.IsReadOnly
  if($outputItem.IsReadOnly){$outputItem.IsReadOnly=$false}

  if(-not(Test-Path $key)){New-Item -Path $key -Force|Out-Null}
  $props=Get-ItemProperty $key
  $oldNo=$props.DoNoWinVerb
  $oldWait=$props.WaitingConvert
  $result.old_DoNoWinVerb=$oldNo
  $result.old_WaitingConvert=$oldWait
  $crashDir=Join-Path $env:LOCALAPPDATA 'CrashDumps'
  $beforeDumps=@(
    Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
  )
  $knownDumps=@($beforeDumps)

  Set-ItemProperty $key -Name DoNoWinVerb -Value 0
  Set-ItemProperty $key -Name WaitingConvert -Value 0

  Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class WordPidNative {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@ -ErrorAction SilentlyContinue

  $wordSpawnedAt=Get-Date
  $word=New-Object -ComObject Word.Application

  # Identify a distinct automation-process candidate before opening the task
  # document; its window PID must then match before any application settings
  # or AxMath macros are changed.
  # Ignore unrelated interactive Word windows launched at the same time.
  # A fresh COM instance on this machine starts as /Automation -Embedding.
  $newWord=@()
  for($attempt=0;$attempt -lt 8;$attempt++){
    $newWord=@(
      Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
      Where-Object {
        $existingWord -notcontains [int]$_.ProcessId -and
        $_.CommandLine -match '(^|\s)/Automation(\s|$)' -and
        $_.CommandLine -match '-Embedding' -and
        $_.CreationDate -ge $wordSpawnedAt.AddSeconds(-2)
      } |
      ForEach-Object { [int]$_.ProcessId }
    )
    if($newWord.Count -gt 0){break}
    Start-Sleep -Milliseconds 250
  }
  if($newWord.Count -eq 1){
    $wordPid=[int]$newWord[0]
    $candidateWordPidValid=$true
    $result.word_pid=$wordPid
    $result.word_pid_owned=$false
  } else {
    try {
      [uint32]$pidValue=0
      [void][WordPidNative]::GetWindowThreadProcessId([IntPtr]$word.Hwnd,[ref]$pidValue)
      if($pidValue -gt 0){
        $wordPid=[int]$pidValue
        $candidateWordPidValid=($newWord -contains $wordPid)
        $result.word_pid=$wordPid
        $result.word_pid_owned=$false
      }
    } catch {}
  }
  if(-not $candidateWordPidValid){
    throw "Could not establish an isolated task-owned Word process; pre-existing Word processes were left untouched."
  }
  $wordStarted=(Get-Process -Id $wordPid).StartTime.ToUniversalTime()
  $wordStartTicks=$wordStarted.Ticks
  $wordCreationUnix=[DateTimeOffset]::new($wordStarted).ToUnixTimeSeconds()

  # Confirm that the opened task document actually belongs to the candidate
  # process before changing Word options, loading AxMath or running a macro.
  $doc=$word.Documents.Open($OutputDocx)
  [uint32]$documentPid=0
  $window=[IntPtr]([int64]$doc.ActiveWindow.Hwnd)
  [void][WordPidNative]::GetWindowThreadProcessId($window,[ref]$documentPid)
  if([int]$documentPid -ne $wordPid){
    throw 'Task document window did not belong to the isolated Word process.'
  }
  $wordPidOwned=$true
  $result.word_pid_owned=$true
  $result.word_window_pid_verified=$true

  $word.Visible=$false
  $word.DisplayAlerts=0
  try{$word.ScreenUpdating=$false}catch{}
  try{$word.Options.CheckSpellingAsYouType=$false}catch{}
  try{$word.Options.CheckGrammarAsYouType=$false}catch{}
  try{$word.Options.SaveNormalPrompt=$false}catch{}
  try{$normalWasSaved=[bool]$word.NormalTemplate.Saved}catch{}

  $found=$false
  foreach($addin in $word.AddIns){
    if($addin.Name -eq 'AxMath.dotm'){
      $found=$true
      if(-not $addin.Installed){$addin.Installed=$true}
    }
  }
  $addin=$null
  if(-not $found){$word.AddIns.Add($TemplatePath,$true)|Out-Null}

  $doc.Activate()
  $result.before_omath=$doc.OMaths.Count
  $result.before_inline_shapes=$doc.InlineShapes.Count
  $result.before_paragraphs=$doc.Paragraphs.Count

  if($ResumeOutput){
    if([int]$resumeState.remaining_omath -ne [int]$doc.OMaths.Count){
      throw "Resume refused: saved OfficeMath count disagrees with checkpoint."
    }
    $batch=[int]$resumeState.last_batch
  } else {
    $batch=0
    Write-ConversionCheckpoint -Remaining ([int]$doc.OMaths.Count) -LastBatch 0 -Status 'started'
  }
  # AxMath itself controls how many equations a macro call converts (~64).
  # There is no arbitrary limit on the number of calls; strict progress
  # checks and the external per-batch watchdog bound failures instead.
  while($doc.OMaths.Count -gt 0){
    $batch++
    $before=$doc.OMaths.Count
    Set-ItemProperty $key -Name WaitingConvert -Value 0

    $preDialogs=@(Get-Process AxMMDlg -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $token=('batch-{0}-{1}-{2}' -f $batch,$before,[DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
    Write-ControlState -Phase 'awaiting_batch_dialog' -Token $token -Batch $batch -Preexisting $preDialogs -Before $before

    $doc.Content.Select()
    $cb=$null
    $sw=[Diagnostics.Stopwatch]::StartNew()
    $word.Run('AMSMML2AM',([ref]$cb))
    $sw.Stop()

    $after=$doc.OMaths.Count
    Write-ControlState -Phase 'saving_batch' -Token $token -Batch $batch -Before $before -After $after

    $converted=$before-$after
    $waitState=(Get-ItemProperty $key).WaitingConvert
    $currentDumps=@(
      Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
      Select-Object -ExpandProperty Name
    )
    $batchDumps=@($currentDumps | Where-Object {$_ -notin $knownDumps})
    $knownDumps=@($currentDumps)
    $batchRec=[ordered]@{
      batch=$batch
      before_omath=$before
      after_omath=$after
      converted=$converted
      seconds=$sw.Elapsed.TotalSeconds
      waiting_convert=$waitState
      save_seconds=0.0
      checkpoint_seconds=0.0
      crash_dumps_new=@($batchDumps)
    }

    if($converted -le 0){
      $result.batches += [pscustomobject]$batchRec
      throw "AxMath batch $batch made no progress ($before -> $after)."
    }

    # Keep the per-batch save. AxMath 2.7.x has produced crash dumps during
    # successful large-document conversions; persisting each completed batch
    # is a recovery boundary, not disposable overhead.
    $saveSw=[Diagnostics.Stopwatch]::StartNew()
    $doc.Save()
    $saveSw.Stop()
    $batchRec.save_seconds=$saveSw.Elapsed.TotalSeconds
    $checkpointSw=[Diagnostics.Stopwatch]::StartNew()
    Write-ConversionCheckpoint -Remaining ([int]$doc.OMaths.Count) -LastBatch $batch
    $checkpointSw.Stop()
    $batchRec.checkpoint_seconds=$checkpointSw.Elapsed.TotalSeconds
    Write-ControlState -Phase 'idle' -Token $token -Batch $batch -Before $before -After $after
    $result.batches += [pscustomobject]$batchRec
    Write-Output ('BATCH {0}: {1}->{2}, converted={3}, macro={4:N2}s, save={5:N2}s, crashes={6}' -f $batch,$before,$after,$converted,$sw.Elapsed.TotalSeconds,$saveSw.Elapsed.TotalSeconds,$batchDumps.Count)
  }

  $result.after_omath_com=$doc.OMaths.Count
  $result.after_inline_shapes_com=$doc.InlineShapes.Count
  $result.after_paragraphs=$doc.Paragraphs.Count
  $result.complete=($doc.OMaths.Count -eq 0)
  if(-not $result.complete){
    throw "Conversion stopped with $($doc.OMaths.Count) OfficeMath equations remaining."
  }

  try{$word.ScreenUpdating=$true}catch{}
  $result.pages=$doc.ComputeStatistics(2)
  $doc.Save()
  Write-ConversionCheckpoint -Remaining 0 -LastBatch $batch -Status 'complete'
  $result.success=$true
}
catch {
  $result.error=$_.Exception.Message
  $result.hresult=$_.Exception.HResult
}
finally {
  try{Write-ControlState -Phase 'finished'}catch{}

  if($doc -ne $null){
    try{$doc.Close($false)}catch{$result.doc_close_error=$_.Exception.Message}
  }

  if($word -ne $null -and $wordPidOwned){
    if(Test-TaskWordIdentity -TaskPid $wordPid -StartTicks $wordStartTicks){
      try{
        if($normalWasSaved -eq $true -and -not [bool]$word.NormalTemplate.Saved){
          $word.NormalTemplate.Saved=$true
        }
      }catch{}
      try{$word.DisplayAlerts=0}catch{}
      try{$word.Quit()}catch{$result.word_quit_error=$_.Exception.Message}
    }else{
      $result.word_quit_skipped_identity_changed=$true
    }
  } elseif($word -ne $null) {
    # Never call Quit() on a COM object whose process ownership could not be
    # proven; it may belong to a user's pre-existing Word session.
    $result.word_quit_skipped_unowned=$true
  }

  $doc=$null
  $word=$null
  $addin=$null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()

  if($null -ne $oldNo){Set-ItemProperty $key -Name DoNoWinVerb -Value $oldNo}
  if($null -ne $oldWait){Set-ItemProperty $key -Name WaitingConvert -Value $oldWait}

  if($wordPid -and $wordPidOwned){
    $deadline=(Get-Date).AddSeconds(6)
    while((Get-Date) -lt $deadline -and (Test-TaskWordIdentity -TaskPid $wordPid -StartTicks $wordStartTicks)){
      Start-Sleep -Milliseconds 250
    }
    if(Test-TaskWordIdentity -TaskPid $wordPid -StartTicks $wordStartTicks){
      try{
        Stop-Process -Id $wordPid -Force -ErrorAction Stop
        $result.word_forced_cleanup=$true
      }catch{
        $result.word_cleanup_error=$_.Exception.Message
      }
    }
  }

  Start-Sleep -Milliseconds 300
  $afterDumps=@(
    Get-ChildItem $crashDir -Filter 'AxMath*.dmp' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty Name
  )
  $result.new_crash_dumps=@($afterDumps|Where-Object{$_ -notin $beforeDumps})
  $total.Stop()
  $result.total_seconds=$total.Elapsed.TotalSeconds
  $macroMeasure=$result.batches | Measure-Object -Property seconds -Sum
  $saveMeasure=$result.batches | Measure-Object -Property save_seconds -Sum
  $checkpointMeasure=$result.batches | Measure-Object -Property checkpoint_seconds -Sum
  $convertedMeasure=$result.batches | Measure-Object -Property converted -Maximum
  $result.macro_seconds_total=[double]$macroMeasure.Sum
  $result.save_seconds_total=[double]$saveMeasure.Sum
  $result.checkpoint_seconds_total=[double]$checkpointMeasure.Sum
  $result.non_macro_seconds=[math]::Max(0.0,$result.total_seconds-$result.macro_seconds_total)
  $result.macro_share_percent=if($result.total_seconds -gt 0){100.0*$result.macro_seconds_total/$result.total_seconds}else{0.0}
  $result.observed_max_batch_converted=[int]$convertedMeasure.Maximum
  $result.batch_limit_owner='AxMath plugin; do not force/bypass WaitingConvert batch semantics'

  $json=$result|ConvertTo-Json -Depth 8
  Set-Content -LiteralPath $ReportPath -Value $json -Encoding UTF8
  Write-Output $json
}
